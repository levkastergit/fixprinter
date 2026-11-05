using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using PDFtoImage;
using SkiaSharp;

namespace FixPrinter.Services;

/// <summary>
/// Печать документов (PDF и Word) напрямую, минуя сбойное окно драйвера.
/// PDF рисуется движком PDFium в растровые страницы; печать и предпросмотр используют
/// ОДНУ геометрию (вписывание/масштаб/поворот), поэтому предпросмотр совпадает с листом.
/// </summary>
public sealed class PrintService
{
    private readonly Logger _log = Logger.Instance;

    public enum Paper { A4, A3 }
    public enum ScaleMode { Fit, Percent }

    public sealed class Options
    {
        public string PrinterName = "";
        public Paper Paper = Paper.A4;
        public int Copies = 1;
        public bool AutoRotate = true;
        public bool Grayscale = false;
        public ScaleMode Scale = ScaleMode.Fit;
        public int ScalePercent = 100;
        public int RenderDpi = 200;
    }

    /// <summary>Размер листа в дюймах (книжная ориентация). Все A-форматы имеют одно соотношение сторон.</summary>
    public static (double W, double H) PaperInches(Paper p) =>
        p == Paper.A3 ? (11.69, 16.54) : (8.27, 11.69);

    public static List<string> ListRealPrinters()
    {
        var result = new List<string>();
        foreach (string p in PrinterSettings.InstalledPrinters) result.Add(p);
        return result;
    }

    // ---------- Разбор диапазона страниц ----------

    /// <summary>«1-3,5,8» → 0-базные индексы (в пределах total). Пусто/некорректно → все страницы.</summary>
    public static List<int> ParsePages(string? spec, int total)
    {
        var all = Enumerable.Range(0, total).ToList();
        if (string.IsNullOrWhiteSpace(spec)) return all;

        var set = new SortedSet<int>();
        foreach (var part in spec.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var dash = part.Split('-', 2);
            if (dash.Length == 2 && int.TryParse(dash[0], out var a) && int.TryParse(dash[1], out var b))
            {
                if (a > b) (a, b) = (b, a);
                for (var i = a; i <= b; i++) if (i >= 1 && i <= total) set.Add(i - 1);
            }
            else if (int.TryParse(part, out var n) && n >= 1 && n <= total) set.Add(n - 1);
        }
        return set.Count > 0 ? set.ToList() : all;
    }

    // ---------- Рендер документа ----------

    /// <summary>Рендерит все страницы файла (Word → PDF при необходимости) в PNG (цветные).</summary>
    public List<byte[]> RenderFile(string path, int dpi, Action<string>? progress = null)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("Файл не найден: " + path);
        var ext = Path.GetExtension(path).ToLowerInvariant();
        string pdfPath = path;
        string? tempPdf = null;
        try
        {
            if (ext is ".doc" or ".docx" or ".rtf" or ".odt" or ".txt")
            {
                progress?.Invoke($"Конвертирую «{Path.GetFileName(path)}» в PDF…");
                pdfPath = tempPdf = ConvertToPdf(path);
            }
            else if (ext != ".pdf")
                throw new NotSupportedException("Поддерживаются PDF и Word (.doc/.docx/.rtf/.odt/.txt).");

            progress?.Invoke($"Готовлю страницы «{Path.GetFileName(path)}»…");
            return RenderPdf(pdfPath, dpi);
        }
        finally
        {
            if (tempPdf != null) { try { File.Delete(tempPdf); } catch { } }
        }
    }

    private List<byte[]> RenderPdf(string pdfPath, int dpi)
    {
        var pages = new List<byte[]>();
        var opts = new RenderOptions { Dpi = dpi, WithAspectRatio = true, WithAnnotations = true, WithFormFill = true };
        using var fs = File.OpenRead(pdfPath);
        foreach (var bmp in Conversion.ToImages(fs, leaveOpen: false, password: null, options: opts))
            using (bmp)
            using (var img = SKImage.FromBitmap(bmp))
            using (var data = img.Encode(SKEncodedImageFormat.Png, 100))
                pages.Add(data.ToArray());
        return pages;
    }

    // ---------- Общая геометрия (в дюймах) ----------

    private static (bool rot, double x, double y, double w, double h) Layout(
        int imgW, int imgH, Options opt, double sheetW, double sheetH)
    {
        var rot = opt.AutoRotate && (imgW > imgH) != (sheetW > sheetH);
        double ew = rot ? imgH : imgW, eh = rot ? imgW : imgH;
        double pageW = ew / opt.RenderDpi, pageH = eh / opt.RenderDpi; // реальный размер страницы, дюймы
        double w, h;
        if (opt.Scale == ScaleMode.Fit)
        {
            var s = Math.Min(sheetW / pageW, sheetH / pageH);
            w = pageW * s; h = pageH * s;
        }
        else
        {
            var p = Math.Clamp(opt.ScalePercent, 5, 400) / 100.0;
            w = pageW * p; h = pageH * p;
        }
        return (rot, (sheetW - w) / 2, (sheetH - h) / 2, w, h);
    }

    // ---------- Предпросмотр листа (Skia) ----------

    /// <summary>Собирает PNG-картинку листа, как он выйдет на печать (для предпросмотра).</summary>
    public byte[] ComposeSheet(byte[] pagePng, Options opt, int longSidePx = 360)
    {
        using var page = SKBitmap.Decode(pagePng);
        var (pwIn, phIn) = PaperInches(opt.Paper);
        var ppi = longSidePx / phIn;
        int sw = Math.Max(1, (int)(pwIn * ppi)), sh = Math.Max(1, (int)(phIn * ppi));

        var sheet = new SKBitmap(sw, sh);
        using (var canvas = new SKCanvas(sheet))
        {
            canvas.Clear(SKColors.White);
            var (rot, xIn, yIn, wIn, hIn) = Layout(page.Width, page.Height, opt, pwIn, phIn);
            var dst = new SKRect((float)(xIn * ppi), (float)(yIn * ppi),
                                 (float)((xIn + wIn) * ppi), (float)((yIn + hIn) * ppi));
            using var paint = new SKPaint { FilterQuality = SKFilterQuality.Medium, IsAntialias = true };
            if (opt.Grayscale)
                paint.ColorFilter = SKColorFilter.CreateColorMatrix(GrayMatrix);

            canvas.Save();
            if (rot)
            {
                // поворот картинки на 90° вокруг центра её места на листе
                canvas.Translate(dst.MidX, dst.MidY);
                canvas.RotateDegrees(90);
                canvas.Translate(-dst.MidY, -dst.MidX);
                var rotDst = new SKRect(dst.Top, dst.Left, dst.Bottom, dst.Right);
                canvas.DrawBitmap(page, rotDst, paint);
            }
            else
            {
                canvas.DrawBitmap(page, dst, paint);
            }
            canvas.Restore();

            using var border = new SKPaint { Style = SKPaintStyle.Stroke, Color = new SKColor(0x33, 0x41, 0x55), StrokeWidth = 1 };
            canvas.DrawRect(0.5f, 0.5f, sw - 1, sh - 1, border);
        }
        using var img = SKImage.FromBitmap(sheet);
        using var data = img.Encode(SKEncodedImageFormat.Png, 90);
        sheet.Dispose();
        return data.ToArray();
    }

    private static readonly float[] GrayMatrix =
    {
        0.299f, 0.587f, 0.114f, 0, 0,
        0.299f, 0.587f, 0.114f, 0, 0,
        0.299f, 0.587f, 0.114f, 0, 0,
        0,      0,      0,      1, 0
    };

    // ---------- Печать ----------

    /// <summary>Печатает заранее отрендеренные страницы (PNG) с выбранными индексами.</summary>
    public void PrintPages(IReadOnlyList<byte[]> pages, IReadOnlyList<int> indices, Options opt, string docName)
    {
        using var pd = new PrintDocument();
        pd.PrinterSettings.PrinterName = opt.PrinterName;
        if (!pd.PrinterSettings.IsValid)
            throw new InvalidOperationException($"Принтер «{opt.PrinterName}» не найден в системе.");
        pd.DocumentName = docName;
        pd.PrinterSettings.Copies = (short)Math.Max(1, opt.Copies);

        var paper = FindPaper(pd.PrinterSettings, opt.Paper);
        pd.DefaultPageSettings.PaperSize = paper;
        pd.DefaultPageSettings.Color = !opt.Grayscale;
        pd.OriginAtMargins = false;
        _log.Info($"Печать «{docName}»: {indices.Count} стр., {paper.PaperName} " +
                  $"({paper.Width}x{paper.Height}), масштаб {(opt.Scale == ScaleMode.Fit ? "вписать" : opt.ScalePercent + "%")}, " +
                  $"цвет: {pd.DefaultPageSettings.Color}, копий: {pd.PrinterSettings.Copies}");

        using var grayAttrs = new ImageAttributes();
        if (opt.Grayscale) grayAttrs.SetColorMatrix(new ColorMatrix(ToJagged(GrayMatrix)));

        var pos = 0;
        pd.PrintPage += (_, e) =>
        {
            var area = e.PageBounds;
            double sheetW = paper.Width / 100.0, sheetH = paper.Height / 100.0;

            using var ms = new MemoryStream(pages[indices[pos]]);
            using var src = Image.FromStream(ms);
            var img = src;
            var (rot, xIn, yIn, wIn, hIn) = Layout(src.Width, src.Height, opt, sheetW, sheetH);
            if (rot) { img = (Image)src.Clone(); img.RotateFlip(RotateFlipType.Rotate90FlipNone); }

            var dst = new Rectangle(
                (int)Math.Round(area.Left + xIn * 100), (int)Math.Round(area.Top + yIn * 100),
                (int)Math.Round(wIn * 100), (int)Math.Round(hIn * 100));

            e.Graphics!.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            if (opt.Grayscale)
                e.Graphics.DrawImage(img, dst, 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, grayAttrs);
            else
                e.Graphics.DrawImage(img, dst);

            if (!ReferenceEquals(img, src)) img.Dispose();
            pos++;
            e.HasMorePages = pos < indices.Count;
        };
        pd.Print();
    }

    private static float[][] ToJagged(float[] m) => new[]
    {
        new[] { m[0], m[1], m[2], m[3], m[4] },
        new[] { m[5], m[6], m[7], m[8], m[9] },
        new[] { m[10], m[11], m[12], m[13], m[14] },
        new[] { m[15], m[16], m[17], m[18], m[19] }
    };

    private static PaperSize FindPaper(PrinterSettings settings, Paper paper)
    {
        var kind = paper == Paper.A3 ? PaperKind.A3 : PaperKind.A4;
        foreach (PaperSize ps in settings.PaperSizes)
            if (ps.Kind == kind) return ps;
        return paper == Paper.A3 ? new PaperSize("A3", 1169, 1654) : new PaperSize("A4", 827, 1169);
    }

    // ---------- Word → PDF ----------

    private string ConvertToPdf(string docPath)
    {
        var outPdf = Path.Combine(Path.GetTempPath(),
            "FixPrinter_" + Path.GetFileNameWithoutExtension(docPath) + "_" + Guid.NewGuid().ToString("N")[..8] + ".pdf");
        if (TryConvertWithWord(docPath, outPdf)) return outPdf;
        if (TryConvertWithLibreOffice(docPath, outPdf)) return outPdf;
        throw new InvalidOperationException(
            "Чтобы печатать Word-документы, нужен установленный Microsoft Word или LibreOffice. " +
            "PDF-файлы печатаются без них.");
    }

    private bool TryConvertWithWord(string docPath, string outPdf)
    {
        var wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null) { _log.Info("Microsoft Word не установлен — пробую LibreOffice."); return false; }
        object? word = null;
        try
        {
            word = Activator.CreateInstance(wordType);
            dynamic app = word!;
            app.Visible = false;
            app.DisplayAlerts = 0;
            dynamic doc = app.Documents.Open(FileName: docPath, ReadOnly: true, AddToRecentFiles: false);
            try { doc.ExportAsFixedFormat(OutputFileName: outPdf, ExportFormat: 17); }
            finally { doc.Close(SaveChanges: false); }
            app.Quit();
            return File.Exists(outPdf);
        }
        catch (Exception ex) { _log.Warn("Конвертация через Word не удалась: " + ex.Message); return false; }
        finally { if (word != null) try { Marshal.ReleaseComObject(word); } catch { } }
    }

    private bool TryConvertWithLibreOffice(string docPath, string outPdf)
    {
        var soffice = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "LibreOffice", "program", "soffice.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "LibreOffice", "program", "soffice.exe")
        }.FirstOrDefault(File.Exists);
        if (soffice == null) { _log.Info("LibreOffice не найден."); return false; }

        var outDir = Path.GetDirectoryName(outPdf)!;
        var psi = new System.Diagnostics.ProcessStartInfo(soffice,
            $"--headless --norestore --convert-to pdf --outdir \"{outDir}\" \"{docPath}\"")
        { UseShellExecute = false, CreateNoWindow = true };
        using var proc = System.Diagnostics.Process.Start(psi);
        if (proc == null) return false;
        proc.WaitForExit(120000);
        var produced = Path.Combine(outDir, Path.GetFileNameWithoutExtension(docPath) + ".pdf");
        if (File.Exists(produced))
        {
            if (!produced.Equals(outPdf, StringComparison.OrdinalIgnoreCase))
                try { File.Move(produced, outPdf, overwrite: true); } catch { return File.Exists(produced); }
            return File.Exists(outPdf);
        }
        return false;
    }
}
