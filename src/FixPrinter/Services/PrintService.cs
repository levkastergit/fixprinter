using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using PDFtoImage;
using SkiaSharp;

namespace FixPrinter.Services;

/// <summary>
/// Печать PDF и Word напрямую, минуя сбойное окно драйвера. Геометрия (вписывание/масштаб/
/// поворот) считается от ПЕЧАТАЕМОЙ области листа (с учётом полей принтера), поэтому край не
/// обрезается, а предпросмотр совпадает с выводом.
/// </summary>
public sealed class PrintService
{
    private readonly Logger _log = Logger.Instance;

    public enum Paper { A4, A3 }
    public enum ScaleMode { Fit, Percent }
    public enum Orient { Auto, Portrait, Landscape }
    public enum DuplexMode { None, Long, Short }

    public sealed class Options
    {
        public string PrinterName = "";
        public Paper Paper = Paper.A4;
        public int Copies = 1;
        public bool Collate = true;
        public Orient Orientation = Orient.Auto;
        public bool Grayscale = false;
        public ScaleMode Scale = ScaleMode.Fit;
        public int ScalePercent = 100;
        public DuplexMode Duplex = DuplexMode.None;
        public string? PaperSource = null;
        public int RenderDpi = 200;
    }

    public static (double W, double H) PaperInches(Paper p) =>
        p == Paper.A3 ? (11.69, 16.54) : (8.27, 11.69);

    public static List<string> ListRealPrinters()
    {
        var result = new List<string>();
        foreach (string p in PrinterSettings.InstalledPrinters) result.Add(p);
        return result;
    }

    public static List<string> GetPaperSources(string printerName)
    {
        var list = new List<string>();
        try
        {
            using var pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = printerName;
            if (!pd.PrinterSettings.IsValid) return list;
            foreach (PaperSource s in pd.PrinterSettings.PaperSources) list.Add(s.SourceName);
        }
        catch { }
        return list;
    }

    public static bool CanDuplex(string printerName)
    {
        try
        {
            using var pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = printerName;
            return pd.PrinterSettings.IsValid && pd.PrinterSettings.CanDuplex;
        }
        catch { return false; }
    }

    // ---------- Диапазон страниц ----------

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

    // ---------- Рендер ----------

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
        finally { if (tempPdf != null) { try { File.Delete(tempPdf); } catch { } } }
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

    // ---------- Общая геометрия (в дюймах, относительно печатаемой области) ----------

    private static (bool rot, double x, double y, double w, double h) Layout(
        int imgW, int imgH, Options opt, double availW, double availH)
    {
        var rot = opt.Orientation switch
        {
            Orient.Portrait => false,
            Orient.Landscape => true,
            _ => imgW > imgH
        };
        double ew = rot ? imgH : imgW, eh = rot ? imgW : imgH;
        double pageW = ew / opt.RenderDpi, pageH = eh / opt.RenderDpi;
        double w, h;
        if (opt.Scale == ScaleMode.Fit)
        {
            var s = Math.Min(availW / pageW, availH / pageH);
            w = pageW * s; h = pageH * s;
        }
        else
        {
            var p = Math.Clamp(opt.ScalePercent, 5, 400) / 100.0;
            w = pageW * p; h = pageH * p;
        }
        return (rot, (availW - w) / 2, (availH - h) / 2, w, h);
    }

    /// <summary>Поля принтера (непечатаемые края) в дюймах: слева/сверху/справа/снизу.</summary>
    public (double L, double T, double R, double B) PrintableInsets(Options opt)
    {
        try
        {
            using var pd = new PrintDocument();
            pd.PrinterSettings.PrinterName = opt.PrinterName;
            if (!pd.PrinterSettings.IsValid) return (0.16, 0.16, 0.16, 0.16);
            var ps = pd.DefaultPageSettings;
            ps.PaperSize = FindPaper(pd.PrinterSettings, opt.Paper);
            var pa = ps.PrintableArea;
            double paperW = ps.PaperSize.Width, paperH = ps.PaperSize.Height;
            double L = pa.X / 100.0, T = pa.Y / 100.0;
            double R = (paperW - (pa.X + pa.Width)) / 100.0;
            double B = (paperH - (pa.Y + pa.Height)) / 100.0;
            L = Math.Clamp(L, 0, 0.6); T = Math.Clamp(T, 0, 0.6);
            R = Math.Clamp(R, 0, 0.6); B = Math.Clamp(B, 0, 0.6);
            if (L + T + R + B < 0.02) return (0.16, 0.16, 0.16, 0.16); // драйвер вернул 0 — берём типичное поле
            return (L, T, R, B);
        }
        catch { return (0.16, 0.16, 0.16, 0.16); }
    }

    // ---------- Предпросмотр листа (Skia) ----------

    public byte[] ComposeSheet(byte[] pagePng, Options opt, (double L, double T, double R, double B) insets, int longSidePx = 360)
    {
        using var page = SKBitmap.Decode(pagePng);
        var (pwIn, phIn) = PaperInches(opt.Paper);
        var ppi = longSidePx / phIn;
        int sw = Math.Max(1, (int)(pwIn * ppi)), sh = Math.Max(1, (int)(phIn * ppi));

        double availW = Math.Max(0.1, pwIn - insets.L - insets.R);
        double availH = Math.Max(0.1, phIn - insets.T - insets.B);
        var (rot, xIn, yIn, wIn, hIn) = Layout(page.Width, page.Height, opt, availW, availH);
        double drawX = insets.L + xIn, drawY = insets.T + yIn;

        var sheet = new SKBitmap(sw, sh);
        using (var canvas = new SKCanvas(sheet))
        {
            canvas.Clear(SKColors.White);
            var dst = new SKRect((float)(drawX * ppi), (float)(drawY * ppi),
                                 (float)((drawX + wIn) * ppi), (float)((drawY + hIn) * ppi));
            using var paint = new SKPaint { FilterQuality = SKFilterQuality.Medium, IsAntialias = true };
            if (opt.Grayscale) paint.ColorFilter = SKColorFilter.CreateColorMatrix(GrayMatrix);

            canvas.Save();
            if (rot)
            {
                canvas.Translate(dst.MidX, dst.MidY);
                canvas.RotateDegrees(90);
                canvas.Translate(-dst.MidY, -dst.MidX);
                canvas.DrawBitmap(page, new SKRect(dst.Top, dst.Left, dst.Bottom, dst.Right), paint);
            }
            else canvas.DrawBitmap(page, dst, paint);
            canvas.Restore();

            // граница печатаемой области (пунктир) и граница листа
            using var margin = new SKPaint { Style = SKPaintStyle.Stroke, Color = new SKColor(0x94, 0xA3, 0xB8, 0x80),
                StrokeWidth = 1, PathEffect = SKPathEffect.CreateDash(new float[] { 4, 4 }, 0) };
            canvas.DrawRect((float)(insets.L * ppi), (float)(insets.T * ppi),
                (float)(availW * ppi), (float)(availH * ppi), margin);
            using var border = new SKPaint { Style = SKPaintStyle.Stroke, Color = new SKColor(0x33, 0x41, 0x55), StrokeWidth = 1 };
            canvas.DrawRect(0.5f, 0.5f, sw - 1, sh - 1, border);
        }
        using var imgOut = SKImage.FromBitmap(sheet);
        using var dataOut = imgOut.Encode(SKEncodedImageFormat.Png, 90);
        sheet.Dispose();
        return dataOut.ToArray();
    }

    private static readonly float[] GrayMatrix =
    {
        0.299f, 0.587f, 0.114f, 0, 0,
        0.299f, 0.587f, 0.114f, 0, 0,
        0.299f, 0.587f, 0.114f, 0, 0,
        0,      0,      0,      1, 0
    };

    // ---------- Печать ----------

    public void PrintPages(IReadOnlyList<byte[]> pages, IReadOnlyList<int> indices, Options opt, string docName)
    {
        using var pd = new PrintDocument();
        pd.PrinterSettings.PrinterName = opt.PrinterName;
        if (!pd.PrinterSettings.IsValid)
            throw new InvalidOperationException($"Принтер «{opt.PrinterName}» не найден в системе.");
        pd.DocumentName = docName;
        pd.PrinterSettings.Copies = (short)Math.Max(1, opt.Copies);
        pd.PrinterSettings.Collate = opt.Collate;
        if (opt.Duplex != DuplexMode.None && pd.PrinterSettings.CanDuplex)
            pd.PrinterSettings.Duplex = opt.Duplex == DuplexMode.Long ? Duplex.Vertical : Duplex.Horizontal;

        var paper = FindPaper(pd.PrinterSettings, opt.Paper);
        pd.DefaultPageSettings.PaperSize = paper;
        pd.DefaultPageSettings.Color = !opt.Grayscale;
        if (opt.PaperSource != null)
            foreach (PaperSource s in pd.PrinterSettings.PaperSources)
                if (s.SourceName == opt.PaperSource) { pd.DefaultPageSettings.PaperSource = s; break; }
        pd.OriginAtMargins = false; // начало координат графики — угол печатаемой области

        _log.Info($"Печать «{docName}»: {indices.Count} стр., {paper.PaperName}, " +
                  $"масштаб {(opt.Scale == ScaleMode.Fit ? "вписать" : opt.ScalePercent + "%")}, " +
                  $"ориентация {opt.Orientation}, цвет {pd.DefaultPageSettings.Color}, " +
                  $"двусторонняя {pd.PrinterSettings.Duplex}, копий {pd.PrinterSettings.Copies}, " +
                  $"лоток {pd.DefaultPageSettings.PaperSource.SourceName}");

        using var grayAttrs = new ImageAttributes();
        if (opt.Grayscale) grayAttrs.SetColorMatrix(new ColorMatrix(ToJagged(GrayMatrix)));

        var pos = 0;
        pd.PrintPage += (_, e) =>
        {
            var pa = e.PageSettings.PrintableArea;
            double availW = pa.Width > 1 ? pa.Width / 100.0 : e.PageBounds.Width / 100.0 - 0.32;
            double availH = pa.Height > 1 ? pa.Height / 100.0 : e.PageBounds.Height / 100.0 - 0.32;

            using var ms = new MemoryStream(pages[indices[pos]]);
            using var src = Image.FromStream(ms);
            var img = src;
            var (rot, xIn, yIn, wIn, hIn) = Layout(src.Width, src.Height, opt, availW, availH);
            if (rot) { img = (Image)src.Clone(); img.RotateFlip(RotateFlipType.Rotate90FlipNone); }

            var dst = new Rectangle(
                (int)Math.Round(xIn * 100), (int)Math.Round(yIn * 100),
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
            "Чтобы печатать Word-документы, нужен установленный Microsoft Word или LibreOffice. PDF-файлы печатаются без них.");
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
