using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using PDFtoImage;
using SkiaSharp;

namespace FixPrinter.Services;

/// <summary>
/// Печать документов (PDF и Word) напрямую, минуя окно настроек драйвера, которое падает.
/// PDF рисуется движком PDFium в растровые страницы, те печатаются через GDI (System.Drawing),
/// формат (A3/A4) задаётся программно. Word предварительно конвертируется в PDF.
/// </summary>
public sealed class PrintService
{
    private readonly Logger _log = Logger.Instance;

    public enum Paper { A4, A3 }

    public sealed class Options
    {
        public string PrinterName = "";
        public Paper Paper = Paper.A4;
        public int Copies = 1;
        public bool AutoRotate = true;      // поворачивать альбомные страницы под лист
        public bool Grayscale = false;      // печать в оттенках серого
        public int Dpi = 200;               // разрешение растеризации PDF
    }

    public static List<string> ListRealPrinters()
    {
        var result = new List<string>();
        foreach (string p in PrinterSettings.InstalledPrinters)
            result.Add(p);
        return result;
    }

    /// <summary>Готовит документ к печати (Word → PDF при необходимости) и печатает. Бросает с понятным текстом.</summary>
    public void PrintDocument(string path, Options opt, Action<string>? progress = null)
    {
        void Report(string m) { _log.Info(m); progress?.Invoke(m); }

        if (!File.Exists(path)) throw new FileNotFoundException("Файл не найден: " + path);
        var ext = Path.GetExtension(path).ToLowerInvariant();

        string pdfPath = path;
        string? tempPdf = null;
        try
        {
            if (ext is ".doc" or ".docx" or ".rtf" or ".odt" or ".txt")
            {
                Report($"Конвертирую «{Path.GetFileName(path)}» в PDF…");
                pdfPath = tempPdf = ConvertToPdf(path);
            }
            else if (ext != ".pdf")
            {
                throw new NotSupportedException("Поддерживаются PDF и Word (.doc/.docx/.rtf/.odt/.txt). Файл: " + ext);
            }

            Report($"Готовлю страницы «{Path.GetFileName(path)}»…");
            var pages = RenderPdf(pdfPath, opt.Dpi, opt.Grayscale);
            if (pages.Count == 0) throw new InvalidOperationException("В документе нет страниц для печати.");
            Report($"Страниц: {pages.Count}. Отправляю на «{opt.PrinterName}» ({opt.Paper}, копий: {opt.Copies})…");

            PrintImages(pages, opt, Path.GetFileName(path));
            Report($"Задание отправлено на принтер: «{Path.GetFileName(path)}».");
        }
        finally
        {
            if (tempPdf != null) { try { File.Delete(tempPdf); } catch { } }
        }
    }

    // ---------- Растеризация PDF ----------

    private List<byte[]> RenderPdf(string pdfPath, int dpi, bool grayscale)
    {
        var pages = new List<byte[]>();
        var opts = new RenderOptions { Dpi = dpi, WithAspectRatio = true, WithAnnotations = true, WithFormFill = true };
        using var fs = File.OpenRead(pdfPath);
        foreach (var bmp in Conversion.ToImages(fs, leaveOpen: false, password: null, options: opts))
        {
            using (bmp)
            {
                var src = bmp;
                SKBitmap? gray = null;
                if (grayscale) { gray = ToGrayscale(bmp); src = gray; }
                using var img = SKImage.FromBitmap(src);
                using var data = img.Encode(SKEncodedImageFormat.Png, 100);
                pages.Add(data.ToArray());
                gray?.Dispose();
            }
        }
        return pages;
    }

    private static SKBitmap ToGrayscale(SKBitmap src)
    {
        var dst = new SKBitmap(src.Width, src.Height);
        using var canvas = new SKCanvas(dst);
        using var paint = new SKPaint();
        var cm = new float[]
        {
            0.299f, 0.587f, 0.114f, 0, 0,
            0.299f, 0.587f, 0.114f, 0, 0,
            0.299f, 0.587f, 0.114f, 0, 0,
            0,      0,      0,      1, 0
        };
        paint.ColorFilter = SKColorFilter.CreateColorMatrix(cm);
        canvas.DrawBitmap(src, 0, 0, paint);
        return dst;
    }

    // ---------- Печать растровых страниц через GDI ----------

    private void PrintImages(List<byte[]> pages, Options opt, string docName)
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
        _log.Info($"Лист: {paper.PaperName} ({paper.Width}x{paper.Height} сотых дюйма), " +
                  $"цвет: {pd.DefaultPageSettings.Color}");

        var idx = 0;
        pd.PrintPage += (_, e) =>
        {
            using var ms = new MemoryStream(pages[idx]);
            using var image = Image.FromStream(ms);
            DrawFitted(e, image, opt.AutoRotate);
            idx++;
            e.HasMorePages = idx < pages.Count;
        };
        pd.Print();
    }

    /// <summary>Находит нужный формат среди форматов принтера; если нет — создаёт свой (в сотых дюйма).</summary>
    private static PaperSize FindPaper(PrinterSettings settings, Paper paper)
    {
        var kind = paper == Paper.A3 ? PaperKind.A3 : PaperKind.A4;
        foreach (PaperSize ps in settings.PaperSizes)
            if (ps.Kind == kind) return ps;
        // Резерв: размеры в сотых долях дюйма (A4 210x297 мм, A3 297x420 мм)
        return paper == Paper.A3
            ? new PaperSize("A3", 1169, 1654)
            : new PaperSize("A4", 827, 1169);
    }

    private static void DrawFitted(PrintPageEventArgs e, Image image, bool autoRotate)
    {
        var area = e.PageBounds; // весь лист, в сотых дюйма
        var img = image;
        if (autoRotate && (img.Width > img.Height) != (area.Width > area.Height))
        {
            img = (Image)image.Clone();
            img.RotateFlip(RotateFlipType.Rotate90FlipNone);
        }

        var scale = Math.Min((float)area.Width / img.Width, (float)area.Height / img.Height);
        var w = img.Width * scale;
        var h = img.Height * scale;
        var x = area.Left + (area.Width - w) / 2f;
        var y = area.Top + (area.Height - h) / 2f;

        e.Graphics!.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        e.Graphics.DrawImage(img, x, y, w, h);
        if (!ReferenceEquals(img, image)) img.Dispose();
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
            "PDF-файлы печатаются без них. Либо сохраните документ как PDF и напечатайте его.");
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
            try
            {
                doc.ExportAsFixedFormat(OutputFileName: outPdf, ExportFormat: 17); // wdExportFormatPDF
            }
            finally
            {
                doc.Close(SaveChanges: false);
            }
            app.Quit();
            _log.Info("Word сконвертировал документ в PDF.");
            return File.Exists(outPdf);
        }
        catch (Exception ex)
        {
            _log.Warn("Конвертация через Word не удалась: " + ex.Message);
            return false;
        }
        finally
        {
            if (word != null) try { Marshal.ReleaseComObject(word); } catch { }
        }
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

        // LibreOffice кладёт файл с тем же именем и расширением .pdf
        var produced = Path.Combine(outDir, Path.GetFileNameWithoutExtension(docPath) + ".pdf");
        if (File.Exists(produced))
        {
            if (!produced.Equals(outPdf, StringComparison.OrdinalIgnoreCase))
            { try { File.Move(produced, outPdf, overwrite: true); } catch { return File.Exists(produced); } }
            _log.Info("LibreOffice сконвертировал документ в PDF.");
            return File.Exists(outPdf);
        }
        return false;
    }
}
