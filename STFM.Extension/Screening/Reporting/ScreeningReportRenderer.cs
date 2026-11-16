// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Globalization;
using Microsoft.Maui.Graphics;
using Microsoft.Maui.Graphics.Skia;
using SkiaSharp;

namespace STFM.Extension.Screening.Reporting;

/// <summary>Creates vector A4 PDFs and matching previews without a separate PDF-reporting library.</summary>
public static class ScreeningReportRenderer
{
    public const float A4Width = 595.276f;
    public const float A4Height = 841.890f;

    public static Task<ScreeningReportDocument> CreateAsync(ScreeningExamination examination,
        string outputDirectory, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(examination);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputDirectory);
        if (examination.Otoscopy is null || examination.Ssq12 is null || examination.QuickSip is null || examination.Audiometry is null)
            throw new InvalidOperationException("Alla fyra steg måste vara klara innan rapporten skapas.");
        return Task.Run(() => Create(examination, outputDirectory, cancellationToken), cancellationToken);
    }

    private static ScreeningReportDocument Create(ScreeningExamination examination, string directory, CancellationToken token)
    {
        Directory.CreateDirectory(directory);
        var snapshot = new ReportSnapshot(examination);
        using var renderer = new PageRenderer(snapshot);
        var name = $"screening-{examination.Id:N}";
        var pdfPath = Path.Combine(directory, name + ".pdf");
        var previews = Enumerable.Range(0, renderer.PageCount).Select(i => Path.Combine(directory, $"{name}-page-{i + 1}.png")).ToArray();
        var temporary = new List<string>();
        try
        {
            var pdfTemporary = pdfPath + ".tmp";
            temporary.Add(pdfTemporary);
            using (var stream = File.Create(pdfTemporary))
                renderer.WritePdf(stream, Enumerable.Range(0, renderer.PageCount).ToArray(), A4Width, A4Height, 0, token);
            for (int i = 0; i < previews.Length; i++)
            {
                token.ThrowIfCancellationRequested();
                var imageTemporary = previews[i] + ".tmp";
                temporary.Add(imageTemporary);
                // 144 dpi makes text legible on tablets while keeping two-page preview memory modest.
                using var bitmap = new SKBitmap((int)Math.Ceiling(A4Width * 2), (int)Math.Ceiling(A4Height * 2));
                using (var canvas = new SKCanvas(bitmap))
                {
                    canvas.Scale(2);
                    renderer.DrawPage(canvas, i);
                }
                using var image = SKImage.FromBitmap(bitmap);
                using var png = image.Encode(SKEncodedImageFormat.Png, 100);
                using var output = File.Create(imageTemporary);
                png.SaveTo(output);
            }
            token.ThrowIfCancellationRequested();
            for (int i = 0; i < previews.Length; i++)
                File.Move(previews[i] + ".tmp", previews[i], overwrite: true);
            File.Move(pdfTemporary, pdfPath, overwrite: true);
            return new ScreeningReportDocument(pdfPath, previews)
            {
                // Snapshot includes copied image bytes/results: reprints cannot accidentally
                // use data entered later or a replacement examination's image paths.
                WritePages = (stream, pages, width, height, margin, ct) =>
                {
                    using var printRenderer = new PageRenderer(snapshot);
                    printRenderer.WritePdf(stream, pages, width, height, margin, ct);
                }
            };
        }
        finally
        {
            foreach (var path in temporary)
                if (File.Exists(path)) File.Delete(path);
        }
    }

    private sealed class ReportSnapshot
    {
        public Guid Id { get; }
        public string Reference { get; }
        public DateTimeOffset StartedAt { get; }
        public DateTimeOffset? CompletedAt { get; }
        public OtoscopyScreeningResult Otoscopy { get; }
        public byte[]? LeftJpeg { get; }
        public byte[]? RightJpeg { get; }
        public Ssq12Result Ssq { get; }
        public QuickSipResult Quick { get; }
        public AudiometryResult Audio { get; }
        public ReportSnapshot(ScreeningExamination source)
        {
            Id = source.Id; Reference = source.Reference; StartedAt = source.StartedAt; CompletedAt = source.CompletedAt;
            var otoscopy = source.Otoscopy!;
            Otoscopy = new OtoscopyScreeningResult
            {
                LeftImagePath = otoscopy.LeftImagePath, RightImagePath = otoscopy.RightImagePath,
                LeftCapturedAt = otoscopy.LeftCapturedAt, RightCapturedAt = otoscopy.RightCapturedAt,
                LeftNotCapturedReason = otoscopy.LeftNotCapturedReason, RightNotCapturedReason = otoscopy.RightNotCapturedReason,
                Notes = otoscopy.Notes
            };
            LeftJpeg = ReadImage(Otoscopy.LeftImagePath); RightJpeg = ReadImage(Otoscopy.RightImagePath);
            Ssq = source.Ssq12! with { Answers = source.Ssq12!.Answers.ToArray() };
            Quick = source.QuickSip!;
            Audio = source.Audiometry! with { Points = source.Audiometry!.Points.ToArray() };
        }
        private static byte[]? ReadImage(string? path) => string.IsNullOrWhiteSpace(path) ? null : File.ReadAllBytes(path);
    }

    private sealed class PageRenderer : IDisposable
    {
        private const float Margin = 36;
        private const float Width = A4Width - 2 * Margin;
        private readonly ReportSnapshot data;
        private readonly SKTypeface regular = SKTypeface.FromFamilyName("sans-serif", SKFontStyle.Normal);
        private readonly SKTypeface bold = SKTypeface.FromFamilyName("sans-serif", SKFontStyle.Bold);
        private readonly SKColor ink = SKColor.Parse("#1C2E3A");
        private readonly SKColor muted = SKColor.Parse("#536572");
        private readonly SKColor accent = SKColor.Parse("#206C7B");
        private readonly List<string> noteLines;
        private const int FirstPageNoteLines = 4;
        private const int ContinuationLines = 44;
        public int PageCount { get; }

        public PageRenderer(ReportSnapshot data)
        {
            this.data = data;
            var notes = new List<string>();
            if (!string.IsNullOrWhiteSpace(data.Otoscopy.Notes)) notes.Add(data.Otoscopy.Notes);
            // Retain full missing-image explanations in the continuation flow too.
            // The image card has limited room, but the report must not lose the reason.
            if (data.LeftJpeg is null) notes.Add("Vänster öra, ingen bild: " + (data.Otoscopy.LeftNotCapturedReason ?? "Orsak ej angiven."));
            if (data.RightJpeg is null) notes.Add("Höger öra, ingen bild: " + (data.Otoscopy.RightNotCapturedReason ?? "Orsak ej angiven."));
            noteLines = Wrap(notes.Count == 0 ? "Inga anteckningar." : string.Join("\n", notes), Width - 24, 10.5f);
            PageCount = 2 + (int)Math.Ceiling(Math.Max(0, noteLines.Count - FirstPageNoteLines) / (double)ContinuationLines);
        }

        public void WritePdf(Stream stream, IReadOnlyList<int> pages, float width, float height, float printerMargin, CancellationToken token)
        {
            if (pages.Count == 0 || pages.Any(page => page < 0 || page >= PageCount))
                throw new ArgumentException("Ogiltigt sidintervall.", nameof(pages));
            using var document = SKDocument.CreatePdf(stream) ?? throw new IOException("PDF-rapporten kunde inte skapas.");
            foreach (var page in pages)
            {
                token.ThrowIfCancellationRequested();
                var canvas = document.BeginPage(width, height);
                canvas.Clear(SKColors.White);
                float scale = Math.Min((width - 2 * printerMargin) / A4Width, (height - 2 * printerMargin) / A4Height);
                canvas.Translate((width - A4Width * scale) / 2, (height - A4Height * scale) / 2);
                canvas.Scale(scale);
                DrawPage(canvas, page);
                document.EndPage();
            }
            document.Close();
        }

        public void DrawPage(SKCanvas canvas, int page)
        {
            canvas.Clear(SKColors.White);
            Header(canvas, page);
            if (page == 0) FirstPage(canvas);
            else if (page == 1) SecondPage(canvas);
            else NotesPage(canvas, page);
            Line(canvas, Margin, 794, A4Width - Margin, 794, SKColor.Parse("#CCD5DB"), .7f);
            Text(canvas, "Resonic Audiology · Hörselscreening", Margin, 811, 8.5f, muted);
            Text(canvas, $"{data.Id.ToString("N")[..12]} · Sida {page + 1} av {PageCount}", A4Width - Margin, 811, 8.5f, muted, right: true);
        }

        private void Header(SKCanvas c, int page)
        {
            Text(c, "HÖRSELSCREENING", Margin, 49, 22, ink, strong: true);
            Text(c, page == 0 ? "Otoskopi och självskattad hörsel" : page == 1 ? "Taluppfattning och automatiserad audiometri" : "Otoskopi · fortsatta anteckningar",
                Margin, 68, 10.5f, muted);
            var referenceLines = Wrap("Referens: " + (string.IsNullOrWhiteSpace(data.Reference) ? "Ej angiven" : data.Reference), Width, 11);
            var y = 92f;
            foreach (var line in referenceLines.Take(3)) { Text(c, line, Margin, y, 11, ink, strong: true); y += 13; }
            Text(c, $"Start: {Date(data.StartedAt)}   |   Klar: {Date(data.CompletedAt)}", Margin, Math.Max(120, y + 2), 9, muted);
            Line(c, Margin, 136, A4Width - Margin, 136, accent, 2);
        }

        private void FirstPage(SKCanvas c)
        {
            Section(c, "1", "Video-otoskopi", 158);
            const float gap = 16;
            float columnWidth = (Width - gap) / 2;
            ImagePanel(c, Margin, 176, columnWidth, 222, "Vänster öra", data.LeftJpeg,
                data.Otoscopy.LeftCapturedAt, data.Otoscopy.LeftNotCapturedReason, SKColor.Parse("#2466A5"));
            ImagePanel(c, Margin + columnWidth + gap, 176, columnWidth, 222, "Höger öra", data.RightJpeg,
                data.Otoscopy.RightCapturedAt, data.Otoscopy.RightNotCapturedReason, SKColor.Parse("#B53838"));
            Text(c, "Personalens anteckningar", Margin, 420, 10, ink, strong: true);
            var y = 437f;
            foreach (var line in noteLines.Take(FirstPageNoteLines)) { Text(c, line, Margin, y, 10.5f, muted); y += 13; }
            if (noteLines.Count > FirstPageNoteLines)
                Text(c, "Fortsättning på sida 3.", Margin, 491, 9, accent);

            Section(c, "2", "SSQ12 · självskattad hörsel", 523);
            Card(c, Margin, 541, Width, 66);
            Text(c, data.Ssq.MeanScore is { } mean ? $"{mean.ToString("0.0", Swedish)} / 10" : "Inget medelvärde", Margin + 14, 571,
                data.Ssq.MeanScore.HasValue ? 23 : 14, accent, strong: true);
            Text(c, $"Medelvärde av {data.Ssq.NumericAnswerCount} numeriska svar av 12", Margin + 190, 563, 10, ink);
            Text(c, "Vet inte räknas inte som noll eller in i medelvärdet.", Margin + 190, 582, 9, muted);
            Text(c, "Svar per fråga (0–10)", Margin, 629, 10, ink, strong: true);
            DrawSsqItems(c, 646);
            Text(c, HearingAidText(data.Ssq), Margin, 759, 9, muted);
        }

        private void SecondPage(SKCanvas c)
        {
            Section(c, "3", "QuickSiP · taluppfattning", 158);
            Card(c, Margin, 176, Width, 93);
            Text(c, data.Quick.ScoreFraction is { } score ? $"{(score * 100).ToString("0", Swedish)} %" : "Resultat saknas",
                Margin + 14, 213, data.Quick.ScoreFraction.HasValue ? 25 : 14, accent, strong: true);
            Text(c, $"{data.Quick.Correct} rätt av {data.Quick.TrialCount} genomförda försök", Margin + 170, 199, 11, ink, strong: true);
            Text(c, $"Planerat: {data.Quick.PlannedTrialCount} · Svarstid överskriden: {data.Quick.TimedOutCount}", Margin + 170, 216, 9, muted);
            TextFit(c, "Startlista: " + data.Quick.StartList, Margin + 170, 233, Width - 184, 9, muted);
            TextFit(c, "Hörtelefon: " + data.Quick.Transducer, Margin + 14, 254, Width - 28, 9, muted);

            Section(c, "4", "Automatiserad audiometri", 296);
            TextFit(c, $"{data.Audio.Protocol} · Luftledning · Screeningnivå {data.Audio.ScreeningLevelDbHl} dB HL", Margin, 317, Width, 9.5f, muted);
            TextFit(c, "Hörtelefon: " + data.Audio.Transducer, Margin, 332, Width, 9, muted);
            Text(c, "dB HL", Margin, 354, 9, muted);
            Text(c, "Frekvens (Hz)", A4Width - Margin, 354, 9, muted, right: true);
            using (var bridge = new SkiaCanvas { Canvas = c })
            {
                new ScreeningAudiogram(data.Audio).DrawScreening(bridge, new RectF(Margin, 359, Width, 278));
            }
            Text(c, "O Höger", Margin + 24, 653, 10, SKColor.Parse("#B53838"), strong: true);
            Text(c, "X Vänster", Margin + 113, 653, 10, SKColor.Parse("#2466A5"), strong: true);
            Text(c, "↓ Ingen respons   ≤ Hört vid screeningnivån", Margin + 214, 653, 9, muted);
            DrawAudiometryTable(c, 671);
        }

        private void NotesPage(SKCanvas c, int page)
        {
            Section(c, "1", "Video-otoskopi · personalens anteckningar", 158);
            int start = FirstPageNoteLines + (page - 2) * ContinuationLines;
            float y = 186;
            foreach (var line in noteLines.Skip(start).Take(ContinuationLines)) { Text(c, line, Margin, y, 10.5f, ink); y += 13; }
        }

        private void DrawSsqItems(SKCanvas c, float y)
        {
            float cell = Width / 12;
            for (int number = 1; number <= 12; number++)
            {
                var answer = data.Ssq.Answers.FirstOrDefault(a => a.Number == number);
                float x = Margin + (number - 1) * cell;
                Text(c, number.ToString(CultureInfo.InvariantCulture), x + cell / 2, y, 9, muted, centered: true);
                var bar = new SKRect(x + cell / 2 - 8, y + 10, x + cell / 2 + 8, y + 70);
                Fill(c, bar, SKColor.Parse("#E8EEF1"));
                if (answer?.Rating is { } value)
                    Fill(c, new SKRect(bar.Left, bar.Bottom - Math.Clamp(value, 0, 10) * 6, bar.Right, bar.Bottom), accent);
                Text(c, answer?.Rating?.ToString(CultureInfo.InvariantCulture) ?? (answer?.DontKnow == true ? "?" : "–"),
                    x + cell / 2, y + 85, 10, ink, strong: true, centered: true);
            }
            Text(c, "? = Vet inte", Margin, y + 101, 8.5f, muted);
        }

        private void DrawAudiometryTable(SKCanvas c, float y)
        {
            var effective = data.Audio.Points.GroupBy(p => (p.Ear, p.FrequencyHz)).Select(g => g.Last()).ToArray();
            var frequencies = effective.Select(p => p.FrequencyHz).Distinct().Order().ToArray();
            float labelWidth = 55;
            float cell = (Width - labelWidth) / Math.Max(1, frequencies.Length);
            Fill(c, new SKRect(Margin, y - 11, A4Width - Margin, y + 6), SKColor.Parse("#E8EEF1"));
            Text(c, "Hz", Margin + 5, y, 8.5f, ink, strong: true);
            Text(c, "Höger", Margin + 5, y + 20, 8.5f, ink, strong: true);
            Text(c, "Vänster", Margin + 5, y + 40, 8.5f, ink, strong: true);
            for (int i = 0; i < frequencies.Length; i++)
            {
                float x = Margin + labelWidth + cell * (i + .5f);
                Text(c, frequencies[i] >= 1000 ? (frequencies[i] / 1000d).ToString("0.#", Swedish) + "k" : frequencies[i].ToString(), x, y, 8.5f, ink, centered: true);
                foreach (var (ear, offset) in new[] { ("Right", 20), ("Left", 40) })
                {
                    var point = effective.LastOrDefault(p => p.Ear.Equals(ear, StringComparison.OrdinalIgnoreCase) && p.FrequencyHz == frequencies[i]);
                    Text(c, PointText(point), x, y + offset, 8.5f, ink, centered: true);
                }
            }
            Text(c, "≤ = tröskeln är högst screeningnivån. > = ingen respons vid angiven nivå.", Margin, y + 61, 8.5f, muted);
            Text(c, "? = osäkert värde. – = ej mätt. * = senaste kontrollmätningen används.", Margin, y + 75, 8.5f, muted);
            Text(c, "Screeningresultat. Klinisk bedömning görs av ansvarig vårdpersonal.", Margin, y + 95, 9, ink);
        }

        private static string PointText(AudiometryPoint? point)
        {
            if (point is null || point.Status == AudiometryPointStatus.NotMeasured || point.LevelDbHl is null) return "–";
            string prefix = point.Status == AudiometryPointStatus.NoResponse ? ">" : point.Status == AudiometryPointStatus.Uncertain ? "?" : point.AtScreeningLevel ? "≤" : "";
            return prefix + point.LevelDbHl.Value.ToString(CultureInfo.InvariantCulture) + (point.IsReliabilityRetest ? "*" : "");
        }

        private void ImagePanel(SKCanvas c, float x, float y, float width, float height, string title,
            byte[]? jpeg, DateTimeOffset? captured, string? missingReason, SKColor color)
        {
            Text(c, title, x, y + 1, 12, color, strong: true);
            var bounds = new SKRect(x, y + 12, x + width, y + height - 18);
            Fill(c, bounds, SKColor.Parse("#F2F5F7"));
            if (jpeg is not null)
            {
                using var image = SKImage.FromEncodedData(jpeg) ?? throw new InvalidDataException("En otoskopibild kunde inte läsas.");
                float scale = Math.Min(bounds.Width / image.Width, bounds.Height / image.Height);
                float w = image.Width * scale, h = image.Height * scale;
                var target = new SKRect(bounds.MidX - w / 2, bounds.MidY - h / 2, bounds.MidX + w / 2, bounds.MidY + h / 2);
                c.DrawImage(image, target);
            }
            else
            {
                Text(c, "Ingen bild", bounds.MidX, bounds.Top + 52, 14, muted, centered: true);
                float lineY = bounds.Top + 76;
                foreach (var line in Wrap(missingReason ?? "Orsak ej angiven.", width - 24, 10).Take(7))
                { Text(c, line, x + 12, lineY, 10, muted); lineY += 13; }
            }
            Text(c, captured is null ? "Ej fotograferat" : Date(captured), x, y + height, 8.5f, muted);
        }

        private void Section(SKCanvas c, string number, string title, float baseline)
        {
            Fill(c, new SKRect(Margin, baseline - 15, Margin + 23, baseline + 6), accent);
            Text(c, number, Margin + 11.5f, baseline + .5f, 11, SKColors.White, strong: true, centered: true);
            Text(c, title, Margin + 34, baseline + .5f, 14, ink, strong: true);
        }
        private void Card(SKCanvas c, float x, float y, float width, float height) => Fill(c, new SKRect(x, y, x + width, y + height), SKColor.Parse("#F0F5F6"));
        private static void Fill(SKCanvas c, SKRect rect, SKColor color) { using var p = new SKPaint { Color = color }; c.DrawRect(rect, p); }
        private static void Line(SKCanvas c, float x1, float y1, float x2, float y2, SKColor color, float width)
        { using var p = new SKPaint { Color = color, StrokeWidth = width, IsAntialias = true }; c.DrawLine(x1, y1, x2, y2, p); }
        private void Text(SKCanvas c, string text, float x, float baseline, float size, SKColor color, bool strong = false, bool right = false, bool centered = false)
        {
            using var font = new SKFont(strong ? bold : regular, size);
            using var paint = new SKPaint { Color = color, IsAntialias = true };
            c.DrawText(text, x, baseline, right ? SKTextAlign.Right : centered ? SKTextAlign.Center : SKTextAlign.Left, font, paint);
        }
        private void TextFit(SKCanvas c, string text, float x, float y, float width, float size, SKColor color)
        {
            using var font = new SKFont(regular, size);
            float measured = font.MeasureText(text);
            Text(c, text, x, y, measured > width ? Math.Max(6, size * width / measured) : size, color);
        }
        private List<string> Wrap(string text, float width, float size)
        {
            using var font = new SKFont(regular, size);
            var lines = new List<string>();
            foreach (var paragraph in text.Replace("\r", "").Split('\n'))
            {
                string line = "";
                foreach (var word in paragraph.Split(' ', StringSplitOptions.RemoveEmptyEntries))
                {
                    var candidate = line.Length == 0 ? word : line + " " + word;
                    if (font.MeasureText(candidate) <= width) { line = candidate; continue; }
                    if (line.Length > 0) lines.Add(line);
                    line = "";
                    // Split unbroken references/notes too; no text is silently clipped.
                    foreach (var character in word)
                    {
                        if (font.MeasureText(line + character) > width && line.Length > 0) { lines.Add(line); line = ""; }
                        line += character;
                    }
                }
                lines.Add(line);
            }
            return lines;
        }
        private static readonly CultureInfo Swedish = CultureInfo.GetCultureInfo("sv-SE");
        private static string Date(DateTimeOffset? value) => value?.ToString("yyyy-MM-dd HH:mm zzz", Swedish) ?? "Ej registrerat";
        private static string HearingAidText(Ssq12Result result) => result.UsesHearingAids switch
        {
            false => "Hörapparat: nej.",
            true => "Hörapparat: ja" + (result.HearingAidLeft || result.HearingAidRight ? " (" + string.Join(", ", new[] { result.HearingAidLeft ? "vänster" : null, result.HearingAidRight ? "höger" : null }.Where(s => s is not null)) + ")." : "."),
            _ => "Hörapparatanvändning: ej angiven."
        };
        public void Dispose() { regular.Dispose(); bold.Dispose(); }
    }
}
