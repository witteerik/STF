// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using STFM.Extension.Otoscopy;
using STFM.Extension.Screening;

namespace STFM.Extension.Views;

public sealed partial class FullScreeningView
{
    private readonly ContentView stageHost = new();
    private readonly Label headingLabel = new() { FontSize = 25, FontAttributes = FontAttributes.Bold, TextColor = Color.FromArgb("#17334F") };
    private readonly Label referenceLabel = new() { FontSize = 13, TextColor = Color.FromArgb("#526477") };
    private readonly Label statusLabel = new() { FontSize = 14, LineBreakMode = LineBreakMode.WordWrap };
    private readonly ActivityIndicator activity = new() { Color = Color.FromArgb("#1B607E"), WidthRequest = 24 };
    private readonly Button continueButton = ActionButton("Fortsätt");
    private readonly Button pauseButton = ActionButton("Pausa / be om hjälp", secondary: true);
    private readonly Button closeButton = ActionButton("Spara och stäng", secondary: true);
    private readonly Button retryButton = ActionButton("Försök igen", secondary: true);
    private Entry? referenceEntry, leftMissingEntry, rightMissingEntry;
    private Editor? notesEditor;
    private CheckBox? readyCheck;
    private Button? printButton, confirmPrintButton, nextPatientButton, exportReportButton;

    private static Button ActionButton(string text, bool secondary = false) => new()
    {
        Text = text, MinimumHeightRequest = 52, FontSize = 17, CornerRadius = 10,
        Padding = new Thickness(18, 10),
        BackgroundColor = Color.FromArgb(secondary ? "#E5EDF2" : "#1B607E"),
        TextColor = secondary ? Color.FromArgb("#17334F") : Colors.White
    };

    private static Label Paragraph(string text, double size = 19) => new()
    {
        Text = text, FontSize = size, TextColor = Color.FromArgb("#24394D"),
        LineBreakMode = LineBreakMode.WordWrap, LineHeight = 1.15
    };

    private static Border Card(View content) => new()
    {
        Content = content, Padding = 20, Margin = new Thickness(0, 0, 0, 14),
        BackgroundColor = Colors.White, Stroke = Color.FromArgb("#D8E3EB"),
        StrokeShape = new Microsoft.Maui.Controls.Shapes.RoundRectangle { CornerRadius = 12 }
    };

    private void BuildLayout()
    {
        BackgroundColor = Color.FromArgb("#F3F6F9");
        var root = new Grid
        {
            Padding = new Thickness(16), RowSpacing = 12,
            RowDefinitions = new RowDefinitionCollection { new(GridLength.Auto), new(GridLength.Star), new(GridLength.Auto) }
        };
        var header = new Grid
        {
            ColumnDefinitions = new ColumnDefinitionCollection { new(GridLength.Star), new(GridLength.Auto) },
            ColumnSpacing = 12
        };
        header.Add(new VerticalStackLayout { Spacing = 4, Children = { headingLabel, referenceLabel } });
        header.Add(closeButton, 1);
        root.Add(header);
        root.Add(stageHost, 0, 1);
        var footer = new VerticalStackLayout { Spacing = 8 };
        var statusRow = new Grid
        {
            ColumnDefinitions = new ColumnDefinitionCollection { new(GridLength.Auto), new(GridLength.Star) },
            ColumnSpacing = 8
        };
        statusRow.Add(activity);
        statusRow.Add(statusLabel, 1);
        footer.Add(statusRow);
        footer.Add(continueButton);
        footer.Add(pauseButton);
        footer.Add(retryButton);
        root.Add(footer, 0, 2);
        Content = root;
        continueButton.Clicked += (_, _) => RunEvent(ContinueAsync);
        pauseButton.Clicked += (_, _) => { testCancellation?.Cancel(); RunEvent(PauseAsync); };
        closeButton.Clicked += async (_, _) => await CloseAsync();
        retryButton.Clicked += (_, _) => RunEvent(RetryAsync);
        retryButton.IsVisible = false;
        UpdateButtons();
    }

    private void UpdateHeader()
    {
        if (examination is null)
        {
            headingLabel.Text = "Hörselscreening";
            referenceLabel.Text = "Öppnar undersökningen…";
            return;
        }
        headingLabel.Text = examination.Stage switch
        {
            ScreeningStage.Setup => "Hörselscreening · förberedelser",
            ScreeningStage.OtoscopyInstructions or ScreeningStage.Otoscopy => "1 av 4 · Video-otoskopi · personal",
            ScreeningStage.Ssq12Instructions or ScreeningStage.Ssq12 => "2 av 4 · Frågeformulär SSQ12",
            ScreeningStage.QuickSipInstructions or ScreeningStage.QuickSip => "3 av 4 · Taluppfattning QuickSiP",
            ScreeningStage.AudiometryInstructions or ScreeningStage.Audiometry => "4 av 4 · Automatiskt tontest",
            _ => "Undersökningen är klar · rapport"
        };
        referenceLabel.Text = string.IsNullOrWhiteSpace(examination.Reference)
            ? "Fyra delar • resultat sparas på enheten"
            : $"Referens: {examination.Reference}  •  {examination.StartedAt:yyyy-MM-dd HH:mm}";
    }

    private void UpdateButtons()
    {
        var available = initialized && !busy && !closed && foreground;
        var running = testSession is not null;
        activity.IsVisible = activity.IsRunning = busy;
        closeButton.IsEnabled = !closed;
        continueButton.IsEnabled = available;
        continueButton.IsVisible = !running && examination?.Stage != ScreeningStage.Report;
        pauseButton.IsVisible = running;
        pauseButton.IsEnabled = available;
        retryButton.IsEnabled = !busy && !closed && foreground;
        if (printButton is not null) printButton.IsEnabled = available && examination?.IsComplete == true;
        if (exportReportButton is not null) exportReportButton.IsEnabled = available && examination?.IsComplete == true;
        if (confirmPrintButton is not null) confirmPrintButton.IsEnabled = available && examination?.PrintRequestedAt is not null;
        if (nextPatientButton is not null) nextPatientButton.IsEnabled = available && examination?.CanStartNext == true;
        if (referenceEntry is not null) referenceEntry.IsEnabled = !busy;
        if (readyCheck is not null) readyCheck.IsEnabled = !busy;
    }

    private void RenderStage()
    {
        if (examination is null)
            return;
        referenceEntry = null;
        readyCheck = null;
        printButton = confirmPrintButton = nextPatientButton = exportReportButton = null;
        UpdateHeader();
        switch (examination.Stage)
        {
            case ScreeningStage.Setup:
                RenderSetup();
                break;
            case ScreeningStage.OtoscopyInstructions:
                ShowInstructions("Till personalen", "Undersökning av öronen",
                    "Kontrollera patientreferensen och följ mottagningens rutin för video-otoskopi.",
                    "Anslut USB-otoskopet. Vänta på en tydlig förhandsvisning och ta en bild märkt Vänster och en märkt Höger. En ny bild ersätter bara samma sida.",
                    "Kontrollera att bilderna är tydliga och rätt sidmärkta. Om en sida inte kan fotograferas, ange orsaken; den följer med i rapporten.",
                    "När otoskopin är klar stoppas kameran. Lämna sedan surfplattan till patienten för de tre återstående delarna.");
                continueButton.Text = "Öppna otoskopet";
                break;
            case ScreeningStage.Otoscopy:
                RenderOtoscopy();
                break;
            case ScreeningStage.Ssq12Instructions:
                ShowInstructions("Till patienten", "Din hörsel i vardagen",
                    "Nu får du svara på 12 frågor om hur du hör i olika situationer.",
                    "Välj ett tal från 0 till 10 som bäst beskriver din egen upplevelse. Läs betydelsen av skalans ändpunkter vid varje fråga.",
                    "Om du inte kan bedöma en situation väljer du Vet ej. Det svaret räknas inte som noll.",
                    "Svara i din egen takt. Du behöver inte ha hörlurarna på för frågeformuläret. Be personalen om hjälp om något är oklart.");
                continueButton.Text = "Jag är redo – börja frågorna";
                break;
            case ScreeningStage.QuickSipInstructions:
                ShowInstructions("Till patienten", "Lyssna efter orden",
                    "Sätt på dig hörlurarna med den röda luren på höger öra. Be personalen kontrollera att de sitter rätt.",
                    "Du får höra korta ord tillsammans med bakgrundsljud. Efter varje ord trycker du på det ord du tror att du hörde bland alternativen på skärmen.",
                    "Du har högst fyra sekunder på dig att svara. Gissa om du är osäker. Om alternativen blinkar rött hann svaret inte registreras.",
                    "Testet innehåller 30 ord och blir efter hand svårare. Tryck på knappen nedan när du är redo.");
                continueButton.Text = "Hörlurarna är på – börja ordtestet";
                break;
            case ScreeningStage.AudiometryInstructions:
                ShowInstructions("Till patienten", "Tryck så länge tonen hörs",
                    "Behåll hörlurarna på. Nu kommer toner med olika tonhöjd och styrka i ett öra i taget.",
                    "Tryck på svarsknappen så snart du hör en ton och håll knappen intryckt så länge tonen hörs.",
                    "Släpp knappen när tonen tystnar. Tryck även när tonen är mycket svag, och vänta när du inte hör någon ton.",
                    "Sitt stilla och var tyst under testet. Du kan pausa och be personalen om hjälp vid behov.");
                continueButton.Text = "Jag förstår – börja tontestet";
                break;
            case ScreeningStage.Report:
                RenderReport();
                break;
        }
        UpdateButtons();
    }

    private void RenderSetup()
    {
        var panel = new VerticalStackLayout { Spacing = 16 };
        panel.Add(Paragraph("Personal – förbered ett nytt besök", 24));
        panel.Add(Paragraph("Video-otoskopi utförs av personalen. Därefter besvarar patienten SSQ12 och genomför QuickSiP samt ett automatiskt tontest.", 19));
        referenceEntry = new Entry
        {
            Placeholder = "Patient- eller undersökningsreferens", Text = examination!.Reference,
            MaxLength = 120, FontSize = 21, TextColor = Color.FromArgb("#17334F"), BackgroundColor = Colors.White
        };
        panel.Add(referenceEntry);
        panel.Add(Paragraph("Använd de konfigurerade, kalibrerade hörlurarna och ljudenheten. Kontrollera anslutningarna och att testmiljön är tillräckligt tyst enligt mottagningens rutin.", 18));
        readyCheck = new CheckBox { Color = Color.FromArgb("#1B607E") };
        var readiness = new Grid { ColumnDefinitions = new ColumnDefinitionCollection { new(GridLength.Auto), new(GridLength.Star) } };
        readiness.Add(readyCheck);
        readiness.Add(Paragraph("Jag har kontrollerat referensen, utrustningen och testmiljön.", 18), 1);
        panel.Add(readiness);
        panel.Add(Paragraph("Varje slutförd del sparas lokalt. Nästa patient kan startas först när personalen har bekräftat att den kompletta pappersrapporten har kommit ut.", 16));
        stageHost.Content = new ScrollView { Content = Card(panel) };
        continueButton.Text = "Kontrollera ljud och börja";
    }

    private void ShowInstructions(string audience, string title, params string[] paragraphs)
    {
        var panel = new VerticalStackLayout { Spacing = 20, MaximumWidthRequest = 900, HorizontalOptions = LayoutOptions.Fill };
        panel.Add(Paragraph(audience.ToUpperInvariant(), 14));
        panel.Add(new Label { Text = title, FontSize = 29, FontAttributes = FontAttributes.Bold, TextColor = Color.FromArgb("#17334F") });
        foreach (var paragraph in paragraphs)
            panel.Add(Paragraph(paragraph, 21));
        stageHost.Content = new ScrollView { Content = Card(panel) };
    }

    private void RenderOtoscopy()
    {
        // Reuse a currently attached camera view rather than opening another engine.
        if (otoscopy is not null)
            return;
        var result = examination!.Otoscopy;
        otoscopy = new OtoscopyView(embedded: true);
        OtoscopyImage? Restore(string? path, DateTimeOffset? at, OtoscopyEar side)
            => path is not null && at.HasValue && File.Exists(path)
                ? new OtoscopyImage(side, File.ReadAllBytes(path), at.Value) : null;
        otoscopy.RestoreImages(Restore(result?.LeftImagePath, result?.LeftCapturedAt, OtoscopyEar.Left),
            Restore(result?.RightImagePath, result?.RightCapturedAt, OtoscopyEar.Right));
        otoscopy.ImageCaptured += OtoscopyImageCaptured;
        // The camera view owns its scrollable preview. Staff notes stay in a separate,
        // compact lower panel; no nested scroll container surrounds live capture.
        var layout = new Grid
        {
            RowSpacing = 8,
            RowDefinitions = new RowDefinitionCollection { new(GridLength.Star), new(GridLength.Auto) }
        };
        layout.Add(otoscopy);
        var notes = new VerticalStackLayout { Spacing = 5 };
        notesEditor = new Editor { Placeholder = "Personalens anteckning (valfritt)", Text = result?.Notes,
            MaxLength = 1000, AutoSize = EditorAutoSizeOption.TextChanges, HeightRequest = 64, FontSize = 16, TextColor = Color.FromArgb("#17334F"), BackgroundColor = Colors.White };
        leftMissingEntry = new Entry { Placeholder = "Orsak om vänster bild saknas", Text = result?.LeftNotCapturedReason, MaxLength = 240, FontSize = 15, TextColor = Color.FromArgb("#17334F"), BackgroundColor = Colors.White };
        rightMissingEntry = new Entry { Placeholder = "Orsak om höger bild saknas", Text = result?.RightNotCapturedReason, MaxLength = 240, FontSize = 15, TextColor = Color.FromArgb("#17334F"), BackgroundColor = Colors.White };
        notes.Add(notesEditor);
        var missing = new Grid { ColumnSpacing = 8, ColumnDefinitions = new ColumnDefinitionCollection { new(GridLength.Star), new(GridLength.Star) } };
        missing.Add(leftMissingEntry);
        missing.Add(rightMissingEntry, 1);
        notes.Add(missing);
        layout.Add(notes, 0, 1);
        stageHost.Content = layout;
        continueButton.Text = "Otoskopin är klar – fortsätt till patientdelen";
        UpdateHeader();
        UpdateButtons();
    }

    private void RenderReport()
    {
        if (examination is null)
            return;
        var panel = new VerticalStackLayout { Spacing = 12 };
        panel.Add(Paragraph("Lämna surfplattan till personalen", 24));
        panel.Add(Paragraph("Alla fyra delar är slutförda. Kontrollera rapporten, skriv ut på A4 och bekräfta att samtliga sidor har kommit ut innan nästa patient börjar.", 18));
        printButton = ActionButton(OperatingSystem.IsWindows()
            ? "Öppna PDF för utskrift (A4)" : "Skicka till skrivare (A4)");
        confirmPrintButton = ActionButton("Utskriften har kommit ut – bekräfta", secondary: true);
        nextPatientButton = ActionButton("Nästa patient");
        exportReportButton = ActionButton("Exportera PDF", secondary: true);
        printButton.Clicked += (_, _) => RunEvent(PrintAsync);
        confirmPrintButton.Clicked += (_, _) => RunEvent(ConfirmPrintAsync);
        nextPatientButton.Clicked += (_, _) => RunEvent(NextPatientAsync);
        exportReportButton.Clicked += (_, _) => RunEvent(ExportReportAsync);
        panel.Add(printButton);
        panel.Add(confirmPrintButton);
        panel.Add(nextPatientButton);
        panel.Add(exportReportButton);
        panel.Add(Paragraph(examination.CanStartNext ? "✓ Pappersrapporten är kontrollerad."
            : "Nästa patient är låst tills pappersrapporten har bekräftats.", 17));
        var paths = report?.PreviewImagePaths ?? examination.ReportPreviewPaths;
        foreach (var path in paths.Where(File.Exists))
        {
            var image = new Image { Source = path, Aspect = Aspect.AspectFit,
                HeightRequest = Math.Clamp((Width > 0 ? Width - 72 : 680) * 1.414, 450, 1200) };
            panel.Add(Card(image));
        }
        if (!paths.Any(File.Exists))
            panel.Add(Paragraph("Rapporten förbereds. Om något går fel kan du försöka skapa den igen med utskriftsknappen.", 17));
        stageHost.Content = new ScrollView { Content = panel };
        UpdateHeader();
        UpdateButtons();
    }
}
