// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

namespace STFM.Extension.Screening.Reporting;

/// <summary>
/// Opens an operating-system print workflow. Neither a successful request nor a
/// completed print dialog proves that paper was received; the caller must ask staff.
/// </summary>
public static partial class ScreeningReportPrinter
{
    public static Task<ScreeningPrintRequestResult> PrintAsync(ScreeningReportDocument document,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(document);
        if (!File.Exists(document.PdfPath))
            throw new FileNotFoundException("Rapportfilen saknas. Skapa rapporten igen.", document.PdfPath);
        cancellationToken.ThrowIfCancellationRequested();
#if ANDROID
        return PrintAndroidAsync(document, cancellationToken);
#elif WINDOWS
        return OpenWindowsReaderAsync(document);
#else
        return Task.FromException<ScreeningPrintRequestResult>(new PlatformNotSupportedException("Utskrift stöds på Android och Windows."));
#endif
    }

#if WINDOWS
    private static async Task<ScreeningPrintRequestResult> OpenWindowsReaderAsync(ScreeningReportDocument document)
    {
        // MAUI has no cross-platform print API. Reuse the installed PDF reader on
        // Windows instead of adding a separate commercial printing/report library.
        // The reader owns printer selection and its print dialog (normally Ctrl+P).
        bool opened = await MainThread.InvokeOnMainThreadAsync(async () =>
        {
            var file = await global::Windows.Storage.StorageFile.GetFileFromPathAsync(document.PdfPath);
            return await global::Windows.System.Launcher.LaunchFileAsync(file);
        });
        return new ScreeningPrintRequestResult(opened, opened, opened
            ? "Rapporten har öppnats i Windows PDF-läsare. Välj Skriv ut (Ctrl+P) och A4. Bekräfta här först när hela rapporten kommit ur skrivaren."
            : "PDF-läsaren kunde inte öppnas. Installera eller välj en PDF-läsare och försök igen.");
    }
#endif
}
