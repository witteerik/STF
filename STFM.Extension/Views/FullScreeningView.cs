// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Diagnostics;
using System.Text.Json;
using CommunityToolkit.Maui.Storage;
using STFM.Extension.Otoscopy;
using STFM.Extension.Screening;
using STFM.Extension.Screening.Reporting;
using STFM.Extension.Screening.Tests;
using STFN.Core;

namespace STFM.Extension.Views;

/// <summary>
/// Staff otoscopy followed by the existing SSQ12, QuickSiP and SAME96 screening tests.
/// All completed results are journalled locally. Closing never resets the active patient.
/// Host navigation should await CloseAsync and replace this view only after Closed.
/// </summary>
public sealed partial class FullScreeningView : ContentView
{
    private static readonly SemaphoreSlim activeWorkflow = new(1, 1);
    private readonly SemaphoreSlim transitions = new(1, 1);
    private readonly ScreeningSessionStore store;
    private ScreeningExamination? examination;
    private AudioSystemSpecification? transducer;
    private IScreeningTestSession? testSession;
    private object? pendingResult;
    private CancellationTokenSource? testCancellation;
    private OtoscopyView? otoscopy;
    private ScreeningReportDocument? report;
    private Page? ownerPage;
    private Window? ownerWindow;
    private bool initialized, closed, closeRequested, busy, foreground = true, ownsLease, sharedContextCaptured;
    private string previousParticipant = string.Empty;
    private STFN.Core.Utils.EnumCollection.Languages previousLanguage;

    public event EventHandler? Closed;

    public FullScreeningView()
    {
        store = new ScreeningSessionStore(Path.Combine(FileSystem.AppDataDirectory, "Screening"));
        BuildLayout();
        Loaded += ViewLoaded;
        Unloaded += ViewUnloaded;
    }

    private void ViewLoaded(object? sender, EventArgs e) => RunEvent(InitializeAsync);

    private async Task InitializeAsync()
    {
        AttachLifecycle();
        foreground = true;
        if (initialized)
            return;
        if (!ownsLease)
        {
            if (!await activeWorkflow.WaitAsync(0))
                throw new InvalidOperationException("En hörselscreening är redan öppen. Stäng den innan en ny vy öppnas.");
            ownsLease = true;
        }
        examination = await store.LoadActiveAsync() ?? await store.CreateAsync(string.Empty);
        TakeSharedContext();
        await store.SaveAsync(examination);
        initialized = true;
        RenderStage();
        if (examination?.Stage == ScreeningStage.Report)
            await EnsureReportAsync();
    }

    private void ViewUnloaded(object? sender, EventArgs e)
    {
        if (closed)
            return;
        foreground = false;
        testCancellation?.Cancel();
        DetachLifecycle();
        RunEvent(PauseAsync);
    }

    private void TakeSharedContext()
    {
        if (!sharedContextCaptured)
        {
            previousParticipant = SharedSpeechTestObjects.CurrentParticipantID;
            previousLanguage = SharedSpeechTestObjects.GuiLanguage;
            sharedContextCaptured = true;
        }
        SharedSpeechTestObjects.GuiLanguage = STFN.Core.Utils.EnumCollection.Languages.Swedish;
        // Never pass a free-text patient reference to legacy code that constructs paths.
        SharedSpeechTestObjects.CurrentParticipantID = examination!.Id.ToString("N");
    }

    private void AttachLifecycle()
    {
        DetachLifecycle();
        for (Element? parent = Parent; parent is not null; parent = parent.Parent)
            if (parent is Page page)
            {
                ownerPage = page;
                page.Disappearing += PageDisappearing;
                page.Appearing += PageAppearing;
                break;
            }
        ownerWindow = Window;
        if (ownerWindow is not null)
        {
            ownerWindow.Stopped += WindowStopped;
            ownerWindow.Resumed += WindowResumed;
            ownerWindow.Destroying += WindowStopped;
        }
    }

    private void DetachLifecycle()
    {
        if (ownerPage is not null)
        {
            ownerPage.Disappearing -= PageDisappearing;
            ownerPage.Appearing -= PageAppearing;
        }
        if (ownerWindow is not null)
        {
            ownerWindow.Stopped -= WindowStopped;
            ownerWindow.Resumed -= WindowResumed;
            ownerWindow.Destroying -= WindowStopped;
        }
        ownerPage = null;
        ownerWindow = null;
    }

    private void PageDisappearing(object? sender, EventArgs e) => OnBackground();
    private void WindowStopped(object? sender, EventArgs e) => OnBackground();
    private void PageAppearing(object? sender, EventArgs e) { foreground = true; UpdateButtons(); }
    private void WindowResumed(object? sender, EventArgs e) { foreground = true; UpdateButtons(); }

    private void OnBackground()
    {
        foreground = false;
        // First-time media/permission pickers also stop the window. Let their
        // setup finish; only an actual patient test must be cancelled here.
        if (testSession is not null)
            testCancellation?.Cancel();
        RunEvent(PauseAsync);
    }

    private async Task PauseAsync()
    {
        if (examination is null)
            return;
        if (pendingResult is not null)
            await CompletePendingTestAsync();
        await SaveDraftAsync();
        if (otoscopy is not null)
            await otoscopy.StopCameraAsync();
        if (testSession is not null)
        {
            await StopTestAsync();
            await SaveChangesAsync(next => next.RecoverInterruptedStage());
            RenderStage();
            SetStatus("Testet pausades. Den pågående delen startas om från instruktionen; tidigare delar är sparade.");
        }
    }

    private async Task SaveChangesAsync(Action<ScreeningExamination> change)
    {
        var source = examination ?? throw new InvalidOperationException("Ingen undersökning är öppnad.");
        // Commit the in-memory state only after the atomic journal write succeeds.
        var bytes = JsonSerializer.SerializeToUtf8Bytes(source, ScreeningJsonContext.Default.ScreeningExamination);
        var next = JsonSerializer.Deserialize(bytes, ScreeningJsonContext.Default.ScreeningExamination)!;
        change(next);
        await store.SaveAsync(next);
        examination = next;
        UpdateHeader();
    }

    private async Task SaveDraftAsync()
    {
        if (examination is null)
            return;
        if (examination.Stage == ScreeningStage.Setup && referenceEntry is not null)
        {
            var reference = referenceEntry.Text?.Trim() ?? string.Empty;
            await SaveChangesAsync(next => next.Reference = reference);
        }
        if (otoscopy is not null)
        {
            if (otoscopy.LeftImage is { } left)
                await PersistImageAsync(left);
            if (otoscopy.RightImage is { } right)
                await PersistImageAsync(right);
            var notes = notesEditor?.Text?.Trim();
            var leftReason = leftMissingEntry?.Text?.Trim();
            var rightReason = rightMissingEntry?.Text?.Trim();
            await SaveChangesAsync(next =>
            {
                next.Otoscopy ??= new OtoscopyScreeningResult();
                next.Otoscopy.Notes = notes;
                if (next.Otoscopy.LeftImagePath is { } leftPath && !File.Exists(leftPath))
                {
                    next.Otoscopy.LeftImagePath = null;
                    next.Otoscopy.LeftCapturedAt = null;
                }
                if (next.Otoscopy.RightImagePath is { } rightPath && !File.Exists(rightPath))
                {
                    next.Otoscopy.RightImagePath = null;
                    next.Otoscopy.RightCapturedAt = null;
                }
                next.Otoscopy.LeftNotCapturedReason = string.IsNullOrWhiteSpace(next.Otoscopy.LeftImagePath) ? leftReason : null;
                next.Otoscopy.RightNotCapturedReason = string.IsNullOrWhiteSpace(next.Otoscopy.RightImagePath) ? rightReason : null;
            });
        }
    }

    private async Task PersistImageAsync(OtoscopyImage image)
    {
        if (examination is null)
            return;
        var current = examination.Otoscopy;
        var savedTime = image.Ear == OtoscopyEar.Left ? current?.LeftCapturedAt : current?.RightCapturedAt;
        var savedPath = image.Ear == OtoscopyEar.Left ? current?.LeftImagePath : current?.RightImagePath;
        if (savedTime == image.CapturedAt && savedPath is not null && File.Exists(savedPath))
            return;
        var path = await store.SaveImageAsync(examination, image.Ear, image.JpegData);
        await SaveChangesAsync(next =>
        {
            next.Otoscopy ??= new OtoscopyScreeningResult();
            if (image.Ear == OtoscopyEar.Left)
            {
                next.Otoscopy.LeftImagePath = path;
                next.Otoscopy.LeftCapturedAt = image.CapturedAt;
                next.Otoscopy.LeftNotCapturedReason = null;
            }
            else
            {
                next.Otoscopy.RightImagePath = path;
                next.Otoscopy.RightCapturedAt = image.CapturedAt;
                next.Otoscopy.RightNotCapturedReason = null;
            }
        });
        SetStatus("Bilden är sparad för denna undersökning.");
    }

    private void OtoscopyImageCaptured(object? sender, OtoscopyImageCapturedEventArgs e) => RunEvent(async () =>
    {
        if (ReferenceEquals(sender, otoscopy))
            await PersistImageAsync(e.Image);
    });

    private async Task ContinueAsync()
    {
        if (examination is null || !foreground)
            return;
        switch (examination.Stage)
        {
            case ScreeningStage.Setup:
                var reference = referenceEntry?.Text?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(reference))
                    throw new InvalidOperationException("Fyll i patientens eller undersökningens referens innan du börjar.");
                if (readyCheck?.IsChecked != true)
                    throw new InvalidOperationException("Bekräfta personalens förberedelser innan du börjar.");
                await SaveChangesAsync(next => next.Reference = reference);
                await PrepareAudioAsync();
                await SaveChangesAsync(next => next.Stage = ScreeningStage.OtoscopyInstructions);
                RenderStage();
                break;
            case ScreeningStage.OtoscopyInstructions:
                await SaveChangesAsync(next => next.Stage = ScreeningStage.Otoscopy);
                if (!foreground || closeRequested)
                {
                    await SaveChangesAsync(next => next.RecoverInterruptedStage());
                    RenderStage();
                    return;
                }
                RenderOtoscopy();
                break;
            case ScreeningStage.Otoscopy:
                await FinishOtoscopyAsync();
                break;
            case ScreeningStage.Ssq12Instructions:
                await StartPatientTestAsync(ScreeningStage.Ssq12, new Ssq12ScreeningSession());
                break;
            case ScreeningStage.QuickSipInstructions:
                await PrepareAudioAsync();
                await StartPatientTestAsync(ScreeningStage.QuickSip, new QuickSipScreeningSession(transducer!));
                break;
            case ScreeningStage.AudiometryInstructions:
                await PrepareAudioAsync();
                await StartPatientTestAsync(ScreeningStage.Audiometry, new AudiometryScreeningSession(transducer!));
                break;
            case ScreeningStage.Report:
                await EnsureReportAsync();
                break;
        }
    }

    private async Task PrepareAudioAsync()
    {
        testCancellation?.Dispose();
        testCancellation = new CancellationTokenSource();
        SetStatus("Kontrollerar ljudenhet, kalibrering och testmaterial…");
        try
        {
            transducer = await ScreeningAudioPreparation.PrepareAsync(testCancellation.Token, transducer);
        }
        catch (ScreeningTransducerSelectionRequiredException selection)
        {
            if (ownerPage is null)
                throw;
            var choices = selection.Choices.Select((item, index) => $"{index + 1}. {item.Name}").ToArray();
            var choice = await ownerPage.DisplayActionSheetAsync("Personal: välj konfigurerade hörlurar", "Avbryt", null, choices);
            int index = Array.IndexOf(choices, choice);
            if (index < 0)
                throw new OperationCanceledException();
            transducer = await ScreeningAudioPreparation.PrepareAsync(testCancellation.Token, selection.Choices[index]);
        }
        testCancellation.Token.ThrowIfCancellationRequested();
        if (!foreground || closeRequested)
            throw new OperationCanceledException();
        SetStatus($"Ljudenhet: {transducer.Name}");
    }

    private async Task FinishOtoscopyAsync()
    {
        if (otoscopy is null || examination is null)
            return;
        await SaveDraftAsync();
        var images = examination.Otoscopy!;
        bool missingLeft = string.IsNullOrWhiteSpace(images.LeftImagePath);
        bool missingRight = string.IsNullOrWhiteSpace(images.RightImagePath);
        if ((missingLeft && string.IsNullOrWhiteSpace(images.LeftNotCapturedReason)) ||
            (missingRight && string.IsNullOrWhiteSpace(images.RightNotCapturedReason)))
            throw new InvalidOperationException("Ta en bild av varje öra. Om en bild inte kan tas, ange orsaken för den sidan nedan.");
        if ((missingLeft || missingRight) && !await AskAsync("Bild saknas",
                "En eller båda öronbilderna saknas. Orsaken kommer att visas i rapporten. Vill du fortsätta?", "Fortsätt", "Tillbaka"))
            return;
        // The preview is not removed until its native camera has released it.
        await otoscopy.ReleaseForHostAsync();
        await SaveDraftAsync();
        await SaveChangesAsync(next => next.Stage = ScreeningStage.Ssq12Instructions);
        otoscopy.ImageCaptured -= OtoscopyImageCaptured;
        otoscopy = null;
        RenderStage();
        SetStatus("Otoskopin är sparad. Lämna nu surfplattan till patienten.");
    }

    private async Task StartPatientTestAsync(ScreeningStage stage, IScreeningTestSession session)
    {
        try
        {
            if (!foreground || closeRequested)
                throw new OperationCanceledException();
            await SaveChangesAsync(next => next.Stage = stage);
            pendingResult = null;
            testSession = session;
            if (!foreground || closeRequested)
                throw new OperationCanceledException();
            testCancellation?.Dispose();
            testCancellation = new CancellationTokenSource();
            session.Completed += TestCompleted;
            session.Failed += TestFailed;
            stageHost.Content = session.View;
            UpdateHeader();
            SetStatus("Testet pågår. Du kan pausa och be personalen om hjälp.");
            UpdateButtons();
            await session.StartAsync(testCancellation.Token);
        }
        catch
        {
            if (ReferenceEquals(testSession, session))
            {
                await StopTestAsync();
                await SaveChangesAsync(next => next.RecoverInterruptedStage());
                RenderStage();
            }
            else
                await session.DisposeAsync();
            throw;
        }
    }

    private void TestCompleted(object? sender, ScreeningTestCompletedEventArgs e) => RunEvent(async () =>
    {
        if (!ReferenceEquals(sender, testSession) || examination is null)
            return;
        pendingResult = e.Result;
        await CompletePendingTestAsync();
    });

    private async Task CompletePendingTestAsync()
    {
        var completedResult = pendingResult ?? throw new InvalidOperationException("Inget resultat väntar på att sparas.");
        // Persist completed data before removing the producing view.
        await SaveChangesAsync(next =>
        {
            switch (completedResult)
            {
                case Ssq12Result ssq when next.Stage == ScreeningStage.Ssq12:
                    next.Ssq12 = ssq;
                    next.Stage = ScreeningStage.QuickSipInstructions;
                    break;
                case QuickSipResult speech when next.Stage == ScreeningStage.QuickSip:
                    next.QuickSip = speech;
                    next.Stage = ScreeningStage.AudiometryInstructions;
                    break;
                case AudiometryResult tones when next.Stage == ScreeningStage.Audiometry:
                    next.Audiometry = tones;
                    next.CompletedAt = DateTimeOffset.Now;
                    next.Stage = ScreeningStage.Report;
                    break;
                default:
                    throw new InvalidOperationException("Testresultatet hör inte till den pågående delen.");
            }
        });
        pendingResult = null;
        await StopTestAsync();
        RenderStage();
        SetStatus("Resultatet är sparat.");
        if (examination?.Stage == ScreeningStage.Report)
            await EnsureReportAsync();
    }

    private void TestFailed(object? sender, ScreeningTestFailedEventArgs e) => RunEvent(async () =>
    {
        if (!ReferenceEquals(sender, testSession))
            return;
        await StopTestAsync();
        await SaveChangesAsync(next => next.RecoverInterruptedStage());
        RenderStage();
        throw new InvalidOperationException(e.Message, e.Exception);
    });

    private async Task StopTestAsync()
    {
        testCancellation?.Cancel();
        var current = testSession;
        if (current is not null)
        {
            // Invalidate queued completion callbacks before stopping timers/audio.
            current.Completed -= TestCompleted;
            current.Failed -= TestFailed;
            await current.StopAsync();
            await current.DisposeAsync();
            testSession = null;
        }
        testCancellation?.Dispose();
        testCancellation = null;
    }

    private async Task EnsureReportAsync()
    {
        if (testSession is not null)
            await StopTestAsync();
        if (otoscopy is not null)
            throw new InvalidOperationException("Avsluta kameran innan rapporten öppnas.");
        if (examination?.IsComplete != true)
            throw new InvalidOperationException("Alla fyra delar måste vara klara innan rapporten skapas.");
        bool replaceMissingReport = !File.Exists(examination.ReportPdfPath);
        if (report is null || !File.Exists(report.PdfPath) || report.PreviewImagePaths.Any(path => !File.Exists(path)))
        {
            SetStatus("Skapar rapporten…");
            report = await ScreeningReportRenderer.CreateAsync(examination,
                Path.Combine(store.GetExaminationDirectory(examination.Id), "Report"));
        }
        if (examination.ReportPdfPath != report.PdfPath || !examination.ReportPreviewPaths.SequenceEqual(report.PreviewImagePaths))
        {
            await SaveChangesAsync(next =>
            {
                next.ReportPdfPath = report.PdfPath;
                next.ReportPreviewPaths = report.PreviewImagePaths.ToArray();
                if (replaceMissingReport)
                {
                    next.PrintRequestedAt = null;
                    next.PrintConfirmedAt = null;
                }
            });
        }
        RenderReport();
        SetStatus(examination.CanStartNext
            ? "Utskriften är bekräftad. Nästa patient kan påbörjas."
            : "Rapporten är sparad på enheten. Skriv ut och bekräfta att alla sidor har kommit ut innan nästa patient.");
    }

    private async Task PrintAsync()
    {
        await EnsureReportAsync();
        // A print request is recorded separately from physical paper confirmation.
        // Cancellation never erases the report or permits a new patient.
        var result = await ScreeningReportPrinter.PrintAsync(report!);
        if (result.Requested)
            await SaveChangesAsync(next => next.RegisterPrintRequest(DateTimeOffset.Now));
        RenderReport();
        SetStatus(result.Message);
    }

    private async Task ConfirmPrintAsync()
    {
        if (examination?.PrintRequestedAt is null)
            throw new InvalidOperationException("Skicka rapporten till utskrift först.");
        if (!await AskAsync("Har rapporten kommit ut?",
                $"Kontrollera att alla {report?.PageCount ?? examination.ReportPreviewPaths.Length} sidor för {examination.Reference} har kommit ut ur skrivaren och är läsbara. " +
                "Bekräfta bara när du har kontrollerat pappersrapporten.", "Ja, alla sidor är kontrollerade", "Nej, behåll patienten"))
            return;
        await SaveChangesAsync(next => next.ConfirmPrinted(DateTimeOffset.Now));
        RenderReport();
        SetStatus("Pappersrapporten är bekräftad. Patientuppgifterna behålls tills du väljer Nästa patient.");
    }

    private async Task NextPatientAsync()
    {
        if (testSession is not null || otoscopy is not null)
            throw new InvalidOperationException("Den föregående testutrustningen måste ha stoppats innan nästa patient börjar.");
        if (examination?.CanStartNext != true)
            throw new InvalidOperationException("Bekräfta först att hela pappersrapporten har kommit ut ur skrivaren.");
        examination = await store.StartNextAsync(examination, string.Empty);
        report = null;
        transducer = null;
        TakeSharedContext();
        RenderStage();
        SetStatus("Ny undersökning. Föregående undersökning finns kvar i appens lokala arkiv.");
    }

    private async Task ExportReportAsync()
    {
        await EnsureReportAsync();
        using var source = File.OpenRead(report!.PdfPath);
        var result = await FileSaver.Default.SaveAsync($"horselscreening-{examination!.Id:N}.pdf", source, CancellationToken.None);
        SetStatus(result.IsSuccessful ? "PDF-rapporten har exporterats. Utskrift behöver fortfarande bekräftas före nästa patient."
            : result.IsCancelled ? "Exporten avbröts. Rapporten finns kvar i appen."
            : $"Exporten kunde inte bekräftas: {result.Exception?.Message}. Rapporten finns kvar i appen; kontrollera målet innan du försöker igen.");
    }

    /// <summary>Saves and pauses this patient; never resets their data or bypasses the print gate.</summary>
    public async Task<bool> CloseAsync()
    {
        if (closed)
            return true;
        closeRequested = true;
        testCancellation?.Cancel();
        var succeeded = await ExecuteAsync(async () =>
        {
            if (initialized)
            {
                if (pendingResult is not null)
                    await CompletePendingTestAsync();
                await SaveDraftAsync();
                await StopTestAsync();
                if (otoscopy is not null)
                {
                    await otoscopy.ReleaseForHostAsync();
                    await SaveDraftAsync();
                    otoscopy.ImageCaptured -= OtoscopyImageCaptured;
                    otoscopy = null;
                }
                await SaveChangesAsync(next => next.RecoverInterruptedStage());
            }
            closed = true;
            DetachLifecycle();
            Loaded -= ViewLoaded;
            Unloaded -= ViewUnloaded;
            if (sharedContextCaptured)
            {
                SharedSpeechTestObjects.CurrentParticipantID = previousParticipant;
                SharedSpeechTestObjects.GuiLanguage = previousLanguage;
            }
            if (ownsLease)
            {
                activeWorkflow.Release();
                ownsLease = false;
            }
            Closed?.Invoke(this, EventArgs.Empty);
        });
        if (!closed)
            closeRequested = false;
        return succeeded;
    }

    private async Task RetryAsync()
    {
        if (!initialized)
        {
            await InitializeAsync();
            return;
        }
        if (pendingResult is not null)
        {
            await CompletePendingTestAsync();
            return;
        }
        if (examination?.Stage == ScreeningStage.Report)
        {
            await EnsureReportAsync();
            return;
        }
        await SaveDraftAsync();
        await StopTestAsync();
        if (otoscopy is not null)
        {
            await otoscopy.ReleaseForHostAsync();
            otoscopy.ImageCaptured -= OtoscopyImageCaptured;
            otoscopy = null;
        }
        await SaveChangesAsync(next => next.RecoverInterruptedStage());
        RenderStage();
        SetStatus("Sparade resultat är kvar. Fortsätt när du är redo.");
    }

    private async Task<bool> AskAsync(string title, string message, string accept, string cancel)
    {
        if (ownerPage is null)
            throw new InvalidOperationException("Öppna vyn i en sida för att bekräfta åtgärden.");
        return await ownerPage.DisplayAlertAsync(title, message, accept, cancel);
    }

    private async Task<bool> ExecuteAsync(Func<Task> action)
    {
        await transitions.WaitAsync();
        if (closed)
        {
            transitions.Release();
            return false;
        }
        busy = true;
        UpdateButtons();
        try
        {
            retryButton.IsVisible = false;
            await action();
            return true;
        }
        catch (OperationCanceledException)
        {
            SetStatus("Åtgärden avbröts. Sparade resultat finns kvar.");
            return false;
        }
        catch (Exception exception)
        {
            Debug.WriteLine(exception);
            SetStatus($"Åtgärden kunde inte slutföras: {exception.Message}", error: true);
            return false;
        }
        finally
        {
            busy = false;
            UpdateButtons();
            transitions.Release();
        }
    }

    private void RunEvent(Func<Task> action)
    {
        async void Run() => await ExecuteAsync(action);
        if (Dispatcher.IsDispatchRequired)
            Dispatcher.Dispatch(Run);
        else
            Run();
    }

    private void SetStatus(string message, bool error = false)
    {
        statusLabel.Text = message;
        retryButton.IsVisible = error;
        statusLabel.TextColor = error ? Color.FromArgb("#A32929") : Color.FromArgb("#37475D");
    }
}
