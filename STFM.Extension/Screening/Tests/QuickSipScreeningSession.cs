// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using STFM.Views;
using STFN.Core;
using static STFN.Core.ResponseViewEvents;

namespace STFM.Extension.Screening.Tests;

/// <summary>
/// Drives the existing QuickSiP engine without constructing a hidden options form.
/// Each trial accepts one response; stale callbacks cannot advance a later trial.
/// </summary>
public sealed class QuickSipScreeningSession(AudioSystemSpecification transducer) : IScreeningTestSession
{
    private readonly ResponseView_Mafc response = new();
    private readonly List<(IDispatcherTimer Timer, EventHandler Handler)> timers = [];
    private QuickSiP_StaticSpeechTest? test;
    private SpeechTest? previousSpeechTest;
    private ScreeningAudioLease? lease;
    private CancellationTokenSource? cancellation;
    private CancellationTokenRegistration cancellationRegistration;
    private Task pendingAdvance = Task.CompletedTask;
    private int trialVersion;
    private DateTime trialStartedAt;
    private bool acceptingResponse;
    private bool started;
    private bool running;
    private bool disposed;

    public View View => response;
    public event EventHandler<ScreeningTestCompletedEventArgs>? Completed;
    public event EventHandler<ScreeningTestFailedEventArgs>? Failed;

    public async Task StartAsync(CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        if (started) throw new InvalidOperationException("Create a new QuickSiP session to restart.");
        cancellationToken.ThrowIfCancellationRequested();
        started = true;
        try
        {
            lease = ScreeningAudioLease.Acquire(transducer);
            previousSpeechTest = SharedSpeechTestObjects.CurrentSpeechTest;
            cancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            test = new QuickSiP_StaticSpeechTest(ScreeningAudioPreparation.QuickSipMaterialName);
            // Some base-class property setters consult this process-wide object.
            SharedSpeechTestObjects.CurrentSpeechTest = test;
            test.SkipGuiUpdates = true;
            test.Transducer = transducer;
            test.PopulateSoundSourceLocationCandidates();
            test.FixedResponseAlternativeCount = 3;
            var startLists = test.SpeechMaterial.GetAllRelativesAtLevel(SpeechMaterialComponent.LinguisticLevels.List, true, false);
            if (startLists.Count == 0) throw new InvalidOperationException("QuickSiP saknar testlistor.");
            test.StartList = startLists[SpeechTest.Randomizer.Next(startLists.Count)].PrimaryStringRepresentation;
            ScreeningAudioPreparation.ValidateQuickSipMaterial(test, test.StartList);
            var initialization = test.InitializeCurrentTest();
            if (!initialization.Item1)
                throw new InvalidOperationException(string.IsNullOrWhiteSpace(initialization.Item2)
                    ? "QuickSiP kunde inte förberedas. Kontrollera testmaterialet." : initialization.Item2);
            var media = test.MediaSet;
            Globals.StfBase.SoundPlayer.FatalPlayerError += FatalPlayerError;
            ScreeningAudioPreparation.ConfigurePlayer(transducer, media.WaveFileSampleRate, media.WaveFileBitDepth,
                media.WaveFileEncoding, test.SoundOverlapDuration);
            response.ResponseGiven += ResponseGiven;
            response.IsEnabled = true;
            running = true;
            cancellationRegistration = cancellationToken.Register(() => MainThread.BeginInvokeOnMainThread(() =>
            {
                StopImmediately();
            }));
            pendingAdvance = AdvanceAsync(null, trialVersion, delay: false);
            await pendingAdvance;
        }
        catch
        {
            await StopAsync();
            throw;
        }
    }

    private void ResponseGiven(object? sender, SpeechTestInputEventArgs args)
    {
        // The existing ResponseView deliberately raises this event on a worker thread.
        MainThread.BeginInvokeOnMainThread(() =>
        {
            if (!running || !acceptingResponse || args.LinguisticResponseTime < trialStartedAt) return;
            acceptingResponse = false;
            StopTimers();
            response.IsEnabled = false;
            pendingAdvance = AdvanceAsync(args, trialVersion, delay: true);
        });
    }

    private async Task AdvanceAsync(SpeechTestInputEventArgs? args, int version, bool delay)
    {
        var token = cancellation?.Token ?? CancellationToken.None;
        try
        {
            if (delay) await Task.Delay(300, token); // Keep the existing selected/timeout feedback visible.
            token.ThrowIfCancellationRequested();
            if (!running || test is null || version != trialVersion) return;
            if (args is not null && test.CurrentTestTrial is not null)
                test.CurrentTestTrial.TimedEventsList.Add(new Tuple<TestTrial.TimedTrialEvents, DateTime>(
                    TestTrial.TimedTrialEvents.ParticipantResponded, args.LinguisticResponseTime));

            // Sound-file loading/mixing must not block the patient-facing UI. Stop
            // cancels timers immediately and waits for this worker before releasing
            // the shared speech-test/audio ownership.
            var reply = await Task.Run(() => test.GetSpeechTestReply(this, args!), token);
            token.ThrowIfCancellationRequested();
            if (!running || version != trialVersion) return;
            switch (reply)
            {
                case SpeechTest.SpeechTestReplies.GotoNextTrial:
                    PresentTrial();
                    break;
                case SpeechTest.SpeechTestReplies.ContinueTrial:
                    acceptingResponse = true;
                    response.IsEnabled = true;
                    break;
                case SpeechTest.SpeechTestReplies.TestIsCompleted:
                    var observed = test.GetObservedTestTrials().ToArray();
                    var result = new QuickSipResult(DateTimeOffset.Now,
                        observed.Count(trial => trial.IsCorrect == true), observed.Length, test.GetTotalTrialCount(),
                        observed.Count(trial => string.IsNullOrEmpty(trial.Response)), test.GetAverageScore(),
                        transducer.Name, test.StartList);
                    StopImmediately();
                    Completed?.Invoke(this, new ScreeningTestCompletedEventArgs(result));
                    break;
                case SpeechTest.SpeechTestReplies.PauseTestingWithCustomInformation:
                    throw new InvalidOperationException(string.IsNullOrWhiteSpace(test.PauseInformation)
                        ? "QuickSiP avbröts och behöver göras om." : test.PauseInformation);
                default:
                    throw new InvalidOperationException("QuickSiP kunde inte slutföras. Be personalen kontrollera utrustningen.");
            }
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { }
        catch (Exception exception)
        {
            if (running) Fail(exception);
        }
    }

    private void PresentTrial()
    {
        if (test?.CurrentTestTrial is null) throw new InvalidOperationException("QuickSiP saknar nästa teststimulus.");
        StopTimers();
        response.InitializeNewTrial();
        response.IsEnabled = true;
        acceptingResponse = false; // Enabled only once response alternatives are visible.
        int version = ++trialVersion;
        trialStartedAt = DateTime.Now;
        test.CurrentTestTrial.TimedEventsList.Add(new Tuple<TestTrial.TimedTrialEvents, DateTime>(
            TestTrial.TimedTrialEvents.TrialStarted, trialStartedAt));
        var trial = test.CurrentTestTrial;
        foreach (var trialEvent in trial.TrialEventList)
        {
            var timer = response.Dispatcher.CreateTimer();
            timer.Interval = TimeSpan.FromMilliseconds(Math.Max(1, trialEvent.TickTime));
            timer.IsRepeating = false;
            timer.Tick += Tick;
            timers.Add((timer, Tick));
            timer.Start();
            void Tick(object? sender, EventArgs args)
            {
                timer.Stop();
                timer.Tick -= Tick;
                if (!running || version != trialVersion) return;
                try
                {
                    switch (trialEvent.Type)
                    {
                        case ResponseViewEvent.ResponseViewEventTypes.PlaySound:
                            trial.TimedEventsList.Add(new Tuple<TestTrial.TimedTrialEvents, DateTime>(
                                TestTrial.TimedTrialEvents.SoundStartedPlay, DateTime.Now));
                            if (!Globals.StfBase.SoundPlayer.SwapOutputSounds(ref trial.Sound))
                                throw new InvalidOperationException("Ljudenheten kunde inte spela QuickSiP-stimulus.");
                            break;
                        case ResponseViewEvent.ResponseViewEventTypes.StopSound:
                            ScreeningAudioPreparation.StopPlayback();
                            break;
                        case ResponseViewEvent.ResponseViewEventTypes.ShowResponseAlternatives:
                            response.ShowResponseAlternatives(trial.ResponseAlternativeSpellings);
                            acceptingResponse = true;
                            break;
                        case ResponseViewEvent.ResponseViewEventTypes.ShowResponseTimesOut:
                            response.ResponseTimesOut();
                            break;
                        case ResponseViewEvent.ResponseViewEventTypes.HideAll:
                            response.HideAllItems();
                            break;
                        default:
                            throw new InvalidOperationException($"Unsupported QuickSiP trial event: {trialEvent.Type}.");
                    }
                }
                catch (Exception exception) { Fail(exception); }
            }
        }
        response.UpdateTestFormProgressbar(test.GetObservedTestTrials().Count + 1, test.GetTotalTrialCount(), 0);
    }

    private void StopTimers()
    {
        foreach (var registration in timers)
        {
            registration.Timer.Stop();
            registration.Timer.Tick -= registration.Handler;
        }
        timers.Clear();
        response.StopAllTimers();
    }

    private void FatalPlayerError() => MainThread.BeginInvokeOnMainThread(() =>
        Fail(new InvalidOperationException("Ljudenheten slutade fungera. QuickSiP har avbrutits och behöver göras om.")));

    private void Fail(Exception exception)
    {
        if (!running) return;
        try { StopImmediately(); }
        catch (Exception stopException) { System.Diagnostics.Debug.WriteLine(stopException); }
        Failed?.Invoke(this, new ScreeningTestFailedEventArgs(exception.Message, exception));
    }

    private void StopImmediately()
    {
        running = false;
        acceptingResponse = false;
        ++trialVersion;
        cancellation?.Cancel();
        StopTimers();
        response.IsEnabled = false;
        response.HideAllItems();
        if (lease is not null) ScreeningAudioPreparation.StopPlayback();
    }

    public async Task StopAsync()
    {
        if (disposed) return;
        StopImmediately();
        // Await pending file/mixing work before restoring shared state. Completion
        // events use async handlers so this does not block their originating worker.
        await pendingAdvance;
        response.ResponseGiven -= ResponseGiven;
        cancellationRegistration.Dispose();
        if (lease is not null)
        {
            Globals.StfBase.SoundPlayer.FatalPlayerError -= FatalPlayerError;
            if (ReferenceEquals(SharedSpeechTestObjects.CurrentSpeechTest, test))
                SharedSpeechTestObjects.CurrentSpeechTest = previousSpeechTest;
            lease.Dispose();
            lease = null;
        }
    }

    public async ValueTask DisposeAsync()
    {
        if (disposed) return;
        await StopAsync();
        disposed = true;
        cancellation?.Dispose();
        cancellation = null;
    }
}
