// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using STFM.Views;
using STFN.Core;
using STFN.Core.Audio.Formats;

namespace STFM.Extension.Screening.Tests;

public sealed class AudiometryScreeningSession(AudioSystemSpecification transducer) : IScreeningTestSession
{
    private readonly ContentView host = new();
    private UoPta? test;
    private UoAudView? audiometer;
    private ScreeningAudioLease? lease;
    private CancellationTokenRegistration cancellationRegistration;
    private bool previousSave;
    private bool previousLog;
    private bool started;
    private bool running;
    private bool disposed;

    public View View => host;
    public event EventHandler<ScreeningTestCompletedEventArgs>? Completed;
    public event EventHandler<ScreeningTestFailedEventArgs>? Failed;

    public Task StartAsync(CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        if (started) throw new InvalidOperationException("Create a new audiometry session to restart.");
        cancellationToken.ThrowIfCancellationRequested();
        started = true;
        try
        {
            lease = ScreeningAudioLease.Acquire(transducer);
            previousSave = UoPta.SaveAudiogramDataToFile;
            previousLog = UoPta.LogTrialData;
            UoPta.SaveAudiogramDataToFile = false;
            UoPta.LogTrialData = false;
            Globals.StfBase.SoundPlayer.FatalPlayerError += FatalPlayerError;
            ScreeningAudioPreparation.ConfigurePlayer(transducer, 48000, 32,
                WaveFormat.WaveFormatEncodings.IeeeFloatingPoints, 0.1);
            test = new UoPta(PtaTestProtocol: UoPta.PtaTestProtocols.SAME96_Screening);
            audiometer = new UoAudView(test) { StopWhenUnloaded = true };
            audiometer.Finished += Finished;
            audiometer.Failed += AudiometerFailed;
            host.Content = audiometer;
            running = true;
            cancellationRegistration = cancellationToken.Register(() =>
                MainThread.BeginInvokeOnMainThread(() => { _ = StopAsync(); }));
            audiometer.StartTest();
        }
        catch
        {
            StopCore();
            throw;
        }
        return Task.CompletedTask;
    }

    private void Finished(object? sender, EventArgs args)
    {
        if (!running || test is null || !test.IsCompleted()) return;
        var result = new AudiometryResult(DateTimeOffset.Now,
            test.GetThresholdSnapshot().Select(point => new AudiometryPoint(
                point.Side == STFN.Core.Utils.EnumCollection.Sides.Left ? "Left" : "Right",
                point.Frequency, point.Level, point.Status switch
                {
                    UoPta.PtaSubTest.ThresholdStatuses.Reached => AudiometryPointStatus.Reached,
                    UoPta.PtaSubTest.ThresholdStatuses.Unreached => AudiometryPointStatus.NoResponse,
                    UoPta.PtaSubTest.ThresholdStatuses.Uncertain => AudiometryPointStatus.Uncertain,
                    _ => AudiometryPointStatus.NotMeasured
                }, point.IsReliabilityCheck, point.AtScreeningLevel)).ToArray(),
            transducer.Name, test.PtaTestProtocol.ToString(), test.ScreeningLevel, test.MaxPresentationLevel);
        StopCore();
        Completed?.Invoke(this, new ScreeningTestCompletedEventArgs(result));
    }

    private void FatalPlayerError() => MainThread.BeginInvokeOnMainThread(() =>
        Fail(new InvalidOperationException("Ljudenheten slutade fungera. Testet har avbrutits och behöver göras om.")));

    private void AudiometerFailed(object? sender, UoAudErrorEventArgs args) => Fail(args.Exception);

    private void Fail(Exception exception)
    {
        if (!running) return;
        try { StopCore(); }
        catch (Exception stopException) { System.Diagnostics.Debug.WriteLine(stopException); }
        Failed?.Invoke(this, new ScreeningTestFailedEventArgs(exception.Message, exception));
    }

    private void StopCore()
    {
        running = false;
        cancellationRegistration.Dispose();
        if (audiometer is not null)
        {
            audiometer.Finished -= Finished;
            audiometer.Failed -= AudiometerFailed;
        }
        try
        {
            if (audiometer is not null) audiometer.StopTest();
            else if (lease is not null) ScreeningAudioPreparation.StopPlayback();
        }
        finally
        {
            if (lease is not null)
            {
                Globals.StfBase.SoundPlayer.FatalPlayerError -= FatalPlayerError;
                UoPta.SaveAudiogramDataToFile = previousSave;
                UoPta.LogTrialData = previousLog;
                lease.Dispose();
                lease = null;
            }
        }
    }

    public Task StopAsync()
    {
        StopCore();
        return Task.CompletedTask;
    }

    public ValueTask DisposeAsync()
    {
        if (!disposed)
        {
            StopCore();
            disposed = true;
            host.Content = null;
        }
        return ValueTask.CompletedTask;
    }
}
