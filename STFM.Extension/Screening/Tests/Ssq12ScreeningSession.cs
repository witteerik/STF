// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using STFM.SpecializedViews.SSQ12;

namespace STFM.Extension.Screening.Tests;

public sealed class Ssq12ScreeningSession : IScreeningTestSession
{
    private readonly SSQ12_MainView questionnaire;
    private readonly bool previousMinimalVersion;
    private CancellationTokenRegistration cancellationRegistration;
    private bool running;
    private bool disposed;

    public Ssq12ScreeningSession()
    {
        previousMinimalVersion = SSQ12_MainView.MinimalVersion;
        SSQ12_MainView.MinimalVersion = true;
        try { questionnaire = new SSQ12_MainView(); }
        catch { SSQ12_MainView.MinimalVersion = previousMinimalVersion; throw; }
        // The standalone questionnaire has desktop-sized side margins. Keep the
        // same questions readable when hosted on a portrait tablet.
        if (questionnaire.Content is ScrollView scroll)
            scroll.Padding = new Thickness(16, 20);
        questionnaire.Finished += Finished;
    }

    public View View => questionnaire;
    public event EventHandler<ScreeningTestCompletedEventArgs>? Completed;
    public event EventHandler<ScreeningTestFailedEventArgs>? Failed;

    public Task StartAsync(CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(disposed, this);
        cancellationToken.ThrowIfCancellationRequested();
        running = true;
        questionnaire.IsEnabled = true;
        cancellationRegistration = cancellationToken.Register(() => MainThread.BeginInvokeOnMainThread(() => { _ = StopAsync(); }));
        return Task.CompletedTask;
    }

    private void Finished(object? sender, EventArgs args)
    {
        if (!running || disposed) return;
        try
        {
            var answers = questionnaire.GetAnswerSnapshot();
            if (answers.Count != 12 || answers.Any(answer => !answer.Answered))
                throw new InvalidOperationException("Alla tolv frågor behöver besvaras, med en siffra eller Vet inte.");
            var hearingAid = questionnaire.GetHearingAidSnapshot();
            var result = new Ssq12Result(DateTimeOffset.Now,
                answers.Select(answer => new Ssq12Answer(answer.Number, answer.Question, answer.Rating, answer.DontKnow)).ToArray(),
                hearingAid.UsesHearingAids, hearingAid.Left, hearingAid.Right);
            running = false;
            questionnaire.IsEnabled = false;
            Completed?.Invoke(this, new ScreeningTestCompletedEventArgs(result));
        }
        catch (Exception exception)
        {
            Failed?.Invoke(this, new ScreeningTestFailedEventArgs(exception.Message, exception));
        }
    }

    public Task StopAsync()
    {
        running = false;
        questionnaire.IsEnabled = false;
        cancellationRegistration.Dispose();
        return Task.CompletedTask;
    }

    public async ValueTask DisposeAsync()
    {
        if (disposed) return;
        await StopAsync();
        disposed = true;
        questionnaire.Finished -= Finished;
        SSQ12_MainView.MinimalVersion = previousMinimalVersion;
    }
}
