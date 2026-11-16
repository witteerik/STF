// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

namespace STFM.Extension.Screening.Tests;

/// <summary>One fresh test run. All methods and completion events are on the UI thread.</summary>
public interface IScreeningTestSession : IAsyncDisposable
{
    View View { get; }
    event EventHandler<ScreeningTestCompletedEventArgs>? Completed;
    event EventHandler<ScreeningTestFailedEventArgs>? Failed;
    Task StartAsync(CancellationToken cancellationToken = default);
    /// <summary>Stops and invalidates the current run. Restart with a new session.</summary>
    Task StopAsync();
}

public sealed class ScreeningTestCompletedEventArgs(object result) : EventArgs
{
    public object Result { get; } = result;
}

public sealed class ScreeningTestFailedEventArgs(string message, Exception? exception = null) : EventArgs
{
    public string Message { get; } = message;
    public Exception? Exception { get; } = exception;
}
