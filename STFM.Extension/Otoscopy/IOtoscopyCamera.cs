// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

namespace STFM.Extension.Otoscopy;

internal sealed record OtoscopyCameraDevice(string Id, string Name)
{
    public override string ToString() => Name;
}

internal sealed class OtoscopyCameraErrorEventArgs(string message) : EventArgs
{
    public string Message { get; } = message;
}

// The view serializes connection, capture and disposal. Device/error notifications can
// arrive on any thread; the view marshals them back to its dispatcher.
internal interface IOtoscopyCamera : IAsyncDisposable
{
    View Preview { get; }
    bool IsPreviewReady { get; }
    bool IsConnected { get; }
    string Diagnostics { get; }
    event EventHandler? DevicesChanged;
    event EventHandler<OtoscopyCameraErrorEventArgs>? Error;
    Task<IReadOnlyList<OtoscopyCameraDevice>> GetCamerasAsync(CancellationToken cancellationToken);
    void StartWatching();
    void StopWatching();
    Task ConnectAsync(OtoscopyCameraDevice device, CancellationToken cancellationToken);
    Task<byte[]> CaptureJpegAsync(CancellationToken cancellationToken);
    Task DisconnectAsync();
}

internal static class OtoscopyCameraFactory
{
    public static IOtoscopyCamera Create()
    {
#if ANDROID
        return new AndroidUsbOtoscopyCamera();
#elif WINDOWS
        return new WindowsUsbOtoscopyCamera();
#else
        return new UnsupportedOtoscopyCamera();
#endif
    }
}

#if !ANDROID && !WINDOWS
internal sealed class UnsupportedOtoscopyCamera : IOtoscopyCamera
{
    public View Preview { get; } = new Label
    {
        Text = "USB otoscopy is available in the Android and Windows apps.",
        TextColor = Colors.White,
        HorizontalTextAlignment = TextAlignment.Center,
        VerticalTextAlignment = TextAlignment.Center
    };

    public bool IsPreviewReady => false;
    public bool IsConnected => false;
    public string Diagnostics => "USB otoscopy requires Android or Windows.";
    public event EventHandler? DevicesChanged { add { } remove { } }
    public event EventHandler<OtoscopyCameraErrorEventArgs>? Error { add { } remove { } }
    public Task<IReadOnlyList<OtoscopyCameraDevice>> GetCamerasAsync(CancellationToken cancellationToken)
        => Task.FromResult<IReadOnlyList<OtoscopyCameraDevice>>([]);
    public void StartWatching() { }
    public void StopWatching() { }
    public Task ConnectAsync(OtoscopyCameraDevice device, CancellationToken cancellationToken)
        => Task.FromException(new PlatformNotSupportedException("USB otoscopy requires Android or Windows."));
    public Task<byte[]> CaptureJpegAsync(CancellationToken cancellationToken)
        => Task.FromException<byte[]>(new PlatformNotSupportedException("USB otoscopy requires Android or Windows."));
    public Task DisconnectAsync() => Task.CompletedTask;
    public ValueTask DisposeAsync() => ValueTask.CompletedTask;
}
#endif
