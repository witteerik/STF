// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

#if WINDOWS
using System.Diagnostics;
using Windows.Devices.Enumeration;
using Windows.Foundation;
using Windows.Graphics.Imaging;
using Windows.Media.Capture;
using Windows.Media.Capture.Frames;
using Windows.Media.MediaProperties;
using Windows.Storage.Streams;

namespace STFM.Extension.Otoscopy;

/// <summary>
/// Windows USB-camera capture using Windows' own media stack. Frames are selected
/// once at connection; preview and snapshots share the same unmirrored image.
/// No additional camera SDK or Windows native redistributable is required.
/// </summary>
internal sealed class WindowsUsbOtoscopyCamera : IOtoscopyCamera
{
    private static readonly object ProcessShutdownLock = new();
    private static Task processShutdown = Task.CompletedTask;
    private static readonly TimeSpan ShutdownTimeout = TimeSpan.FromSeconds(15);
    private readonly Image preview = new() { Aspect = Aspect.AspectFit };
    private readonly object frameLock = new();
    private MediaCapture? capture;
    private MediaFrameReader? reader;
    private DeviceWatcher? watcher;
    private bool watcherEnumerated;
    private volatile bool acceptingFrames;
    private volatile bool connected;
    private Task frameTask = Task.CompletedTask;
    private Task shutdownTask = Task.CompletedTask;
    private bool shutdownWaitTimedOut;
    private TaskCompletionSource<bool>? firstFrame;
    private byte[]? latestJpeg;
    private long lastPreviewTimestamp;
    private long lastFrameTimestamp;
    private int previewUpdateQueued;
    private int generation;
    private string details = "No USB camera has been opened yet.";
    private string? lastError;

    public View Preview => preview;
    public bool IsPreviewReady => preview.Handler is not null;
    public bool IsConnected => connected;
    public string Diagnostics => $"{details}\nCapture backend: Windows MediaCapture\nLast camera error: {lastError ?? "none"}";
    public event EventHandler? DevicesChanged;
    public event EventHandler<OtoscopyCameraErrorEventArgs>? Error;

    public async Task<IReadOnlyList<OtoscopyCameraDevice>> GetCamerasAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var groups = await MediaFrameSourceGroup.FindAllAsync();
        cancellationToken.ThrowIfCancellationRequested();
        return groups.Where(IsUsbColourGroup)
            .Select(group => new OtoscopyCameraDevice(group.Id,
                string.IsNullOrWhiteSpace(group.DisplayName) ? "USB camera" : group.DisplayName))
            .OrderBy(device => device.Name, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    private static bool IsUsbColourGroup(MediaFrameSourceGroup group)
        => group.SourceInfos.Any(info => info.SourceKind == MediaFrameSourceKind.Color && IsUsbSource(info));

    private static bool IsUsbSource(MediaFrameSourceInfo info)
    {
        var device = info.DeviceInformation;
        if (device is null)
            return false;
        // USB interface paths contain a VID and PID. Exclude cameras Windows
        // explicitly identifies as mounted in the device's enclosure.
        return device.Id.Contains("vid_", StringComparison.OrdinalIgnoreCase)
            && device.Id.Contains("pid_", StringComparison.OrdinalIgnoreCase)
            && (device.EnclosureLocation is null || device.EnclosureLocation.Panel == Panel.Unknown);
    }

    public void StartWatching()
    {
        if (watcher is not null)
            return;
        watcherEnumerated = false;
        watcher = DeviceInformation.CreateWatcher(MediaFrameSourceGroup.GetDeviceSelector());
        watcher.Added += CameraAdded;
        watcher.Removed += CameraRemoved;
        watcher.EnumerationCompleted += CameraEnumerationCompleted;
        watcher.Start();
    }

    public void StopWatching()
    {
        var old = watcher;
        watcher = null;
        if (old is null)
            return;
        old.Added -= CameraAdded;
        old.Removed -= CameraRemoved;
        old.EnumerationCompleted -= CameraEnumerationCompleted;
        if (old.Status is DeviceWatcherStatus.Started or DeviceWatcherStatus.EnumerationCompleted)
            old.Stop();
    }

    private void CameraAdded(DeviceWatcher sender, DeviceInformation args)
    {
        if (ReferenceEquals(sender, watcher) && watcherEnumerated)
            DevicesChanged?.Invoke(this, EventArgs.Empty);
    }

    private void CameraRemoved(DeviceWatcher sender, DeviceInformationUpdate args)
    {
        if (ReferenceEquals(sender, watcher))
            DevicesChanged?.Invoke(this, EventArgs.Empty);
    }

    private void CameraEnumerationCompleted(DeviceWatcher sender, object args)
    {
        // Initial GetCamerasAsync already obtains the complete list. Do not
        // trigger another connect while that first connection is starting.
        if (ReferenceEquals(sender, watcher))
            watcherEnumerated = true;
    }

    public async Task ConnectAsync(OtoscopyCameraDevice selected, CancellationToken cancellationToken)
    {
        await DisconnectAsync();
        await WaitForPendingShutdownAsync(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        lastError = null;
        details = $"Camera: {selected.Name}\nWindows device ID: {selected.Id}";
        var groups = await MediaFrameSourceGroup.FindAllAsync();
        var group = groups.FirstOrDefault(item => item.Id == selected.Id && IsUsbColourGroup(item))
            ?? throw new InvalidOperationException("The USB camera is no longer available. Reconnect it and try again.");
        details += $"\nUSB interface: {group.SourceInfos.First(item => IsUsbSource(item)).DeviceInformation.Id}";
        cancellationToken.ThrowIfCancellationRequested();
        await WaitForPendingShutdownAsync(cancellationToken);
        var newCapture = new MediaCapture();
        capture = newCapture;
        newCapture.Failed += CaptureFailed;
        try
        {
            await newCapture.InitializeAsync(new MediaCaptureInitializationSettings
            {
                SourceGroup = group,
                StreamingCaptureMode = StreamingCaptureMode.Video,
                SharingMode = MediaCaptureSharingMode.ExclusiveControl,
                MemoryPreference = MediaCaptureMemoryPreference.Cpu
            });
            cancellationToken.ThrowIfCancellationRequested();
            var source = newCapture.FrameSources.Values.FirstOrDefault(item =>
                item.Info.SourceKind == MediaFrameSourceKind.Color && IsUsbSource(item.Info))
                ?? throw new InvalidOperationException("The camera does not expose a supported colour video stream.");
            var format = source.SupportedFormats.Where(item => item.VideoFormat is not null)
                .OrderBy(item => Math.Abs((long)item.VideoFormat.Width * item.VideoFormat.Height - 1280L * 720))
                .ThenBy(item => Math.Abs(item.FrameRate.Denominator == 0 ? 30 :
                    (double)item.FrameRate.Numerator / item.FrameRate.Denominator - 30))
                .FirstOrDefault();
            if (format is not null)
            {
                await source.SetFormatAsync(format);
                details += $"\nSelected mode: {format.VideoFormat.Width}x{format.VideoFormat.Height} " +
                    $"{format.FrameRate.Numerator}/{format.FrameRate.Denominator} fps {format.Subtype}";
            }
            cancellationToken.ThrowIfCancellationRequested();
            // The subtype overload returns independent frame copies. Holding a
            // frame while encoding a snapshot therefore cannot exhaust the driver pool.
            var newReader = await newCapture.CreateFrameReaderAsync(source, MediaEncodingSubtypes.Bgra8);
            reader = newReader;
            newReader.AcquisitionMode = MediaFrameReaderAcquisitionMode.Realtime;
            firstFrame = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            latestJpeg = null;
            lastPreviewTimestamp = 0;
            lastFrameTimestamp = 0;
            acceptingFrames = true;
            newReader.FrameArrived += FrameArrived;
            var status = await newReader.StartAsync();
            if (status != MediaFrameReaderStartStatus.Success)
                throw new InvalidOperationException($"The USB video stream could not start: {status}.");
            await firstFrame.Task.WaitAsync(TimeSpan.FromSeconds(20), cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (lastError is not null)
                throw new InvalidOperationException(lastError);
            connected = true;
        }
        catch (Exception exception)
        {
            lastError ??= exception.Message;
            await DisconnectAsync();
            throw;
        }
    }

    private void FrameArrived(MediaFrameReader sender, MediaFrameArrivedEventArgs args)
    {
        lock (frameLock)
        {
            if (!acceptingFrames || !ReferenceEquals(sender, reader) || !frameTask.IsCompleted)
                return;
            long timestamp = Stopwatch.GetTimestamp();
            // A ten-frame/s preview limits allocation/encoding cost while keeping
            // full camera resolution for the image retained by the Capture buttons.
            if (lastPreviewTimestamp != 0 && Stopwatch.GetElapsedTime(lastPreviewTimestamp, timestamp).TotalMilliseconds < 100)
                return;
            lastPreviewTimestamp = timestamp;
            frameTask = ReadFrameAsync(sender, generation);
        }
    }

    private async Task ReadFrameAsync(MediaFrameReader source, int session)
    {
        try
        {
            using var frame = source.TryAcquireLatestFrame();
            using var sourceBitmap = frame?.VideoMediaFrame?.SoftwareBitmap;
            if (sourceBitmap is null)
                return;
            using var bitmap = SoftwareBitmap.Convert(sourceBitmap, BitmapPixelFormat.Bgra8, BitmapAlphaMode.Ignore);
            using var stream = new InMemoryRandomAccessStream();
            var options = new BitmapPropertySet
            {
                ["ImageQuality"] = new BitmapTypedValue(0.95f, PropertyType.Single)
            };
            var encoder = await BitmapEncoder.CreateAsync(BitmapEncoder.JpegEncoderId, stream, options);
            encoder.SetSoftwareBitmap(bitmap);
            await encoder.FlushAsync();
            stream.Seek(0);
            using var dataReader = new DataReader(stream.GetInputStreamAt(0));
            var bytes = new byte[checked((int)stream.Size)];
            await dataReader.LoadAsync((uint)bytes.Length);
            dataReader.ReadBytes(bytes);
            if (!acceptingFrames || session != Volatile.Read(ref generation))
                return;
            Volatile.Write(ref latestJpeg, bytes);
            Volatile.Write(ref lastFrameTimestamp, Stopwatch.GetTimestamp());
            firstFrame?.TrySetResult(true);
            if (Interlocked.Exchange(ref previewUpdateQueued, 1) == 0)
            {
                // Queue at most one UI update, so a busy UI cannot accumulate a
                // separate full-resolution JPEG for every pending preview frame.
                MainThread.BeginInvokeOnMainThread(() =>
                {
                    Interlocked.Exchange(ref previewUpdateQueued, 0);
                    var current = Volatile.Read(ref latestJpeg);
                    if (acceptingFrames && session == generation && current is not null)
                        preview.Source = ImageSource.FromStream(() => new MemoryStream(current, writable: false));
                });
            }
        }
        catch (Exception exception)
        {
            if (acceptingFrames && session == Volatile.Read(ref generation))
                SetError(exception.Message);
        }
    }

    public Task<byte[]> CaptureJpegAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var bytes = Volatile.Read(ref latestJpeg);
        if (!connected || bytes is null)
            throw new InvalidOperationException("Wait for a live USB camera preview before taking a picture.");
        if (Stopwatch.GetElapsedTime(Volatile.Read(ref lastFrameTimestamp)).TotalSeconds > 2)
            throw new InvalidOperationException("No recent camera frame is available. Reconnect the USB camera and try again.");
        return Task.FromResult((byte[])bytes.Clone());
    }

    private void CaptureFailed(MediaCapture sender, MediaCaptureFailedEventArgs args)
    {
        if (ReferenceEquals(sender, capture))
            SetError(args.Message);
    }

    private void SetError(string message)
    {
        lastError = message;
        connected = false;
        firstFrame?.TrySetException(new InvalidOperationException(message));
        Error?.Invoke(this, new OtoscopyCameraErrorEventArgs(message));
    }

    public async Task DisconnectAsync()
    {
        connected = false;
        Task pendingFrame;
        MediaFrameReader? oldReader;
        MediaCapture? oldCapture;
        lock (frameLock)
        {
            acceptingFrames = false;
            generation++;
            oldReader = reader;
            reader = null;
            pendingFrame = frameTask;
            oldCapture = capture;
            capture = null;
        }
        if (oldReader is not null)
            oldReader.FrameArrived -= FrameArrived;
        if (oldCapture is not null)
            oldCapture.Failed -= CaptureFailed;
        if (oldReader is not null || oldCapture is not null)
        {
            shutdownWaitTimedOut = false;
            var completion = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (ProcessShutdownLock)
                processShutdown = processShutdown.IsCompletedSuccessfully ? completion.Task : Task.WhenAll(processShutdown, completion.Task);
            shutdownTask = completion.Task;
            _ = ShutdownAsync(oldReader, oldCapture, pendingFrame, completion);
        }
        if (shutdownWaitTimedOut && !shutdownTask.IsCompleted)
            throw new InvalidOperationException("The USB camera is still stopping. Save your images and restart the app.");
        try
        {
            await shutdownTask.WaitAsync(ShutdownTimeout);
            latestJpeg = null;
            preview.Source = null;
        }
        catch (TimeoutException) when (!shutdownTask.IsFaulted)
        {
            shutdownWaitTimedOut = true;
            lastError = "USB camera shutdown has not finished. Save your images and restart the app.";
            throw new InvalidOperationException(lastError);
        }
        catch (Exception exception)
        {
            lastError = "USB camera shutdown failed. Save your images and restart the app.";
            throw new InvalidOperationException(lastError, exception);
        }
    }

    private static async Task ShutdownAsync(MediaFrameReader? oldReader, MediaCapture? oldCapture,
        Task pendingFrame, TaskCompletionSource<bool> completion)
    {
        try
        {
            // Do not dispose a reader while a frame callback is still encoding.
            await pendingFrame;
            if (oldReader is not null)
                await oldReader.StopAsync();
            oldReader?.Dispose();
            oldCapture?.Dispose();
            completion.TrySetResult(true);
        }
        catch (Exception exception)
        {
            completion.TrySetException(exception);
        }
    }

    private async Task WaitForPendingShutdownAsync(CancellationToken cancellationToken)
    {
        var elapsed = Stopwatch.StartNew();
        while (true)
        {
            Task prior;
            lock (ProcessShutdownLock)
                prior = processShutdown;
            try
            {
                var remaining = ShutdownTimeout - elapsed.Elapsed;
                if (remaining <= TimeSpan.Zero)
                    throw new TimeoutException();
                await prior.WaitAsync(remaining, cancellationToken);
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                throw;
            }
            catch (Exception exception)
            {
                lastError = "An earlier USB camera has not shut down cleanly. Save your images and restart the app.";
                throw new InvalidOperationException(lastError, exception);
            }
            lock (ProcessShutdownLock)
            {
                if (ReferenceEquals(prior, processShutdown))
                    return;
            }
        }
    }

    public async ValueTask DisposeAsync()
    {
        StopWatching();
        await DisconnectAsync();
    }
}
#endif
