// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

#if ANDROID
using Android.Content;
using Android.Hardware.Usb;
using System.Globalization;
using SkiaSharp;
using VisioForge.Core;
using VisioForge.Core.GStreamer.Android.UVC;
using VisioForge.Core.Types;
using VisioForge.Core.Types.Events;
using VisioForge.Core.Types.X.Sources.AndroidUVC;
using VisioForge.Core.VideoCaptureX;
using VideoView = VisioForge.Core.UI.MAUI.VideoView;

namespace STFM.Extension.Otoscopy;

/// <summary>Direct USB/UVC capture. It never falls back to a built-in camera.</summary>
internal sealed class AndroidUsbOtoscopyCamera : IOtoscopyCamera
{
    private static readonly SemaphoreSlim SdkInitializationLock = new(1, 1);
    private readonly VideoView preview = new()
    {
        HorizontalOptions = LayoutOptions.Fill,
        VerticalOptions = LayoutOptions.Fill
    };
    private readonly Context context = global::Android.App.Application.Context;
    private UsbChangeReceiver? receiver;
    private VideoCaptureCoreX? core;
    private string? connectedDeviceId;
    private HashSet<string> knownDeviceIds = [];
    private volatile bool isConnected;
    private int deviceGeneration;
    private string? lastError;
    private string deviceDetails = "No USB camera has been opened yet.";

    public View Preview => preview;
    public bool IsPreviewReady => preview.GetVideoView() is not null;
    public bool IsConnected => isConnected;
    public IReadOnlyList<OtoscopyCameraMode> Modes { get; private set; } = [];
    public string? RequestedModeDescription { get; private set; }
    public string Diagnostics => $"{deviceDetails}\nRequested mode: {RequestedModeDescription ?? "not started"}\n" +
        $"Last SDK error: {lastError ?? "none"}\n" +
        "Advertised modes usable by this SDK (MJPEG/YUY2):\n" +
        (Modes.Count == 0 ? "No modes could be read." : string.Join("\n", Modes.Select(mode => mode.Name)));
    public event EventHandler? DevicesChanged;
    public event EventHandler? ModesChanged;
    public event EventHandler<OtoscopyCameraErrorEventArgs>? Error;

    public IReadOnlyList<OtoscopyCameraDevice> GetCameras()
        => AndroidUVCDevices.FindCameras()
            .Select(device => new OtoscopyCameraDevice(device.DeviceName,
                string.IsNullOrWhiteSpace(device.ProductName) ? "USB camera" : device.ProductName,
                $"{device.VendorId:X4}:{device.ProductId:X4}"))
            .OrderBy(device => device.Id, StringComparer.Ordinal)
            .ToArray();

    public void StartWatching()
    {
        if (receiver is not null)
            return;

        knownDeviceIds = GetCameras().Select(device => device.Id).ToHashSet(StringComparer.Ordinal);
        var newReceiver = new UsbChangeReceiver(UsbDevicesChanged);
        var filter = new IntentFilter();
        filter.AddAction(UsbManager.ActionUsbDeviceAttached);
        filter.AddAction(UsbManager.ActionUsbDeviceDetached);
        try
        {
            if (OperatingSystem.IsAndroidVersionAtLeast(33))
                context.RegisterReceiver(newReceiver, filter, ReceiverFlags.NotExported);
            else
                context.RegisterReceiver(newReceiver, filter);
            receiver = newReceiver;
        }
        catch
        {
            newReceiver.Dispose();
            throw;
        }
    }

    public void StopWatching()
    {
        var oldReceiver = receiver;
        receiver = null;
        if (oldReceiver is null)
            return;
        try
        {
            context.UnregisterReceiver(oldReceiver);
        }
        finally
        {
            oldReceiver.Dispose();
        }
    }

    public async Task ConnectAsync(OtoscopyCameraDevice selected, string? modeId, CancellationToken cancellationToken)
    {
        await DisconnectAsync();
        cancellationToken.ThrowIfCancellationRequested();
        lastError = null;
        RequestedModeDescription = null;
        Modes = [];
        deviceDetails = $"Camera: {selected.Name}";
        var nativeView = preview.GetVideoView()
            ?? throw new InvalidOperationException("The camera preview is not ready. Tap Reconnect to try again.");

        // Android also requires CAMERA permission for direct USB video capture.
        var permission = await Permissions.CheckStatusAsync<Permissions.Camera>();
        if (permission != PermissionStatus.Granted)
            permission = await Permissions.RequestAsync<Permissions.Camera>();
        cancellationToken.ThrowIfCancellationRequested();
        if (permission != PermissionStatus.Granted)
            throw new InvalidOperationException("Allow camera access in Android settings, then tap Reconnect.");

        var device = AndroidUVCDevices.FindCameras().FirstOrDefault(camera => camera.DeviceName == selected.Id)
            ?? throw new InvalidOperationException("The USB camera was disconnected. Reconnect the cable to try again.");
        deviceDetails = $"Camera: {selected.Name}\nUSB VID:PID: {device.VendorId:X4}:{device.ProductId:X4}\n" +
            $"Android: {(global::Android.OS.Build.VERSION.Release)}; tablet: {(global::Android.OS.Build.Model)}\n" +
            $"VisioForge managed SDK: {VideoCaptureCoreX.SDK_Version}";

        if (!await AndroidUVCDevices.RequestPermissionAsync(device, cancellationToken))
            throw new InvalidOperationException("USB camera access was not granted. Tap Reconnect to try again.");
        cancellationToken.ThrowIfCancellationRequested();

        await EnsureSdkInitializedAsync(cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();

        // GetModes briefly opens the USB device. Only enumerate after disposing the
        // previous engine, before starting the next one, and off the UI thread.
        // Retain the SDK's exact VideoFrameRate (including fractional rates).
        var advertised = await Task.Run(() => AndroidUVCDevices.GetModes(device), cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var modesById = advertised
            .GroupBy(ModeId, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.First(), StringComparer.Ordinal);
        Modes = modesById.Select(pair => new OtoscopyCameraMode(pair.Key, pair.Value.ToString())).ToArray();
        ModesChanged?.Invoke(this, EventArgs.Empty);

        AndroidUVCCameraMode? requestedMode;
        if (!string.IsNullOrEmpty(modeId))
        {
            if (modesById.Count == 0)
                throw new InvalidOperationException("Could not read the camera modes. Unplug/reconnect the camera to retry the selected mode.");
            if (!modesById.TryGetValue(modeId, out requestedMode))
                throw new InvalidOperationException("That video mode is no longer available. Select another mode or reconnect with Automatic.");
        }
        else
        {
            // Preserve the original preference for the working camera: closest pixel
            // count to 1280x720, then MJPEG, then the advertised rate nearest 30 fps.
            requestedMode = advertised
                .OrderBy(mode => Math.Abs((long)mode.Width * mode.Height - 1280L * 720))
                .ThenBy(mode => mode.Format == AndroidUVCFrameFormat.MJPEG ? 0 : 1)
                .ThenBy(mode => Math.Abs(mode.FrameRate.ToDouble(CultureInfo.InvariantCulture) - 30))
                .FirstOrDefault();
        }

        var source = new AndroidUVCSourceSettings
        {
            Device = device,
            Width = 1280,
            Height = 720,
            FrameRate = new VideoFrameRate(30)
        };
        if (requestedMode is not null)
        {
            source.Width = requestedMode.Width;
            source.Height = requestedMode.Height;
            source.FrameRate = requestedMode.FrameRate;
            source.Format = requestedMode.Format;
            RequestedModeDescription = requestedMode.ToString();
        }
        else
        {
            // An empty list may also mean a temporary device-open failure. Preserve
            // the previous automatic attempt; diagnostics distinguish this case.
            RequestedModeDescription = "Automatic: 1280x720 @ 30 (mode list unavailable)";
        }
        global::Android.Util.Log.Info("Otoscopy", Diagnostics);

        connectedDeviceId = selected.Id;
        var generation = Volatile.Read(ref deviceGeneration);

        var newCore = new VideoCaptureCoreX(nativeView)
        {
            Video_Source = source,
            Video_Play = true,
            Audio_Play = false,
            Audio_Record = false,
            Snapshot_Grabber_Enabled = true,
            Video_Renderer_IsSync = false,
            Video_Output_IsSync = false
        };

        core = newCore;
        newCore.OnError += CoreOnError;
        try
        {
            // StartAsync and snapshots have no cancellable overload. Keep ownership until
            // each native operation finishes; never dispose an engine underneath it.
            if (!await newCore.StartAsync())
                throw new InvalidOperationException(lastError ?? "Unable to start the USB camera. Close other camera apps and tap Reconnect.");
            cancellationToken.ThrowIfCancellationRequested();
            if (!await newCore.WaitForStartAsync(TimeSpan.FromSeconds(20), cancellationToken))
                throw new InvalidOperationException(lastError ?? "The camera did not start. Check the USB connection and tap Reconnect.");
            cancellationToken.ThrowIfCancellationRequested();

            if (generation != Volatile.Read(ref deviceGeneration) || lastError is not null ||
                !AndroidUVCDevices.FindCameras().Any(camera => camera.DeviceName == selected.Id))
                throw new InvalidOperationException(lastError ?? "The USB camera was disconnected while starting.");

            isConnected = true;
        }
        catch
        {
            await DisconnectAsync();
            throw;
        }
    }

    private static string ModeId(AndroidUVCCameraMode mode)
        => $"{mode.Format}:{mode.Width}:{mode.Height}:" +
           mode.FrameRate.ToDouble(CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture);

    public async Task<byte[]> CaptureJpegAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var activeCore = core;
        var generation = Volatile.Read(ref deviceGeneration);
        if (!isConnected || activeCore is null || !IsDevicePresent())
            throw new InvalidOperationException("Connect the USB camera before taking a picture.");

        using var bitmap = await activeCore.Snapshot_GetSKAsync();
        cancellationToken.ThrowIfCancellationRequested();
        if (bitmap is null)
            throw new InvalidOperationException("No camera frame is available yet. Wait for the preview and try again.");

        // Encode the original camera frame, not a screenshot of the scaled preview.
        var jpeg = await Task.Run(() =>
        {
            using var encoded = bitmap.Encode(SKEncodedImageFormat.Jpeg, 95)
                ?? throw new InvalidOperationException("The camera image could not be encoded.");
            return encoded.ToArray();
        }, cancellationToken);

        cancellationToken.ThrowIfCancellationRequested();
        if (!isConnected || generation != Volatile.Read(ref deviceGeneration) || !IsDevicePresent())
            throw new InvalidOperationException("The camera disconnected before the picture was captured. Please try again.");
        return jpeg;
    }

    public async Task DisconnectAsync()
    {
        isConnected = false;
        connectedDeviceId = null;
        var oldCore = core;
        core = null;
        if (oldCore is null)
            return;
        oldCore.OnError -= CoreOnError;
        await oldCore.DisposeAsync();
    }

    public async ValueTask DisposeAsync()
    {
        StopWatching();
        await DisconnectAsync();
        // VisioForgeX is shared across the process. Do not destroy another view's SDK.
    }

    private bool IsDevicePresent()
        => connectedDeviceId is not null &&
           AndroidUVCDevices.FindCameras().Any(device => device.DeviceName == connectedDeviceId);

    private static async Task EnsureSdkInitializedAsync(CancellationToken cancellationToken)
    {
        await SdkInitializationLock.WaitAsync(cancellationToken);
        try
        {
            if (!VisioForgeX.IsSDKInitialized)
                await VisioForgeX.InitSDKAsync();
        }
        finally
        {
            SdkInitializationLock.Release();
        }
    }

    private void CoreOnError(object? sender, ErrorsEventArgs args)
    {
        if (!ReferenceEquals(sender, core))
            return;
        lastError = args.Message;
        global::Android.Util.Log.Warn("Otoscopy", Diagnostics);
        isConnected = false;
        Interlocked.Increment(ref deviceGeneration);
        Error?.Invoke(this, new OtoscopyCameraErrorEventArgs(args.Message));
    }

    private void UsbDevicesChanged(Intent? intent)
    {
        try
        {
            var current = GetCameras().Select(device => device.Id).ToHashSet(StringComparer.Ordinal);
            bool changed = !knownDeviceIds.SetEquals(current);
            knownDeviceIds = current;
            // A quickly reconnected camera can reuse its device path before enumeration.
            // Honour the detach notification as well as the current list in that case.
            bool wasDetached = false;
            if (intent?.Action == UsbManager.ActionUsbDeviceDetached && connectedDeviceId is not null)
            {
                UsbDevice? removedDevice;
                if (OperatingSystem.IsAndroidVersionAtLeast(33))
                    removedDevice = intent.GetParcelableExtra(UsbManager.ExtraDevice,
                        Java.Lang.Class.FromType(typeof(UsbDevice))) as UsbDevice;
                else
                {
#pragma warning disable CS0618 // Typed overload is available only from Android 13.
                    removedDevice = intent.GetParcelableExtra(UsbManager.ExtraDevice) as UsbDevice;
#pragma warning restore CS0618
                }
                wasDetached = removedDevice is null || removedDevice.DeviceName == connectedDeviceId;
            }
            // Invalidate immediately, before the view queues its serialized disconnect.
            bool disconnected = connectedDeviceId is not null &&
                (wasDetached || !current.Contains(connectedDeviceId));
            if (disconnected)
            {
                isConnected = false;
                Interlocked.Increment(ref deviceGeneration);
            }
            if (changed || disconnected)
                DevicesChanged?.Invoke(this, EventArgs.Empty);
        }
        catch (Exception ex)
        {
            isConnected = false;
            Interlocked.Increment(ref deviceGeneration);
            Error?.Invoke(this, new OtoscopyCameraErrorEventArgs(ex.Message));
        }
    }

    private sealed class UsbChangeReceiver(Action<Intent?> changed) : BroadcastReceiver
    {
        public override void OnReceive(Context? context, Intent? intent) => changed(intent);
    }
}
#endif
