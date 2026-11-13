// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Diagnostics;
using STFM.Extension.Otoscopy;

namespace STFM.Extension.Views;

/// <summary>
/// Previews an Android USB camera and retains one JPEG per ear. Captures are not saved
/// automatically: the host can use LeftImage, RightImage or ImageCaptured to save them.
/// Call ClearImages before reusing this instance for another examination.
/// </summary>
public partial class OtoscopyView : ContentView
{
    private readonly IOtoscopyCamera camera;
    private readonly SemaphoreSlim operations = new(1, 1);
    private CancellationTokenSource? sessionCancellation;
    private Page? ownerPage;
    private Window? ownerWindow;
    private bool loaded;
    private bool pageVisible = true;
    private bool windowActive = true;
    private bool manuallyStopped;
    private bool active;
    private bool busy;
    private bool watching;
    private bool autoConnectAttempted;
    private bool updatingPicker;
    private string? selectedCameraId;
    private int sessionVersion;
    private int examinationVersion;

    public OtoscopyView()
    {
        InitializeComponent();
        camera = OtoscopyCameraFactory.Create();
        PreviewHost.Add(camera.Preview);
        Loaded += ViewLoaded;
        Unloaded += ViewUnloaded;
        SizeChanged += ViewSizeChanged;
        camera.Preview.Loaded += PreviewReady;
        camera.Preview.HandlerChanged += PreviewReady;
#if !ANDROID
        StatusLabel.Text = "USB otoscopy is available in the Android app.";
        ReconnectButton.IsEnabled = false;
#endif
    }

    public OtoscopyImage? LeftImage { get; private set; }
    public OtoscopyImage? RightImage { get; private set; }

    /// <summary>Raised on the UI thread after a successful capture replaces the selected ear's image.</summary>
    public event EventHandler<OtoscopyImageCapturedEventArgs>? ImageCaptured;

    /// <summary>Clears both results and invalidates any capture already in progress. Call on the UI thread.</summary>
    public void ClearImages()
    {
        examinationVersion++;
        LeftImage = null;
        RightImage = null;
        LeftPreview.Source = null;
        RightPreview.Source = null;
        LeftPlaceholder.IsVisible = true;
        RightPlaceholder.IsVisible = true;
        CaptureLeftButton.Text = "Capture left";
        CaptureRightButton.Text = "Capture right";
        OnPropertyChanged(nameof(LeftImage));
        OnPropertyChanged(nameof(RightImage));
        UpdateButtons();
    }

    /// <summary>
    /// Releases the camera without clearing snapshots. Call on the UI thread before
    /// replacing the view in a custom host. Reconnect resumes it while still displayed.
    /// </summary>
    public async Task StopCameraAsync()
    {
        manuallyStopped = true;
        await UpdateActivityAsync();
    }

    private void ViewLoaded(object? sender, EventArgs e) => RunSafely(async () =>
    {
        loaded = true;
        manuallyStopped = false;
        pageVisible = true;
        windowActive = true;
        AttachLifecycle();
        await UpdateActivityAsync();
    });

    private void ViewUnloaded(object? sender, EventArgs e) => RunSafely(async () =>
    {
        loaded = false;
        DetachLifecycle();
        await UpdateActivityAsync();
    });

    private void AttachLifecycle()
    {
        DetachLifecycle();
        for (Element? parent = Parent; parent != null; parent = parent.Parent)
        {
            if (parent is Page page)
            {
                ownerPage = page;
                ownerPage.Appearing += PageAppearing;
                ownerPage.Disappearing += PageDisappearing;
                break;
            }
        }

        ownerWindow = Window;
        if (ownerWindow != null)
        {
            ownerWindow.Stopped += WindowStopped;
            ownerWindow.Resumed += WindowResumed;
            ownerWindow.Destroying += WindowDestroying;
        }
    }

    private void DetachLifecycle()
    {
        if (ownerPage != null)
        {
            ownerPage.Appearing -= PageAppearing;
            ownerPage.Disappearing -= PageDisappearing;
            ownerPage = null;
        }
        if (ownerWindow != null)
        {
            ownerWindow.Stopped -= WindowStopped;
            ownerWindow.Resumed -= WindowResumed;
            ownerWindow.Destroying -= WindowDestroying;
            ownerWindow = null;
        }
    }

    private void PageAppearing(object? sender, EventArgs e) => RunSafely(async () =>
    {
        pageVisible = true;
        await UpdateActivityAsync();
    });

    private void PageDisappearing(object? sender, EventArgs e) => RunSafely(async () =>
    {
        pageVisible = false;
        await UpdateActivityAsync();
    });

    private void WindowStopped(object? sender, EventArgs e) => RunSafely(async () =>
    {
        windowActive = false;
        await UpdateActivityAsync();
    });

    private void WindowResumed(object? sender, EventArgs e) => RunSafely(async () =>
    {
        windowActive = true;
        await UpdateActivityAsync();
    });

    private void WindowDestroying(object? sender, EventArgs e) => RunSafely(async () =>
    {
        loaded = false;
        DetachLifecycle();
        await UpdateActivityAsync();
    });

    private async Task UpdateActivityAsync()
    {
#if ANDROID
        bool shouldBeActive = loaded && pageVisible && windowActive && !manuallyStopped;
        if (shouldBeActive == active)
        {
            if (active)
                await TryAutoConnectAsync();
            else
                await DisconnectAsync("Camera paused.");
            return;
        }

        active = shouldBeActive;
        if (active)
        {
            camera.DevicesChanged += DevicesChanged;
            camera.Error += CameraError;
            try
            {
                camera.StartWatching();
                watching = true;
            }
            catch
            {
                camera.DevicesChanged -= DevicesChanged;
                camera.Error -= CameraError;
                active = false;
                watching = false;
                throw;
            }
            autoConnectAttempted = false;
            await TryAutoConnectAsync();
        }
        else
        {
            try
            {
                if (watching)
                    camera.StopWatching();
            }
            finally
            {
                if (watching)
                {
                    camera.DevicesChanged -= DevicesChanged;
                    camera.Error -= CameraError;
                    watching = false;
                }
                await DisconnectAsync("Camera paused.");
            }
        }
#endif
        UpdateButtons();
    }

    private void PreviewReady(object? sender, EventArgs e) => RunSafely(TryAutoConnectAsync);

    private Task TryAutoConnectAsync()
    {
        if (!active || autoConnectAttempted || !camera.IsPreviewReady)
            return Task.CompletedTask;

        autoConnectAttempted = true;
        return ConnectAsync();
    }

    private void DevicesChanged(object? sender, EventArgs e) => RunSafely(async () =>
    {
        if (active && camera.IsPreviewReady)
        {
            autoConnectAttempted = true;
            await ConnectAsync();
        }
    });

    private void CameraError(object? sender, OtoscopyCameraErrorEventArgs e)
    {
        // A native callback may already be queued when the user starts a new session.
        var version = Volatile.Read(ref sessionVersion);
        RunSafely(async () =>
        {
            if (active && version == sessionVersion)
                await DisconnectAsync($"Camera stopped: {e.Message} Tap Reconnect to try again.");
        });
    }

    private void ReconnectClicked(object? sender, EventArgs e) => RunSafely(async () =>
    {
        manuallyStopped = false;
        if (!active)
        {
            await UpdateActivityAsync();
        }
        else if (camera.IsPreviewReady)
        {
            autoConnectAttempted = true;
            await ConnectAsync(forceReconnect: true);
        }
    });

    private void CameraSelectionChanged(object? sender, EventArgs e)
    {
        if (updatingPicker || CameraPicker.SelectedItem is not OtoscopyCameraDevice device)
            return;
        selectedCameraId = device.Id;
        RunSafely(() => active ? ConnectAsync(forceReconnect: true) : Task.CompletedTask);
    }

    private async Task ConnectAsync(bool forceReconnect = false)
    {
        var previousCancellation = sessionCancellation;
        sessionCancellation = new CancellationTokenSource();
        var cancellation = sessionCancellation;
        var token = cancellation.Token;
        var version = ++sessionVersion;
        previousCancellation?.Cancel();
        busy = true;
        UpdateButtons();
        bool entered = false;
        try
        {
            await operations.WaitAsync(token);
            entered = true;
            previousCancellation?.Dispose();
            previousCancellation = null;
            if (!active || version != sessionVersion)
                return;

            var cameras = camera.GetCameras();
            var selected = cameras.FirstOrDefault(item => item.Id == selectedCameraId) ?? cameras.FirstOrDefault();
            bool sameDevice = selected?.Id == selectedCameraId;
            selectedCameraId = selected?.Id;
            UpdateCameraPicker(cameras, selected);

            if (!forceReconnect && sameDevice && camera.IsConnected)
            {
                StatusLabel.Text = $"Live preview — {selected!.Name}";
                return;
            }

            await camera.DisconnectAsync();
            token.ThrowIfCancellationRequested();
            if (selected == null)
            {
                StatusLabel.Text = "No USB camera found. Connect a camera to begin.";
                return;
            }

            StatusLabel.Text = $"Connecting to {selected.Name}…";
            await camera.ConnectAsync(selected, token);
            token.ThrowIfCancellationRequested();
            StatusLabel.Text = $"Live preview — {selected.Name}";
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            // A newer connection or a lifecycle event owns the next status update.
        }
        catch (Exception exception)
        {
            Debug.WriteLine(exception);
            if (version == sessionVersion && active)
                StatusLabel.Text = $"{exception.Message} Tap Reconnect to try again.";
            if (entered)
                await camera.DisconnectAsync();
        }
        finally
        {
            if (entered)
                operations.Release();
            previousCancellation?.Dispose();
            if (version == sessionVersion)
            {
                busy = false;
                UpdateButtons();
            }
        }
    }

    private async Task DisconnectAsync(string status)
    {
        var version = ++sessionVersion;
        var cancellation = sessionCancellation;
        sessionCancellation = null;
        cancellation?.Cancel();
        busy = true;
        UpdateButtons();
        await operations.WaitAsync();
        try
        {
            // A resumed view may already have requested a fresh session.
            if (version == sessionVersion)
            {
                await camera.DisconnectAsync();
                StatusLabel.Text = status;
            }
        }
        finally
        {
            operations.Release();
            cancellation?.Dispose();
            if (version == sessionVersion)
            {
                busy = false;
                UpdateButtons();
            }
        }
    }

    private void UpdateCameraPicker(IReadOnlyList<OtoscopyCameraDevice> cameras, OtoscopyCameraDevice? selected)
    {
        updatingPicker = true;
        try
        {
            CameraPicker.ItemsSource = cameras.ToList();
            CameraPicker.SelectedItem = selected;
            CameraPicker.IsVisible = cameras.Count > 1;
        }
        finally
        {
            updatingPicker = false;
        }
    }

    private void CaptureLeftClicked(object? sender, EventArgs e) => RunSafely(() => CaptureAsync(OtoscopyEar.Left));
    private void CaptureRightClicked(object? sender, EventArgs e) => RunSafely(() => CaptureAsync(OtoscopyEar.Right));

    private async Task CaptureAsync(OtoscopyEar ear)
    {
        if (!active || busy || !camera.IsConnected || sessionCancellation == null)
            return;

        var token = sessionCancellation.Token;
        var version = sessionVersion;
        var examination = examinationVersion;
        OtoscopyImage? captured = null;
        busy = true;
        UpdateButtons();
        bool entered = false;
        try
        {
            await operations.WaitAsync(token);
            entered = true;
            var jpeg = await camera.CaptureJpegAsync(token);
            token.ThrowIfCancellationRequested();
            if (!active || !camera.IsConnected || version != sessionVersion || examination != examinationVersion)
                return;

            captured = new OtoscopyImage(ear, jpeg, DateTimeOffset.Now);
            var source = ImageSource.FromStream(captured.OpenRead);
            if (ear == OtoscopyEar.Left)
            {
                LeftImage = captured;
                LeftPreview.Source = source;
                LeftPlaceholder.IsVisible = false;
                CaptureLeftButton.Text = "Retake left";
                OnPropertyChanged(nameof(LeftImage));
            }
            else
            {
                RightImage = captured;
                RightPreview.Source = source;
                RightPlaceholder.IsVisible = false;
                CaptureRightButton.Text = "Retake right";
                OnPropertyChanged(nameof(RightImage));
            }
            if (examination == examinationVersion)
                StatusLabel.Text = $"{ear} ear image captured. Preview remains live.";
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
        }
        catch (Exception exception)
        {
            Debug.WriteLine(exception);
            if (version == sessionVersion && active)
                StatusLabel.Text = $"Image was not captured: {exception.Message}";
        }
        finally
        {
            if (entered)
                operations.Release();
            if (version == sessionVersion)
            {
                busy = false;
                UpdateButtons();
            }
        }

        // Notify outside the camera lock so host code can stop the view or save the image.
        // PropertyChanged handlers above can synchronously clear or replace an exam.
        if (captured != null && examination == examinationVersion &&
            ReferenceEquals(ear == OtoscopyEar.Left ? LeftImage : RightImage, captured))
            ImageCaptured?.Invoke(this, new OtoscopyImageCapturedEventArgs(captured));
    }

    private void ClearClicked(object? sender, EventArgs e) => RunSafely(async () =>
    {
        if (ownerPage != null && await ownerPage.DisplayAlertAsync(
                "Clear otoscopy images?", "Both images will be removed from this view.", "Clear", "Cancel"))
            ClearImages();
    });

    private void UpdateButtons()
    {
        bool canCapture = active && camera.IsConnected && !busy;
        CaptureLeftButton.IsEnabled = canCapture;
        CaptureRightButton.IsEnabled = canCapture;
        CameraPicker.IsEnabled = active && !busy;
        ClearButton.IsEnabled = !busy && (LeftImage != null || RightImage != null);
        BusyIndicator.IsVisible = busy;
        BusyIndicator.IsRunning = busy;
#if ANDROID
        ReconnectButton.IsEnabled = loaded && pageVisible && windowActive && !busy;
#endif
    }

    private void ViewSizeChanged(object? sender, EventArgs e)
    {
        if (Width > 0)
            PreviewFrame.HeightRequest = Math.Clamp((Width - 32) * 0.625, 200, 420);
    }

    // Event handlers intentionally funnel through one observed async boundary. Native
    // callbacks may run on a streaming thread; no UI or camera operation happens there.
    private void RunSafely(Func<Task> action)
    {
        async void Run()
        {
            try
            {
                await action();
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
                StatusLabel.Text = $"Camera unavailable: {exception.Message}";
                UpdateButtons();
            }
        }

        if (Dispatcher.IsDispatchRequired)
            Dispatcher.Dispatch(Run);
        else
            Run();
    }
}
