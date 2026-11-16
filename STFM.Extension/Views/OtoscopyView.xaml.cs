// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Diagnostics;
using CommunityToolkit.Maui.Storage;
using STFM.Extension.Otoscopy;

namespace STFM.Extension.Views;

/// <summary>
/// Previews an Android USB camera and retains one JPEG per ear. Save / export writes
/// an examination ZIP to a user-selected destination; captures are not saved automatically.
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
    private bool resultActionInProgress;
    private bool closed;
    private string? selectedCameraId;
    private int sessionVersion;
    private int examinationVersion;
    private int resultVersion;
    private int exportedVersion = -1;

    public OtoscopyView() : this(embedded: false) { }

    /// <summary>
    /// Embedded mode delegates examination storage and navigation to the screening host.
    /// The host must await ReleaseForHostAsync before removing this view.
    /// </summary>
    public OtoscopyView(bool embedded)
    {
        IsEmbedded = embedded;
        InitializeComponent();
        camera = OtoscopyCameraFactory.Create();
        PreviewHost.Add(camera.Preview);
        CloseButton.IsVisible = !embedded;
        SaveExportButton.IsVisible = !embedded;
        ReferenceEntry.IsVisible = !embedded;
        ClearButton.IsVisible = !embedded;
        ResultStatusLabel.IsVisible = !embedded;
        Loaded += ViewLoaded;
        Unloaded += ViewUnloaded;
        SizeChanged += ViewSizeChanged;
        camera.Preview.Loaded += PreviewReady;
        camera.Preview.HandlerChanged += PreviewReady;
#if !ANDROID && !WINDOWS
        StatusLabel.Text = "USB otoscopy is available in the Android and Windows apps.";
        ReconnectButton.IsEnabled = false;
#endif
    }

    public bool IsEmbedded { get; }

    public OtoscopyImage? LeftImage { get; private set; }
    public OtoscopyImage? RightImage { get; private set; }

    /// <summary>Raised on the UI thread after a successful capture replaces the selected ear's image.</summary>
    public event EventHandler<OtoscopyImageCapturedEventArgs>? ImageCaptured;

    /// <summary>Raised on the UI thread after CloseAsync stops the camera and permits closing.</summary>
    public event EventHandler? Closed;

    private bool HasImages => LeftImage is not null || RightImage is not null;
    private bool HasUnexportedImages => HasImages && resultVersion != exportedVersion;

    /// <summary>
    /// Restores a host's persisted captures before this view is attached or loaded.
    /// Original capture timestamps and ear identities are preserved. Call on the UI thread.
    /// </summary>
    public void RestoreImages(OtoscopyImage? left, OtoscopyImage? right)
    {
        if (loaded || active || busy || closed || resultActionInProgress)
            throw new InvalidOperationException("Restore images before opening the otoscopy view.");
        if (left is not null && left.Ear != OtoscopyEar.Left)
            throw new ArgumentException("The left image must belong to the left ear.", nameof(left));
        if (right is not null && right.Ear != OtoscopyEar.Right)
            throw new ArgumentException("The right image must belong to the right ear.", nameof(right));

        ClearImages();
        LeftImage = left;
        RightImage = right;
        LeftPreview.Source = left is null ? null : ImageSource.FromStream(left.OpenRead);
        RightPreview.Source = right is null ? null : ImageSource.FromStream(right.OpenRead);
        LeftPlaceholder.IsVisible = left is null;
        RightPlaceholder.IsVisible = right is null;
        CaptureLeftButton.Text = left is null ? "Capture left" : "Retake left";
        CaptureRightButton.Text = right is null ? "Capture right" : "Retake right";
        OnPropertyChanged(nameof(LeftImage));
        OnPropertyChanged(nameof(RightImage));
        ResultStatusLabel.Text = HasImages ? "Previous images restored." : "No images captured yet.";
        UpdateButtons();
    }

    /// <summary>Clears both results and invalidates any capture already in progress. Call on the UI thread.</summary>
    public void ClearImages()
    {
        examinationVersion++;
        resultVersion++;
        ReferenceEntry.Text = string.Empty;
        ResultStatusLabel.Text = "No images saved for this examination.";
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
        bool wasActive = active;
        manuallyStopped = true;
        await UpdateActivityAsync();
        if (!wasActive)
            await DisconnectAsync("Camera paused. Tap Reconnect to resume.");
    }

    /// <summary>
    /// Ends an embedded step without a save/discard dialog or clearing its images.
    /// Call on the UI thread. Remove this view only after successful completion;
    /// a shutdown failure leaves its preview attached and throws to the host.
    /// </summary>
    public async Task ReleaseForHostAsync()
    {
        if (!IsEmbedded)
            throw new InvalidOperationException("Use CloseAsync for a standalone otoscopy view.");
        if (closed)
            return;
        if (resultActionInProgress)
            throw new InvalidOperationException("An otoscopy operation is already in progress.");
        resultActionInProgress = true;
        UpdateButtons();
        try
        {
            await StopCameraAsync();
            await camera.DisposeAsync();
            FinishClosing();
        }
        finally
        {
            resultActionInProgress = false;
            UpdateButtons();
        }
    }

    private void ViewLoaded(object? sender, EventArgs e) => RunSafely(async () =>
    {
        if (closed)
            return;
        loaded = true;
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
#if ANDROID || WINDOWS
        bool shouldBeActive = loaded && pageVisible && windowActive && !manuallyStopped && !closed;
        if (shouldBeActive == active)
        {
            if (active)
                await TryAutoConnectAsync();
            UpdateButtons();
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
                await DisconnectAsync("Camera paused. Tap Reconnect to resume.");
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
                await DisconnectAsync(CameraFailureMessage(e.Message));
        });
    }

    private static string LivePreviewMessage(string name)
        => $"Live preview — {name}. Video format selected automatically.";

    private static string CameraFailureMessage(string message)
    {
        if (message.Contains("delivered no frames", StringComparison.OrdinalIgnoreCase))
            return "No video frames received in automatic mode. Unplug/reconnect the otoscope, then tap Reconnect. " +
                "Copy camera details below for troubleshooting.";
        return $"Camera stopped: {message} Tap Reconnect to try again.";
    }

    private void ReconnectClicked(object? sender, EventArgs e) => RunSafely(async () =>
    {
        if (closed || resultActionInProgress)
            return;
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
        if (closed || resultActionInProgress || updatingPicker || CameraPicker.SelectedItem is not OtoscopyCameraDevice device)
            return;
        selectedCameraId = device.Id;
        RunSafely(() => active ? ConnectAsync(forceReconnect: true) : Task.CompletedTask);
    }

    private async Task ConnectAsync(bool forceReconnect = false)
    {
        if (closed || resultActionInProgress || !active)
            return;
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

            var cameras = await camera.GetCamerasAsync(token);
            token.ThrowIfCancellationRequested();
            var selected = cameras.FirstOrDefault(item => item.Id == selectedCameraId) ?? cameras.FirstOrDefault();
            bool sameDevice = selected?.Id == selectedCameraId;
            selectedCameraId = selected?.Id;
            UpdateCameraPicker(cameras, selected);

            if (!forceReconnect && sameDevice && camera.IsConnected)
            {
                StatusLabel.Text = LivePreviewMessage(selected!.Name);
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
            CopyDetailsButton.IsVisible = true;
            await camera.ConnectAsync(selected, token);
            token.ThrowIfCancellationRequested();
            StatusLabel.Text = LivePreviewMessage(selected.Name);
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            // A newer connection or a lifecycle event owns the next status update.
        }
        catch (Exception exception)
        {
            Debug.WriteLine(exception);
            if (version == sessionVersion && active)
            {
                StatusLabel.Text = CameraFailureMessage(exception.Message);
            }
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

    private void CopyDetailsClicked(object? sender, EventArgs e) => RunSafely(async () =>
    {
        // Device identity and SDK errors only; never include images or the reference.
        await Clipboard.Default.SetTextAsync(camera.Diagnostics);
        if (ownerPage is not null)
            await ownerPage.DisplayAlertAsync("Camera details copied", "The USB camera details and last SDK error are on the clipboard.", "OK");
    });

    private void CaptureLeftClicked(object? sender, EventArgs e) => RunSafely(() => CaptureAsync(OtoscopyEar.Left));
    private void CaptureRightClicked(object? sender, EventArgs e) => RunSafely(() => CaptureAsync(OtoscopyEar.Right));

    private async Task CaptureAsync(OtoscopyEar ear)
    {
        if (closed || resultActionInProgress || !active || busy || !camera.IsConnected || sessionCancellation == null)
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
            {
                resultVersion++;
                ResultStatusLabel.Text = "Images have unsaved changes.";
                StatusLabel.Text = $"{ear} ear image captured. Preview remains live.";
            }
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
        if (closed || resultActionInProgress || busy)
            return;
        resultActionInProgress = true;
        UpdateButtons();
        try
        {
            if (ownerPage != null && await ownerPage.DisplayAlertAsync(
                    "Clear otoscopy images?", "Both images and the reference will be removed from this view. Saved files are kept.",
                    "Clear", "Cancel"))
                ClearImages();
        }
        finally
        {
            resultActionInProgress = false;
            UpdateButtons();
        }
    });

    private void ReferenceChanged(object? sender, TextChangedEventArgs e)
    {
        resultVersion++;
        if (HasImages)
            ResultStatusLabel.Text = "Images have unsaved changes.";
    }

    private void SaveExportClicked(object? sender, EventArgs e) => RunSafely(async () =>
    {
        if (closed || resultActionInProgress || busy || !HasImages)
            return;
        resultActionInProgress = true;
        UpdateButtons();
        try
        {
            await PauseForResultsAsync();
            await ExportImagesAsync();
        }
        finally
        {
            resultActionInProgress = false;
            UpdateButtons();
        }
    });

    // Export must remain available even after a native camera shutdown fails.
    // It only reads already captured JPEGs; it does not touch the USB camera.
    private async Task PauseForResultsAsync()
    {
        try
        {
            await StopCameraAsync();
        }
        catch (Exception exception)
        {
            Debug.WriteLine(exception);
            StatusLabel.Text = $"Camera could not stop cleanly: {exception.Message} Captured images can still be saved.";
        }
    }

    private async Task<bool> ExportImagesAsync()
    {
        if (!HasImages)
            return false;

        var left = LeftImage;
        var right = RightImage;
        var reference = ReferenceEntry.Text;
        var version = resultVersion;
        ResultStatusLabel.Text = "Preparing images for export…";
        try
        {
            // The camera lock is never held while the Android file picker is open.
            // Lifecycle notifications can therefore finish without deadlocking it.
            using var archive = new MemoryStream();
            var export = await Task.Run(() => OtoscopyExport.CreateArchiveAsync(archive, left, right, reference));
            archive.Position = 0;
            var result = await FileSaver.Default.SaveAsync(export.SuggestedFileName, archive, CancellationToken.None);
            if (result.IsSuccessful)
            {
                if (version == resultVersion)
                    exportedVersion = version;
                ResultStatusLabel.Text = version == resultVersion
                    ? $"Images saved: {export.SuggestedFileName}"
                    : "The selected images were saved. Newer changes still need saving.";
                return version == resultVersion;
            }
            if (result.IsCancelled)
            {
                ResultStatusLabel.Text = "Save cancelled. Your images are still here.";
                return false;
            }
            throw result.Exception ?? new IOException("The destination did not confirm the save.");
        }
        catch (Exception exception)
        {
            Debug.WriteLine(exception);
            // Some document providers can fail after writing. Never infer success
            // from a path, delete the captures, or mark them saved on an error.
            ResultStatusLabel.Text = $"Save could not be confirmed: {exception.Message} " +
                "Your images are still here. Check the destination before retrying.";
            return false;
        }
    }

    private void CloseClicked(object? sender, EventArgs e) => RunSafely(async () => { await CloseAsync(); });

    /// <summary>
    /// Stops capture, offers to export unsaved images, then raises Closed. Returns false
    /// if cancelled or saving fails. Call on the UI thread; the host removes this view
    /// only after Closed. A cancelled close leaves the camera paused until Reconnect.
    /// </summary>
    public async Task<bool> CloseAsync()
    {
        if (closed)
            return true;
        if (resultActionInProgress)
            return false;
        resultActionInProgress = true;
        UpdateButtons();
        try
        {
            // Cancel/join any capture before asking about its results. Otherwise a
            // late frame could arrive after the user decides what to save/discard.
            await PauseForResultsAsync();
            if (HasUnexportedImages)
            {
                if (ownerPage is null)
                {
                    ResultStatusLabel.Text = "Images have not been saved. Open the view to save or discard them before closing.";
                    return false;
                }
                var choice = await ownerPage.DisplayActionSheetAsync(
                    "Save images before closing?", "Cancel", "Discard images", "Save / export images");
                if (choice == "Save / export images")
                {
                    if (!await ExportImagesAsync())
                        return false;
                }
                else if (choice != "Discard images")
                    return false;
            }

            try
            {
                await camera.DisposeAsync();
            }
            catch (Exception exception)
            {
                Debug.WriteLine(exception);
                StatusLabel.Text = $"The camera has not shut down cleanly: {exception.Message}";
                if (ownerPage is not null)
                    await ownerPage.DisplayAlertAsync("Camera is still unavailable",
                        "The view will stay open because the camera has not confirmed shutdown. " +
                        "Your captured images are still here and can be exported. " +
                        "Try Close again once shutdown finishes, or save your images and restart the app.", "OK");
                // A barrier against reopening is not proof that the renderer has
                // released this surface. Keep the preview attached on failed or
                // pending cleanup rather than disposing it underneath native code.
                return false;
            }
            FinishClosing();
            IsVisible = false;
            Closed?.Invoke(this, EventArgs.Empty);
            return true;
        }
        finally
        {
            resultActionInProgress = false;
            UpdateButtons();
        }
    }

    private void FinishClosing()
    {
        closed = true;
        loaded = false;
        DetachLifecycle();
        Loaded -= ViewLoaded;
        Unloaded -= ViewUnloaded;
        SizeChanged -= ViewSizeChanged;
        camera.Preview.Loaded -= PreviewReady;
        camera.Preview.HandlerChanged -= PreviewReady;
    }

    private void UpdateButtons()
    {
        bool available = !closed && !resultActionInProgress && !busy;
        bool canCapture = available && active && camera.IsConnected;
        CaptureLeftButton.IsEnabled = canCapture;
        CaptureRightButton.IsEnabled = canCapture;
        CameraPicker.IsEnabled = available && active;
        CopyDetailsButton.IsEnabled = available;
        ClearButton.IsEnabled = available && HasImages;
        SaveExportButton.IsEnabled = available && HasImages;
        ReferenceEntry.IsEnabled = available;
        // Allow Close to cancel a connection or capture already in progress.
        CloseButton.IsEnabled = !closed && !resultActionInProgress;
        BusyIndicator.IsVisible = busy || resultActionInProgress;
        BusyIndicator.IsRunning = busy || resultActionInProgress;
#if ANDROID || WINDOWS
        ReconnectButton.IsEnabled = available && loaded && pageVisible && windowActive;
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
