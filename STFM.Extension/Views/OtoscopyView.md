# USB otoscopy view

`STFM.Extension.Views.OtoscopyView` is a reusable MAUI `ContentView` for an Android
USB/UVC webcam or otoscope. It uses **VisioForge Video Capture SDK .NET**, with one
JPEG result for the left ear and one for the right ear.

## Try it

The supplied changes already register `UseOtoscopy()` in `StfResearchSuite` and add
the Android camera permission. No VisioForge licence key is needed for its trial.
Build the Android target normally in Visual Studio and run it on your physical
tablet. There is no Android Studio, Gradle or manual native-library build step.

Place the view in whichever page should host otoscopy:

```xml
<ContentPage xmlns="http://schemas.microsoft.com/dotnet/2021/maui"
             xmlns:x="http://schemas.microsoft.com/winfx/2009/xaml"
             xmlns:views="clr-namespace:STFM.Extension.Views;assembly=STFM.Extension">
    <views:OtoscopyView x:Name="Otoscopy" />
</ContentPage>
```

Or, in an existing page's code-behind:

```csharp
var otoscopy = new STFM.Extension.Views.OtoscopyView();
Content = otoscopy;
```

For a quick test in `StfResearchSuite`, put the two lines above at the **end** of
the `MainPage` constructor. This temporarily replaces its welcome screen. The
module does not change the existing speech-test navigation automatically.

1. Connect the USB camera using the tablet's USB host/OTG connection.
2. Open the view and allow Android camera access and USB-device access when asked.
3. Wait for the live preview. Capture the left and right ears with their separate
   buttons; a subsequent capture replaces only that ear's image.
4. Use **Clear both images** before starting another examination. The button asks
   before clearing. **Reconnect** retries after a denied permission or camera error.

The view selects the first USB camera automatically. If several are attached, a
selector appears. It never switches to the tablet's built-in camera. Plugging a
camera in while the view is open also starts connection automatically.

**Video mode** lists the resolutions, frame rates and MJPEG/YUY2 formats reported
by the connected camera. **Automatic** keeps the original selection preference:
nearest pixel count to 1280 × 720, then MJPEG, then the advertised rate nearest
30 fps. A manual choice stops and disposes the current capture before opening a
new one with that mode's exact dimensions, format and frame rate. The choice is
remembered by USB vendor/product ID for this view, including unplug/replug; it is
not saved as an app setting. A different camera model starts with Automatic,
while returning to a previously selected model restores its choice. Two cameras
with the same vendor/product IDs share the choice.

If an older camera connects but reports that it delivered no frames:

1. Choose one of its advertised modes explicitly. For a scope reported as
   720 × 720 on a PC, use that size only if it appears in the Android list.
2. Try a lower **listed** frame rate at the same size, and the alternative format
   (MJPEG or YUY2) if offered. The module never invents unsupported frame rates.
3. If the device remains stalled, unplug/reconnect it and tap **Reconnect**.
4. Use **Copy camera details** to obtain the USB vendor/product IDs, tablet and SDK
   versions, requested mode, advertised modes and complete last SDK error. This
   does not copy images or patient data. The same details are logged with the
   Android logcat tag `Otoscopy`.

The status shows the mode requested from the SDK, not a measurement of frames
received. VisioForge may still negotiate another mode; its logcat line
`USB camera ready: WxH@fps FORMAT` reports the negotiated mode. An empty mode list
can mean the camera could not be opened (for example another app owns it), not
only that it lacks supported formats. In that case Automatic retains the old
1280 × 720 / 30 fps request and lets the SDK negotiate.

This adds mode control and diagnostics, not a confirmed fix for every older
otoscope. Working as a Windows webcam does not prove compatibility with Android's
USB transport or the SDK's native bridge. No camera-specific identity or mode is
assumed, and no SDK package upgrade is required for these changes.

## Get the results

`LeftImage` and `RightImage` contain the latest `OtoscopyImage`, or `null` if that
side has not been captured. Each image provides:

- `Ear`: `OtoscopyEar.Left` or `OtoscopyEar.Right`.
- `CapturedAt`: local capture date/time including the UTC offset.
- `JpegData`: read-only JPEG bytes.
- `OpenRead()`: a new read-only stream owned by the caller.

`ImageCaptured` fires on the UI thread after a successful capture. Its
`e.Image` is the captured result. For example:

```csharp
otoscopy.ImageCaptured += (_, e) =>
{
    System.Diagnostics.Debug.WriteLine(
        $"{e.Image.Ear}: {e.Image.JpegData.Length} bytes at {e.Image.CapturedAt}");
};
```

The view retains results **in memory**. It does not choose a patient record or save
files automatically. To save the current left image from your own Save command:

```csharp
var image = otoscopy.LeftImage;
if (image is not null)
{
    // Choose the destination according to the host application's storage workflow.
    string path = Path.Combine(FileSystem.AppDataDirectory, "otoscopy-left.jpg");
    using var source = image.OpenRead();
    await using var destination = File.Create(path);
    await source.CopyToAsync(destination);
}
```

This example overwrites `otoscopy-left.jpg`; a real examination-saving command
should choose its own destination. Use `RightImage` in the same way for the right
ear. Retaking or clearing an image in the view does not delete files your host has
already saved. `ClearImages()` clears both in-memory results without a dialog;
call it on the UI thread when your host begins a new examination.

The camera stops when the view unloads, its page disappears, or the app goes into
the background. It reconnects on return and keeps the captured images. Custom
hosts can explicitly `await otoscopy.StopCameraAsync()` before replacing the view;
the Reconnect button resumes a manually stopped view. Clear results before
reusing an instance for a different examination.

## Dependencies and configuration

The code matches this repository's .NET 10 MAUI targets and Android 12 minimum.
Windows builds display an Android-only placeholder without loading VisioForge.

Pinned Android dependencies:

| Package | Version |
| --- | --- |
| `VisioForge.DotNet.VideoCapture` | `2026.11.12` |
| `VisioForge.DotNet.Core.UI.MAUI` | `2026.11.12` |
| `VisioForge.CrossPlatform.Core.Android` | `2026.11.5` |

The MAUI Controls and Compatibility references are aligned to **10.0.110** in
STFM.Core, STFM.Extension and StfResearchSuite, as required by this SDK UI package.
The projects continue targeting .NET 10. The managed SDK update includes vendor
fixes for overlapping start/stop/dispose operations and an MJPEG camera memory
leak on repeated starts.

The small official Android companion project is included in
`ThirdParty/VisioForge.Android`; its source revision and notices are documented
there. NuGet supplies the compiled native runtime. In another host app, call
`.UseOtoscopy()` on the `MauiAppBuilder` and add these manifest entries:

```xml
<uses-permission android:name="android.permission.CAMERA" />
<uses-feature android:name="android.hardware.camera" android:required="false" />
<uses-feature android:name="android.hardware.camera.autofocus" android:required="false" />
<uses-feature android:name="android.hardware.usb.host" android:required="false" />
```

AndroidX Lifecycle packages are aligned to 2.10.0.2 and SavedState.Ktx to 1.4.0.2 in
the project file to match the preview control's dependencies.

The initial automatic preference is **1280 × 720 at 30 fps**, resolved against
the advertised modes as described above. Snapshots encode the original camera frame as JPEG at
quality 95, rather than photographing the on-screen preview. There is no audio or
video recording. JPEG quality is set in `Otoscopy/AndroidUsbOtoscopyCamera.cs`.
The USB camera and tablet determine the modes available in the picker.

The SDK runs in its trial mode without a certificate. For a later licensed build,
use VisioForge's `SetLicenseCertificateAsync` on each newly created engine before
starting it; keep licence material out of source control. The module initializes
the SDK lazily and disposes each camera engine, while keeping the shared SDK
runtime available for the process lifetime.

Every disconnect now awaits `StopAsync()` before `DisposeAsync()`. The view's
operation semaphore keeps this sequence ahead of the next camera open. Repeated
cleanup calls join the same shutdown task. If it faults, subsequent connections
on that camera service remain blocked; restart the app to recover. A 15-second
wait timeout reports the problem but does not cancel native shutdown or permit
a new stream while it is still running. A later retry can proceed only if that
shutdown actually completed successfully. No forced USB reset or arbitrary
settling delay is used.

## Device verification

The Android project can be compile-checked without a connected camera. Actual
streaming needs a physical USB host device. On your tablet, verify preview and
both captures, retaking one side, unplug/replug, and leaving/returning to the page
or app. Also check that denied permissions give a useful retry message and that
another camera app is closed before this one claims the USB device.

Official references:

- [Android USB capture guide](https://www.visioforge.com/help/docs/dotnet/general/guides/android-usb-camera/)
- [MAUI integration](https://www.visioforge.com/help/docs/dotnet/install/maui/)
- [Video Capture SDK .NET](https://www.visioforge.com/video-capture-sdk-net)
