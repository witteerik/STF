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
| `VisioForge.DotNet.VideoCapture` | `2026.11.6` |
| `VisioForge.DotNet.Core.UI.MAUI` | `2026.11.6` |
| `VisioForge.CrossPlatform.Core.Android` | `2026.11.5` |

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

The initial request is **1280 × 720 at 30 fps**, with the SDK choosing a supported
mode and preferring MJPEG. Snapshots encode the original camera frame as JPEG at
quality 95, rather than photographing the on-screen preview. There is no audio or
video recording. The requested resolution and JPEG quality are in
`Otoscopy/AndroidUsbOtoscopyCamera.cs` if you need to change them after trying your
camera. The USB camera and tablet still determine the modes actually available.

The SDK runs in its trial mode without a certificate. For a later licensed build,
use VisioForge's `SetLicenseCertificateAsync` on each newly created engine before
starting it; keep licence material out of source control. The module initializes
the SDK lazily and disposes each camera engine, while keeping the shared SDK
runtime available for the process lifetime.

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
