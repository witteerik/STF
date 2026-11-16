# USB otoscopy view

`STFM.Extension.Views.OtoscopyView` is a reusable MAUI `ContentView` for an Android or Windows
USB/UVC webcam or otoscope. Android uses **VisioForge Video Capture SDK .NET**;
Windows uses native `MediaCapture` without an additional camera SDK. Both retain
one JPEG result for the left ear and one for the right ear.

## Try it

`StfResearchSuite` already registers `UseOtoscopy()` and the Android camera
permission. Open otoscopy from `WelcomePageR`; closing the view restores the
welcome content. Build the Android target normally in Visual Studio and run it
on a physical tablet. There is no Android Studio, Gradle or manual native-library
build step.

1. Connect the USB camera using the tablet's USB host/OTG connection.
2. Open otoscopy and allow Android camera access and USB-device access when asked.
3. Wait for the live preview. Capture the left and right ears with their separate
   buttons; a subsequent capture replaces only that ear's image.
4. Optionally enter an examination reference, then use **Save / export** to save
   the current images together. A single captured ear is sufficient.
5. Use **Reconnect** to resume the camera after exporting. Use **Clear both
   images** before starting another examination, or **Close** to return to the
   welcome screen.

On Windows, attach a USB webcam or otoscope supported by the installed Windows
driver. Windows camera access must be enabled for the app. Only colour video
sources with USB VID/PID identifiers are listed; sources explicitly identified
as mounted in the computer enclosure are excluded. Drivers without a compatible
MediaFrameReader colour stream report an error rather than using a built-in camera.

The view selects the first USB camera automatically. If several are attached, a
camera selector appears. It never switches to the tablet's built-in camera.
Plugging a camera in while the view is active also starts connection automatically.

Resolution switching has been removed. The source requests **1280 × 720 at
30 fps** with `AndroidUVCFrameFormat.Unknown`, leaving the SDK to select a
supported mode. This is a preference, not a promise of that output size or rate.
The app does not call the native `GetModes` probe or cycle through alternative
modes. This avoids the extra native camera opens and closes used by the previous
mode picker. On Windows, a supported format closest to the same size/rate
preference is selected once before streaming starts. Its preview is limited to
10 frames/second to reduce encoding and UI overhead; snapshots retain the full
selected camera resolution.

If a camera connects but delivers no frames, unplug/reconnect it and tap
**Reconnect**. **Copy camera details** copies USB identifiers, tablet and SDK
versions, the automatic request and the last SDK error. It excludes captured
images and the examination reference. It does not probe or list advertised
modes. The details are also logged with Android logcat tag `Otoscopy`.

The requested mode shown in diagnostics is not a measurement of frames received.
VisioForge's logcat line `USB camera ready: WxH@fps FORMAT` reports its negotiated
mode. An otoscope working as a Windows webcam does not establish compatibility
with Android's USB transport or the SDK's native bridge. Removing mode switching
is not a confirmed fix for tablet power-offs or every older otoscope; physical
device testing is still required.

## Save and export

**Save / export** creates one ZIP archive and opens the system destination picker
through CommunityToolkit.Maui `FileSaver`. Choose a destination offered by the
device. The filename contains a UTC timestamp and a unique export ID, for example
`otoscopy-20261113-141011-<export-id>.zip`. The optional reference is stored inside
the archive, never used as a filename or path.

The archive contains:

- `left.jpg`, if a left image was captured.
- `right.jpg`, if a right image was captured.
- `manifest.json`, with format version 1, export ID, UTC export timestamp,
  optional reference, and each ear's capture flag, filename and original capture
  timestamp with its UTC offset. Missing ears have `captured: false` and null
  filename/timestamp fields.

The JPEG bytes are copied unchanged. Export does not resize, re-encode or annotate
the images, and it does not include camera diagnostics. It does not assign a
patient record or save an examination automatically.

The camera stops before the destination picker opens and stays paused afterward,
including if the picker is cancelled. Tap **Reconnect** when ready to capture
again. Saving retains both captured images and the reference on screen. Cancelling
or failing to save also retains them so that the export can be retried.

Destination support depends on Android, the installed document provider and the
Toolkit. A provider/Toolkit error can occur even after a file has been written;
check the chosen destination before retrying. A reported failure never marks the
examination as exported or discards its captures.

**Clear both images** asks before clearing the current results. Clearing or
retaking an image does not delete any archive already saved. **Close** offers
**Save / export**, **Discard** and **Cancel** when there are unexported results.
A cancelled or failed export keeps the view open. A successful export permits
closing after camera cleanup; **Discard** closes without saving the current
results after camera cleanup. If cleanup fails or is still running, the view stays
attached so its preview is not removed while native code may still use it. Images
remain available for export. Retry Close after cleanup finishes, or save your
images and restart the app if shutdown failed. Cancelling Close leaves the camera
paused until **Reconnect**.

## Host integration and results

To host the view in another page:

```xml
<ContentPage xmlns="http://schemas.microsoft.com/dotnet/2021/maui"
             xmlns:x="http://schemas.microsoft.com/winfx/2009/xaml"
             xmlns:views="clr-namespace:STFM.Extension.Views;assembly=STFM.Extension">
    <views:OtoscopyView x:Name="Otoscopy" />
</ContentPage>
```

Or create it in code and handle `Closed` to restore the host's previous content:

```csharp
var previousContent = Content;
var otoscopy = new STFM.Extension.Views.OtoscopyView();
otoscopy.Closed += (_, _) => Content = previousContent;
Content = otoscopy;
```

The full screening suite instead creates `new OtoscopyView(embedded: true)`.
Embedded mode hides the standalone export, reference, clear and close controls;
the suite owns examination storage and navigation. After copying or saving the
captured images, the host must `await otoscopy.ReleaseForHostAsync()` **before**
removing the view. It throws if shutdown is incomplete or failed, leaving the
preview attached. On success, its camera and lifecycle handlers are released and
`LeftImage`/`RightImage` remain available. Do not call standalone `CloseAsync()` to
advance an embedded examination.

`WelcomePageR` uses this pattern. A custom host can call
`await otoscopy.CloseAsync()` to run the same save/discard/cancel workflow as the
Close button. It returns `true` when closed and `false` if closing was cancelled
or could not complete. `Closed` tells the host to remove or replace the view; the
view does not choose a navigation route itself. Route custom back/navigation
actions through `CloseAsync()` when they should protect unexported results.

`LeftImage` and `RightImage` contain the latest `OtoscopyImage`, or `null` if that
side has not been captured. Each image provides:

- `Ear`: `OtoscopyEar.Left` or `OtoscopyEar.Right`.
- `CapturedAt`: local capture date/time including the UTC offset.
- `JpegData`: read-only JPEG bytes.
- `OpenRead()`: a new read-only stream owned by the caller.

`ImageCaptured` fires on the UI thread after a successful capture. Its `e.Image`
is the captured result:

```csharp
otoscopy.ImageCaptured += (_, e) =>
{
    System.Diagnostics.Debug.WriteLine(
        $"{e.Image.Ear}: {e.Image.JpegData.Length} bytes at {e.Image.CapturedAt}");
};
```

Captures are held in memory until cleared or the view is discarded. They do not
survive process termination. `ClearImages()` clears the in-memory results without
a dialog; call it on the UI thread when the host begins a new examination.

The camera stops when the view unloads, its page disappears, or the app goes into
the background. Normal return reconnects it while preserving captures. Explicit
pauses, including export or `await otoscopy.StopCameraAsync()`, require
**Reconnect**. `StopCameraAsync()` remains available for custom hosts; it stops
capture without closing the view or saving its images.

Custom storage workflows can also use the pure .NET archive helper:

```csharp
using var archive = new MemoryStream();
var export = await OtoscopyExport.CreateArchiveAsync(
    archive, otoscopy.LeftImage, otoscopy.RightImage, reference: "Exam 42");
archive.Position = 0;
// Copy archive to your chosen destination using export.SuggestedFileName.
```

The helper requires an empty writable destination and at least one image,
validates each image's ear, and leaves the destination open. Discard incomplete
output if the operation is cancelled or fails.

## Dependencies and configuration

The code matches this repository's .NET 10 MAUI targets and Android 12 minimum.
Windows camera support uses the operating system media APIs and needs the
`webcam` device capability in a packaged host app manifest; it does not load
VisioForge. The supplied StfResearchSuite manifest includes that capability.

Pinned Android dependencies:

| Package | Version |
| --- | --- |
| `VisioForge.DotNet.VideoCapture` | `2026.11.12` |
| `VisioForge.DotNet.Core.UI.MAUI` | `2026.11.12` |
| `VisioForge.CrossPlatform.Core.Android` | `2026.11.5` |

The MAUI Controls and Compatibility references are aligned to **10.0.110** in
STFM.Core, STFM.Extension and StfResearchSuite. STFM.Core already supplies
CommunityToolkit.Maui **14.1.1**, and `StfResearchSuite` initializes it with
`UseMauiCommunityToolkit()`; export uses its `FileSaver` API.

The official Android companion project is included in
`ThirdParty/VisioForge.Android`; its source revision and notices are documented
there. NuGet supplies the compiled native runtime. In another host app, register
the Toolkit and call `.UseOtoscopy()` on the `MauiAppBuilder`, and add these
manifest entries:

```xml
<uses-permission android:name="android.permission.CAMERA" />
<uses-feature android:name="android.hardware.camera" android:required="false" />
<uses-feature android:name="android.hardware.camera.autofocus" android:required="false" />
<uses-feature android:name="android.hardware.usb.host" android:required="false" />
```

AndroidX Lifecycle packages are aligned to 2.10.0.2 and SavedState.Ktx to 1.4.0.2 in
the project file. Snapshots encode the original camera frame as JPEG at quality
95, rather than photographing the on-screen preview. There is no audio or video
recording. JPEG quality is set in `Otoscopy/AndroidUsbOtoscopyCamera.cs`.

The SDK runs in trial mode without a certificate. For a licensed build, use
VisioForge's `SetLicenseCertificateAsync` on each newly created engine before
starting it; keep licence material out of source control. The module initializes
the SDK lazily and disposes each camera engine, while keeping the shared SDK
runtime available for the process lifetime.

Every disconnect awaits `StopAsync()` before `DisposeAsync()`. The view serializes
camera operations, and a process-wide shutdown barrier also prevents a newly
created view from starting a camera while a previous view's native shutdown is
pending or has failed. Repeated cleanup joins the same shutdown task. A 15-second
wait timeout reports the problem; it does not cancel native shutdown or permit a
new stream while shutdown is still pending. A retry can proceed only after that
shutdown succeeds. A failed shutdown blocks new connections until the app is
restarted. No forced USB reset or arbitrary settling delay is used.

The shutdown wait bound does not cover a native start or snapshot that never
returns. Camera operations remain serialized rather than disposing an engine
underneath an unfinished native call, so Close cannot guarantee recovery from
every native hang.

## Verification

Run the archive checks without MAUI workloads or a USB camera:

```sh
dotnet run --project Tests/OtoscopyExport/OtoscopyExport.Tests.csproj --configuration Release
```

These checks use the production export/image sources and verify unchanged JPEG
bytes, side labels, capture timestamps and offsets, one-ear exports, safe unique
filenames, destination ownership, invalid inputs and cancellation. JSON reflection
is disabled in the test project to exercise the source-generated serializer used
by MAUI. The checks require no test-framework NuGet packages.

On a physical tablet, verify preview and both captures, retaking one side,
unplug/replug, and leaving/returning to the app. Test saving both ears and a single
ear, cancelling the destination picker, export failure, explicit reconnect after
export, and all three Close choices. Open saved ZIPs and inspect their JPEGs and
manifest. Also check denied permissions and ensure another camera app is closed
before this one claims the USB device. Compilation and archive checks do not
verify USB-driver behavior or resolve the reported tablet power-offs.

Official references:

- [Android USB capture guide](https://www.visioforge.com/help/docs/dotnet/general/guides/android-usb-camera/)
- [MAUI integration](https://www.visioforge.com/help/docs/dotnet/install/maui/)
- [Video Capture SDK .NET](https://www.visioforge.com/video-capture-sdk-net)
- [CommunityToolkit.Maui FileSaver](https://learn.microsoft.com/dotnet/communitytoolkit/maui/essentials/file-saver)
