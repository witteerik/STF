# Full hearing screening

`STFM.Extension.Views.FullScreeningView` hosts one examination containing staff
video-otoscopy followed by three patient-operated tests. Instructions and the
combined report are in Swedish, matching the existing Swedish SSQ12 and QuickSiP
materials in this repository.

## Start the module

On `WelcomePageR`, choose **Fullständig hörselscreening**. The existing screening
checkbox and Start route in `StfResearchSuite` also open the new module. The
standalone **Otoscopy** button remains available for camera-only use.

A staff preparation screen asks for a patient/examination reference and confirms
that equipment and the test environment have been checked. Audio preparation
uses the configured headphone transducer, output routing, calibration and speech
material. If a choice of configured headphones is needed, staff selects it. A
missing device, calibration entry or required QuickSiP sound file produces an
error that can be corrected and retried; the module does not substitute another
output device or the tablet speaker.

The steps are:

1. **Video-otoscopy — staff.** Read the staff instructions, connect the USB scope,
   and capture a correctly labelled image of each ear. Retaking replaces that
   side's current result. Staff can add a note. If a side cannot be photographed,
   enter a reason and explicitly confirm continuing without its image. Capture
   files are saved locally as they are taken. The camera must release its preview
   before the workflow moves on.
2. **SSQ12 — patient.** Read the instructions and answer the existing 12-question
   SSQ12 view, in its minimal version. Numeric answers, including zero, contribute
   to the mean; **Vet ej** is retained as missing/not rated and is excluded from
   the numeric mean. If every answer is unknown, the report shows no numeric mean.
3. **QuickSiP — patient.** The instructions explain headphone placement, the
   response alternatives and the four-second response window. The module runs
   the existing `QuickSiP_StaticSpeechTest` with the Swedish QuickSiP material and
   a randomly selected start list.
4. **Automated audiometry — patient.** The instructions explain pressing and
   holding the response button while a tone is audible. The module runs the
   existing `UoPta` **SAME96_Screening** protocol through `UoAudView`.

Each patient-operated step starts only after its instruction button is pressed.
A pause/help control stops the current test. Returning from the background or
resuming an interrupted examination restarts that unfinished step from its
instructions; it does not resume midway through a tone or speech trial. Previously
completed steps remain saved.

## Existing test behavior

This module changes orchestration, persistence and presentation, not the clinical
algorithms or interpretation cutoffs. It does not add pass/fail thresholds or a
diagnostic classification.

QuickSiP's existing engine handles an unanswered/timed-out trial using its own
random-guess scoring rule. The report retains that engine score and includes the
number of timed-out responses; it must not be interpreted as a new scoring
implementation that counts every timeout as incorrect.

`UoAudView` retains the existing **DD65v2 RETSPL table**. The availability of a
configured transducer and numeric calibration entries does not establish that a
different headphone model is appropriate for that table. Staff must use equipment
and calibration suitable for the existing audiometry implementation; this patch
does not introduce transducer-specific RETSPL management.

The audiogram distinguishes a response at the screening level from an exact
threshold, and retains no-response, uncertain/unmeasured and reliability-retest
information. It uses the repository's `PlotBase` logarithmic frequency drawing
and coordinate functions. It does not use the placeholder diagonal in the old
`AudiogramView`. The report's threshold table preserves the available detail,
including reliability retests, rather than silently treating every value as an
ordinary threshold.

## Local results and recovery

The active examination is stored under:

```text
FileSystem.AppDataDirectory/Screening/
    active-examination.txt
    <examination-guid>/
        examination.json
        left-<capture-guid>.jpg
        right-<capture-guid>.jpg
        Report/
            screening-<examination-guid>.pdf
            screening-<examination-guid>-page-<number>.png
```

These are app-private files. The patient reference is stored inside the journal
and report; it is never used as a directory name. Each retake writes a new image
file before changing the journal, so a failed journal write cannot alter the
previously committed photograph or its timestamp. Older capture files and
previous patients' directories are retained. This version has no archive browser
or automatic deletion/retention policy.

Journal writes and the active-examination pointer are replaced atomically. The
UI advances only after the changed examination has been saved. A failed save
keeps the current patient and makes retry available. Invalid or missing active
journal data produces an error instead of silently creating an empty patient.

Opening the module again resumes the same active patient. A completed report and
its print-confirmation state can also be resumed. **Spara och stäng** and the
host's hardware Back route save/pause and return to the previous view; they do
not reset the patient. An unfinished patient test restarts from instructions on
return. If native camera cleanup has not finished, the preview remains attached
and closing/advancing is refused rather than detaching a surface that native code
may still use. A native operation that never returns can still prevent cleanup.

App-private storage is not an external backup. Export the PDF when a separate
copy is needed. Deleting application data or uninstalling the application is not
a supported way to reset an examination.

## Report, export and printing

After the four steps, the module creates an A4 PDF and matching on-screen preview
pages. The report contains:

- The patient/examination reference and examination timestamps.
- Left/right otoscopy images, capture timestamps and any missing-image reasons
  or staff notes.
- SSQ12 mean and answered-item information, plus individual item ratings.
- The existing QuickSiP score, trial counts and timeout information.
- An audiogram and structured audiometry results with the configured transducer
  and protocol information.

A typical report occupies two A4 pages. Long staff notes continue onto additional
pages instead of being silently truncated. The same rendering code is used for
preview and PDF/print output. Printing owns a snapshot of the results and images,
so a queued print job cannot pick up a later patient's data.

**Exportera PDF** opens the system save destination picker. Export does not count
as printing and does not unlock the next patient. Cancelling or failing export
leaves the local report and examination intact. A document provider can report an
error after writing; if the app cannot confirm an export, check its destination
before retrying.

### Android

**Skicka till skrivare (A4)** opens Android's native `PrintManager` workflow with
A4 selected. The device needs an installed and configured print service that can
reach the chosen printer. Printer discovery, transport and printer-specific
support come from that service; the app does not add a printer driver. A local
printer/service can be used without an internet connection, but support for a
particular USB or network printer must be tested on the intended tablet.

The print adapter handles layout, requested page ranges and cancellation. Closing
the print dialog, submitting a job or a printer service accepting the document
never confirms that paper has physically arrived.

### Windows

Printing opens the generated report in Windows' configured PDF reader. Choose
**Print / Ctrl+P**, select the printer and A4, then return to the screening app.
This is an explicit PDF-reader workflow, not silent direct printing. A suitable
PDF reader and printer configuration are required.

### Confirmation before the next patient

After requesting printing, staff selects **Utskriften har kommit ut – bekräfta**.
The app asks staff to check that **all pages for the correct reference have come
out and are readable**. Choosing No keeps the patient active. Cancelling or
failing the print operation also keeps the examination and report available for
reprinting.

**Nästa patient** remains locked until that explicit confirmation has been saved.
The storage layer independently checks the completed examination, the report file
and durable print confirmation before changing the active patient. Creating or
exporting a PDF alone cannot bypass this gate. Starting the next patient retains
the previous patient's files. If a missing report has to be recreated, printing
and receipt must be confirmed for the replacement before moving on.

## Camera and platform support

Android uses the latest USB-otoscopy implementation on this branch, including
VisioForge's automatic format selection. It does not add resolution switching or
a separate native mode-enumeration open/close cycle. Copying camera details is
still available.

The embedded `OtoscopyView` hides its standalone export/close controls because
the screening workflow owns the examination. `ReleaseForHostAsync()` is awaited
before changing the displayed step. Existing captures can be restored when an
interrupted examination is reopened. The camera service retains its shutdown
barrier across views so a failed/pending cleanup cannot be bypassed by opening a
new view.

Windows otoscopy uses Windows `MediaCapture` and colour-frame sources exposed by
the installed USB-camera driver. It filters USB device identities and avoids
cameras Windows identifies as built into the device. It needs a compatible
Windows driver and camera privacy access. A camera that only exposes a proprietary
vendor interface, or cannot supply usable colour frames through Windows' media
stack, may not work. Windows camera integration has compile verification only;
physical-camera compatibility still needs testing.

Patient-test adapters own their timers, event handlers and audio sessions. They
stop on pause/close and restore shared test settings after disposal. Configured
Android audio-device loss is an error; playback must not silently fall back to
the default speaker.

## Host integration

The supplied welcome and main-page changes already launch and close the module.
For another host:

```csharp
var previousContent = Content;
var screening = new STFM.Extension.Views.FullScreeningView();
screening.Closed += (_, _) => Content = previousContent;
Content = screening;
```

For navigation or a host Back handler, call:

```csharp
bool closed = await screening.CloseAsync();
```

Call on the UI thread. `CloseAsync()` returns false if saving or cleanup prevents
closing. Remove or replace the view only after its `Closed` event, not merely
because a close attempt started. The supplied `StfResearchSuite.MainPage` routes
hardware Back through this method for both launch routes. The module allows only
one active full-screening workflow in the process.

## Dependencies and verification

The report uses the existing SkiaSharp drawing stack, with these package versions:

| Package | Version | Purpose |
| --- | --- | --- |
| `Microsoft.Maui.Graphics.Skia` | `10.0.100` | Draw the existing MAUI/`PlotBase` audiogram onto the report canvas |
| `SkiaSharp` | `3.119.4` | PDF and preview rendering; pinned to avoid an unintended SkiaSharp 4 upgrade |

Android printing and Windows camera/printing use operating-system APIs. There is
no additional commercial reporting or printing SDK. The existing Android
VisioForge packages and licensing requirements still apply to USB otoscopy.

The following checks run without a physical camera, headphones or printer:

```sh
dotnet run --project Tests/OtoscopyExport/OtoscopyExport.Tests.csproj --configuration Release
dotnet run --project Tests/ScreeningWorkflow/ScreeningWorkflow.Tests.csproj --configuration Release
dotnet run --project Tests/ScreeningReport/ScreeningReport.Tests.csproj --configuration Release
```

The workflow checks exercise durable recovery, original-image preservation,
corrupt-state handling, patient isolation and the print/reset gate. The report
checks render sample PDFs/previews, exercise missing results, long-note pages,
page selection and cancellation, and verify that print snapshots retain their
own data. The report test project includes a Linux Skia native asset only when
running its headless checks on Linux; that asset is not an application dependency.

Build the Android application normally in Visual Studio. Windows compilation
checks do not exercise the native camera, PDF reader or printer. Neither compile
checks nor these regression checks establish physical-device compatibility. No
physical tablet, otoscope, calibrated audio setup or printer was available during
implementation.

Before using the workflow on the intended equipment, test:

1. A complete patient journey through all four steps, including known SSQ12
   answers, QuickSiP response/timeout behavior and tone press/hold/release behavior.
2. Both otoscopy images, retaking one side, explicit missing-image reasons and
   restored photographs after closing/reopening the view.
3. Pausing/backgrounding during each patient test, resuming at instructions and
   verifying that earlier completed results remain available.
4. Camera unplug/replug, missing audio hardware, and audio-device removal during
   a test. Verify that no test sound is rerouted to the tablet speaker.
5. PDF preview/export, both A4 pages plus any note continuation pages, printer
   margins, and readable left/right symbols and images on the actual paper.
6. Cancelling export/printing, printer failure, reprinting, rejecting paper
   confirmation, and verifying that the next-patient control stays locked.
7. Confirming all printed pages, starting the next patient, and verifying that
   no previous reference or result appears in the new examination.
8. Closing/reopening and restarting the application at instruction, result and
   report stages. Verify that the active patient and print gate recover correctly.
