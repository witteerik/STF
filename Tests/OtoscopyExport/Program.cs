using System.IO.Compression;
using System.Text.Json;
using STFM.Extension.Otoscopy;

await BothEarsArePreserved();
await OneEarIsExplicitlyMissing();
await InvalidInputsDoNotWrite();
await CancellationDoesNotReportSuccess();
Console.WriteLine("All otoscopy export checks passed.");

static async Task BothEarsArePreserved()
{
    byte[] leftBytes = [0xff, 0xd8, 0x00, 0x13, 0xff, 0xd9];
    byte[] rightBytes = [0xff, 0xd8, 0x42, 0x18, 0xff, 0xd9];
    var leftTime = new DateTimeOffset(2026, 11, 13, 15, 10, 11, TimeSpan.FromHours(1));
    var rightTime = leftTime.AddMinutes(2);
    var left = new OtoscopyImage(OtoscopyEar.Left, leftBytes, leftTime);
    var right = new OtoscopyImage(OtoscopyEar.Right, rightBytes, rightTime);
    const string reference = "../exam\\outside:<42>";
    using var output = new MemoryStream();
    var info = await OtoscopyExport.CreateArchiveAsync(output, left, right, reference);
    Check(output.CanWrite, "The destination must remain open.");
    Check(!info.SuggestedFileName.Contains("exam") && !info.SuggestedFileName.Contains('/') &&
          !info.SuggestedFileName.Contains('\\'), "The reference must never become a path.");
    Check(info.SuggestedFileName.EndsWith($"-{info.ExportId:N}.zip", StringComparison.Ordinal), "Unique filename ID.");
    Check(info.CreatedAtUtc.Offset == TimeSpan.Zero, "Export creation time must be UTC.");

    output.Position = 0;
    using var zip = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true);
    Check(zip.Entries.Select(entry => entry.FullName).Order().SequenceEqual(
        new[] { "left.jpg", "manifest.json", "right.jpg" }), "Unexpected archive entry names.");
    Check((await ReadEntry(zip, "left.jpg")).SequenceEqual(leftBytes), "Left image bytes changed.");
    Check((await ReadEntry(zip, "right.jpg")).SequenceEqual(rightBytes), "Right image bytes changed.");
    using var document = JsonDocument.Parse(await ReadEntry(zip, "manifest.json"));
    var manifest = document.RootElement;
    Check(manifest.GetProperty("formatVersion").GetInt32() == 1, "Format version missing.");
    Check(manifest.GetProperty("exportId").GetGuid() == info.ExportId, "Archive/filename ID mismatch.");
    Check(manifest.GetProperty("createdAtUtc").GetDateTimeOffset() == info.CreatedAtUtc, "Export time mismatch.");
    Check(manifest.GetProperty("reference").GetString() == reference, "Reference changed.");
    foreach (var (side, expectedTime) in new[] { ("left", leftTime), ("right", rightTime) })
    {
        var ear = manifest.GetProperty(side);
        Check(ear.GetProperty("captured").GetBoolean(), $"{side} capture flag missing.");
        Check(ear.GetProperty("fileName").GetString() == $"{side}.jpg", $"{side} filename mismatch.");
        var actual = ear.GetProperty("capturedAt").GetDateTimeOffset();
        Check(actual == expectedTime && actual.Offset == expectedTime.Offset, $"{side} timestamp/offset changed.");
    }

    using var secondOutput = new MemoryStream();
    var secondInfo = await OtoscopyExport.CreateArchiveAsync(secondOutput, left, right);
    Check(secondInfo.ExportId != info.ExportId && secondInfo.SuggestedFileName != info.SuggestedFileName,
        "Repeated exports must not reuse filenames.");
    Check(left.JpegData.Span.SequenceEqual(leftBytes) && right.JpegData.Span.SequenceEqual(rightBytes),
        "Source images were modified.");
    Console.WriteLine("PASS: Both ears, exact JPEG bytes, timestamps, unique filenames, and reference isolation.");
}

static async Task OneEarIsExplicitlyMissing()
{
    foreach (var capturedSide in new[] { OtoscopyEar.Left, OtoscopyEar.Right })
    {
        var image = new OtoscopyImage(capturedSide, [0xff, 0xd8, 0xff, 0xd9], DateTimeOffset.Now);
        using var output = new MemoryStream();
        await OtoscopyExport.CreateArchiveAsync(output,
            capturedSide == OtoscopyEar.Left ? image : null,
            capturedSide == OtoscopyEar.Right ? image : null, "  ");
        output.Position = 0;
        using var zip = new ZipArchive(output, ZipArchiveMode.Read);
        Check(zip.Entries.Count == 2, "One-ear export must contain only one JPEG and the manifest.");
        var missingSide = capturedSide == OtoscopyEar.Left ? "right" : "left";
        Check(zip.GetEntry($"{missingSide}.jpg") is null, "Missing ear must not have an invented image.");
        using var document = JsonDocument.Parse(await ReadEntry(zip, "manifest.json"));
        var missing = document.RootElement.GetProperty(missingSide);
        Check(!missing.GetProperty("captured").GetBoolean(), "Missing-ear flag incorrect.");
        Check(missing.GetProperty("fileName").ValueKind == JsonValueKind.Null &&
              missing.GetProperty("capturedAt").ValueKind == JsonValueKind.Null, "Missing ear metadata must be null.");
        Check(document.RootElement.GetProperty("reference").ValueKind == JsonValueKind.Null, "Blank reference must be null.");
    }
    Console.WriteLine("PASS: Left-only and right-only exports explicitly describe the missing ear.");
}

static async Task InvalidInputsDoNotWrite()
{
    using var output = new MemoryStream();
    var left = new OtoscopyImage(OtoscopyEar.Left, [1, 2, 3], DateTimeOffset.Now);
    var right = new OtoscopyImage(OtoscopyEar.Right, [4, 5], DateTimeOffset.Now);
    await Throws<InvalidOperationException>(() => OtoscopyExport.CreateArchiveAsync(output, null, null));
    await Throws<ArgumentException>(() => OtoscopyExport.CreateArchiveAsync(output, right, null));
    await Throws<ArgumentException>(() => OtoscopyExport.CreateArchiveAsync(output, null, left));
    await Throws<ArgumentException>(() => OtoscopyExport.CreateArchiveAsync(output,
        new OtoscopyImage(OtoscopyEar.Left, [], DateTimeOffset.Now), null));
    Check(output.Length == 0, "Validation failure must precede all writes.");
    using var readOnly = new MemoryStream([1], writable: false);
    await Throws<ArgumentException>(() => OtoscopyExport.CreateArchiveAsync(readOnly, left, null));
    output.WriteByte(123);
    await Throws<ArgumentException>(() => OtoscopyExport.CreateArchiveAsync(output, left, null));
    Check(output.Length == 1 && output.ToArray()[0] == 123, "Nonempty destination was altered.");
    Console.WriteLine("PASS: Empty exams, swapped sides, empty images and invalid destinations are rejected without writes.");
}

static async Task CancellationDoesNotReportSuccess()
{
    using var output = new MemoryStream();
    using var cancellation = new CancellationTokenSource();
    cancellation.Cancel();
    var left = new OtoscopyImage(OtoscopyEar.Left, [1, 2, 3], DateTimeOffset.Now);
    await Throws<OperationCanceledException>(() => OtoscopyExport.CreateArchiveAsync(output, left, null,
        cancellationToken: cancellation.Token));
    Check(output.Length == 0 && output.CanWrite, "Pre-cancellation must not write or close the destination.");
    Console.WriteLine("PASS: Cancellation is observed before writing and cannot be reported as a successful export.");
}

static async Task<byte[]> ReadEntry(ZipArchive zip, string name)
{
    using var input = (zip.GetEntry(name) ?? throw new Exception($"Missing {name}.")).Open();
    using var output = new MemoryStream();
    await input.CopyToAsync(output);
    return output.ToArray();
}

static void Check(bool condition, string message)
{
    if (!condition)
        throw new Exception(message);
}

static async Task Throws<T>(Func<Task> action) where T : Exception
{
    try { await action(); }
    catch (T) { return; }
    throw new Exception($"Expected {typeof(T).Name}.");
}
