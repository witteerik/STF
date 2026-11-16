using STFM.Extension.Otoscopy;
using STFM.Extension.Screening;

var root = Path.Combine(Path.GetTempPath(), "stf-screening-checks-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    await PersistenceAndPrintGate();
    await RecoveryAndCorruption();
    await RetakeAndFailureSafety();
    await ConcurrentHosts();
    Console.WriteLine("Screening workflow checks passed.");
}
finally { Directory.Delete(root, recursive: true); }

async Task PersistenceAndPrintGate()
{
    var directory = Path.Combine(root, "print");
    var store = new ScreeningSessionStore(directory);
    Equal(null, await store.LoadActiveAsync(), "No exam is invented before setup");
    var exam = await store.CreateAsync("../patient-reference-is-not-a-path");
    Equal(Path.Combine(directory, exam.Id.ToString("N")), store.GetExaminationDirectory(exam.Id), "Patient reference is not a directory");
    await Throws<InvalidOperationException>(() => store.CreateAsync("replacement"), "Cannot replace an active exam");
    await Throws<InvalidOperationException>(() => store.StartNextAsync(exam, "new"), "Incomplete exam cannot reset");
    MakeComplete(exam);
    exam.ReportPdfPath = Path.Combine(store.GetExaminationDirectory(exam.Id), "report.pdf");
    await File.WriteAllTextAsync(exam.ReportPdfPath, "%PDF-test");
    await store.SaveAsync(exam);
    await Throws<InvalidOperationException>(() => store.StartNextAsync(exam, "new"), "Report creation is not printing");
    exam.RegisterPrintRequest(DateTimeOffset.UtcNow);
    await store.SaveAsync(exam);
    Check(!exam.CanStartNext, "Submitting a print job cannot unlock the next patient");
    await Throws<InvalidOperationException>(() => store.StartNextAsync(exam, "new"), "Submitted job is not received paper");
    exam.ConfirmPrinted(DateTimeOffset.UtcNow);
    Check(exam.CanStartNext, "Explicit confirmation enables the model gate");
    await Throws<InvalidOperationException>(() => store.StartNextAsync(exam, "new"), "Confirmation must be durably saved");
    await store.SaveAsync(exam);
    var restored = await new ScreeningSessionStore(directory).LoadActiveAsync();
    Check(restored!.CanStartNext, "Print confirmation survives restart");
    Equal(5.0, restored.Ssq12!.MeanScore, "SSQ unknown answers are excluded from mean after serialization");
    Equal(2, restored.Ssq12.NumericAnswerCount, "SSQ numeric count survives serialization");
    Equal(AudiometryPointStatus.NoResponse, restored.Audiometry!.Points[0].Status, "No-response is not serialized as a normal threshold");
    File.Delete(exam.ReportPdfPath);
    await Throws<InvalidOperationException>(() => store.StartNextAsync(exam, "new"), "Missing report cannot reset");
    await File.WriteAllTextAsync(exam.ReportPdfPath, "%PDF-test");
    var next = await store.StartNextAsync(exam, "next patient");
    Check(next.Id != exam.Id && !next.CanStartNext, "Next patient has fresh independent state");
    Check(File.Exists(exam.ReportPdfPath), "New patient preserves old PDF");
    Check(File.Exists(Path.Combine(store.GetExaminationDirectory(exam.Id), "examination.json")), "New patient preserves old journal");
    Equal(next.Id, (await store.LoadActiveAsync())!.Id, "Pointer changes to the new patient");
    await Throws<InvalidOperationException>(() => store.SaveAsync(exam), "A late old-patient event cannot overwrite new active patient");
}

async Task RecoveryAndCorruption()
{
    var directory = Path.Combine(root, "recovery");
    var store = new ScreeningSessionStore(directory);
    var exam = await store.CreateAsync("Exam A");
    exam.Ssq12 = Ssq();
    foreach (var (running, instructions) in new[] {
        (ScreeningStage.Otoscopy, ScreeningStage.OtoscopyInstructions),
        (ScreeningStage.Ssq12, ScreeningStage.Ssq12Instructions),
        (ScreeningStage.QuickSip, ScreeningStage.QuickSipInstructions),
        (ScreeningStage.Audiometry, ScreeningStage.AudiometryInstructions) })
    {
        exam.Stage = running;
        await store.SaveAsync(exam);
        var loaded = await new ScreeningSessionStore(directory).LoadActiveAsync();
        Equal(instructions, loaded!.Stage, "Interrupted tests resume at instructions, not an active stimulus");
        Equal(5.0, loaded.Ssq12!.MeanScore, "Recovery preserves completed result data");
    }
    var journalPath = Path.Combine(store.GetExaminationDirectory(exam.Id), "examination.json");
    var validJournal = await File.ReadAllBytesAsync(journalPath);
    await File.WriteAllTextAsync(journalPath, "{broken");
    await Throws<InvalidDataException>(() => store.LoadActiveAsync(), "A corrupt journal cannot silently reset the patient");
    await Throws<InvalidDataException>(() => store.CreateAsync("replacement"), "Create cannot bypass corrupt active data");
    await File.WriteAllBytesAsync(journalPath, validJournal);
    await File.WriteAllTextAsync(Path.Combine(directory, "active-examination.txt"), "../escape");
    await Throws<InvalidDataException>(() => store.LoadActiveAsync(), "Pointer traversal is rejected");
    Check(File.Exists(journalPath), "Recovery errors preserve the journal");
}

async Task RetakeAndFailureSafety()
{
    var directory = Path.Combine(root, "images");
    var store = new ScreeningSessionStore(directory);
    var exam = await store.CreateAsync("Images");
    byte[] first = [0xff, 0xd8, 0x11, 0x22, 0xff, 0xd9];
    byte[] second = [0xff, 0xd8, 0x99, 0xff, 0xd9];
    var firstPath = await store.SaveImageAsync(exam, OtoscopyEar.Left, first);
    exam.Otoscopy = new() { LeftImagePath = firstPath, LeftCapturedAt = DateTimeOffset.Now, RightNotCapturedReason = "Not performed" };
    await store.SaveAsync(exam);
    var journalPath = Path.Combine(store.GetExaminationDirectory(exam.Id), "examination.json");
    var before = await File.ReadAllBytesAsync(journalPath);
    var secondPath = await store.SaveImageAsync(exam, OtoscopyEar.Left, second);
    Check(firstPath != secondPath, "Retakes never overwrite the journal's previous JPEG");
    Check((await File.ReadAllBytesAsync(firstPath)).SequenceEqual(first), "Original JPEG bytes are preserved");
    Check((await File.ReadAllBytesAsync(secondPath)).SequenceEqual(second), "New JPEG bytes are preserved");
    exam.Otoscopy.LeftImagePath = secondPath;
    using var cancelled = new CancellationTokenSource();
    cancelled.Cancel();
    await Throws<OperationCanceledException>(() => store.SaveAsync(exam, cancelled.Token), "Cancelled saves are reported");
    Check((await File.ReadAllBytesAsync(journalPath)).SequenceEqual(before), "Cancelled saves leave the previous complete journal untouched");
    Equal(firstPath, (await store.LoadActiveAsync())!.Otoscopy!.LeftImagePath, "Uncommitted retake is not substituted for the previous result");
    exam.ReportPdfPath = Path.Combine(root, "foreign-patient.pdf");
    await Throws<InvalidDataException>(() => store.SaveAsync(exam), "Cross-patient file paths are rejected");
    Check((await File.ReadAllBytesAsync(journalPath)).SequenceEqual(before), "Validation failure preserves the previous journal");
    Check(!Directory.EnumerateFiles(store.GetExaminationDirectory(exam.Id), "*.tmp").Any(), "No cancelled temporary files are left behind");
    await Throws<ArgumentException>(() => store.SaveImageAsync(exam, OtoscopyEar.Right, ReadOnlyMemory<byte>.Empty), "Empty image is not saved");
}

async Task ConcurrentHosts()
{
    var directory = Path.Combine(root, "concurrent");
    var first = new ScreeningSessionStore(directory);
    var second = new ScreeningSessionStore(directory + Path.DirectorySeparatorChar);
    var attempt1 = first.CreateAsync("First host");
    var attempt2 = second.CreateAsync("Second host");
    try { await Task.WhenAll(attempt1, attempt2); }
    catch (InvalidOperationException) { }
    Check(attempt1.IsCompletedSuccessfully != attempt2.IsCompletedSuccessfully,
        "Two hosts cannot simultaneously create different active patients");
    var winner = attempt1.IsCompletedSuccessfully ? attempt1.Result : attempt2.Result;
    Equal(winner.Id, (await first.LoadActiveAsync())!.Id, "Concurrent hosts retain the one successful patient");
}

static Ssq12Result Ssq() => new(DateTimeOffset.Now,
    [new(1, "One", 4, false), new(2, "Two", null, true), new(3, "Three", 6, false)], null, false, false);
static void MakeComplete(ScreeningExamination exam)
{
    exam.Stage = ScreeningStage.Report;
    exam.CompletedAt = DateTimeOffset.Now;
    exam.Otoscopy = new() { LeftNotCapturedReason = "Not performed", RightNotCapturedReason = "Not performed" };
    exam.Ssq12 = Ssq();
    exam.QuickSip = new(DateTimeOffset.Now, 20, 30, 30, 2, 20.0 / 30, "Test transducer", "List 1");
    exam.Audiometry = new(DateTimeOffset.Now,
        [new("Right", 1000, 60, AudiometryPointStatus.NoResponse, false, false)], "Test transducer", "SAME96_Screening", 20, 60);
}
static void Check(bool condition, string description)
{
    if (!condition) throw new Exception("FAILED: " + description);
}
static void Equal<T>(T expected, T actual, string description) => Check(EqualityComparer<T>.Default.Equals(expected, actual), description);
static async Task Throws<T>(Func<Task> action, string description) where T : Exception
{
    try { await action(); }
    catch (T) { return; }
    throw new Exception("FAILED: " + description);
}
