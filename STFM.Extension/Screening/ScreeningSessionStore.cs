// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Text;
using System.Text.Json;
using System.Collections.Concurrent;
using STFM.Extension.Otoscopy;

namespace STFM.Extension.Screening;

/// <summary>
/// Local durable examination storage. The active pointer changes only after its
/// new journal is complete. Previously saved examinations and captures are retained.
/// A failed save never changes the active examination to a new patient.
/// </summary>
public sealed class ScreeningSessionStore
{
    // Hosts may recreate a view while an earlier journal write is completing.
    // Serialize stores that refer to the same directory across that transition.
    private static readonly ConcurrentDictionary<string, SemaphoreSlim> DirectoryGates = new(
        OperatingSystem.IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
    private readonly string baseDirectory;
    private readonly string activePointer;
    private readonly SemaphoreSlim gate;

    public ScreeningSessionStore(string baseDirectory)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(baseDirectory);
        this.baseDirectory = Path.TrimEndingDirectorySeparator(Path.GetFullPath(baseDirectory));
        activePointer = Path.Combine(this.baseDirectory, "active-examination.txt");
        gate = DirectoryGates.GetOrAdd(this.baseDirectory, _ => new SemaphoreSlim(1, 1));
    }

    public string GetExaminationDirectory(Guid id)
    {
        if (id == Guid.Empty)
            throw new ArgumentException("An examination must have an ID.", nameof(id));
        return Path.Combine(baseDirectory, id.ToString("N"));
    }

    /// <summary>Returns null only when no active pointer exists; corrupt/missing journals throw.</summary>
    public async Task<ScreeningExamination?> LoadActiveAsync(CancellationToken cancellationToken = default)
    {
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var examination = await ReadActiveAsync(cancellationToken).ConfigureAwait(false);
            examination?.RecoverInterruptedStage();
            return examination;
        }
        finally { gate.Release(); }
    }

    /// <summary>Creates the initial patient only. Use StartNextAsync for a subsequent patient.</summary>
    public async Task<ScreeningExamination> CreateAsync(string reference, CancellationToken cancellationToken = default)
    {
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (await ReadActiveAsync(cancellationToken).ConfigureAwait(false) is not null)
                throw new InvalidOperationException("An examination is already active. Resume it before starting another patient.");
            return await CreateAndActivateAsync(reference, cancellationToken).ConfigureAwait(false);
        }
        finally { gate.Release(); }
    }

    /// <summary>Save is serialized. The caller must not mutate examination while this task is running.</summary>
    public async Task SaveAsync(ScreeningExamination examination, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(examination);
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await RequireActiveAsync(examination.Id, cancellationToken).ConfigureAwait(false);
            await WriteExaminationAsync(examination, cancellationToken).ConfigureAwait(false);
        }
        finally { gate.Release(); }
    }

    /// <summary>
    /// Saves a new immutable JPEG file. Unique names keep an earlier journal's image
    /// intact if saving the replacement journal fails. The caller then saves its new path.
    /// </summary>
    public async Task<string> SaveImageAsync(ScreeningExamination examination, OtoscopyEar side,
        ReadOnlyMemory<byte> jpeg, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(examination);
        if (side is not OtoscopyEar.Left and not OtoscopyEar.Right)
            throw new ArgumentOutOfRangeException(nameof(side));
        if (jpeg.IsEmpty)
            throw new ArgumentException("The captured image is empty.", nameof(jpeg));
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            await RequireActiveAsync(examination.Id, cancellationToken).ConfigureAwait(false);
            var name = $"{side.ToString().ToLowerInvariant()}-{Guid.NewGuid():N}.jpg";
            var path = Path.Combine(GetExaminationDirectory(examination.Id), name);
            await WriteAtomicAsync(path, stream => stream.WriteAsync(jpeg, cancellationToken).AsTask(),
                cancellationToken).ConfigureAwait(false);
            return path;
        }
        finally { gate.Release(); }
    }

    /// <summary>
    /// Starts another patient only after a durable completed report and explicit paper
    /// confirmation. Save the confirmation with SaveAsync before calling this method.
    /// The old patient's files are preserved, never reset or deleted.
    /// </summary>
    public async Task<ScreeningExamination> StartNextAsync(ScreeningExamination completedExamination,
        string reference, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(completedExamination);
        await gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            var saved = await RequireActiveAsync(completedExamination.Id, cancellationToken).ConfigureAwait(false);
            if (!saved.CanStartNext)
                throw new InvalidOperationException("Confirm that the complete printed report has been received before starting another patient.");
            if (!File.Exists(saved.ReportPdfPath))
                throw new InvalidOperationException("The saved report is missing. Regenerate it and confirm its printed copy before starting another patient.");
            return await CreateAndActivateAsync(reference, cancellationToken).ConfigureAwait(false);
        }
        finally { gate.Release(); }
    }

    private async Task<ScreeningExamination> CreateAndActivateAsync(string reference, CancellationToken cancellationToken)
    {
        var examination = new ScreeningExamination { Reference = reference ?? string.Empty };
        await WriteExaminationAsync(examination, cancellationToken).ConfigureAwait(false);
        var pointerBytes = Encoding.UTF8.GetBytes(examination.Id.ToString("N"));
        await WriteAtomicAsync(activePointer, stream => stream.WriteAsync(pointerBytes, cancellationToken).AsTask(),
            cancellationToken).ConfigureAwait(false);
        return examination;
    }

    private async Task<ScreeningExamination> RequireActiveAsync(Guid id, CancellationToken cancellationToken)
    {
        var current = await ReadActiveAsync(cancellationToken).ConfigureAwait(false);
        if (current is null || current.Id != id)
            throw new InvalidOperationException("This examination is no longer the active patient.");
        return current;
    }

    private async Task<ScreeningExamination?> ReadActiveAsync(CancellationToken cancellationToken)
    {
        if (!File.Exists(activePointer))
        {
            if (Directory.Exists(activePointer))
                throw new InvalidDataException("The active examination pointer is invalid. Existing patient files have been preserved.");
            return null;
        }
        var text = await File.ReadAllTextAsync(activePointer, cancellationToken).ConfigureAwait(false);
        if (!Guid.TryParseExact(text.Trim(), "N", out var id) || id == Guid.Empty)
            throw new InvalidDataException("The active examination pointer is unreadable. Existing patient files have been preserved.");
        var path = Path.Combine(GetExaminationDirectory(id), "examination.json");
        try
        {
            await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read,
                4096, FileOptions.Asynchronous | FileOptions.SequentialScan);
            var examination = await JsonSerializer.DeserializeAsync(stream,
                ScreeningJsonContext.Default.ScreeningExamination, cancellationToken).ConfigureAwait(false)
                ?? throw new InvalidDataException("The examination journal is empty.");
            if (examination.Id != id)
                throw new InvalidDataException("The examination journal does not match its active patient ID.");
            Validate(examination);
            return examination;
        }
        catch (Exception exception) when (exception is JsonException or FileNotFoundException or DirectoryNotFoundException)
        {
            throw new InvalidDataException("The active examination could not be read. Patient files have been preserved; do not start a new patient until this is resolved.", exception);
        }
    }

    private async Task WriteExaminationAsync(ScreeningExamination examination, CancellationToken cancellationToken)
    {
        Validate(examination);
        var path = Path.Combine(GetExaminationDirectory(examination.Id), "examination.json");
        await WriteAtomicAsync(path, stream => JsonSerializer.SerializeAsync(stream, examination,
            ScreeningJsonContext.Default.ScreeningExamination, cancellationToken), cancellationToken).ConfigureAwait(false);
    }

    private void Validate(ScreeningExamination examination)
    {
        if (examination.FormatVersion != 1 || examination.Id == Guid.Empty || examination.Reference is null ||
            examination.ReportPreviewPaths is null || !Enum.IsDefined(examination.Stage) ||
            (examination.Ssq12 is not null && examination.Ssq12.Answers is null) ||
            (examination.Audiometry is not null && examination.Audiometry.Points is null))
            throw new InvalidDataException("The examination journal has an unsupported or invalid format.");
        var directory = GetExaminationDirectory(examination.Id) + Path.DirectorySeparatorChar;
        foreach (var path in new[] { examination.ReportPdfPath, examination.Otoscopy?.LeftImagePath,
                     examination.Otoscopy?.RightImagePath }.Concat(examination.ReportPreviewPaths))
        {
            if (string.IsNullOrWhiteSpace(path))
                continue;
            var comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (!Path.GetFullPath(path).StartsWith(directory, comparison))
                throw new InvalidDataException("An examination file points outside its patient directory.");
        }
    }

    private static async Task WriteAtomicAsync(string path, Func<FileStream, Task> write,
        CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            await using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write,
                FileShare.None, 8192, FileOptions.Asynchronous))
            {
                await write(stream).ConfigureAwait(false);
                await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                stream.Flush(flushToDisk: true);
            }
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporary, path, overwrite: true);
        }
        finally
        {
            // Failure/cancellation leaves the previous complete destination untouched.
            if (File.Exists(temporary))
                File.Delete(temporary);
        }
    }
}
