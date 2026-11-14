// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Globalization;
using System.IO.Compression;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace STFM.Extension.Otoscopy;

/// <summary>Identifies a completed archive and suggests a filename without patient information.</summary>
public sealed record OtoscopyExportInfo(Guid ExportId, DateTimeOffset CreatedAtUtc, string SuggestedFileName);

/// <summary>Packages the current ear images without resizing or re-encoding their JPEG data.</summary>
public static class OtoscopyExport
{
    /// <summary>
    /// Writes left.jpg and/or right.jpg plus manifest.json to an empty destination stream.
    /// Missing ears are explicitly marked in the manifest. The caller owns the destination;
    /// it remains open and is positioned after the archive on success. Rewind it before reading.
    /// On cancellation or an I/O failure, discard the incomplete destination.
    /// </summary>
    public static async Task<OtoscopyExportInfo> CreateArchiveAsync(
        Stream destination,
        OtoscopyImage? left,
        OtoscopyImage? right,
        string? reference = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite)
            throw new ArgumentException("The export destination must be writable.", nameof(destination));
        if (destination.CanSeek && (destination.Position != 0 || destination.Length != 0))
            throw new ArgumentException("The export destination must be empty.", nameof(destination));
        if (left is null && right is null)
            throw new InvalidOperationException("Capture at least one ear before exporting.");

        ValidateImage(left, OtoscopyEar.Left, nameof(left));
        ValidateImage(right, OtoscopyEar.Right, nameof(right));
        cancellationToken.ThrowIfCancellationRequested();

        var exportId = Guid.NewGuid();
        var createdAtUtc = DateTimeOffset.UtcNow;
        var fileName = $"otoscopy-{createdAtUtc.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture)}-{exportId:N}.zip";
        var manifest = new OtoscopyExportManifest(
            FormatVersion: 1,
            ExportId: exportId,
            CreatedAtUtc: createdAtUtc,
            Reference: string.IsNullOrWhiteSpace(reference) ? null : reference.Trim(),
            Left: DescribeImage(left, "left.jpg"),
            Right: DescribeImage(right, "right.jpg"));

        using (var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true))
        {
            if (left is not null)
                await AddImageAsync(archive, "left.jpg", left, createdAtUtc, cancellationToken).ConfigureAwait(false);
            if (right is not null)
                await AddImageAsync(archive, "right.jpg", right, createdAtUtc, cancellationToken).ConfigureAwait(false);

            var manifestEntry = archive.CreateEntry("manifest.json", CompressionLevel.Fastest);
            manifestEntry.LastWriteTime = createdAtUtc;
            await using var manifestStream = manifestEntry.Open();
            await JsonSerializer.SerializeAsync(manifestStream, manifest,
                OtoscopyExportJsonContext.Default.OtoscopyExportManifest, cancellationToken).ConfigureAwait(false);
        }

        cancellationToken.ThrowIfCancellationRequested();
        return new OtoscopyExportInfo(exportId, createdAtUtc, fileName);
    }

    private static void ValidateImage(OtoscopyImage? image, OtoscopyEar expectedEar, string parameterName)
    {
        if (image is null)
            return;
        if (image.Ear != expectedEar)
            throw new ArgumentException($"The {parameterName} image must belong to the {expectedEar} ear.", parameterName);
        if (image.JpegData.IsEmpty)
            throw new ArgumentException("The image contains no JPEG data.", parameterName);
    }

    private static OtoscopyExportEar DescribeImage(OtoscopyImage? image, string fileName)
        => new(image is not null, image is null ? null : fileName, image?.CapturedAt);

    private static async Task AddImageAsync(ZipArchive archive, string fileName, OtoscopyImage image,
        DateTimeOffset createdAtUtc, CancellationToken cancellationToken)
    {
        // JPEG data is already compressed. Preserve every byte of the original snapshot.
        var entry = archive.CreateEntry(fileName, CompressionLevel.NoCompression);
        entry.LastWriteTime = createdAtUtc;
        using var source = image.OpenRead();
        await using var destination = entry.Open();
        await source.CopyToAsync(destination, cancellationToken).ConfigureAwait(false);
    }
}

internal sealed record OtoscopyExportManifest(
    int FormatVersion,
    Guid ExportId,
    DateTimeOffset CreatedAtUtc,
    string? Reference,
    OtoscopyExportEar Left,
    OtoscopyExportEar Right);

internal sealed record OtoscopyExportEar(bool Captured, string? FileName, DateTimeOffset? CapturedAt);

// Source generation keeps archive export usable with MAUI trimming and AOT.
[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(OtoscopyExportManifest))]
internal partial class OtoscopyExportJsonContext : JsonSerializerContext;
