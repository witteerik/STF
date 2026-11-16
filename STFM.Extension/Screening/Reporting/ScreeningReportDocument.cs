// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

namespace STFM.Extension.Screening.Reporting;

/// <summary>A report and previews generated from the same immutable examination snapshot.</summary>
public sealed record ScreeningReportDocument(string PdfPath, IReadOnlyList<string> PreviewImagePaths)
{
    internal Action<Stream, IReadOnlyList<int>, float, float, float, CancellationToken>? WritePages { get; init; }
    public int PageCount => PreviewImagePaths.Count;
}

/// <summary>Indicates only that the operating system print UI/reader was opened, never paper output.</summary>
public sealed record ScreeningPrintRequestResult(bool Requested, bool OpenedInPdfReader, string Message);
