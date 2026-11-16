// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using System.Text.Json.Serialization;

namespace STFM.Extension.Screening;

public enum ScreeningStage
{
    Setup,
    OtoscopyInstructions,
    Otoscopy,
    Ssq12Instructions,
    Ssq12,
    QuickSipInstructions,
    QuickSip,
    AudiometryInstructions,
    Audiometry,
    Report
}

public sealed class OtoscopyScreeningResult
{
    public string? LeftImagePath { get; set; }
    public string? RightImagePath { get; set; }
    public DateTimeOffset? LeftCapturedAt { get; set; }
    public DateTimeOffset? RightCapturedAt { get; set; }
    public string? LeftNotCapturedReason { get; set; }
    public string? RightNotCapturedReason { get; set; }
    public string? Notes { get; set; }
}

/// <summary>
/// Durable patient examination state. Submission to a print service and confirmation
/// that paper was received are deliberately separate events.
/// </summary>
public sealed class ScreeningExamination
{
    public int FormatVersion { get; set; } = 1;
    public Guid Id { get; init; } = Guid.NewGuid();
    public string Reference { get; set; } = string.Empty;
    public DateTimeOffset StartedAt { get; init; } = DateTimeOffset.Now;
    public DateTimeOffset? CompletedAt { get; set; }
    public ScreeningStage Stage { get; set; }
    public OtoscopyScreeningResult? Otoscopy { get; set; }
    public Ssq12Result? Ssq12 { get; set; }
    public QuickSipResult? QuickSip { get; set; }
    public AudiometryResult? Audiometry { get; set; }
    public string? ReportPdfPath { get; set; }
    public string[] ReportPreviewPaths { get; set; } = [];
    public DateTimeOffset? PrintRequestedAt { get; set; }
    public DateTimeOffset? PrintConfirmedAt { get; set; }

    [JsonIgnore]
    public bool IsComplete => Stage == ScreeningStage.Report && CompletedAt.HasValue &&
        Otoscopy is not null && Ssq12 is not null && QuickSip is not null && Audiometry is not null;

    [JsonIgnore]
    public bool CanStartNext => IsComplete && !string.IsNullOrWhiteSpace(ReportPdfPath) &&
        PrintRequestedAt.HasValue && PrintConfirmedAt.HasValue && PrintConfirmedAt >= PrintRequestedAt;

    /// <summary>Interrupted patient tests restart from their instructions, never from an active tone/trial.</summary>
    public ScreeningStage RecoverInterruptedStage()
    {
        Stage = Stage switch
        {
            ScreeningStage.Otoscopy => ScreeningStage.OtoscopyInstructions,
            ScreeningStage.Ssq12 => ScreeningStage.Ssq12Instructions,
            ScreeningStage.QuickSip => ScreeningStage.QuickSipInstructions,
            ScreeningStage.Audiometry => ScreeningStage.AudiometryInstructions,
            _ => Stage
        };
        return Stage;
    }

    public void RegisterPrintRequest(DateTimeOffset requestedAt)
    {
        if (!IsComplete || string.IsNullOrWhiteSpace(ReportPdfPath))
            throw new InvalidOperationException("A completed examination and report are required before printing.");
        PrintRequestedAt = requestedAt;
        PrintConfirmedAt = null;
    }

    /// <summary>Call only after staff explicitly confirms receiving and checking the printed report.</summary>
    public void ConfirmPrinted(DateTimeOffset confirmedAt)
    {
        if (!IsComplete || !PrintRequestedAt.HasValue || confirmedAt < PrintRequestedAt.Value)
            throw new InvalidOperationException("The report must be sent to a printer before receipt can be confirmed.");
        PrintConfirmedAt = confirmedAt;
    }
}
