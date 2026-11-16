// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

namespace STFM.Extension.Screening;

public sealed record Ssq12Answer(int Number, string Question, int? Rating, bool DontKnow);

/// <summary>Numeric answers are averaged; "don't know" is excluded, never scored zero.</summary>
public sealed record Ssq12Result(DateTimeOffset CompletedAt, Ssq12Answer[] Answers,
    bool? UsesHearingAids, bool HearingAidLeft, bool HearingAidRight)
{
    public int NumericAnswerCount => Answers.Count(answer => answer.Rating.HasValue);
    public double? MeanScore => NumericAnswerCount == 0 ? null
        : Answers.Where(answer => answer.Rating.HasValue).Average(answer => (double)answer.Rating!.Value);
}

/// <summary>Preserves the existing QuickSiP engine's score, including its handling of timeouts.</summary>
public sealed record QuickSipResult(DateTimeOffset CompletedAt, int Correct, int TrialCount,
    int PlannedTrialCount, int TimedOutCount, double? ScoreFraction, string Transducer, string StartList);

public enum AudiometryPointStatus { NotMeasured, Reached, NoResponse, Uncertain }

/// <summary>
/// AtScreeningLevel is an upper bound on threshold (heard at the screening level),
/// not an exact threshold. NoResponse is a lower bound; uncertain values retain that status.
/// </summary>
public sealed record AudiometryPoint(string Ear, int FrequencyHz, int? LevelDbHl,
    AudiometryPointStatus Status, bool IsReliabilityRetest, bool AtScreeningLevel);

/// <summary>Points retain the separate reliability retest; the last point per ear/frequency is used for plotting.</summary>
public sealed record AudiometryResult(DateTimeOffset CompletedAt, AudiometryPoint[] Points,
    string Transducer, string Protocol, int ScreeningLevelDbHl, int MaximumLevelDbHl);
