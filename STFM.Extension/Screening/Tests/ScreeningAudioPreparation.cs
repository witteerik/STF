// Copyright (c) 2026 Erik Witte. See License.txt in the repository root.

using STFN.Core;
using STFN.Core.Audio.Formats;

namespace STFM.Extension.Screening.Tests;

/// <summary>Uses the host's configured hardware and calibration, never an arbitrary replacement output.</summary>
public static class ScreeningAudioPreparation
{
    public const string QuickSipMaterialName = "The Swedish SiP-test (Quick-SiP)";

    public static async Task<AudioSystemSpecification> PrepareAsync(CancellationToken cancellationToken = default,
        AudioSystemSpecification? requestedTransducer = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (!STFM.StfmBase.IsInitialized)
            await STFM.StfmBase.InitializeSTFM(CloseAppOnAudioFailure: false);
        cancellationToken.ThrowIfCancellationRequested();
        if (Globals.StfBase.SoundPlayer is null)
            throw new InvalidOperationException("Ljudspelaren kunde inte förberedas. Kontrollera appens ljudinställningar.");
        var transducer = requestedTransducer ?? Globals.StfBase.SoundPlayer.GetMixer()?.GetParentTransducerSpecification;
        if (transducer is null)
        {
            var candidates = Globals.StfBase.AvaliableTransducers.Where(item => item.CanPlay && item.IsHeadphones()).ToArray();
            if (candidates.Length > 1) throw new ScreeningTransducerSelectionRequiredException(candidates);
            transducer = candidates.SingleOrDefault()
                ?? throw new InvalidOperationException("Ingen hörlursutgång är konfigurerad. Kontrollera appens ljudinställningar och anslut den kalibrerade utrustningen.");
        }
        ValidateTransducer(transducer);
        Globals.StfBase.LoadAvailableSpeechMaterialSpecifications();
        if (!Globals.StfBase.AvailableSpeechMaterials.Any(material => material.Name == QuickSipMaterialName))
            throw new InvalidOperationException($"QuickSiP-materialet saknas: {QuickSipMaterialName}. Installera materialet via appens ordinarie inställningar.");
        var test = new QuickSiP_StaticSpeechTest(QuickSipMaterialName);
        ValidateQuickSipMaterial(test);
        return transducer;
    }

    public static void ValidateTransducer(AudioSystemSpecification transducer)
    {
        ArgumentNullException.ThrowIfNull(transducer);
        if (!transducer.CanPlay || transducer.Mixer is null || !transducer.IsHeadphones())
            throw new InvalidOperationException("Screeningen kräver en konfigurerad och tillgänglig hörlursutgång. Kontrollera ljudenheten och kalibreringen.");
        if (transducer.HardwareOutputChannels.Count < 2 || transducer.HardwareOutputChannels.Distinct().Count() < 2 ||
            transducer.CalibrationGain.Count < 2 || transducer.CalibrationGain.Any(value => !double.IsFinite(value)))
            throw new InvalidOperationException("Ljudgivaren saknar giltig stereoroutning eller kalibreringsdata.");
        foreach (int frequency in new UoPta(PtaTestProtocol: UoPta.PtaTestProtocols.SAME96_Screening).GetTestFrequencies())
        {
            if (!transducer.PtaCalibrationGain.TryGetValue(frequency, out var gain) || !double.IsFinite(gain))
                throw new InvalidOperationException($"Tonaudiometrins kalibrering saknas eller är ogiltig vid {frequency} Hz.");
        }
#if ANDROID
        string outputDevice = transducer.ParentAudioApiSettings.GetSelectedOutputDeviceName();
        if (string.IsNullOrWhiteSpace(outputDevice) || !STFM.AndroidAudioTrackPlayer.CheckIfDeviceExists(outputDevice, true))
            throw new InvalidOperationException($"Den konfigurerade ljudenheten är inte ansluten: {outputDevice}. Anslut den innan testet startas.");
#endif
    }

    internal static void ValidateQuickSipMaterial(QuickSiP_StaticSpeechTest test, string? startList = null)
    {
        if (test.SpeechMaterial is null || test.AvailableMediasets.Count == 0)
            throw new InvalidOperationException("QuickSiP saknar ett läsbart talmaterial eller mediaset.");
        var media = test.AvailableMediasets[0];
        var lists = test.SpeechMaterial.GetAllRelativesAtLevel(SpeechMaterialComponent.LinguisticLevels.List, true, false);
        if (startList is not null)
            lists = lists.Where(list => list.PrimaryStringRepresentation == startList).ToList();
        if (lists.Count == 0)
            throw new InvalidOperationException("QuickSiP saknar en giltig testlista.");
        string initialSound = Path.Combine(test.SpeechMaterial.ParentTestSpecification.GetTestRootPath(), media.LombardNoisePath);
        if (!File.Exists(initialSound))
            throw new FileNotFoundException("QuickSiP:s inledande ljudfil saknas.", initialSound);
        foreach (var list in lists)
        {
            var words = list.GetAllDescenentsAtLevel(SpeechMaterialComponent.LinguisticLevels.Sentence);
            if (words.Count == 0)
                throw new InvalidOperationException($"QuickSiP-listan {list.PrimaryStringRepresentation} är tom.");
            foreach (var word in words)
            {
                string path = word.GetSoundPath(ref media, 0);
                if (!File.Exists(path))
                    throw new FileNotFoundException("En ljudfil som behövs för QuickSiP saknas.", path);
                if (string.IsNullOrWhiteSpace(word.GetCategoricalVariableValue("Spelling")) ||
                    word.GetCategoricalVariableValue("Alternatives").Split(',').Count(value => !string.IsNullOrWhiteSpace(value)) != 3)
                    throw new InvalidOperationException($"QuickSiP-listan {list.PrimaryStringRepresentation} har ofullständiga svarsalternativ.");
            }
        }
    }

    internal static void ConfigurePlayer(AudioSystemSpecification transducer, int sampleRate, int bitDepth,
        WaveFormat.WaveFormatEncodings encoding, double overlapDuration)
    {
        ValidateTransducer(transducer);
        Globals.StfBase.SoundPlayer.ChangePlayerSettings(transducer.ParentAudioApiSettings, sampleRate, bitDepth,
            encoding, overlapDuration, transducer.Mixer,
            STFN.Core.Audio.SoundPlayers.iSoundPlayer.SoundDirections.PlaybackOnly,
            ReOpenStream: true, ReStartStream: true);
    }

    internal static void StopPlayback()
    {
        var player = Globals.StfBase.SoundPlayer;
        if (player is null) return;
        try { player.FadeOutPlayback(); }
        catch { player.CloseStream(); }
    }
}

// The STF sound player and CurrentSpeechTest are process-wide. A second hosted
// acoustic test must not silently take over while the first is running/stopping.
internal sealed class ScreeningAudioLease : IDisposable
{
    private static int held;
    private bool released;
    private readonly STFN.Core.Audio.AudioSettings settings;
    private readonly bool? previousAllowDefaultOutput;
    private ScreeningAudioLease(AudioSystemSpecification transducer)
    {
        settings = transducer.ParentAudioApiSettings;
        previousAllowDefaultOutput = settings.AllowDefaultOutputDevice;
        // The Android player's existing device monitor raises FatalPlayerError
        // when this is false. Never let a disconnected USB output fall back to
        // the tablet speaker under the headphones' calibration.
        settings.AllowDefaultOutputDevice = false;
    }
    public static ScreeningAudioLease Acquire(AudioSystemSpecification transducer)
    {
        if (Interlocked.CompareExchange(ref held, 1, 0) != 0)
            throw new InvalidOperationException("Ett annat ljudtest pågår. Avsluta det innan ett nytt test startas.");
        try { return new ScreeningAudioLease(transducer); }
        catch { Volatile.Write(ref held, 0); throw; }
    }
    public void Dispose()
    {
        if (released) return;
        released = true;
        settings.AllowDefaultOutputDevice = previousAllowDefaultOutput;
        Volatile.Write(ref held, 0);
    }
}


public sealed class ScreeningTransducerSelectionRequiredException(IReadOnlyList<AudioSystemSpecification> choices)
    : InvalidOperationException("Välj vilken konfigurerad hörlursutgång som ska användas för screeningen.")
{
    public IReadOnlyList<AudioSystemSpecification> Choices { get; } = choices;
}
