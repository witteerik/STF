' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core
Imports STFN.Core.SipTest
Imports STFN.Core.Audio.SoundScene
Imports STFN.Core.Utils


Public Class AdaptiveSip_TsfcStudy
    Inherits SipBaseSpeechTest

    Public Overrides ReadOnly Property FilePathRepresentation As String = "SiP_Tsfc"

    Private BallparkLength As Integer = 4
    Private TestSectionTripletCount As Integer = 40

    Private StartAdaptiveLevel As Double = 0
    Private BallparkStepSize As Double = 5

    Private MinPNR As Double = -10
    Private MaxPNR As Double = 15


    'Private IsTSFC As Boolean = False


    'Setting the sound source locations
    Private SipTargetStimulusLocations As SoundSourceLocation() = {New SoundSourceLocation With {.HorizontalAzimuth = 0, .Distance = 1.45}}
    Private SipMaskerLocations As SoundSourceLocation() = {
            New SoundSourceLocation With {.HorizontalAzimuth = -30, .Distance = 1.45},
            New SoundSourceLocation With {.HorizontalAzimuth = 30, .Distance = 1.45}}
    Private SipBackgroundLocations As SoundSourceLocation() = {
            New SoundSourceLocation With {.HorizontalAzimuth = -60, .Distance = 1.45},
            New SoundSourceLocation With {.HorizontalAzimuth = -60, .Distance = 1.45},
            New SoundSourceLocation With {.HorizontalAzimuth = 150, .Distance = 1.45},
            New SoundSourceLocation With {.HorizontalAzimuth = -150, .Distance = 1.45}}


    Public Sub New(ByVal SpeechMaterialName As String)
        MyBase.New(SpeechMaterialName)
        ApplyTestSpecificSettings()
    End Sub

    Public Shadows Sub ApplyTestSpecificSettings()

        TesterInstructions = "--  SiP  --" & vbCrLf & vbCrLf &
             "För detta test behövs inga inställningar." & vbCrLf & vbCrLf &
             "1. Informera deltagaren om hur testet går till." & vbCrLf &
             "2. Vänd skärmen till deltagaren. Be sedan deltagaren klicka på start för att starta testet."

        ParticipantInstructions = "--  SiP  --" & vbCrLf & vbCrLf &
             "Lyssna efter enstaviga ord som uttalas i en stadsmiljö." & vbCrLf &
             "Välj från svarsalternativen på skärmen det ord du uppfattade. " & vbCrLf &
             "Gissa om du är osäker." & vbCrLf &
             "Efter varje ord har du maximalt " & MaximumResponseTime & " sekunder på dig att ange ditt svar." & vbCrLf &
             "Om svarsalternativen blinkar i röd färg har du inte svarat i tid." & vbCrLf &
             "Testet består av totalt 30 ord, som blir svårare och svårare ju längre testet går." & vbCrLf &
             "Starta testet genom att klicka på knappen 'Start'"

        ParticipantInstructionsButtonText = "Deltagarinstruktion"

        'ShowGuiChoice_SoundFieldSimulation = True

        DirectionalSimulationSet = "ARC - Harcellen - HATS - SiP - HR"
        'DirectionalSimulationSet = "ARC - Harcellen - HATS - SiP"
        ' DirectionalSimulationSet = "ARC - Harcellen - HATS 256 - 48kHz"

        PopulateSoundSourceLocationCandidates()
        SimulatedSoundField = True

        ShowGuiChoice_TargetLocations = False
        ShowGuiChoice_MaskerLocations = False
        ShowGuiChoice_BackgroundNonSpeechLocations = False
        ShowGuiChoice_BackgroundSpeechLocations = False

        SupportsManualPausing = True

        GuiResultType = GuiResultTypes.VisualResults

    End Sub

    Public Overrides ReadOnly Property ShowGuiChoice_TargetSNRLevel As Boolean = False


    Private PresetName As String = "QuickSiP"

    'Private ResultsSummary As SortedList(Of Double, Tuple(Of SipTestList, Double))

    Public Overrides Function InitializeCurrentTest() As Tuple(Of Boolean, String)

        'Transducer = AvaliableTransducers(1)

        CurrentSipTestMeasurement = New SipMeasurement(CurrentParticipantID, SpeechMaterial.ParentTestSpecification, SiPTestProcedure.AdaptiveTypes.Fixed, SelectedTestparadigm)

        CurrentSipTestMeasurement.ExportTrialSoundFiles = False

        'Setting SimulatedSoundField to depend on the type of transducer selected
        If Transducer.IsHeadphones = True Then
            'SiP should only be used with simulated sound field in headphones
            SimulatedSoundField = True
        Else
            'Or without simulated sound field (of course) in sound field presentation.
            SimulatedSoundField = False
        End If

        If SimulatedSoundField = True Then
            SelectedSoundPropagationType = SoundPropagationTypes.SimulatedSoundField

            'Dim AvailableSets = DirectionalSimulator.GetAvailableDirectionalSimulationSets(SelectedTransducer)
            'DirectionalSimulator.TrySetSelectedDirectionalSimulationSet(AvailableSets(1), SelectedTransducer, False)

            Dim FoundDirSimulator As Boolean = Globals.StfBase.DirectionalSimulator.TrySetSelectedDirectionalSimulationSet(DirectionalSimulationSet, Transducer, False)
            If FoundDirSimulator = False Then
                Return New Tuple(Of Boolean, String)(False, "Unable to find the directional simulation set " & DirectionalSimulationSet)
            End If

        Else
            SelectedSoundPropagationType = SoundPropagationTypes.PointSpeakers
        End If

        'Setting up test trials to run
        PlanSiPTrials(SelectedSoundPropagationType, RandomSeed)

        If CurrentSipTestMeasurement.HasSimulatedSoundFieldTrials = True And Globals.StfBase.DirectionalSimulator.SelectedDirectionalSimulationSetName = "" Then
            Return New Tuple(Of Boolean, String)(False, "The measurement requires a directional simulation set to be selected!")
        End If

        'Setting test protocol, including estimated slope and target score
        TestProtocol = New STFN.Core.BrandKollmeier2002_TestProtocol
        DirectCast(TestProtocol, BrandKollmeier2002_TestProtocol).Slope = 0.034 'TODO: Check this value
        DirectCast(TestProtocol, BrandKollmeier2002_TestProtocol).TargetScore = 2 / 3
        DirectCast(TestProtocol, BrandKollmeier2002_TestProtocol).EnsureSentenceTest = False

        'TODO: each test unit needs its own TestProtocol!
        TestProtocol.InitializeProtocol(New TestProtocol.NextTaskInstruction With {.AdaptiveValue = 0, .TestLength = CurrentSipTestMeasurement.TestUnits(0).PlannedTrials.Count - BallparkLength})

        Return New Tuple(Of Boolean, String)(True, "")

    End Function


    Public Overrides Function GetScorePerLevel() As Tuple(Of String, SortedList(Of Double, Double))

        Throw New NotImplementedException

        'Dim PnrScoresList = GetPnrScores()
        'Dim OutputList = New SortedList(Of Double, Double)

        'For Each kvp In PnrScoresList
        '    OutputList.Add(kvp.Key, kvp.Value.Item2)
        'Next

        'Return New Tuple(Of String, SortedList(Of Double, Double))("PNR (dB)", OutputList)

    End Function

    'Public Function GetPnrScores() As SortedList(Of Double, Tuple(Of SipTestList, Double))

    '    Dim ResultList As New List(Of Tuple(Of SipTestList, Double)) ' QuickSipList, MeanScore

    '    For i = 0 To SipTestLists.Count - 1

    '        Dim CurrentSipTestList = SipTestLists(i)
    '        Dim CurrentScoresList As New List(Of Double)

    '        For Each Trial In CurrentSipTestMeasurement.ObservedTrials
    '            If Trial.MediaSet Is CurrentSipTestList.MediaSet And
    '                    Trial.PNR = CurrentSipTestList.PNR And
    '                    Trial.SpeechMaterialComponent.ParentComponent Is CurrentSipTestList.SMC Then

    '                If Trial.IsCorrect = True Then
    '                    CurrentScoresList.Add(1)
    '                Else
    '                    CurrentScoresList.Add(0)
    '                End If

    '            End If
    '        Next

    '        Dim AverageScore As Double = Double.NaN
    '        If CurrentScoresList.Count > 0 Then
    '            AverageScore = CurrentScoresList.Average
    '        End If

    '        ResultList.Add(New Tuple(Of SipTestList, Double)(CurrentSipTestList, AverageScore))
    '    Next

    '    Dim PnrSortedList As New SortedList(Of Double, Tuple(Of SipTestList, Double))
    '    For Each Result In ResultList
    '        PnrSortedList.Add(Result.Item1.PNR, Result)
    '    Next

    '    Return PnrSortedList

    'End Function

    Public Shared Function GetTestPnrs() As List(Of Double)

        'TODO: Adjust for constant stimuli test
        Dim PNRs As New List(Of Double) From {-5, 0, 5}
        Return PNRs

    End Function

    Private Sub PlanSiPTrials(ByVal SoundPropagationType As SoundPropagationTypes, Optional ByVal RandomSeed As Integer? = Nothing)

        Dim AllMediaSets As List(Of MediaSet) = AvailableMediasets

        Dim SelectedMediaSets As New List(Of MediaSet)
        Dim IncludedMediaSetNames As New List(Of String) From {"City-Talker1-RVE"}
        For Each AvailableMediaSet In AllMediaSets
            If IncludedMediaSetNames.Contains(AvailableMediaSet.MediaSetName) Then
                SelectedMediaSets.Add(AvailableMediaSet)
            End If
        Next

        'Creating a new random if seed is supplied
        If RandomSeed.HasValue Then CurrentSipTestMeasurement.Randomizer = New Random(RandomSeed)

        'Getting the preset
        Dim Preset = CurrentSipTestMeasurement.ParentTestSpecification.SpeechMaterial.Presets.GetPreset(PresetName).Members

        'Adding trials into TestUnits. One TestUnit for each test word group.
        For Each TWG In Preset

            'Creating a new test unit
            Dim NewTestUnit = New SiPTestUnit(CurrentSipTestMeasurement, TWG.PrimaryStringRepresentation)

            'Getting the test words (i.e. sentence level components)
            Dim TestWords = TWG.GetChildren

            'Planning a maximum number of trials
            Dim MaxTrialCount As Integer = BallparkLength + 3 * TestSectionTripletCount

            'Adding ballpark trials
            'TODO: Implement sequence restrictions
            For i = 0 To BallparkLength - 1
                'Drawing test words at random
                Dim RandomIndex As Integer = CurrentSipTestMeasurement.Randomizer.Next(0, TestWords.Count)

                'Creating the trial
                Dim NewTestTrial As New SipTrial(NewTestUnit, TestWords(RandomIndex), SelectedMediaSets.First, SoundPropagationType,
                                                 SipTargetStimulusLocations, SipMaskerLocations, SipBackgroundLocations, CurrentSipTestMeasurement.Randomizer)

                'Adding the trial
                NewTestUnit.PlannedTrials.Add(NewTestTrial)

            Next

            'Adding test section trials
            For i = 0 To TestSectionTripletCount - 1

                'Getting a vector of test word indices to draw
                Dim RandomIndices = STFN.Core.DSP.SampleWithoutReplacement(TestWords.Count, 0, TestWords.Count)

                For Each RandomIndex In RandomIndices

                    'Creating the trial
                    Dim NewTestTrial As New SipTrial(NewTestUnit, TestWords(RandomIndex), SelectedMediaSets.First, SoundPropagationType,
                                                 SipTargetStimulusLocations, SipMaskerLocations, SipBackgroundLocations, CurrentSipTestMeasurement.Randomizer)

                    'Adding the trial
                    NewTestUnit.PlannedTrials.Add(NewTestTrial)

                Next
            Next

            'Adding the test unit
            CurrentSipTestMeasurement.TestUnits.Add(NewTestUnit)

        Next


        'Adding the trials to CurrentSipTestMeasurement PlannedTrials (from which they can be drawn during testing)
        For ui = 0 To CurrentSipTestMeasurement.TestUnits.Count - 1
            Dim Unit As SiPTestUnit = CurrentSipTestMeasurement.TestUnits(ui)
            For Each Trial In Unit.PlannedTrials
                CurrentSipTestMeasurement.PlannedTrials.Add(Trial)
            Next
        Next

    End Sub

    Protected Overrides Sub InitiateTestByPlayingSound()

        'Sets the measurement datetime
        CurrentSipTestMeasurement.MeasurementDateTime = DateTime.Now

        'Cretaing a context sound without any test stimulus, that runs for approx TestSetup.PretestSoundDuration seconds, using audio from the first selected MediaSet
        Dim AllMediaSets As List(Of MediaSet) = AvailableMediasets

        Dim SelectedMediaSets As New List(Of MediaSet)
        Dim IncludedMediaSetNames As New List(Of String) From {"City-Talker1-RVE"}
        For Each AvailableMediaSet In AllMediaSets
            If IncludedMediaSetNames.Contains(AvailableMediaSet.MediaSetName) Then
                SelectedMediaSets.Add(AvailableMediaSet)
            End If
        Next

        Dim TestSound As STFN.Core.Audio.Sound = CreateInitialSound(SelectedMediaSets.First)

        'Plays sound
        Globals.StfBase.SoundPlayer.SwapOutputSounds(TestSound)

    End Sub


    Public Function CreateInitialSound(ByRef SelectedMediaSet As MediaSet, Optional ByVal Duration As Double? = Nothing) As STFN.Core.Audio.Sound

        Try

            'Setting up the SiP-trial sound mix
            Dim MixStopWatch As New Stopwatch
            MixStopWatch.Start()

            'Sets a List of SoundSceneItem in which to put the sounds to mix
            Dim ItemList = New List(Of SoundSceneItem)

            Dim SoundWaveFormat As STFN.Core.Audio.Formats.WaveFormat = Nothing

            'Getting a background non-speech sound
            Dim BackgroundNonSpeech_Sound As STFN.Core.Audio.Sound = SpeechMaterial.GetBackgroundNonspeechSound(SelectedMediaSet, 0)

            'Stores the sample rate and the wave format
            Dim CurrentSampleRate As Integer = BackgroundNonSpeech_Sound.WaveFormat.SampleRate
            SoundWaveFormat = BackgroundNonSpeech_Sound.WaveFormat

            'Sets a total pretest sound length
            Dim TrialSoundLength As Integer
            If Duration.HasValue Then
                TrialSoundLength = Duration * SoundWaveFormat.SampleRate
            Else
                TrialSoundLength = (PretestSoundDuration + 4) * CurrentSampleRate 'Adds 4 seconds to allow for potential delay caused by the mixing time of the first test trial sounds
            End If

            'Copies copies random sections of the background non-speech sound into two sounds
            Dim Background1 = BackgroundNonSpeech_Sound.CopySection(1, Randomizer.Next(0, BackgroundNonSpeech_Sound.WaveData.SampleData(1).Length - TrialSoundLength - 2), TrialSoundLength)
            Dim Background2 = BackgroundNonSpeech_Sound.CopySection(1, Randomizer.Next(0, BackgroundNonSpeech_Sound.WaveData.SampleData(1).Length - TrialSoundLength - 2), TrialSoundLength)
            Dim Background3 = BackgroundNonSpeech_Sound.CopySection(1, Randomizer.Next(0, BackgroundNonSpeech_Sound.WaveData.SampleData(1).Length - TrialSoundLength - 2), TrialSoundLength)
            Dim Background4 = BackgroundNonSpeech_Sound.CopySection(1, Randomizer.Next(0, BackgroundNonSpeech_Sound.WaveData.SampleData(1).Length - TrialSoundLength - 2), TrialSoundLength)

            'Sets up fading specifications for the background signals
            Dim FadeSpecs_Background = New List(Of DSP.FadeSpecifications)
            FadeSpecs_Background.Add(New DSP.FadeSpecifications(Nothing, 0, 0, CurrentSampleRate * 1))
            FadeSpecs_Background.Add(New DSP.FadeSpecifications(0, Nothing, -CurrentSampleRate * 0.01))

            'Adds the background (non-speech) signals, with fade, duck and location specifications
            Dim LevelGroup As Integer = 1 ' The level group value is used to set the added sound level of items sharing the same (arbitrary) LevelGroup value to the indicated sound level. (Thus, the sounds with the same LevelGroup value are measured together.)

            ItemList.Add(New SoundSceneItem(Background1, 1, SelectedMediaSet.BackgroundNonspeechRealisticLevel, LevelGroup,
                                            SipBackgroundLocations(0), SoundSceneItem.SoundSceneItemRoles.BackgroundNonspeech, 0,,,, FadeSpecs_Background))
            ItemList.Add(New SoundSceneItem(Background2, 1, SelectedMediaSet.BackgroundNonspeechRealisticLevel, LevelGroup,
                                            SipBackgroundLocations(1), SoundSceneItem.SoundSceneItemRoles.BackgroundNonspeech, 0,,,, FadeSpecs_Background))
            ItemList.Add(New SoundSceneItem(Background3, 1, SelectedMediaSet.BackgroundNonspeechRealisticLevel, LevelGroup,
                                            SipBackgroundLocations(2), SoundSceneItem.SoundSceneItemRoles.BackgroundNonspeech, 0,,,, FadeSpecs_Background))
            ItemList.Add(New SoundSceneItem(Background4, 1, SelectedMediaSet.BackgroundNonspeechRealisticLevel, LevelGroup,
                                            SipBackgroundLocations(3), SoundSceneItem.SoundSceneItemRoles.BackgroundNonspeech, 0,,,, FadeSpecs_Background))
            LevelGroup += 1

            MixStopWatch.Stop()
            If LogToConsole = True Then Console.WriteLine("Prepared sounds in " & MixStopWatch.ElapsedMilliseconds & " ms.")
            MixStopWatch.Restart()

            'Creating the mix by calling CreateSoundScene of the current Mixer
            Dim MixedInitialSound As STFN.Core.Audio.Sound = Transducer.Mixer.CreateSoundScene(ItemList, False, False, SelectedSoundPropagationType, Transducer.LimiterThreshold)

            If LogToConsole = True Then Console.WriteLine("Mixed sound in " & MixStopWatch.ElapsedMilliseconds & " ms.")

            Return MixedInitialSound

        Catch ex As Exception
            Logging.SendInfoToLog(ex.ToString, "ExceptionsDuringTesting")
            Return Nothing
        End Try

    End Function

    Public Overrides Function GetSpeechTestReply(sender As Object, e As SpeechTestInputEventArgs) As SpeechTestReplies

        If e IsNot Nothing Then

            'This is an incoming test trial response

            'Corrects the trial response, based on the given response

            'Resets the CurrentTestTrial.ScoreList
            'And also storing SiP-test type data
            CurrentTestTrial.ScoreList = New List(Of Integer)
            Select Case e.LinguisticResponses(0)
                Case CurrentTestTrial.SpeechMaterialComponent.GetCategoricalVariableValue("Spelling")
                    CurrentTestTrial.ScoreList.Add(1)
                    DirectCast(CurrentTestTrial, SipTrial).Result = SipTrial.PossibleResults.Correct
                    CurrentTestTrial.IsCorrect = True

                Case ""
                    CurrentTestTrial.ScoreList.Add(0)
                    DirectCast(CurrentTestTrial, SipTrial).Result = SipTrial.PossibleResults.Missing

                    'Randomizing IsCorrect with a 1/3 chance for True
                    Dim ChanceList As New List(Of Boolean) From {True, False, False}
                    Dim RandomIndex As Integer = Randomizer.Next(ChanceList.Count)
                    CurrentTestTrial.IsCorrect = ChanceList(RandomIndex)

                Case Else
                    CurrentTestTrial.ScoreList.Add(0)
                    DirectCast(CurrentTestTrial, SipTrial).Result = SipTrial.PossibleResults.Incorrect
                    CurrentTestTrial.IsCorrect = False

            End Select

            DirectCast(CurrentTestTrial, SipTrial).Response = e.LinguisticResponses(0)

            'Moving to trial history
            CurrentSipTestMeasurement.MoveTrialToHistory(CurrentTestTrial)

            'Taking a dump of the SpeechTest
            CurrentTestTrial.SpeechTestPropertyDump = Logging.ListObjectPropertyValues(Me.GetType, Me)


        Else
            'Nothing to correct (this should be the start of a new test)
            'Playing initial sound, and premixing trials
            InitiateTestByPlayingSound()

        End If

        'TODO: We must store the responses and response times!!!

        'Returning if no more trials to present
        If CurrentSipTestMeasurement.PlannedTrials.Count = 0 Then
            'Test is completed
            Return SpeechTestReplies.TestIsCompleted
        End If

        'Preparing the next trial
        CurrentTestTrial = CurrentSipTestMeasurement.GetNextTrial()

        'Calculating the speech level
        Dim ProtocolReply As TestProtocol.NextTaskInstruction = Nothing

        Dim CurrentSipTrials = DirectCast(CurrentTestTrial, SipTrial).ParentTestUnit.ObservedTrials

        Dim EvaluationTrials As New TestTrialCollection
        For Each Trial In CurrentSipTrials
            EvaluationTrials.Add(Trial)
        Next

        If EvaluationTrials.Count < BallparkLength Then

            'We are in the ballpark stage of this test unit
            ProtocolReply = New TestProtocol.NextTaskInstruction()
            ProtocolReply.Decision = SpeechTestReplies.GotoNextTrial

            If EvaluationTrials.Count = 0 Then
                'We get the start level
                ProtocolReply.AdaptiveValue = StartAdaptiveLevel
                ProtocolReply.AdaptiveStepSize = BallparkStepSize
            Else

                If EvaluationTrials.Last.IsTSFCTrial = True Then

                    'Implementing- the TSFC logic for the ballpark stage
                    'Determining step size
                    Dim StepSize As Double = BallparkStepSize * (3 * EvaluationTrials.Last.GradedResponse - 1) / 2 - (BallparkStepSize / 2)

                    'Storing the step size
                    ProtocolReply.AdaptiveStepSize = StepSize

                    'Calculating the new level
                    ProtocolReply.AdaptiveValue = DirectCast(EvaluationTrials.Last, SipTrial).PNR + ProtocolReply.AdaptiveStepSize

                Else

                    'Setting Adaptive level based on the observed binary responses in the ballpark stage of this test word group / TestUnit
                    Dim LastTrial As SipTrial = EvaluationTrials.Last
                    ProtocolReply.AdaptiveStepSize = BallparkStepSize

                    If LastTrial.IsCorrect = True Then
                        ProtocolReply.AdaptiveValue = LastTrial.PNR - ProtocolReply.AdaptiveStepSize
                    Else
                        ProtocolReply.AdaptiveValue = LastTrial.PNR + ProtocolReply.AdaptiveStepSize
                    End If

                End If
            End If

        Else
            'We are in the test stage of this test unit

            'Using the selected test protocol
            'Determining if the level should be update. 
            If EvaluationTrials.Count > BallparkLength And (EvaluationTrials.Count - (BallparkLength + 3)) Mod 3 = 0 Then

                'Getting the last three trials that occur at an update point after the 7th trial (i.e. the ballpark stage of 4 trials and the first test stage of 3 trials
                EvaluationTrials = EvaluationTrials.GetRange(EvaluationTrials.Count - 4, 3)

                'Setting EvaluationTrialCount so that EvaluationTrials.GetObservedScore returns the average of the three last trials in the test unti
                EvaluationTrials.EvaluationTrialCount = 3
                ProtocolReply = TestProtocol.NewResponse(EvaluationTrials)

            Else
                'Simply reusing the level from the previous trial
                ProtocolReply = New TestProtocol.NextTaskInstruction()
                ProtocolReply.Decision = SpeechTestReplies.GotoNextTrial
                ProtocolReply.AdaptiveValue = DirectCast(EvaluationTrials.Last, SipTrial).PNR
                ProtocolReply.AdaptiveStepSize = 0
            End If
        End If

        'Clamping the
        ProtocolReply.AdaptiveValue = Math.Clamp(ProtocolReply.AdaptiveValue.Value, MinPNR, MaxPNR)

        'Preparing next trial if needed
        If ProtocolReply.Decision = SpeechTestReplies.GotoNextTrial Then
            PrepareNextTrial(ProtocolReply)
        End If

        Return ProtocolReply.Decision

    End Function


    Protected Overrides Sub PrepareNextTrial(ByVal NextTaskInstruction As TestProtocol.NextTaskInstruction)


        CurrentTestTrial.TestStage = NextTaskInstruction.TestStage

        'Setting levels of the SiP trial
        DirectCast(CurrentTestTrial, SipTrial).SetLevels(ReferenceLevel, NextTaskInstruction.AdaptiveValue)

        CurrentTestTrial.Tasks = 1
        CurrentTestTrial.ResponseAlternativeSpellings = New List(Of List(Of SpeechTestResponseAlternative))
        Dim ResponseAlternatives As New List(Of SpeechTestResponseAlternative)

        'Adding the current word spelling as a response alternative
        ResponseAlternatives.Add(New SpeechTestResponseAlternative With {.Spelling = CurrentTestTrial.SpeechMaterialComponent.GetCategoricalVariableValue("Spelling"), .IsScoredItem = CurrentTestTrial.SpeechMaterialComponent.IsKeyComponent, .ParentTestTrial = CurrentTestTrial})

        'Picking random response alternatives from all available test words
        Dim AllContrastingWords = CurrentTestTrial.SpeechMaterialComponent.GetSiblingsExcludingSelf()
        For Each ContrastingWord In AllContrastingWords
            ResponseAlternatives.Add(New SpeechTestResponseAlternative With {.Spelling = ContrastingWord.GetCategoricalVariableValue("Spelling"), .IsScoredItem = ContrastingWord.IsKeyComponent, .ParentTestTrial = CurrentTestTrial})
        Next

        'Shuffling the order of response alternatives
        ResponseAlternatives = DSP.Shuffle(ResponseAlternatives, Randomizer).ToList

        'Adding the response alternatives
        CurrentTestTrial.ResponseAlternativeSpellings.Add(ResponseAlternatives)

        'Mixing trial sound
        DirectCast(CurrentTestTrial, SipTrial).MixSound(Transducer, MinimumStimulusOnsetTime, MaximumStimulusOnsetTime, Randomizer, TrialSoundMaxDuration, UseBackgroundSpeech)

        'Storing the LinguisticSoundStimulusStartTime and the LinguisticSoundStimulusDuration 
        CurrentTestTrial.LinguisticSoundStimulusStartTime = DirectCast(CurrentTestTrial, SipTrial).TestWordStartTime
        CurrentTestTrial.LinguisticSoundStimulusDuration = DirectCast(CurrentTestTrial, SipTrial).TestWordCompletedTime - CurrentTestTrial.LinguisticSoundStimulusStartTime
        CurrentTestTrial.MaximumResponseTime = MaximumResponseTime

        'And also the background level directly from the MediaSet (note that this is the level without any ducking applied!)
        BackgroundLevel = DirectCast(CurrentTestTrial, SipTrial).MediaSet.BackgroundNonspeechRealisticLevel

        'Setting visual que intervals
        Dim ShowVisualQueTimer_Interval As Double
        Dim HideVisualQueTimer_Interval As Double
        Dim ShowResponseAlternativePositions_Interval As Integer
        Dim ShowResponseAlternativesTimer_Interval As Double
        Dim MaxResponseTimeTimer_Interval As Double

        If UseVisualQue = True Then
            ShowVisualQueTimer_Interval = System.Math.Max(1, DirectCast(CurrentTestTrial, SipTrial).TestWordStartTime * 1000)
            HideVisualQueTimer_Interval = System.Math.Max(2, DirectCast(CurrentTestTrial, SipTrial).TestWordCompletedTime * 1000)
            ShowResponseAlternativesTimer_Interval = HideVisualQueTimer_Interval + 1000 * ResponseAlternativeDelay 'TestSetup.CurrentEnvironment.TestSoundMixerSettings.ResponseAlternativeDelay * 1000
            MaxResponseTimeTimer_Interval = System.Math.Max(1, ShowResponseAlternativesTimer_Interval + 1000 * MaximumResponseTime)  ' TestSetup.CurrentEnvironment.TestSoundMixerSettings.MaximumResponseTime * 1000
        Else
            ShowResponseAlternativePositions_Interval = ShowResponseAlternativePositionsTime * 1000
            ShowResponseAlternativesTimer_Interval = System.Math.Max(1, DirectCast(CurrentTestTrial, SipTrial).TestWordStartTime * 1000) + 1000 * ResponseAlternativeDelay
            MaxResponseTimeTimer_Interval = System.Math.Max(2, DirectCast(CurrentTestTrial, SipTrial).TestWordCompletedTime * 1000) + 1000 * MaximumResponseTime
        End If

        'Setting trial events
        CurrentTestTrial.TrialEventList = New List(Of ResponseViewEvent)
        CurrentTestTrial.TrialEventList.Add(New ResponseViewEvent With {.TickTime = 1, .Type = ResponseViewEvent.ResponseViewEventTypes.PlaySound})

        If UseVisualQue = False Then
            ' Test word alternatives on the sides are only supported when the visual que is not shown
            If ShowTestSide = True Then
                CurrentTestTrial.TrialEventList.Add(New ResponseViewEvent With {.TickTime = ShowResponseAlternativePositions_Interval, .Type = ResponseViewEvent.ResponseViewEventTypes.ShowResponseAlternativePositions})
            End If
        Else
            CurrentTestTrial.TrialEventList.Add(New ResponseViewEvent With {.TickTime = ShowVisualQueTimer_Interval, .Type = ResponseViewEvent.ResponseViewEventTypes.ShowVisualCue})
            CurrentTestTrial.TrialEventList.Add(New ResponseViewEvent With {.TickTime = HideVisualQueTimer_Interval, .Type = ResponseViewEvent.ResponseViewEventTypes.HideVisualCue})
        End If

        CurrentTestTrial.TrialEventList.Add(New ResponseViewEvent With {.TickTime = ShowResponseAlternativesTimer_Interval, .Type = ResponseViewEvent.ResponseViewEventTypes.ShowResponseAlternatives})
        CurrentTestTrial.TrialEventList.Add(New ResponseViewEvent With {.TickTime = MaxResponseTimeTimer_Interval, .Type = ResponseViewEvent.ResponseViewEventTypes.ShowResponseTimesOut})

    End Sub


    Public Overrides Function GetAverageScore(Optional IncludeTrialsFromEnd As Integer? = Nothing) As Double?

        'Getting all trials
        Dim ObservedTrials = GetObservedTestTrials().ToList

        'Returning Nothing if no trials have been observed
        If ObservedTrials.Count = 0 Then Return Nothing

        'Creating a list to store observed trials to include
        Dim TrialsToInclude As New List(Of TestTrial)

        If IncludeTrialsFromEnd.HasValue Then

            'Getting the IncludeTrialsFromEnd last test trials, or all if ObservedTrials is shorter than IncludeTrialsFromEnd
            If ObservedTrials.Count < IncludeTrialsFromEnd + 1 Then
                TrialsToInclude.AddRange(ObservedTrials)
            Else
                TrialsToInclude.AddRange(ObservedTrials.GetRange(ObservedTrials.Count - IncludeTrialsFromEnd, IncludeTrialsFromEnd))
            End If

        Else
            'Getting all trials
            TrialsToInclude.AddRange(ObservedTrials)
        End If

        'Calculating average score
        Dim ScoreList As New List(Of Integer)
        For Each Trial In TrialsToInclude

            'Getting all results
            If Trial.IsCorrect = True Then
                ScoreList.Add(1)
            Else
                ScoreList.Add(0)
            End If

        Next

        Return ScoreList.Average


    End Function


    Public Overrides Function GetResultStringForGui() As String

        Return "Not yet implemented"

        'Dim Output As New List(Of String)

        'Dim AverageScore As Double? = GetAverageScore()
        'If AverageScore.HasValue Then
        '    Output.Add("Overall score: " & DSP.Rounding(100 * GetAverageScore()) & "%")
        'End If

        'ResultsSummary = GetPnrScores()

        'If ResultsSummary IsNot Nothing Then
        '    Output.Add("Scores per PNR level:")
        '    Output.Add("PNR (dB)" & vbTab & "Score" & vbTab & "List")
        '    For Each kvp In ResultsSummary
        '        Output.Add(kvp.Value.Item1.PNR & vbTab & DSP.Rounding(100 * kvp.Value.Item2) & " %" & vbTab & kvp.Value.Item1.SMC.PrimaryStringRepresentation)
        '    Next
        'End If

        'Return String.Join(vbCrLf, Output)

    End Function

    ''' <summary>
    ''' This function should list the names of variables included SpeechTestDump of each test trial to be exported in the "selected-variables" export file.
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function GetSelectedExportVariables() As List(Of String)
        Return New List(Of String)
    End Function

    Public Overrides Function GetSubGroupResults() As List(Of Tuple(Of String, Double))

        Return New List(Of Tuple(Of String, Double))
        'Throw New NotImplementedException()
    End Function
End Class