' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core
Imports STFN.Core.Audio.SoundScene
Imports STFN.Core.SipTest
Imports STFN.Core.Utils


Public Class AdaptiveSip_TsfcStudy
    Inherits SipBaseSpeechTest

    Public Overrides ReadOnly Property FilePathRepresentation As String = "SiP_Tsfc"

    Private TestSectionTripletCount As Integer = 40

    Private StartAdaptiveLevel As Double = 0

    Private MaximumTestTrials As Integer = 200

    Private MinPNR As Double = -20
    Private MaxPNR As Double = 15
    Private IgnoreNegativeCoordinates As Boolean = False

    Private TestProtocols As Dictionary(Of String, TestProtocol)
    Private TestProtocolCompleted As Dictionary(Of String, Boolean)
    Private AdaptiveLevelHistory As Dictionary(Of String, List(Of Double))



    Public IsTSFC As Boolean = False


    'Setting the sound source locations
    Private SipTargetStimulusLocations As SoundSourceLocation() = {New SoundSourceLocation With {.HorizontalAzimuth = 0, .Distance = 3}}
    'Private SipMaskerLocations As SoundSourceLocation() = {
    '        New SoundSourceLocation With {.HorizontalAzimuth = -30, .Distance = 3},
    '        New SoundSourceLocation With {.HorizontalAzimuth = 30, .Distance = 3}}
    Private SipMaskerLocations As SoundSourceLocation() = {
            New SoundSourceLocation With {.HorizontalAzimuth = 0, .Distance = 3}}

    Private SipBackgroundLocations As SoundSourceLocation() = {
            New SoundSourceLocation With {.HorizontalAzimuth = 60, .Distance = 3},
            New SoundSourceLocation With {.HorizontalAzimuth = -60, .Distance = 3},
            New SoundSourceLocation With {.HorizontalAzimuth = 150, .Distance = 3},
            New SoundSourceLocation With {.HorizontalAzimuth = -150, .Distance = 3}}


    Public Overrides Property TestProtocol As TestProtocol
        Get
            Return GetCurrentTestProtocol()
        End Get
        Set(value As TestProtocol)
            'Ignores any calls here since TestProtocol should never be set
            Throw New NotImplementedException("Test protocol should not be set in this way is the adaptive TSFC test")
            OnPropertyChanged()
        End Set
    End Property

    Private Function GetCurrentTestProtocol() As TestProtocol

        'Returns Nothing if there is no current trial
        If CurrentTestTrial Is Nothing Then Return Nothing

        'Try to get the test protocol from the current trial, based on the TWG PrimaryStringRepresentation
        If TestProtocols Is Nothing Then Return Nothing
        If TestProtocols.ContainsKey(CurrentTestTrial.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation) = False Then Return Nothing

        'Gets the test protocol
        Dim CurrentTestProtocol = TestProtocols(CurrentTestTrial.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation)

        'If failed to get the test protocol
        If CurrentTestProtocol Is Nothing Then Return Nothing

        Return CurrentTestProtocol

    End Function


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

        DirectionalSimulationSet = "OL-HEAD_BuK_ECEbl - 48kHz"
        'DirectionalSimulationSet = "ARC - Harcellen - HATS - SiP - HR"
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

        TestProtocols = New Dictionary(Of String, TestProtocol)
        AdaptiveLevelHistory = New Dictionary(Of String, List(Of Double))

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

            'Adding test section trials
            For i = 0 To TestSectionTripletCount - 1

                'Getting a vector of test word indices to draw
                Dim RandomIndices = STFN.Core.DSP.SampleWithoutReplacement(TestWords.Count, 0, TestWords.Count)

                For Each RandomIndex In RandomIndices

                    'Creating the trial
                    Dim NewTestTrial As New SipTrial(NewTestUnit, TestWords(RandomIndex), SelectedMediaSets.First, SoundPropagationType,
                                                 SipTargetStimulusLocations, SipMaskerLocations, SipBackgroundLocations, CurrentSipTestMeasurement.Randomizer)

                    'Storing IsTSFC in every trial
                    NewTestTrial.IsTSFCTrial = IsTSFC

                    'Adding the trial
                    NewTestUnit.PlannedTrials.Add(NewTestTrial)

                Next
            Next

            'Adding the test unit
            CurrentSipTestMeasurement.TestUnits.Add(NewTestUnit)

            'Also creating test protocols
            'Setting test protocol, including estimated slope and target score
            Dim NewTestProtocol = New STFN.Core.BrandKollmeier2002_TestProtocol
            NewTestProtocol.Slope = 0.034
            If IgnoreNegativeCoordinates = True Then
                NewTestProtocol.TargetScore = 0.5
            Else
                NewTestProtocol.TargetScore = 2 / 3
            End If
            NewTestProtocol.EnsureSentenceTest = False
            'Storing the TWG.PrimaryStringRepresentation in the test protocol's TargetStimulusSet
            NewTestProtocol.TargetStimulusSet = TWG.PrimaryStringRepresentation

            TestProtocols.Add(TWG.PrimaryStringRepresentation, NewTestProtocol)

            'Also adds the key to AdaptiveLevelHistory 
            AdaptiveLevelHistory.Add(TWG.PrimaryStringRepresentation, New List(Of Double))

        Next

        'Adding test test trials. Interleaving trials across test units: for each trial position, adding one trial per unit in random order
        'Adding the trials to CurrentSipTestMeasurement PlannedTrials (from which they can be drawn during testing)
        For testUnitTrialIndex = 0 To CurrentSipTestMeasurement.TestUnits(0).PlannedTrials.Count - 1

            'Getting a vector of test unit indices to draw
            Dim RandomIndices = STFN.Core.DSP.SampleWithoutReplacement(CurrentSipTestMeasurement.TestUnits.Count, 0, CurrentSipTestMeasurement.TestUnits.Count)

            'Adding to CurrentSipTestMeasurement.PlannedTrials in random order, each testUnitTrialIndex at a time
            For testUnitIndex = 0 To RandomIndices.Count - 1
                CurrentSipTestMeasurement.PlannedTrials.Add(CurrentSipTestMeasurement.TestUnits(RandomIndices(testUnitIndex)).PlannedTrials(testUnitTrialIndex))
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

            If CurrentTestTrial.IsTSFCTrial = False Then

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

            Else

                Dim ResponseList = TryCast(e.Box, SortedList(Of String, Double))
                If ResponseList IsNot Nothing Then
                    ' Use responseList here

                    'Getting the coordinate for the correct test word
                    For Each ResponseCandidate In ResponseList
                        Dim CorrectResponse = CurrentTestTrial.SpeechMaterialComponent.GetCategoricalVariableValue("Spelling")
                        If CorrectResponse = ResponseCandidate.Key Then

                            If IgnoreNegativeCoordinates = True Then

                                'The incoming coordinate (for the correct response alternative) is here streched linearly from the range 1/3 - 1 to 0 - 1
                                Dim BarycentricCoordinate As Double = (3 * ResponseCandidate.Value - 1) / 2

                                'And clamped to the range 0 - 1
                                BarycentricCoordinate = Math.Clamp(BarycentricCoordinate, 0, 1)

                                'And storing it in the trial
                                CurrentTestTrial.GradedResponse = BarycentricCoordinate

                            Else

                                'Stores the barycentric coordinate as it is (but clamped to the range 0 - 1, just in case it should be slightly off this range - which it should never be, but just in case something went wrong with the limits in the TSFC-GUI) 
                                CurrentTestTrial.GradedResponse = Math.Clamp(ResponseCandidate.Value, 0, 1)

                            End If

                            'Exiting the loop
                            Exit For

                        End If
                    Next

                Else
                    Throw New Exception("This is a bug!")
                End If
            End If

            'Moving to trial history
            CurrentSipTestMeasurement.MoveTrialToHistory(CurrentTestTrial)

            'Taking a dump of the SpeechTest
            CurrentTestTrial.SpeechTestPropertyDump = Logging.ListObjectPropertyValues(Me.GetType, Me)


        Else
            'Nothing to correct (this should be the start of a new test)
            'Playing initial sound, and premixing trials
            InitiateTestByPlayingSound()

        End If

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

        'Initializing the test protocol of this test unit
        If EvaluationTrials.Count = 0 Then
            Dim TestUnitPlannedTrials = DirectCast(CurrentTestTrial, SipTrial).ParentTestUnit.PlannedTrials
            TestProtocols(TestUnitPlannedTrials.First.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation).InitializeProtocol(
            New TestProtocol.NextTaskInstruction With {.AdaptiveValue = StartAdaptiveLevel, .TestLength = TestUnitPlannedTrials.Count})

            'We are in the ballpark stage of this test unit
            ProtocolReply = New TestProtocol.NextTaskInstruction()
            ProtocolReply.Decision = SpeechTestReplies.GotoNextTrial
            ProtocolReply.AdaptiveValue = StartAdaptiveLevel
            ProtocolReply.AdaptiveStepSize = 0 'TODO. Check if this value is relevant!

        Else

            'Using the selected test protocol
            'Determining if the level should be update. 
            Dim AdaptiveEvaluationLength As Integer = 1
            If EvaluationTrials.Last.IsTSFCTrial = False Then
                'TODO Consider if this should be 3 in the MAFC test
                AdaptiveEvaluationLength = 1
            End If

            If (EvaluationTrials.Count - AdaptiveEvaluationLength) Mod AdaptiveEvaluationLength = 0 Then

                'Setting EvaluationTrialCount so that EvaluationTrials.GetObservedScore returns the average of the three last trials in the test unti
                EvaluationTrials.EvaluationTrialCount = AdaptiveEvaluationLength
                ProtocolReply = TestProtocols(EvaluationTrials.Last.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation).NewResponse(EvaluationTrials)

                'Adding the adaptive level to the AdaptiveLevelHistory 
                AdaptiveLevelHistory(EvaluationTrials.Last.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation).Add(ProtocolReply.AdaptiveValue)

                'Stopping after X reversals
                If Math.Abs(ProtocolReply.AdaptiveReversalCount.Value) > 15 Then

                    'Noting the test protocol / test unit as completed
                    TestProtocolCompleted(EvaluationTrials.Last.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation) = True

                    If AllProtocolsCompleted() Then
                        Return SpeechTestReplies.TestIsCompleted
                    Else
                        Return SpeechTestReplies.GotoNextTrial
                    End If

                End If

                'Or stopping when the adaptive levels plateau
                Dim AdaptiveLevelStepStoppingCriteriumLength As Integer = 5
                Dim EvaluationList = AdaptiveLevelHistory(EvaluationTrials.Last.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation)
                If EvaluationList.Count > AdaptiveLevelStepStoppingCriteriumLength Then
                    Dim LastLevelSteps = EvaluationList.GetRange(EvaluationList.Count - AdaptiveLevelStepStoppingCriteriumLength, AdaptiveLevelStepStoppingCriteriumLength)

                    'Stopping when the adaptive level range falls under 1 dB  
                    If Math.Abs(LastLevelSteps.Max - LastLevelSteps.Min) < 1 Then

                        'Noting the test protocol / test unit as completed
                        TestProtocolCompleted(EvaluationTrials.Last.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation) = True

                        If AllProtocolsCompleted() Then
                            Return SpeechTestReplies.TestIsCompleted
                        Else
                            Return SpeechTestReplies.GotoNextTrial
                        End If

                    End If
                End If

                'Or if the trial length limit is reached
                If EvaluationTrials.Count > MaximumTestTrials Then

                    'Noting the test protocol / test unit as completed
                    TestProtocolCompleted(EvaluationTrials.Last.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation) = True

                    If AllProtocolsCompleted() Then
                        Return SpeechTestReplies.TestIsCompleted
                    Else
                        Return SpeechTestReplies.GotoNextTrial
                    End If

                End If

            Else
                'Simply reusing the level from the previous trial
                ProtocolReply = New TestProtocol.NextTaskInstruction()
                ProtocolReply.Decision = SpeechTestReplies.GotoNextTrial
                ProtocolReply.AdaptiveValue = DirectCast(EvaluationTrials.Last, SipTrial).PNR
                ProtocolReply.AdaptiveStepSize = 0
            End If

        End If


        'Clamping the adaptive value
        ProtocolReply.AdaptiveValue = Math.Clamp(ProtocolReply.AdaptiveValue.Value, MinPNR, MaxPNR)

        'Preparing next trial if needed
        If ProtocolReply.Decision = SpeechTestReplies.GotoNextTrial Then
            PrepareNextTrial(ProtocolReply)
        End If

        Return ProtocolReply.Decision

    End Function

    Public Function AllProtocolsCompleted() As Boolean

        For Each protocol In TestProtocolCompleted
            If protocol.Value = False Then Return False
        Next
        Return True

    End Function

    Protected Overrides Sub PrepareNextTrial(ByVal NextTaskInstruction As TestProtocol.NextTaskInstruction)


        CurrentTestTrial.TestStage = NextTaskInstruction.TestStage

        'Setting levels of the SiP trial
        DirectCast(CurrentTestTrial, SipTrial).SetLevels(ReferenceLevel, NextTaskInstruction.AdaptiveValue)

        'String the adaptive protocol value in the trial (this is needed in the test protocol class)
        CurrentTestTrial.AdaptiveProtocolValue = DirectCast(CurrentTestTrial, SipTrial).PNR

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

    Private TwgProtocolCombinations As New SortedList(Of String, String)

    Public Overrides Function GetResultStringForGui() As String

        'This is not used

        Return ""

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