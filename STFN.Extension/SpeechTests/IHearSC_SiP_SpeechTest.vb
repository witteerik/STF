﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core
Imports STFN.Core.SipTest
Imports STFN.Core.Audio
Imports STFN.Core.Audio.SoundScene
Imports STFN.Core.Utils

Public Class IHearSC_SiP_SpeechTest

    Inherits SipBaseSpeechTest

    Public Overrides ReadOnly Property FilePathRepresentation As String = "SipTest"

    Public Sub New(ByVal SpeechMaterialName As String)
        MyBase.New(SpeechMaterialName)
        ApplyTestSpecificSettings()
    End Sub

    Public Shadows Sub ApplyTestSpecificSettings()

        TesterInstructions = "För detta test behövs inga inställningar." & vbCrLf & vbCrLf &
                "1. Informera patienten om hur testet går till." & vbCrLf &
                "2. Vänd skärmen till patienten. Be sedan patienten klicka på start för att starta testet."


        ParticipantInstructions = "Patientens uppgift: " & vbCrLf & vbCrLf &
                " - Patienten startar testet genom att klicka på knappen 'Start'" & vbCrLf &
                " - Under testet ska patienten lyssna efter enstaviga ord i olika ljudmiljöer och efter varje ord ange på skärmen vilket ord hen uppfattade. " & vbCrLf &
                " - Patienten ska gissa om hen är osäker. Många ord är mycket svåra att höra!" & vbCrLf &
                " - Efter varje ord har patienten maximalt " & MaximumResponseTime & " sekunder på sig att ange sitt svar." & vbCrLf &
                " - Om svarsalternativen blinkar i röd färg har patienten inte svarat i tid."

        'SupportsManualPausing = False

        ShowGuiChoice_TargetLocations = False
        ShowGuiChoice_MaskerLocations = False
        ShowGuiChoice_BackgroundNonSpeechLocations = False
        ShowGuiChoice_BackgroundSpeechLocations = False

        MinimumStimulusOnsetTime = 0.3 + 0.3 ' 0.3 in sound field
        MaximumStimulusOnsetTime = 0.8 + 0.3 ' 0.3 in sound field

        'DirectionalSimulationSet = "ARC - Harcellen - HATS - SiP"

    End Sub

    Public Overrides ReadOnly Property ShowGuiChoice_TargetSNRLevel As Boolean = False


    Private PresetName As String = "IHeAR_CS"


    Public Overrides Function InitializeCurrentTest() As Tuple(Of Boolean, String)

        Transducer = Globals.StfBase.AvaliableTransducers(0)

        CurrentSipTestMeasurement = New SipTest.SipMeasurement(CurrentParticipantID, SpeechMaterial.ParentTestSpecification, SiPTestProcedure.AdaptiveTypes.Fixed, SelectedTestparadigm)

        CurrentSipTestMeasurement.ExportTrialSoundFiles = False

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


    Private Sub PlanSiPTrials(ByVal SoundPropagationType As SoundPropagationTypes, Optional ByVal RandomSeed As Integer? = Nothing)

        'Clearing any trials that may have been planned by a previous call
        CurrentSipTestMeasurement.ClearTrials()

        'Creating a new random if seed is supplied
        If RandomSeed.HasValue Then CurrentSipTestMeasurement.Randomizer = New Random(RandomSeed)

        'Sampling a MediaSet
        'Dim MediaSet = SelectedMediaSet
        Dim SelectedMediaSets As List(Of MediaSet) = AvailableMediasets
        Dim MediaSet = SelectedMediaSets(0)

        'Getting all lists 

        'Getting the preset
        Dim TestLists As List(Of SpeechMaterialComponent) = Nothing
        If IsPractiseTest Then
            TestLists = CurrentSipTestMeasurement.ParentTestSpecification.SpeechMaterial.GetAllRelativesAtLevel(SpeechMaterialComponent.LinguisticLevels.List, False, True)
        Else
            TestLists = CurrentSipTestMeasurement.ParentTestSpecification.SpeechMaterial.Presets.GetPreset(PresetName).Members 'TODO! Specify correct members in text file
        End If

        'Getting the sound source locations
        'Head slightly turned right (i.e. Speech on left side)
        Dim TargetStimulusLocations_HeadTurnedRight = {New SoundSourceLocation With {.HorizontalAzimuth = -10, .Distance = 1.45}}
        Dim MaskerLocations_HeadTurnedRight = {New SoundSourceLocation With {.HorizontalAzimuth = 170, .Distance = 1.45}}
        Dim BackgroundLocations_HeadTurnedRight = {New SoundSourceLocation With {.HorizontalAzimuth = -10, .Distance = 1.45}, New SoundSourceLocation With {.HorizontalAzimuth = 170, .Distance = 1.45}}

        'Head slightly turned left (i.e. Speech on right side)
        Dim TargetStimulusLocations_HeadTurnedLeft = {New SoundSourceLocation With {.HorizontalAzimuth = 10, .Distance = 1.45}}
        Dim MaskerLocations_HeadTurnedLeft = {New SoundSourceLocation With {.HorizontalAzimuth = 190, .Distance = 1.45}}
        Dim BackgroundLocations_HeadTurnedLeft = {New SoundSourceLocation With {.HorizontalAzimuth = 10, .Distance = 1.45}, New SoundSourceLocation With {.HorizontalAzimuth = 190, .Distance = 1.45}}

        'Setting up test trials
        Dim TempLeftTurnTrials As New List(Of Object)
        Dim TempRightTurnTrials As New List(Of Object)

        'Creating a test unit
        Dim CurrentTestUnit = New SiPTestUnit(CurrentSipTestMeasurement)
        CurrentSipTestMeasurement.TestUnits.Add(CurrentTestUnit)

        For i = 0 To TestLists.Count - 1

            Dim TestWords = TestLists(i).GetAllDescenentsAtLevel(SpeechMaterialComponent.LinguisticLevels.Sentence)

            Dim PNR As Double
            If IsPractiseTest Then
                PNR = 15
            Else
                PNR = 0 'TODO! Select correct PNR
            End If

            For c = 0 To TestWords.Count - 1
                Dim NewLeftTurnTrial As SipTrial = New SipTrial(CurrentTestUnit, TestWords(c), MediaSet, SoundPropagationType, TargetStimulusLocations_HeadTurnedLeft.ToArray, MaskerLocations_HeadTurnedLeft.ToArray, BackgroundLocations_HeadTurnedLeft, CurrentTestUnit.ParentMeasurement.Randomizer)
                Dim NewRightTurnTrial As SipTrial = New SipTrial(CurrentTestUnit, TestWords(c), MediaSet, SoundPropagationType, TargetStimulusLocations_HeadTurnedRight.ToArray, MaskerLocations_HeadTurnedRight.ToArray, BackgroundLocations_HeadTurnedRight, CurrentTestUnit.ParentMeasurement.Randomizer)

                'Setting levels
                NewLeftTurnTrial.SetLevels(ReferenceLevel, PNR)
                NewRightTurnTrial.SetLevels(ReferenceLevel, PNR)

                'Adding the trials to the test unit
                CurrentTestUnit.PlannedTrials.Add(NewLeftTurnTrial)
                CurrentTestUnit.PlannedTrials.Add(NewRightTurnTrial)

                TempLeftTurnTrials.Add(NewLeftTurnTrial)
                TempRightTurnTrials.Add(NewRightTurnTrial)

            Next
        Next

        'Randomizing the order within trial lists
        TempLeftTurnTrials = DSP.Shuffle(TempLeftTurnTrials, Randomizer)
        TempRightTurnTrials = DSP.Shuffle(TempRightTurnTrials, Randomizer)

        'Putting the Trials in a list of the correct type
        Dim LeftTurnTrials As New List(Of SipTrial)
        For Each Item As SipTrial In TempLeftTurnTrials
            LeftTurnTrials.Add(Item)
        Next
        Dim RightTurnTrials As New List(Of SipTrial)
        For Each Item As SipTrial In TempRightTurnTrials
            RightTurnTrials.Add(Item)
        Next

        'Putting equal number trials to the left and then to the right in CurrentSipTestMeasurement (from which they can be drawn during testing)
        'Randomizing the side to start with
        Dim AddSideFirst As Integer = Randomizer.Next(0, 2)
        Dim TrialsBeforeSideSwap As Integer = 5 ' The number of trials to present on each head turn

        Dim AddedSets As Integer = 0
        For i = 0 To TempLeftTurnTrials.Count - 1 Step TrialsBeforeSideSwap

            Dim TrialsToGet As Integer = TrialsBeforeSideSwap
            If AddedSets * TrialsBeforeSideSwap + TrialsBeforeSideSwap > TempLeftTurnTrials.Count Then
                'Get the remaining number of trials
                TrialsToGet = TempLeftTurnTrials.Count - AddedSets * TrialsBeforeSideSwap
            End If

            If AddSideFirst = 0 Then
                CurrentSipTestMeasurement.PlannedTrials.AddRange(LeftTurnTrials.GetRange(i, TrialsToGet))
                CurrentSipTestMeasurement.PlannedTrials.AddRange(RightTurnTrials.GetRange(i, TrialsToGet))
            Else
                CurrentSipTestMeasurement.PlannedTrials.AddRange(RightTurnTrials.GetRange(i, TrialsToGet))
                CurrentSipTestMeasurement.PlannedTrials.AddRange(LeftTurnTrials.GetRange(i, TrialsToGet))
            End If

            AddedSets += 1

        Next


    End Sub

    Protected Overrides Sub InitiateTestByPlayingSound()

        'Sets the measurement datetime
        CurrentSipTestMeasurement.MeasurementDateTime = DateTime.Now

        'Cretaing a context sound without any test stimulus, that runs for approx TestSetup.PretestSoundDuration seconds, using audio from the first selected MediaSet
        Dim SelectedMediaSets As List(Of MediaSet) = AvailableMediasets

        Dim TestSound As Sound = CreateInitialSound(SelectedMediaSets(0))

        'Plays sound
        Globals.StfBase.SoundPlayer.SwapOutputSounds(TestSound)

        'Premixing the first 10 sounds 
        CurrentSipTestMeasurement.PreMixTestTrialSoundsOnNewTread(Transducer, MinimumStimulusOnsetTime, MaximumStimulusOnsetTime, Randomizer, TrialSoundMaxDuration, UseBackgroundSpeech, 10)

    End Sub


    Public Function CreateInitialSound(ByRef SelectedMediaSet As MediaSet, Optional ByVal Duration As Double? = Nothing) As Sound

        Try

            'Setting up the SiP-trial sound mix
            Dim MixStopWatch As New Stopwatch
            MixStopWatch.Start()

            'Sets a List of SoundSceneItem in which to put the sounds to mix
            Dim ItemList = New List(Of SoundSceneItem)

            Dim SoundWaveFormat As Formats.WaveFormat = Nothing

            'Getting a background non-speech sound
            Dim BackgroundNonSpeech_Sound As Sound = SpeechMaterial.GetBackgroundNonspeechSound(SelectedMediaSet, 0)

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

            'Sets up fading specifications for the background signals
            Dim FadeSpecs_Background = New List(Of DSP.FadeSpecifications)
            FadeSpecs_Background.Add(New DSP.FadeSpecifications(Nothing, 0, 0, CurrentSampleRate * 1))
            FadeSpecs_Background.Add(New DSP.FadeSpecifications(0, Nothing, -CurrentSampleRate * 0.01))

            'Adds the background (non-speech) signals, with fade, duck and location specifications
            Dim LevelGroup As Integer = 1 ' The level group value is used to set the added sound level of items sharing the same (arbitrary) LevelGroup value to the indicated sound level. (Thus, the sounds with the same LevelGroup value are measured together.)

            Dim BackgroundLocations_HeadTurnedRight = {New SoundSourceLocation With {.HorizontalAzimuth = -10, .Distance = 1.45}, New SoundSourceLocation With {.HorizontalAzimuth = 170, .Distance = 1.45}}
            ItemList.Add(New SoundSceneItem(Background1, 1, SelectedMediaSet.BackgroundNonspeechRealisticLevel, LevelGroup,
                                            BackgroundLocations_HeadTurnedRight(0), SoundSceneItem.SoundSceneItemRoles.BackgroundNonspeech, 0,,,, FadeSpecs_Background))
            ItemList.Add(New SoundSceneItem(Background2, 1, SelectedMediaSet.BackgroundNonspeechRealisticLevel, LevelGroup,
                                            BackgroundLocations_HeadTurnedRight(1), SoundSceneItem.SoundSceneItemRoles.BackgroundNonspeech, 0,,,, FadeSpecs_Background))
            LevelGroup += 1

            MixStopWatch.Stop()
            If LogToConsole = True Then Console.WriteLine("Prepared sounds in " & MixStopWatch.ElapsedMilliseconds & " ms.")
            MixStopWatch.Restart()

            'Creating the mix by calling CreateSoundScene of the current Mixer
            Dim MixedInitialSound As Sound = Transducer.Mixer.CreateSoundScene(ItemList, False, False, SelectedSoundPropagationType, Transducer.LimiterThreshold)

            If LogToConsole = True Then Console.WriteLine("Mixed sound in " & MixStopWatch.ElapsedMilliseconds & " ms.")

            'TODO: Here we can simulate and/or compensate for hearing loss:
            'SimulateHearingLoss,
            'CompensateHearingLoss

            Return MixedInitialSound

        Catch ex As Exception
            Logging.SendInfoToLog(ex.ToString, "ExceptionsDuringTesting")
            Return Nothing
        End Try

    End Function


    Private Sub PrepareTestTrialSound()

        Try

            If (CurrentSipTestMeasurement.ObservedTrials.Count + 3) Mod 10 = 0 Then
                'Premixing the next 10 sounds, starting three trials before the next is needed 
                CurrentSipTestMeasurement.PreMixTestTrialSoundsOnNewTread(Transducer, MinimumStimulusOnsetTime, MaximumStimulusOnsetTime, Randomizer, TrialSoundMaxDuration, UseBackgroundSpeech, 10)
            End If

            'Waiting for the background thread to finish mixing
            Dim WaitPeriods As Integer = 0
            While CurrentTestTrial.Sound Is Nothing
                WaitPeriods += 1
                Threading.Thread.Sleep(100)
                If LogToConsole = True Then Console.WriteLine("Waiting for sound to mix: " & WaitPeriods * 100 & " ms")
            End While

        Catch ex As Exception
            'Ignores any exceptions...
            'Utils.SendInfoToLog(ex.ToString, "ExceptionsDuringTesting")
        End Try

    End Sub


    Protected Overrides Sub PrepareNextTrial(ByVal NextTaskInstruction As TestProtocol.NextTaskInstruction)

        'Preparing the next trial
        CurrentTestTrial = CurrentSipTestMeasurement.PlannedTrials(0) ' GetNextTrial()
        CurrentTestTrial.TestStage = NextTaskInstruction.TestStage
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
        PrepareTestTrialSound()

        'Storing the LinguisticSoundStimulusStartTime and the LinguisticSoundStimulusDuration 
        CurrentTestTrial.LinguisticSoundStimulusStartTime = DirectCast(CurrentTestTrial, SipTrial).TestWordStartTime
        CurrentTestTrial.LinguisticSoundStimulusDuration = DirectCast(CurrentTestTrial, SipTrial).TestWordCompletedTime - CurrentTestTrial.LinguisticSoundStimulusStartTime
        CurrentTestTrial.MaximumResponseTime = MaximumResponseTime

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


    Private Function GetAverageHeadTurnScores(Optional ByVal TurnedRight As Boolean? = Nothing)

        Dim TrialScoreList As New List(Of Integer)
        For Each Trial In CurrentSipTestMeasurement.ObservedTrials

            'Skipping to next if it's a practise trial
            If Trial.IsTestTrial = False Then Continue For

            If TurnedRight.HasValue = False Then

                'Getting all results
                If Trial.IsCorrect = True Then
                    TrialScoreList.Add(1)
                Else
                    TrialScoreList.Add(0)
                End If

            Else

                Dim TrialIsTurnRight As Boolean
                If Trial.TargetStimulusLocations(0).HorizontalAzimuth = -10 Then
                    TrialIsTurnRight = True
                ElseIf Trial.TargetStimulusLocations(0).HorizontalAzimuth = 10 Then
                    TrialIsTurnRight = False
                Else
                    Throw New Exception("Incompatible head-turn data. This is a bug!")
                End If

                'Getting results only from the indicated head turn
                If TurnedRight = True And TrialIsTurnRight = True Then
                    If Trial.IsCorrect = True Then
                        TrialScoreList.Add(1)
                    Else
                        TrialScoreList.Add(0)
                    End If
                End If

                If TurnedRight = False And TrialIsTurnRight = False Then
                    If Trial.IsCorrect = True Then
                        TrialScoreList.Add(1)
                    Else
                        TrialScoreList.Add(0)
                    End If
                End If

            End If

        Next

        If TrialScoreList.Count > 0 Then
            Return TrialScoreList.Average
        Else
            Return -1
        End If

    End Function


    Public Overrides Function GetResultStringForGui() As String

        Dim TestResultSummaryLines = New List(Of String)
        TestResultSummaryLines.Add("Resultat: " & vbTab & DSP.Rounding(100 * GetAverageHeadTurnScores(Nothing)) & " % rätt")
        'TestResult.TestResultSummaryLines.Add("Head turned left: " & Math.Rounding(100 * GetAverageHeadTurnScores(False)) & " %")
        'TestResult.TestResultSummaryLines.Add("Head turned right : " & Math.Rounding(100 * GetAverageHeadTurnScores(True)) & " %")

        Return String.Join(vbCrLf, TestResultSummaryLines)

    End Function

    ''' <summary>
    ''' This function should list the names of variables included SpeechTestDump of each test trial to be exported in the "selected-variables" export file.
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function GetSelectedExportVariables() As List(Of String)
        Return New List(Of String)
    End Function

    Public Overrides Function GetSubGroupResults() As List(Of Tuple(Of String, Double))
        Return Nothing
    End Function

    Public Overrides Function GetScorePerLevel() As Tuple(Of String, SortedList(Of Double, Double))
        Return Nothing
    End Function


End Class