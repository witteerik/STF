﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core.SipTest
Imports STFN.Core.Audio.SoundScene
Imports STFN.Core.Utils

Public Class QuickSiP
    Inherits SipBaseSpeechTest

    Public Overrides ReadOnly Property FilePathRepresentation As String = "QuickSiP"

    Public Sub New(ByVal SpeechMaterialName As String)
        MyBase.New(SpeechMaterialName)
        ApplyTestSpecificSettings()
    End Sub

    Public Shadows Sub ApplyTestSpecificSettings()

        TesterInstructions = "--  Quick SiP  --" & vbCrLf & vbCrLf &
             "För detta test behövs inga inställningar." & vbCrLf & vbCrLf &
             "1. Informera deltagaren om hur testet går till." & vbCrLf &
             "2. Vänd skärmen till deltagaren. Be sedan deltagaren klicka på start för att starta testet."

        ParticipantInstructions = "--  Quick SiP  --" & vbCrLf & vbCrLf &
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

    Private ResultsSummary As SortedList(Of Double, Tuple(Of QuickSipList, Double))

    Public Overrides Function InitializeCurrentTest() As Tuple(Of Boolean, String)

        'Transducer = AvaliableTransducers(1)

        CurrentSipTestMeasurement = New SipMeasurement(CurrentParticipantID, SpeechMaterial.ParentTestSpecification, SiPTestProcedure.AdaptiveTypes.Fixed, SelectedTestparadigm)

        CurrentSipTestMeasurement.ExportTrialSoundFiles = False

        'Setting SimulatedSoundField to depend on the type of transducer selected
        If Transducer.IsHeadphones = True Then
            'Quick SiP should only be used with simulated sound field in headphones
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
        PlanQuickSiPTrials(SelectedSoundPropagationType, RandomSeed)

        If CurrentSipTestMeasurement.HasSimulatedSoundFieldTrials = True And Globals.StfBase.DirectionalSimulator.SelectedDirectionalSimulationSetName = "" Then
            Return New Tuple(Of Boolean, String)(False, "The measurement requires a directional simulation set to be selected!")
        End If

        Return New Tuple(Of Boolean, String)(True, "")

    End Function

    Dim SipTestLists As New List(Of QuickSipList)

    Public Class QuickSipList
        Public SMC As SpeechMaterialComponent
        Public MediaSet As MediaSet
        Public PNR As Double
    End Class

    Public Overrides Function GetScorePerLevel() As Tuple(Of String, SortedList(Of Double, Double))

        Dim PnrScoresList = GetPnrScores()
        Dim OutputList = New SortedList(Of Double, Double)

        For Each kvp In PnrScoresList
            OutputList.Add(kvp.Key, kvp.Value.Item2)
        Next

        Return New Tuple(Of String, SortedList(Of Double, Double))("PNR (dB)", OutputList)

    End Function

    Public Function GetPnrScores() As SortedList(Of Double, Tuple(Of QuickSipList, Double))

        Dim ResultList As New List(Of Tuple(Of QuickSipList, Double)) ' QuickSipList, MeanScore

        For i = 0 To SipTestLists.Count - 1

            Dim CurrentSipTestList = SipTestLists(i)
            Dim CurrentScoresList As New List(Of Double)

            For Each Trial In CurrentSipTestMeasurement.ObservedTrials
                If Trial.MediaSet Is CurrentSipTestList.MediaSet And
                        Trial.PNR = CurrentSipTestList.PNR And
                        Trial.SpeechMaterialComponent.ParentComponent Is CurrentSipTestList.SMC Then

                    If Trial.IsCorrect = True Then
                        CurrentScoresList.Add(1)
                    Else
                        CurrentScoresList.Add(0)
                    End If

                End If
            Next

            Dim AverageScore As Double = Double.NaN
            If CurrentScoresList.Count > 0 Then
                AverageScore = CurrentScoresList.Average
            End If

            ResultList.Add(New Tuple(Of QuickSipList, Double)(CurrentSipTestList, AverageScore))
        Next

        Dim PnrSortedList As New SortedList(Of Double, Tuple(Of QuickSipList, Double))
        For Each Result In ResultList
            PnrSortedList.Add(Result.Item1.PNR, Result)
        Next

        Return PnrSortedList

    End Function

    Public Overrides Function GetSubGroupResults() As List(Of Tuple(Of String, Double))

        Dim ObservedTrials = GetObservedTestTrials().ToList
        If ObservedTrials.Count = 0 Then Return Nothing

        'Adding keys in the intended order 
        Dim GroupScoreList As New Dictionary(Of String, List(Of Integer))
        GroupScoreList.Add("Vokaler", New List(Of Integer))
        GroupScoreList.Add("Konsonanter", New List(Of Integer))

        Dim DirectionsList As New SortedSet(Of String)
        For Each Trial In ObservedTrials
            Dim CurrentHorizontalAzimuth As Double = DirectCast(Trial, SipTrial).TargetStimulusLocations(0).HorizontalAzimuth ' Using the first target location only
            Select Case CurrentHorizontalAzimuth
                Case > 0
                    DirectionsList.Add("Höger (" & CurrentHorizontalAzimuth & "°)")
                Case < 0
                    DirectionsList.Add("Vänster (" & CurrentHorizontalAzimuth & "°)")
                Case Else
                    'This should not be used in Quick SiP
                    DirectionsList.Add("Framifrån")
            End Select
        Next

        For Each Trial In CurrentSipTestMeasurement.PlannedTrials
            Dim CurrentHorizontalAzimuth As Double = DirectCast(Trial, SipTrial).TargetStimulusLocations(0).HorizontalAzimuth ' Using the first target location only
            Select Case CurrentHorizontalAzimuth
                Case > 0
                    DirectionsList.Add("Höger (" & CurrentHorizontalAzimuth & "°)")
                Case < 0
                    DirectionsList.Add("Vänster (" & CurrentHorizontalAzimuth & "°)")
                Case Else
                    'This should not be used in Quick SiP
                    DirectionsList.Add("Framifrån")
            End Select
        Next

        For Each item In DirectionsList
            GroupScoreList.Add(item, New List(Of Integer))
        Next

        GroupScoreList.Add("Man", New List(Of Integer))
        GroupScoreList.Add("Kvinna", New List(Of Integer))

        'Adding data
        For Each Trial In ObservedTrials

            'Getting the sub score 
            Dim Score As Double = Trial.ScoreList.Average

            'Step 1 - Consonants / vowels
            'Getting the group
            Dim SubGroupName1 As String
            Select Case Trial.SpeechMaterialComponent.ParentComponent.PrimaryStringRepresentation
                Case "fyr_skyr_syr", "kil_fil_sil", "tuff_tuss_tusch"
                    SubGroupName1 = "Konsonanter"
                Case "sitt_sytt_sött", "mark_märk_mörk"
                    SubGroupName1 = "Vokaler"
                Case Else
                    Throw New Exception("Unexpected test list in Quick SiP")
            End Select

            'Adding the score to the sub-score list
            GroupScoreList(SubGroupName1).Add(Score)


            'Step 2 - Left / right side azimuth
            Dim SubGroupName2 As String
            Dim CurrentHorizontalAzimuth As Double = DirectCast(Trial, SipTrial).TargetStimulusLocations(0).HorizontalAzimuth ' Using the first target location only
            Select Case CurrentHorizontalAzimuth
                Case > 0
                    SubGroupName2 = "Höger (" & CurrentHorizontalAzimuth & "°)"
                Case < 0
                    SubGroupName2 = "Vänster (" & CurrentHorizontalAzimuth & "°)"
                Case Else
                    SubGroupName2 = "Framifrån"
            End Select

            'Adding the score to the sub-score list
            GroupScoreList(SubGroupName2).Add(Score)


            'Step 3 - Male / female
            Dim SubGroupName3 As String
            Select Case DirectCast(Trial, SipTrial).MediaSet.TalkerGender
                Case MediaSet.Genders.Male
                    SubGroupName3 = "Man"
                Case MediaSet.Genders.Female
                    SubGroupName3 = "Kvinna"
                Case Else
                    Throw New Exception("Unexpected talker gender in Quick SiP")
            End Select

            'Adding the score to the sub-score list
            GroupScoreList(SubGroupName3).Add(Score)

        Next

        Dim Output As New List(Of Tuple(Of String, Double))
        For Each kvp In GroupScoreList
            If kvp.Value.Any Then
                Output.Add(New Tuple(Of String, Double)(kvp.Key, kvp.Value.Average))
            Else
                Output.Add(New Tuple(Of String, Double)(kvp.Key, 0))
            End If
        Next

        Return Output

    End Function

    Public Shared Function GetTestPnrs() As List(Of Double)

        Dim PNRs As New List(Of Double)
        Dim TempPnr As Double = 15
        For i = 0 To 9
            PNRs.Add(TempPnr)
            TempPnr -= 3.5
        Next

        Return PNRs

    End Function

    Private Sub PlanQuickSiPTrials(ByVal SoundPropagationType As SoundPropagationTypes, Optional ByVal RandomSeed As Integer? = Nothing)

        Dim AllMediaSets As List(Of MediaSet) = AvailableMediasets

        Dim SelectedMediaSets As New List(Of MediaSet)
        Dim IncludedMediaSetNames As New List(Of String) From {"City-Talker1-RVE", "City-Talker2-RVE"}
        For Each AvailableMediaSet In AllMediaSets
            If IncludedMediaSetNames.Contains(AvailableMediaSet.MediaSetName) Then
                SelectedMediaSets.Add(AvailableMediaSet)
            End If
        Next


        'Creating a new random if seed is supplied
        If RandomSeed.HasValue Then CurrentSipTestMeasurement.Randomizer = New Random(RandomSeed)

        'Getting the preset
        Dim Preset = CurrentSipTestMeasurement.ParentTestSpecification.SpeechMaterial.Presets.GetPreset(PresetName).Members

        'Ordering presets as intended
        'mark_märk_mörk, fyr_skyr_syr, sitt_sytt_sött, kil_fil_sil
        'Dim IntendedPresetOrder As New List(Of String) From {"mark_märk_mörk", "fyr_skyr_syr", "sitt_sytt_sött", "kil_fil_sil"}
        Dim IntendedPresetOrder As New List(Of String) From {"fyr_skyr_syr", "mark_märk_mörk", "kil_fil_sil", "tuff_tuss_tusch", "sitt_sytt_sött"}
        Dim TempPresets As New List(Of SpeechMaterialComponent)
        For i = 0 To 4
            For j = 0 To Preset.Count - 1
                If Preset(j).PrimaryStringRepresentation = IntendedPresetOrder(i) Then
                    TempPresets.Add(Preset(j))
                    Exit For
                End If
            Next
        Next

        Preset = TempPresets

        'Getting the sound source locations
        'Head slightly turned right (i.e. Speech on left side)
        Dim TargetStimulusLocations_HeadTurnedRight As SoundSourceLocation()
        Dim MaskerLocations_HeadTurnedRight As SoundSourceLocation()
        Dim BackgroundLocations_HeadTurnedRight As SoundSourceLocation()

        'Head slightly turned left (i.e. Speech on right side)
        Dim TargetStimulusLocations_HeadTurnedLeft As SoundSourceLocation()
        Dim MaskerLocations_HeadTurnedLeft As SoundSourceLocation()
        Dim BackgroundLocations_HeadTurnedLeft As SoundSourceLocation()

        'Placing trials with 10 degrees of rotation is no longer (after 2025-01-05) used only with SimulatedSoundField, but also with loadspeakers at 0 / 180 degrees.
        'Note, however, that the actual audio signal will be rounded to the 0 / 180 degrees speakers, only if no closer speaker exist in the loudspeaker setup used. 
        ' Thus, in sound field, this test can only be performed if speakers are more than 20 degrees (equaliiy spaces) apart (with at least one in front and back positions).
        ' Note that for calculating speaker distance, not only the azimuth, but also the distance (and elevation?) is used.

        'Head slightly turned right (i.e. Speech on left side)
        TargetStimulusLocations_HeadTurnedRight = {New SoundSourceLocation With {.HorizontalAzimuth = -10, .Distance = 1.45}}
        MaskerLocations_HeadTurnedRight = {New SoundSourceLocation With {.HorizontalAzimuth = 170, .Distance = 1.45}}
        BackgroundLocations_HeadTurnedRight = {New SoundSourceLocation With {.HorizontalAzimuth = -10, .Distance = 1.45}, New SoundSourceLocation With {.HorizontalAzimuth = 170, .Distance = 1.45}}

        'Head slightly turned left (i.e. Speech on right side)
        TargetStimulusLocations_HeadTurnedLeft = {New SoundSourceLocation With {.HorizontalAzimuth = 10, .Distance = 1.45}}
        MaskerLocations_HeadTurnedLeft = {New SoundSourceLocation With {.HorizontalAzimuth = 190, .Distance = 1.45}}
        BackgroundLocations_HeadTurnedLeft = {New SoundSourceLocation With {.HorizontalAzimuth = 10, .Distance = 1.45}, New SoundSourceLocation With {.HorizontalAzimuth = 190, .Distance = 1.45}}


        'Clearing any trials that may have been planned by a previous call
        CurrentSipTestMeasurement.ClearTrials()

        'Talker in front
        Dim PNRs As List(Of Double) = GetTestPnrs()

        SipTestLists.Add(New QuickSipList With {.SMC = Preset(0), .MediaSet = SelectedMediaSets(1), .PNR = PNRs(0)})
        SipTestLists.Add(New QuickSipList With {.SMC = Preset(1), .MediaSet = SelectedMediaSets(0), .PNR = PNRs(1)})

        SipTestLists.Add(New QuickSipList With {.SMC = Preset(2), .MediaSet = SelectedMediaSets(1), .PNR = PNRs(2)})
        SipTestLists.Add(New QuickSipList With {.SMC = Preset(3), .MediaSet = SelectedMediaSets(0), .PNR = PNRs(3)})

        SipTestLists.Add(New QuickSipList With {.SMC = Preset(4), .MediaSet = SelectedMediaSets(1), .PNR = PNRs(4)})
        SipTestLists.Add(New QuickSipList With {.SMC = Preset(0), .MediaSet = SelectedMediaSets(0), .PNR = PNRs(5)})

        SipTestLists.Add(New QuickSipList With {.SMC = Preset(1), .MediaSet = SelectedMediaSets(1), .PNR = PNRs(6)})
        SipTestLists.Add(New QuickSipList With {.SMC = Preset(2), .MediaSet = SelectedMediaSets(0), .PNR = PNRs(7)})

        SipTestLists.Add(New QuickSipList With {.SMC = Preset(3), .MediaSet = SelectedMediaSets(1), .PNR = PNRs(8)})
        SipTestLists.Add(New QuickSipList With {.SMC = Preset(4), .MediaSet = SelectedMediaSets(0), .PNR = PNRs(9)})


        'A list that determines which SipTestLists that will be randomized together
        Dim NewTestUnitIndices As New List(Of Integer) From {0, 2, 4, 6, 8}

        Dim CurrentTestUnit As SiPTestUnit = Nothing
        Dim TrialsAdded As Integer = 0

        For i = 0 To SipTestLists.Count - 1

            If NewTestUnitIndices.Contains(i) Then
                If CurrentTestUnit IsNot Nothing Then
                    CurrentSipTestMeasurement.TestUnits.Add(CurrentTestUnit)
                End If
                CurrentTestUnit = New SiPTestUnit(CurrentSipTestMeasurement)
            End If

            Dim PresetComponent = SipTestLists(i).SMC
            Dim MediaSet = SipTestLists(i).MediaSet
            Dim PNR = SipTestLists(i).PNR

            Dim TestWords = PresetComponent.GetAllDescenentsAtLevel(SpeechMaterialComponent.LinguisticLevels.Sentence)
            CurrentTestUnit.SpeechMaterialComponents.AddRange(TestWords)

            For c = 0 To TestWords.Count - 1
                Dim NewTrial As SipTrial = Nothing

                'Adding left and right head turns systematically to every other trial
                If TrialsAdded Mod 2 = 0 Then
                    NewTrial = New SipTrial(CurrentTestUnit, TestWords(c), MediaSet, SoundPropagationType, TargetStimulusLocations_HeadTurnedRight.ToArray, MaskerLocations_HeadTurnedRight.ToArray, BackgroundLocations_HeadTurnedRight, CurrentTestUnit.ParentMeasurement.Randomizer)
                Else
                    NewTrial = New SipTrial(CurrentTestUnit, TestWords(c), MediaSet, SoundPropagationType, TargetStimulusLocations_HeadTurnedLeft.ToArray, MaskerLocations_HeadTurnedLeft.ToArray, BackgroundLocations_HeadTurnedLeft, CurrentTestUnit.ParentMeasurement.Randomizer)
                End If

                NewTrial.SetLevels(ReferenceLevel, PNR)
                CurrentTestUnit.PlannedTrials.Add(NewTrial)

                TrialsAdded += 1
            Next
        Next

        'Adds the last unit
        CurrentSipTestMeasurement.TestUnits.Add(CurrentTestUnit)


        'Randomizing the order within units, in blocks of three trials (so that the head doesn't have to be moved between every trial). Note that this latter is a modification from the initial Quick-SiP data collection in 2024.

        'Randomizing the start order, left Right
        Dim DirectionIndex As Integer = CurrentSipTestMeasurement.Randomizer.Next(0, 2)

        For ui = 0 To CurrentSipTestMeasurement.TestUnits.Count - 1
            Dim Unit As SiPTestUnit = CurrentSipTestMeasurement.TestUnits(ui)

            Dim UnitLeftTurnTrials = New List(Of SipTrial)
            Dim UnitRightTurnTrials = New List(Of SipTrial)

            For Each Trial In Unit.PlannedTrials
                If Trial.TargetStimulusLocations(0).HorizontalAzimuth < 0 Then
                    ' Sound source is to the left of the head -> head is right turned in comparison
                    UnitRightTurnTrials.Add(Trial)
                Else
                    ' Sound source is to the right of the head -> head is left turned in comparison
                    UnitLeftTurnTrials.Add(Trial)
                End If
            Next

            Dim RandomList As New List(Of SipTrial)
            For r = 0 To 1

                If DirectionIndex = 0 Then

                    Do Until UnitRightTurnTrials.Count = 0
                        Dim RandomIndex As Integer = CurrentSipTestMeasurement.Randomizer.Next(0, UnitRightTurnTrials.Count)
                        RandomList.Add(UnitRightTurnTrials(RandomIndex))
                        UnitRightTurnTrials.RemoveAt(RandomIndex)
                    Loop

                    'Swapping the value of DirectionIndex, so that the other side gets included next time
                    DirectionIndex = 1
                Else

                    Do Until UnitLeftTurnTrials.Count = 0
                        Dim RandomIndex As Integer = CurrentSipTestMeasurement.Randomizer.Next(0, UnitLeftTurnTrials.Count)
                        RandomList.Add(UnitLeftTurnTrials(RandomIndex))
                        UnitLeftTurnTrials.RemoveAt(RandomIndex)
                    Loop

                    'Swapping the value of DirectionIndex, so that the other side gets included next time
                    DirectionIndex = 0
                End If

            Next

            Unit.PlannedTrials = RandomList
        Next

        'Adding the trials CurrentSipTestMeasurement (from which they can be drawn during testing)
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
        Dim SelectedMediaSets As List(Of MediaSet) = AvailableMediasets

        Dim TestSound As Audio.Sound = CreateInitialSound(SelectedMediaSets(0))

        'Plays sound
        Globals.StfBase.SoundPlayer.SwapOutputSounds(TestSound)

        'Premixing the first 10 sounds 
        CurrentSipTestMeasurement.PreMixTestTrialSoundsOnNewTread(Transducer, MinimumStimulusOnsetTime, MaximumStimulusOnsetTime, Randomizer, TrialSoundMaxDuration, UseBackgroundSpeech, 10)

    End Sub


    Public Function CreateInitialSound(ByRef SelectedMediaSet As MediaSet, Optional ByVal Duration As Double? = Nothing) As Audio.Sound

        Try

            'Setting up the SiP-trial sound mix
            Dim MixStopWatch As New Stopwatch
            MixStopWatch.Start()

            'Sets a List of SoundSceneItem in which to put the sounds to mix
            Dim ItemList = New List(Of SoundSceneItem)

            Dim SoundWaveFormat As Audio.Formats.WaveFormat = Nothing

            'Getting a background non-speech sound
            Dim BackgroundNonSpeech_Sound As Audio.Sound = SpeechMaterial.GetBackgroundNonspeechSound(SelectedMediaSet, 0)

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
            Dim MixedInitialSound As Audio.Sound = Transducer.Mixer.CreateSoundScene(ItemList, False, False, SelectedSoundPropagationType, Transducer.LimiterThreshold)

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
        CurrentTestTrial = CurrentSipTestMeasurement.GetNextTrial()
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

        Dim Output As New List(Of String)

        Dim AverageScore As Double? = GetAverageScore()
        If AverageScore.HasValue Then
            Output.Add("Overall score: " & DSP.Rounding(100 * GetAverageScore()) & "%")
        End If

        ResultsSummary = GetPnrScores()

        If ResultsSummary IsNot Nothing Then
            Output.Add("Scores per PNR level:")
            Output.Add("PNR (dB)" & vbTab & "Score" & vbTab & "List")
            For Each kvp In ResultsSummary
                Output.Add(kvp.Value.Item1.PNR & vbTab & DSP.Rounding(100 * kvp.Value.Item2) & " %" & vbTab & kvp.Value.Item1.SMC.PrimaryStringRepresentation)
            Next
        End If

        Return String.Join(vbCrLf, Output)

    End Function

    ''' <summary>
    ''' This function should list the names of variables included SpeechTestDump of each test trial to be exported in the "selected-variables" export file.
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function GetSelectedExportVariables() As List(Of String)
        Return New List(Of String)
    End Function


End Class