﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core.Audio.SoundScene

Namespace SipTest



    Public Class SiPTestUnit

        Public Property Description As String = ""

        Public Property ParentMeasurement As SipMeasurement

        Public Property SpeechMaterialComponents As New List(Of SpeechMaterialComponent)

        Public Property PlannedTrials As New List(Of SipTrial)

        Public Property ObservedTrials As New List(Of SipTrial)

        Public Property AdaptiveValue As Double

        Public Property TestProtocol As TestProtocol

        Public Sub New(ByRef ParentMeasurement As SipMeasurement, Optional ByVal Description As String = "")
            Me.ParentMeasurement = ParentMeasurement
            Me.Description = Description
        End Sub

        Public Sub PlanTrials(ByRef MediaSet As MediaSet, ByVal SoundPropagationType As Utils.SoundPropagationTypes,
            ByVal TargetLocation As SoundSourceLocation,
                              ByVal MaskerLocations As SoundSourceLocation(),
                              ByVal BackgroundLocations As SoundSourceLocation())

            PlannedTrials.Clear()

            Select Case ParentMeasurement.TestProcedure.AdaptiveType
                Case SiPTestProcedure.AdaptiveTypes.Fixed

                    'For n = 1 To ParentMeasurement.TestProcedure.LengthReduplications ' Should this be done here, or at a higher level?
                    For c = 0 To SpeechMaterialComponents.Count - 1
                        Dim NewTrial As New SipTrial(Me, SpeechMaterialComponents(c), MediaSet, SoundPropagationType, {TargetLocation}, MaskerLocations, BackgroundLocations, ParentMeasurement.Randomizer)
                        PlannedTrials.Add(NewTrial)
                    Next

                    'Next

                    'Case AdaptiveTypes.SimpleUpDown

                    '    'For n = 1 To ParentMeasurement.TestProcedure.LengthReduplications ' Should this be done here, or at a higher level?
                    '    For c = 0 To SpeechMaterialComponents.Count - 1
                    '        Dim NewTrial As New SipTrial(Me, SpeechMaterialComponents(c), MediaSet)
                    '        PlannedTrials.Add(NewTrial)
                    '    Next
                    '    'Next


            End Select

        End Sub

        Public Function GetNextTrial(ByRef rnd As Random) As SipTrial

            Select Case ParentMeasurement.TestProcedure.AdaptiveType
                Case SiPTestProcedure.AdaptiveTypes.Fixed

                    If PlannedTrials.Count = 0 Then Return Nothing

                    If ParentMeasurement.TestProcedure.RandomizeOrder = True Then
                        Dim RandomIndex = rnd.Next(0, PlannedTrials.Count)
                        Dim NextTrial = PlannedTrials(RandomIndex)
                        'Removing the trial from PlannedTrials
                        PlannedTrials.RemoveAt(RandomIndex)
                        Return NextTrial
                    Else
                        Dim NextTrial = PlannedTrials(0)
                        'Removing the trial from PlannedTrials
                        PlannedTrials.RemoveAt(0)
                        Return NextTrial
                    End If

                Case Else

                    Throw New NotImplementedException

            End Select

        End Function

        Public Function IsCompleted() As Boolean

            Select Case ParentMeasurement.TestProcedure.AdaptiveType
                Case SiPTestProcedure.AdaptiveTypes.Fixed

                    If PlannedTrials.Count = 0 Then
                        Return True
                    Else
                        Return False
                    End If

                Case Else

                    Throw New NotImplementedException

            End Select

        End Function


    End Class



End Namespace