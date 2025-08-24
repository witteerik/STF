﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Public Class HagermanKinnefors1995_TestProtocol
    Inherits TestProtocol

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Hagerman-Kinnefors (1995)"
        End Get
    End Property

    Public Overrides ReadOnly Property Information As String
        Get
            Return "Hagerman-Kinnefors (1995)"
        End Get
    End Property

    Public Overrides Function GetPatientInstructions() As String
        Return ""
    End Function

    Public Overrides Function GetSuggestedStartlevel(Optional ReferenceValue As Double? = Nothing) As Double
        Throw New NotImplementedException()
    End Function

    Public Overrides Property StoppingCriterium As StoppingCriteria
        Get
            Return StoppingCriteria.TrialCount
        End Get
        Set(value As StoppingCriteria)
            MsgBox("An attempt was made to change the test protocol stopping criterium. The " & Name & " procedure requires that the stopping criterium is 'ThresholdReached'. Ignoring the attempt to set the new value.", , "Unsupported stopping criterium.")
        End Set
    End Property


    Private NextAdaptiveLevel As Double = 0
    Private FinalThreshold As Double? = Nothing
    Private TestLength As Integer

    Public Enum AdaptiveTypes
        ThresholdInNoise
        ThresholdInSilence
    End Enum

    Public Property AdaptiveType As AdaptiveTypes = AdaptiveTypes.ThresholdInNoise


    ''' <summary>
    ''' Note that after initialization, AdaptiveType also need to be set!
    ''' </summary>
    ''' <param name="InitialTaskInstruction"></param>
    ''' <returns></returns>
    Public Overrides Function InitializeProtocol(ByRef InitialTaskInstruction As NextTaskInstruction) As Boolean

        ' Setting the target score, which is 40 % in the the Hagerman procedure
        TargetScore = 0.4

        'Setting the (initial) adaptive level (should be the speech level) specified by the calling code
        NextAdaptiveLevel = InitialTaskInstruction.AdaptiveValue

        'Setting the Test length
        TestLength = InitialTaskInstruction.TestLength

        Return True

    End Function

    ''' <summary>
    ''' Returns the number of trials remaining or -1 if this is not possible to determine.
    ''' </summary>
    ''' <returns></returns>
    Public Overrides Function TotalTrialCount() As Integer
        Return TestLength
    End Function

    Public Overrides Function NewResponse(ByRef TrialHistory As TestTrialCollection) As NextTaskInstruction

        If TrialHistory.Count = 0 Then

            If AdaptiveType = AdaptiveTypes.ThresholdInNoise And IsInPretestMode = True Then
                'This is the start of the test, returns the start SNR
                Return New NextTaskInstruction With {.AdaptiveValue = 20, .Decision = SpeechTest.SpeechTestReplies.GotoNextTrial}
            Else
                'This is the start of the test, returns the initial settings
                Return New NextTaskInstruction With {.AdaptiveValue = NextAdaptiveLevel, .Decision = SpeechTest.SpeechTestReplies.GotoNextTrial}
            End If
        End If

        'Ensuring that it's a sentence test
        If TrialHistory(TrialHistory.Count - 1).Tasks = 1 Then
            'The test needs to be aborted since it's not a sentence test. (TODO: No info about the error is sent. Maybe this should instead throw an exception, or info need to be sent somehow.)
            Return New NextTaskInstruction With {.Decision = SpeechTest.SpeechTestReplies.AbortTest}
        End If

        Dim LastTrial_TasksCorrect As Double = TrialHistory.Last.ScoreList.Sum

        'Skips out of pretest mode if two or less correct answers were given
        If AdaptiveType = AdaptiveTypes.ThresholdInNoise And IsInPretestMode = True Then
            If LastTrial_TasksCorrect < 3 Then
                IsInPretestMode = False
            End If
        End If


        'Calculate the stepsize, DeltaL 
        Dim DeltaL As Double

        Select Case AdaptiveType
            Case AdaptiveTypes.ThresholdInNoise

                If IsInPretestMode = True Then

                    'Also here, the speech level is adjusted (with opposite sign as in Hagerman and Kinnefors)
                    Dim PreviousAdaptiveLevel As Double = NextAdaptiveLevel
                    Select Case TrialHistory.Count
                        Case 1
                            NextAdaptiveLevel = 10
                            DeltaL = NextAdaptiveLevel - PreviousAdaptiveLevel
                        Case 2
                            NextAdaptiveLevel = 5
                            DeltaL = NextAdaptiveLevel - PreviousAdaptiveLevel
                        Case 3
                            NextAdaptiveLevel = 0
                            DeltaL = NextAdaptiveLevel - PreviousAdaptiveLevel
                        Case 4
                            NextAdaptiveLevel = -5
                            DeltaL = NextAdaptiveLevel - PreviousAdaptiveLevel
                        Case 5
                            NextAdaptiveLevel = -8
                            DeltaL = NextAdaptiveLevel - PreviousAdaptiveLevel

                            'Skips out of the pretest mode
                            IsInPretestMode = False

                        Case Else
                            'Skips out of the pretest mode (this code should never be reached, but is left here just in case...
                            IsInPretestMode = False

                    End Select

                Else

                    'Here the speech level i adjusted
                    DeltaL = -(LastTrial_TasksCorrect - 2)

                    'This is acctually the opposite sign as in Hagerman and Kinnefors, but it used as such here since the Adaptve value should always signal the SNR (positive changes should make it easier, and vice versa)
                    'This would be the original: DeltaL = LastTrial_TasksCorrect - 2

                    'Modifying the NextAdaptiveLevel by DeltaL
                    NextAdaptiveLevel += DeltaL

                End If


            Case AdaptiveTypes.ThresholdInSilence

                'Here the speech level i adjusted
                DeltaL = -(LastTrial_TasksCorrect - 2) * 2

                'Modifying the NextAdaptiveLevel by DeltaL
                NextAdaptiveLevel += DeltaL

            Case Else
                Throw New Exception("Unkown adaptive type")
        End Select


        'Checking if test is complete (presenting max number of trials)
        If TrialHistory.Count >= TestLength Then

            'Finalizing the protocol
            FinalizeProtocol(TrialHistory)

            'Exits the test
            Return New NextTaskInstruction With {.AdaptiveValue = NextAdaptiveLevel, .AdaptiveStepSize = DeltaL, .Decision = SpeechTest.SpeechTestReplies.TestIsCompleted}

        End If

        'Continues the test
        Return New NextTaskInstruction With {.AdaptiveValue = NextAdaptiveLevel, .AdaptiveStepSize = DeltaL, .Decision = SpeechTest.SpeechTestReplies.GotoNextTrial}

    End Function

    Private Sub FinalizeProtocol(ByRef TrialHistory As TestTrialCollection)

        If TrialHistory.Count >= TestLength Then

            'Calculating the threshold
            Dim LevelList As New List(Of Double)
            For i = TrialHistory.Count - 10 To TrialHistory.Count - 1
                'Adding the adaptive levels for each of the last nine trials
                LevelList.Add(TrialHistory(i).AdaptiveProtocolValue)
            Next
            'Adding the NextAdaptiveLevel variable from the last (tenth) trial which has not yet been presented
            LevelList.Add(NextAdaptiveLevel)

            'Geting the average as the threshold
            FinalThreshold = LevelList.Average
        Else

            'Returning NaN if not all trials have been tested
            FinalThreshold = Double.NaN
        End If

    End Sub

    Public Overrides Function GetFinalResultType() As String
        Return "SRT"
    End Function

    Public Overrides Function GetFinalResultValue() As Double?

        Return FinalThreshold

    End Function

    Public Overrides Sub AbortAheadOfTime(ByRef TrialHistory As TestTrialCollection)
        FinalizeProtocol(TrialHistory)
    End Sub

    Public Overrides Function GetCurrentAdaptiveValue() As Double?
        Return NextAdaptiveLevel
    End Function

    Public Overrides Sub OverrideCurrentAdaptiveValue(ByVal NewValue As Double)
        NextAdaptiveLevel = NewValue
    End Sub

End Class