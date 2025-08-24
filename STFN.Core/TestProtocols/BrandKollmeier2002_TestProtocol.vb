﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Public Class BrandKollmeier2002_TestProtocol
    Inherits TestProtocol

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Brand-Kollmeier (2002)"
        End Get
    End Property

    Public Overrides ReadOnly Property Information As String
        Get
            Return "Brand-Kollmeier (2002)"
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
    Public Property Slope As Double = 0.16 'This is the slope of the psychometric function of the test material (0.16 is the slope for the Swedish Hagerman test according to Kollmeier et al 2015)' This parameter could be read from the MediaSet specification file to allow for different speech materials 

    Public Overrides Function InitializeProtocol(ByRef InitialTaskInstruction As NextTaskInstruction) As Boolean

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
            'This is the start of the test, returns the initial settings
            Return New NextTaskInstruction With {.AdaptiveValue = NextAdaptiveLevel, .Decision = SpeechTest.SpeechTestReplies.GotoNextTrial}
        End If

        'Ensuring that it's a sentence test
        If TrialHistory(TrialHistory.Count - 1).Tasks = 1 Then
            'The test needs to be aborted since it's not a sentence test. (TODO: No info about the error is sent. Maybe this should instead throw an exception, or info need to be sent somehow.)
            Return New NextTaskInstruction With {.Decision = SpeechTest.SpeechTestReplies.AbortTest}
        End If

        'Calculating adaptive step size
        'First counting the number of level reversals
        'Determining if level went up, remained unchanged or went down
        Dim DirectionList As New List(Of Utils.ElevationChange)

        For i = 1 To TrialHistory.Count - 1
            If (TrialHistory(i).AdaptiveProtocolValue > TrialHistory(i - 1).AdaptiveProtocolValue) Then
                DirectionList.Add(Utils.EnumCollection.ElevationChange.Ascending)
            ElseIf TrialHistory(i).AdaptiveProtocolValue < TrialHistory(i - 1).AdaptiveProtocolValue Then
                DirectionList.Add(Utils.EnumCollection.ElevationChange.Descendning)
            Else
                'TODO: how should Unchanged be counted?
                'As for now, unchanged is ignored
                'DirectionList.Add(Utils.Constants.ElevationChange.Unchanged)
            End If
        Next

        Dim ReversalCount As Integer = 0
        For i = 1 To DirectionList.Count - 1
            'The following three lines can be used as long as 'ElevationChange.Unchanged' is not stored in DirectionList
            If DirectionList(i) <> DirectionList(i - 1) Then
                ReversalCount += 1
            End If
        Next

        'Calculates f(i)
        Dim fi As Double = Math.Max(0.1, 1.5 * 1.41 ^ (-ReversalCount))
        'Calculates the 'prev' (proportion correct taks in the previous trial)
        Dim LastTrial_ProportionTasksCorrect As Double = TrialHistory.Last.GetProportionTasksCorrect
        'Calculated the stepsize, DeltaL 
        Dim DeltaL As Double = -(fi * (LastTrial_ProportionTasksCorrect - TargetScore)) / Slope ' N.B. TargetScore is the 'tar' variable in Brand and Kollmeier 2002, indicating the target percentage correct

        'Modifying the NextAdaptiveLevel by DeltaL
        NextAdaptiveLevel += DeltaL

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

    Public Overrides Function GetCurrentAdaptiveValue() As Double?
        Return NextAdaptiveLevel
    End Function

    Public Overrides Sub OverrideCurrentAdaptiveValue(ByVal NewValue As Double)
        NextAdaptiveLevel = NewValue
    End Sub

    Private Sub FinalizeProtocol(ByRef TrialHistory As TestTrialCollection)

        If TrialHistory.Count > 0 Then
            'Using the NextAdaptiveLevel variable from the last trial which has not yet been presented
            FinalThreshold = NextAdaptiveLevel
        Else
            'Returning NaN is no test trials have been stored
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
        'Setting FinalThreshold to NaN, since it's not possible to finalize early
        FinalThreshold = Double.NaN
    End Sub
End Class