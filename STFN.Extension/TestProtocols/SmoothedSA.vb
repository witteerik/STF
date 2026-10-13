
' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core


Public Class SmoothedSA
    Inherits TestProtocol

    Public Overrides ReadOnly Property Name As String
        Get
            Return "Dingemanse-Goedegebure (2020)"
        End Get
    End Property

    Public Overrides ReadOnly Property Information As String
        Get
            Return "Dingemanse-Goedegebure (2020)"
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

    ''' <summary>
    ''' b is the step size parameter.
    ''' </summary>
    ''' <returns></returns>
    Public Property b As Double

    ''' <summary>
    ''' alpha is the step size decreas rate parameter.
    ''' </summary>
    ''' <returns></returns>
    Public Property alpha As Double

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

        'Calculate next adaptive level based on the previous trials

        'Calculating the average of the adaptive levels
        Dim x_List As New List(Of Double)
        For i = 0 To TrialHistory.Count - 1
            x_List.Add(TrialHistory(i).AdaptiveProtocolValue)
        Next
        Dim x_avg As Double = x_List.Average()

        'Calculating the average of the observed scores
        Dim y_List As New List(Of Double)
        For i = 0 To TrialHistory.Count - 1
            If TrialHistory(i).IsTSFCTrial = False Then
                y_List.Add(TrialHistory(i).GetProportionTasksCorrect)
            Else
                y_List.Add(TrialHistory(i).GradedResponse)
            End If
        Next
        Dim y_avg As Double = y_List.Average()

        Dim n As Double = TrialHistory.Count
        ' NOTE, Pt is implemented as TargetScore, which is the target proportion correct. This is consistent with the original Dingemanse & Goedegebure (2020) paper, where they use Pt to denote the target proportion correct. In this implementation, we are using TargetScore to represent the same concept.
        Dim a_n As Double = b / (n ^ alpha)
        Dim xnext As Double = x_avg + n * a_n * (TargetScore - y_avg)

        'Storing the next adaptive level for the next trial
        NextAdaptiveLevel = xnext

        'Checking if test is complete (presenting max number of trials)
        If TrialHistory.Count >= TestLength Then

            'Finalizing the protocol
            FinalizeProtocol(TrialHistory)

            'Exits the test
            Return New NextTaskInstruction With {.AdaptiveValue = NextAdaptiveLevel, .Decision = SpeechTest.SpeechTestReplies.TestIsCompleted}

        End If

        'Continues the test
        Return New NextTaskInstruction With {.AdaptiveValue = NextAdaptiveLevel, .Decision = SpeechTest.SpeechTestReplies.GotoNextTrial}

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
