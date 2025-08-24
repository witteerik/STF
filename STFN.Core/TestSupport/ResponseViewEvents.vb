' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte
Public Module ResponseViewEvents

    Public Class ResponseViewEvent

        ''' <summary>
        ''' The time (in ms) relative to the onset of the test trial that the event should take place
        ''' </summary>
        Public TickTime As Integer

        Public Type As ResponseViewEventTypes

        Public Enum ResponseViewEventTypes
            PlaySound
            StopSound
            ShowVisualSoundSources
            ShowResponseAlternatives
            ShowResponseAlternativePositions
            ShowVisualCue
            HideVisualCue
            ShowResponseTimesOut
            ShowMessage
            HideAll
        End Enum

        Public Box As Object

    End Class


End Module