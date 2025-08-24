' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Public Class SpeechTestInputEventArgs
    Inherits EventArgs

    Public Property LinguisticResponses As New List(Of String)

    Public Property LinguisticResponseTime As DateTime

    Public Property DirectionResponseLocations As New List(Of Audio.SoundScene.SoundSourceLocation)

    Public Property DirectionResponseName As String

    Public Property DirectionResponseTime As DateTime

End Class


