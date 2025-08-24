' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Public Class SpeechTestResponseAlternative
    Public Spelling As String = ""
    Public IsScoredItem As Boolean = True
    Public SoundSourceLocation As Audio.SoundScene.SoundSourceLocation
    Public ParentTestTrial As TestTrial
    Public TrialPresentationIndex As Integer = 0
    Public IsVisible As Boolean = True
End Class
