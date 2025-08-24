' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte
Public Module SharedSpeechTestObjects

    Public CurrentParticipantID As String = ""

    Public Const NoTestId As String = "zz9999"

    Public CurrentSpeechTest As SpeechTest

    Public GuiLanguage As Utils.Languages = Utils.EnumCollection.Languages.Swedish

    Public TestResultsRootFolder As String = ""

End Module
