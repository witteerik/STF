' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte


Namespace Audio

    Public MustInherit Class AudioSettings

        Property SelectedInputDevice As Integer?
        Property SelectedOutputDevice As Integer?
        Property FramesPerBuffer As Integer
        MustOverride Overrides Function ToString() As String
        MustOverride Property NumberOfInputChannels() As Integer?
        MustOverride Property NumberOfOutputChannels() As Integer?
        Property AllowDefaultOutputDevice As Boolean?
        Property AllowDefaultInputDevice As Boolean?
        Public MustOverride Function GetSelectedOutputDeviceName() As String

    End Class

    Public Class AndroidAudioTrackPlayerSettings
        Inherits AudioSettings

        Public SelectedOutputDeviceName As String = ""
        Public SelectedInputDeviceName As String = ""

        Private _NumberOfInputChannels As Integer? = Nothing

        Public Overrides Property NumberOfInputChannels As Integer?
            Get
                Return _NumberOfInputChannels
            End Get
            Set(value As Integer?)
                _NumberOfInputChannels = value
            End Set
        End Property

        Private _NumberOfOutputChannels As Integer? = Nothing

        Public Overrides Property NumberOfOutputChannels As Integer?
            Get
                Return _NumberOfOutputChannels
            End Get
            Set(value As Integer?)
                _NumberOfOutputChannels = value
            End Set
        End Property

        Public Overrides Function ToString() As String
            Dim OutputString As String = "Selected sound settings:"

            If SelectedOutputDeviceName <> "" Then OutputString &= "Selected input device:" & vbLf & SelectedOutputDeviceName & vbCrLf
            If SelectedInputDeviceName <> "" Then OutputString &= "Selected output device:" & vbLf & SelectedInputDeviceName & vbCrLf
            OutputString &= "FramesPerBuffer: " & FramesPerBuffer

            Return OutputString
        End Function

        Public Overrides Function GetSelectedOutputDeviceName() As String
            Return SelectedOutputDeviceName
        End Function
    End Class


End Namespace
