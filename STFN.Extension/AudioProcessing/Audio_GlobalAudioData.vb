' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte


Namespace Audio
    Public Module GlobalAudioData

        ''' <summary>
        ''' Holds the level of Raised vocal effort (in dB SPL), according to SII, ANSI S3.5 1997.
        ''' </summary>
        Public Const RaisedVocalEffortLevel As Double = 68.34

        ''' <summary>
        ''' Values for TDH type, IEC 60318-3
        ''' </summary>
        Public HL_2_SPL As New SortedList(Of Double, Double) From {{125, 45.0}, {250, 27.0}, {500, 13.5}, {750, 9.0}, {1000, 7.5}, {1500, 7.5}, {2000, 9.0}, {3000, 11.5}, {4000, 12.0}, {6000, 16.0}, {8000, 15.5}, {375, 20.25}} ' Value for 375 Hz is linearly interpolated

        Public ReferenceSoundIntensityLevel As Double = 10 ^ (-12)

        Public Enum SoundTransducerModes
            SoundField
            HeadPhones
        End Enum

        Public CurrentSoundTransducerMode As Audio.SoundTransducerModes = SoundTransducerModes.SoundField


    End Module

End Namespace
