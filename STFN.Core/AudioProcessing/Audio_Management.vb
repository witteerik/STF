﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Namespace Audio
    Public Module AudioManagement

        ''' <summary>
        ''' A helper function to determine which channels to modify on a sound.
        ''' </summary>
        Public Class AudioOutputConstructor

            Public Property FirstChannelIndex As Integer
            Public Property LastChannelIndex As Integer
            Private outputSoundFormat As Formats.WaveFormat

            ''' <summary>
            ''' Creates a new instance of AudioOutputConstructor
            ''' </summary>
            ''' <param name="format">The wave format of the source sound.</param>
            ''' <param name="channel">The channel to modify, as indicated by the calling function.</param>
            Public Sub New(ByRef format As Formats.WaveFormat, Optional ByVal channel As Integer? = Nothing)

                If channel = 0 Then Throw New ArgumentException("Channel value cannot be lower than 1.")

                If channel IsNot Nothing Then
                    FirstChannelIndex = channel
                    LastChannelIndex = channel
                Else
                    FirstChannelIndex = 1
                    LastChannelIndex = format.Channels
                End If

                Dim outputChannelCount As Integer = LastChannelIndex - FirstChannelIndex + 1
                outputSoundFormat = New Formats.WaveFormat(format.SampleRate, format.BitDepth, outputChannelCount,, format.Encoding)

            End Sub

            ''' <summary>
            ''' Creates a new sound based on the AudioOutputConstructor settings.
            ''' </summary>
            ''' <returns></returns>
            Public Function GetNewOutputSound() As Sound
                Dim outputSound = New Sound(outputSoundFormat)
                Return outputSound
            End Function

        End Class

        ''' <summary>
        ''' 'This sub checks to see if the values given for startSample and sectionLength is too high or below zero.
        ''' If startSample is below 0, it is set to 0. If it is higher than the length of the array, it is set to the last sampe in the array.
        ''' If sectionlength is below zero, it is set to 0. If it is too long for the array, it changed to fit within the length of the array.
        ''' </summary>
        ''' <param name="inputArrayLength">The length of the input array.</param>
        ''' <param name="startSample">The index of the start sample. If a negative value is supplied, the start sample will be changed to be startSample samples from the end of the array (-1 will mean the last index in the array, -2 will mean the second last sample, and so on.) </param>
        ''' <param name="sectionLength">The length of the section in samples. If the input value is Nothing, then sectionLength is set to the length of the rest of the sound.</param>
        ''' <returns>Returns True if any of startSample or sectionLength was corrected, and False if no correction was made.</returns>
        Public Function CheckAndCorrectSectionLength(ByVal inputArrayLength As Double, ByRef startSample As Integer, ByRef sectionLength As Integer?) As Boolean

            'N.B. This sub was changed 2022-09-19 to support negative start samples (in typical Python style), with the meaning start startSample samples before the (exclusive) end of the array.

            Dim modified As Integer = 0

            'New from 2022-09-19
            If startSample < 0 Then
                startSample = inputArrayLength - Math.Abs(startSample)
            End If
            'End new from 2022-09-19

            If startSample < 0 Then
                startSample = 0
                modified += 1
            End If

            If startSample > inputArrayLength - 1 Then
                startSample = inputArrayLength - 1
                modified += 1
            End If

            If sectionLength Is Nothing Then
                sectionLength = inputArrayLength
                modified += 1
            End If

            If sectionLength < 0 Then
                sectionLength = 0
                modified += 1
            End If

            If sectionLength > inputArrayLength - startSample Then
                sectionLength = inputArrayLength - startSample
                modified += 1
            End If

            If modified > 0 Then
                Return True
            Else
                Return False
            End If

        End Function



        ''' <summary>
        ''' Returns a value indicating if a audio bit depth for a given encoding is supported by the current audio library.
        ''' </summary>
        ''' <param name="WaveEncoding"></param>
        ''' <param name="BitDepth"></param>
        ''' <returns></returns>
        Public Function CheckIfBitDepthIsSupported(ByVal WaveEncoding As Formats.WaveFormat.WaveFormatEncodings, ByVal BitDepth As Integer) As Boolean
            Select Case WaveEncoding
                Case Formats.WaveFormat.WaveFormatEncodings.PCM
                    Select Case BitDepth
                        Case 16, 32
                            Return True
                        Case Else
                            Return False
                    End Select

                Case Formats.WaveFormat.WaveFormatEncodings.IeeeFloatingPoints
                    Select Case BitDepth
                        Case 32
                            Return True
                        Case Else
                            Return False
                    End Select
                Case Else
                    Return False
            End Select
        End Function

        Public Sub CheckChannelValue(ByRef channelValue As Integer, ByVal upperBound As Integer)
            If channelValue < 1 Then Throw New Exception("Channel value cannot be lower than 1.")
            If channelValue > upperBound Then Throw New Exception("The referred instance does not have " & channelValue & " channels.")
        End Sub


    End Module

End Namespace
