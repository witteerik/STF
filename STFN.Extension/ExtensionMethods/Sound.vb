' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports System.Runtime.CompilerServices
Imports STFN.Core
Imports STFN.Core.Audio
Imports STFN.Core.Audio.Formats

'This file contains extension methods for STFN.Core.Audio.Sound

Partial Public Module Extensions

    ''' <summary>
    ''' Creates a multi channel sound out of a mono sound (or the first channel in a multichannel sound).
    ''' </summary>
    ''' <param name="NewChannelCount"></param>
    ''' <returns></returns>
    <Extension>
    Public Function ConvertMonoToMultiChannel(obj As Sound, Optional NewChannelCount As Integer = 2, Optional ShallowChannelCopies As Boolean = False) As Sound

        Dim OutputSound As New Sound(New Formats.WaveFormat(obj.WaveFormat.SampleRate, obj.WaveFormat.BitDepth, NewChannelCount,, obj.WaveFormat.Encoding))

        If ShallowChannelCopies = True Then
            For c = 1 To NewChannelCount
                OutputSound.WaveData.SampleData(c) = obj.WaveData.SampleData(1)
            Next
        Else
            For c = 1 To NewChannelCount
                OutputSound.WaveData.SampleData(c) = obj.WaveData.SampleData(1).Take(obj.WaveData.SampleData(1).Length).ToArray
            Next
        End If

        Return OutputSound

    End Function





    ''' <summary>
    ''' Creates a new sound converted to the specified format. The SMA object neither copied or referenced.
    ''' </summary>
    ''' <returns></returns>
    <Extension()>
    Public Function BitDepthConversion(obj As Sound, ByRef NewWaveFormat As Formats.WaveFormat) As Sound

        'Setting the conversion factor 
        Dim SampleConversionFactor As Double = 1
        Select Case obj.WaveFormat.BitDepth
            Case 16
                Select Case NewWaveFormat.BitDepth
                    Case 16
                            'Same, does nothing
                    Case 32
                        SampleConversionFactor = 1 / obj.WaveFormat.PositiveFullScale
                    Case Else
                        Throw New NotImplementedException
                End Select

            Case 32
                Select Case NewWaveFormat.BitDepth
                    Case 16
                        SampleConversionFactor = 1 * NewWaveFormat.PositiveFullScale
                    Case 32
                        'Same, does nothing
                    Case Else
                        Throw New NotImplementedException
                End Select
            Case Else
                Throw New NotImplementedException
        End Select

        'Creating an output sound
        Dim OutputSound As New Sound(NewWaveFormat)

        For c = 1 To OutputSound.WaveFormat.Channels

            'Copying and converting the sample data
            Dim NewChannelArray(obj.WaveData.SampleData(1).Length - 1) As Single
            OutputSound.WaveData.SampleData(c) = NewChannelArray
            For n = 0 To OutputSound.WaveData.SampleData(c).Length - 1
                OutputSound.WaveData.SampleData(c)(n) = obj.WaveData.SampleData(1)(n) * SampleConversionFactor
            Next
        Next

        Return OutputSound

    End Function


    ''' <summary>
    ''' Creates copy of the current sound which is zero-padded to a certain duration (in seconds).
    ''' <param name="TotalDuration">The duration in seconds to which the sound should be zero padded. If the sound is longer than this duration, no padding is done, but instead a copy of the original sound will be returned.</param>
    ''' <param name="CreateNewSound">If set to false, nothing will be returned, but instead the original sound will be zero padded.</param>
    ''' </summary>
    ''' <returns>Returns a new sound, which is a zero-padded copy of the current sound.</returns>
    <Extension()>
    Public Function ZeroPad(obj As Sound, ByVal TotalDuration As Double, Optional ByVal CreateNewSound As Boolean = False) As Sound

        Dim TotalLength As Integer = TotalDuration * obj.WaveFormat.SampleRate
        Dim PaddingLength As Integer = TotalLength - obj.WaveData.ShortestChannelSampleCount

        If PaddingLength <= 0 Then
            If CreateNewSound = True Then
                Return obj.CreateCopy
            Else
                Return Nothing
            End If
        Else
            Return obj.ZeroPad(0, PaddingLength, CreateNewSound)
        End If

    End Function

    ''' <summary>
    ''' Creates copy of the current sound which is zero-padded to a certain length (in samples).
    ''' <param name="TotalLength">The length in samples to which the sound should be zero padded. If the shortest channel of the sound is longer than this length, no padding is done, but instead a copy of the original sound will be returned.</param>
    ''' <param name="CreateNewSound">If set to false, nothing will be returned, but instead the original sound will be zero padded.</param>
    ''' </summary>
    ''' <returns>Returns a new sound, which is a zero-padded copy of the current sound.</returns>
    <Extension()>
    Public Function ZeroPad(obj As Sound, ByVal TotalLength As Integer, Optional ByVal CreateNewSound As Boolean = False) As Sound

        Dim PaddingLength As Integer = TotalLength - obj.WaveData.ShortestChannelSampleCount

        If PaddingLength <= 0 Then
            If CreateNewSound = True Then
                Return obj.CreateCopy
            Else
                Return Nothing
            End If
        Else
            Return obj.ZeroPad(0, PaddingLength, CreateNewSound)
        End If

    End Function


    ''' <summary>
    ''' Inverts the sound in the indicated channel. Throws an error if the indicated channel number is higher than the channel count in the WaveFormat object of the sound.
    ''' </summary>
    ''' <param name="Channel"></param>
    <Extension()>
    Public Sub InvertChannel(obj As Sound, ByVal Channel As Integer)

        If Channel < 1 Then Throw New ArgumentException("The channel value cannot be lower than 1!")
        If Channel > obj.WaveFormat.Channels Then Throw New ArgumentException("The current sound object does not have " & Channel & " channels!")

        Dim ArrayToInvert = obj.WaveData.SampleData(Channel)
        For s = 0 To ArrayToInvert.Length - 1
            ArrayToInvert(s) = -ArrayToInvert(s)
        Next

    End Sub


    ''' <summary>
    ''' Create and returns a new sound that can be used for debugging etc.
    ''' </summary>
    ''' <param name="obj"></param>
    <Extension()>
    Public Function GetTestSound(obj As STFN.Core.Audio.Sound)

        Dim TestSound = DSP.CreateSineWave(New STFN.Core.Audio.Formats.WaveFormat(48000, 32, 1, , STFN.Core.Audio.Formats.WaveFormat.WaveFormatEncodings.IeeeFloatingPoints), 1, 500, 0.1, 3)

        TestSound.SMA = New STFN.Core.Audio.Sound.SpeechMaterialAnnotation
        TestSound.SMA.AddChannelData(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.CHANNEL, Nothing) With {.OrthographicForm = "Test sound, channel 1", .PhoneticForm = "Test sound, channel 1 (phonetic form)"})

        TestSound.SMA.ChannelData(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.SENTENCE, TestSound.SMA.ChannelData(1)) With {.OrthographicForm = "Sentence 1 (orthographic form)", .PhoneticForm = "Sentence 1 (phonetic form)"})
        TestSound.SMA.ChannelData(1)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.WORD, TestSound.SMA.ChannelData(1)(0)) With {.OrthographicForm = "Word 1 (orthographic form)", .PhoneticForm = "Word 1 (phonetic form)"})
        TestSound.SMA.ChannelData(1)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.WORD, TestSound.SMA.ChannelData(1)(0)) With {.OrthographicForm = "Word 2 (orthographic form)", .PhoneticForm = "Word 2 (phonetic form)"})
        TestSound.SMA.ChannelData(1)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.WORD, TestSound.SMA.ChannelData(1)(0)) With {.OrthographicForm = "Word 3 (orthographic form)", .PhoneticForm = "Word 3 (phonetic form)"})

        TestSound.SMA.ChannelData(1)(0)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(0)) With {.OrthographicForm = "G1", .PhoneticForm = "P1"})
        TestSound.SMA.ChannelData(1)(0)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(0)) With {.OrthographicForm = "G2", .PhoneticForm = "P2"})
        TestSound.SMA.ChannelData(1)(0)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(0)) With {.OrthographicForm = "G3", .PhoneticForm = "P3"})

        TestSound.SMA.ChannelData(1)(0)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(1)) With {.OrthographicForm = "G4", .PhoneticForm = "P4"})
        TestSound.SMA.ChannelData(1)(0)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(1)) With {.OrthographicForm = "G5", .PhoneticForm = "P5"})
        TestSound.SMA.ChannelData(1)(0)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(1)) With {.OrthographicForm = "G6", .PhoneticForm = "P6"})

        TestSound.SMA.ChannelData(1)(0)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(2)) With {.OrthographicForm = "G7", .PhoneticForm = "P7"})
        TestSound.SMA.ChannelData(1)(0)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(2)) With {.OrthographicForm = "G8", .PhoneticForm = "P8"})
        TestSound.SMA.ChannelData(1)(0)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(2)) With {.OrthographicForm = "G9", .PhoneticForm = "P9"})
        TestSound.SMA.ChannelData(1)(0)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(0)(2)) With {.OrthographicForm = "G10", .PhoneticForm = "P10"})

        TestSound.SMA.ChannelData(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.SENTENCE, TestSound.SMA.ChannelData(1)) With {.OrthographicForm = "Sentence 2 (orthographic form)", .PhoneticForm = "Sentence 2 (phonetic form)"})
        TestSound.SMA.ChannelData(1)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.WORD, TestSound.SMA.ChannelData(1)(1)) With {.OrthographicForm = "Word 4 (orthographic form)", .PhoneticForm = "Word 4 (phonetic form)"})
        TestSound.SMA.ChannelData(1)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.WORD, TestSound.SMA.ChannelData(1)(1)) With {.OrthographicForm = "Word 5 (orthographic form)", .PhoneticForm = "Word 5 (phonetic form)"})
        TestSound.SMA.ChannelData(1)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.WORD, TestSound.SMA.ChannelData(1)(1)) With {.OrthographicForm = "Word 6 (orthographic form)", .PhoneticForm = "Word 6 (phonetic form)"})

        TestSound.SMA.ChannelData(1)(1)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(0)) With {.OrthographicForm = "G11", .PhoneticForm = "P11"})
        TestSound.SMA.ChannelData(1)(1)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(0)) With {.OrthographicForm = "G12", .PhoneticForm = "P12"})
        TestSound.SMA.ChannelData(1)(1)(0).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(0)) With {.OrthographicForm = "G13", .PhoneticForm = "P13"})

        TestSound.SMA.ChannelData(1)(1)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(1)) With {.OrthographicForm = "G14", .PhoneticForm = "P14"})
        TestSound.SMA.ChannelData(1)(1)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(1)) With {.OrthographicForm = "G15", .PhoneticForm = "P15"})
        TestSound.SMA.ChannelData(1)(1)(1).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(1)) With {.OrthographicForm = "G16", .PhoneticForm = "P16"})

        TestSound.SMA.ChannelData(1)(1)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(2)) With {.OrthographicForm = "G17", .PhoneticForm = "P17"})
        TestSound.SMA.ChannelData(1)(1)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(2)) With {.OrthographicForm = "G18", .PhoneticForm = "P18"})
        TestSound.SMA.ChannelData(1)(1)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(2)) With {.OrthographicForm = "G19", .PhoneticForm = "P19"})
        TestSound.SMA.ChannelData(1)(1)(2).Add(New STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaComponent(TestSound.SMA, STFN.Core.Audio.Sound.SpeechMaterialAnnotation.SmaTags.PHONE, TestSound.SMA.ChannelData(1)(1)(2)) With {.OrthographicForm = "G20", .PhoneticForm = "P20"})

        Return TestSound

    End Function


End Module