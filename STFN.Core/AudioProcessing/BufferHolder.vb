' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports System.Threading

Namespace Audio

    Public Class BufferHolder
        Public InterleavedSampleArray As Single()
        Public ChannelDataList As List(Of Single())
        Public ChannelCount As Integer
        Public FrameCount As Integer

        ''' <summary>
        ''' Holds the (0-based) index of the first sample in the current BufferHolder
        ''' </summary>
        Public StartSample As Integer

        Public Sub New(ByVal ChannelCount As Integer, ByVal FrameCount As Integer)
            Me.ChannelCount = ChannelCount
            Me.FrameCount = FrameCount
            Dim NewInterleavedBuffer(ChannelCount * FrameCount - 1) As Single
            InterleavedSampleArray = NewInterleavedBuffer
        End Sub

        Public Sub New(ByVal ChannelCount As Integer, ByVal FrameCount As Integer, ByRef InterleavedSampleArray As Single())
            Me.ChannelCount = ChannelCount
            Me.FrameCount = FrameCount
            Me.InterleavedSampleArray = InterleavedSampleArray
        End Sub

        Public Sub ConvertToChannelData(ByRef DuplexMixer As Audio.SoundScene.DuplexMixer)
            Throw New NotImplementedException
        End Sub

        Public Shared Function CreateBufferHoldersOnNewThread(ByRef InputSound As Sound, ByRef Mixer As Audio.SoundScene.DuplexMixer, ByVal FramesPerBuffer As Integer, ByRef NumberOfOutputChannels As Integer, Optional ByVal BitdepthScaling As Double = 1, Optional ByVal BuffersOnMainThread As Integer = 10) As BufferHolder()

            Dim BufferCount As Integer = Int(InputSound.WaveData.LongestChannelSampleCount / FramesPerBuffer) + 1

            Dim Output(BufferCount - 1) As BufferHolder

            'Initializing the BufferHolders
            For b = 0 To Output.Length - 1
                Output(b) = New BufferHolder(NumberOfOutputChannels, FramesPerBuffer)
            Next

            'Creating the BuffersOnMainThread first buffers
            'Limiting the number of main thread buffers if the sound is very short
            If (Output.Length - 1) < BuffersOnMainThread Then
                BuffersOnMainThread = Math.Max(0, Output.Length - 1)
            End If


            Dim CurrentChannelInterleavedPosition As Integer
            For Each OutputRouting In Mixer.OutputRouting

                If OutputRouting.Value = 0 Then Continue For

                If OutputRouting.Value > InputSound.WaveFormat.Channels Then Continue For

                'Skipping if channel contains no data
                If InputSound.WaveData.SampleData(OutputRouting.Value).Length = 0 Then Continue For

                'Calculates the calibration gain
                Dim CalibrationGainFactor As Double = BitdepthScaling * 10 ^ (Mixer.CalibrationGain(OutputRouting.Key) / 20)

                CurrentChannelInterleavedPosition = OutputRouting.Key - 1

                'Going through buffer by buffer
                For BufferIndex = 0 To BuffersOnMainThread - 1

                    'Setting start sample and time
                    Output(BufferIndex).StartSample = BufferIndex * FramesPerBuffer

                    'Shuffling samples from the input sound to the interleaved array
                    Dim CurrentWriteSampleIndex As Integer = 0
                    For Sample = BufferIndex * FramesPerBuffer To (BufferIndex + 1) * FramesPerBuffer - 1

                        Dim x = CurrentWriteSampleIndex * NumberOfOutputChannels + CurrentChannelInterleavedPosition

                        Output(BufferIndex).InterleavedSampleArray(CurrentWriteSampleIndex * NumberOfOutputChannels + CurrentChannelInterleavedPosition) = InputSound.WaveData.SampleData(OutputRouting.Value)(Sample) * CalibrationGainFactor
                        CurrentWriteSampleIndex += 1
                    Next
                Next
            Next

            'Fixes the rest of the buffers on a new thread, allowing the new sound to start playing
            Dim ThreadWork As New BufferCreaterOnNewThread(InputSound, Output, BuffersOnMainThread, NumberOfOutputChannels, Mixer, FramesPerBuffer, BitdepthScaling)

            Return Output

        End Function

    End Class



    Public Class BufferCreaterOnNewThread
        Implements IDisposable

        Private InputSound As Sound
        Private Output As BufferHolder()
        Private BuffersOnMainThread As Integer
        Private NumberOfOutputChannels As Integer
        Private Mixer As Audio.SoundScene.DuplexMixer
        Private FramesPerBuffer As UInteger
        Private BitdepthScaling As Double

        Public Sub New(ByRef InputSound As Sound, ByRef Output As BufferHolder(), ByVal BuffersOnMainThread As Integer,
                         ByVal NumberOfOutputChannels As Integer, ByRef Mixer As Audio.SoundScene.DuplexMixer, ByVal FramesPerBuffer As UInteger, Optional ByVal BitdepthScaling As Double = 1)
            Me.InputSound = InputSound
            Me.Output = Output
            Me.BuffersOnMainThread = BuffersOnMainThread
            Me.NumberOfOutputChannels = NumberOfOutputChannels
            Me.Mixer = Mixer
            Me.FramesPerBuffer = FramesPerBuffer
            Me.BitdepthScaling = BitdepthScaling

            'Starting the new worker thread
            Dim NewThred As New Thread(AddressOf DoWork)
            NewThred.IsBackground = True
            NewThred.Start()

        End Sub

        Private Sub DoWork()

            Dim CurrentChannelInterleavedPosition As Integer
            For Each OutputRouting In Mixer.OutputRouting

                If OutputRouting.Value = 0 Then Continue For

                If OutputRouting.Value > InputSound.WaveFormat.Channels Then Continue For

                'Skipping if channel contains no data
                If InputSound.WaveData.SampleData(OutputRouting.Value).Length = 0 Then Continue For

                'Calculates the calibration gain
                Dim CalibrationGainFactor As Double = BitdepthScaling * 10 ^ (Mixer.CalibrationGain(OutputRouting.Key) / 20)

                CurrentChannelInterleavedPosition = OutputRouting.Key - 1

                'Going through buffer by buffer
                For BufferIndex = BuffersOnMainThread To Output.Length - 2

                    'Setting start sample 
                    Output(BufferIndex).StartSample = BufferIndex * FramesPerBuffer

                    'Shuffling samples from the input sound to the interleaved array, and also applied the calibration gain for the output hardware channel
                    Dim CurrentWriteSampleIndex As Integer = 0
                    For Sample = BufferIndex * FramesPerBuffer To (BufferIndex + 1) * FramesPerBuffer - 1

                        Output(BufferIndex).InterleavedSampleArray(CurrentWriteSampleIndex * NumberOfOutputChannels + CurrentChannelInterleavedPosition) = InputSound.WaveData.SampleData(OutputRouting.Value)(Sample) * CalibrationGainFactor
                        CurrentWriteSampleIndex += 1
                    Next
                Next

                'Reading the last bit
                'Setting start sample 
                Output(Output.Length - 1).StartSample = (Output.Length - 1) * FramesPerBuffer

                'Shuffling samples from the input sound to the interleaved array
                Dim CurrentWriteSampleIndexB As Integer = 0
                For Sample = FramesPerBuffer * (Output.Length - 1) To InputSound.WaveData.SampleData(OutputRouting.Value).Length - 1

                    Output(Output.Length - 1).InterleavedSampleArray(CurrentWriteSampleIndexB * NumberOfOutputChannels + CurrentChannelInterleavedPosition) = InputSound.WaveData.SampleData(OutputRouting.Value)(Sample) * CalibrationGainFactor
                    CurrentWriteSampleIndexB += 1
                Next
            Next

            'Disposing Me
            Me.Dispose()

        End Sub

#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region
    End Class


End Namespace
