﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports System.Threading
Imports STFN.Core.Audio


Public Class DSP

    ''' <summary>
    ''' Calculates the signal to noise ratio between the signal level and the noise level.
    ''' </summary>
    ''' <param name="SignalLevel"></param>
    ''' <param name="NoiseLevel"></param>
    ''' <returns></returns>
    Public Shared Function SignalToNoiseRatio(ByVal SignalLevel As Double, ByVal NoiseLevel As Double)
        Return SignalLevel - NoiseLevel
    End Function



    ''' <summary>
    ''' Measures the sound level of a section of the input sound.
    ''' </summary>
    ''' <param name="InputSound"></param>
    ''' <param name="channel"></param>
    ''' <param name="StartSample"></param>
    ''' <param name="SectionLength"></param>
    ''' <param name="outputUnit"></param>
    ''' <param name="SoundMeasurementType"></param>
    ''' <param name="Frequencyweighting">The frequency Weighting to be applied before the sound measurement.</param>
    ''' <param name="ReturnLinearMeanSquareData">If set to true, the linear mean square of the measured section is returned (I.e. Any values set for outputUnit and SoundMeasurementType are overridden.).</param>
    ''' <param name="LinearSquareData">If ReturnLinearMeanSquareData is set to True, LinearSquareData will contain item1 = linear sum of square, and item2 = length of the measurement section in samples.</param>
    ''' <returns></returns>
    Public Shared Function MeasureSectionLevel(ByRef InputSound As Sound, ByVal channel As Integer,
                                        Optional ByVal StartSample As Integer = 0, Optional ByVal SectionLength As Integer? = Nothing,
                                        Optional ByVal outputUnit As SoundDataUnit = SoundDataUnit.dB,
                                        Optional ByVal SoundMeasurementType As SoundMeasurementType = SoundMeasurementType.RMS,
                                        Optional ByVal Frequencyweighting As FrequencyWeightings = FrequencyWeightings.Z,
                                        Optional ByVal ReturnLinearMeanSquareData As Boolean = False,
                                        Optional ByRef LinearSquareData As Tuple(Of Double, Integer) = Nothing) As Double?

        Try

            CheckAndCorrectSectionLength(InputSound.WaveData.SampleData(channel).Length, StartSample, SectionLength)

            'Preparing an array to do measurements on
            Dim MeasurementArray() As Single

            If Frequencyweighting = FrequencyWeightings.Z Then
                'Just referencing the input sound array
                MeasurementArray = InputSound.WaveData.SampleData(channel)

            Else
                'Preparing a new sound with only the measurement section, to be filterred
                Dim tempSound As New Sound(InputSound.WaveFormat)
                ReDim MeasurementArray(SectionLength - 1)
                For s = StartSample To StartSample + SectionLength - 1
                    MeasurementArray(s - StartSample) = InputSound.WaveData.SampleData(channel)(s)
                Next
                tempSound.WaveData.SampleData(channel) = MeasurementArray

                'Filterring for Weighing
                tempSound = DSP.IIRFilter(tempSound, Frequencyweighting, channel)
                If tempSound Is Nothing Then
                    Throw New Exception("Something went wrong during IIR-filterring")
                    Return Nothing 'Aborting and return vb null if something went wrong here
                End If

                'Referencing the MeasurementArray again (since the reference is broken during the IIR filtering)
                MeasurementArray = tempSound.WaveData.SampleData(channel)

                'Setting startsample to 0 since all sound before the startsample has been excluded from MeasurementArray
                StartSample = 0

            End If

            'Overrides soundmeasurmenttype
            If ReturnLinearMeanSquareData = True Then SoundMeasurementType = SoundMeasurementType.RMS

            Select Case SoundMeasurementType
                Case SoundMeasurementType.RMS

                    'Calculates RMS value of the section

                    Dim RMS As Double? = Nothing

                    Dim SumOfSquare As Double = 0

                    SumOfSquare = CalculateSumOfSquare(MeasurementArray, StartSample, SectionLength)

                    ''Calculating the sum of sqares in libostfdsp
                    'SumOfSquare = LibOstfDsp_VB.calculateFloatSumOfSquare(MeasurementArray, MeasurementArray.Length, StartSample, SectionLength)

                    'Returns the mean square (MR) if ReturnLinearMeanSquareData is True
                    If ReturnLinearMeanSquareData = False Then

                        'Calculates RMS
                        RMS = (SumOfSquare / SectionLength) ^ (1 / 2)

                    Else

                        'Stores LinearSquareData
                        LinearSquareData = New Tuple(Of Double, Integer)(SumOfSquare, SectionLength)

                        'Returns the ReturnLinearMeanSquareData
                        Return SumOfSquare / SectionLength

                    End If


                    Select Case outputUnit
                        Case SoundDataUnit.dB
                            Dim sectionLevel As Double = dBConversion(RMS, dBConversionDirection.to_dB, InputSound.WaveFormat)
                            Return sectionLevel
                        Case SoundDataUnit.linear
                            Return RMS
                    End Select

                Case SoundMeasurementType.AbsolutePeakAmplitude
                    'Calculates the absolute max amplitude of the section

                    Dim LocalArrayCopy(SectionLength - 1) As Single
                    Array.Copy(MeasurementArray, LocalArrayCopy, SectionLength.Value)

                    Dim peak_pos As Double = LocalArrayCopy.Max
                    Dim peak_neg As Double = LocalArrayCopy.Min

                    Dim peak As Double = 0

                    If peak_pos > -peak_neg Then
                        peak = peak_pos
                    Else
                        peak = -peak_neg
                    End If

                    Select Case outputUnit
                        Case SoundDataUnit.dB
                            Dim sectionLevel As Double = dBConversion(peak, dBConversionDirection.to_dB, InputSound.WaveFormat)
                            Return sectionLevel
                        Case SoundDataUnit.linear
                            Return peak
                    End Select

                Case SoundMeasurementType.averageAbsoluteAmplitude

                    'Calculates the average absolute amplitude of the section
                    Dim AccumulativeSoundLevel As Double

                    'MsgBox(inputArray.Length & " " & startSample & " " & sectionLength)

                    For n = StartSample To StartSample + SectionLength - 1
                        AccumulativeSoundLevel = AccumulativeSoundLevel + Math.Abs(MeasurementArray(n))
                    Next

                    Dim averageAbsoluteAmplitude = AccumulativeSoundLevel / SectionLength

                    Select Case outputUnit
                        Case SoundDataUnit.dB
                            Dim sectionLevel As Double = dBConversion(averageAbsoluteAmplitude, dBConversionDirection.to_dB, InputSound.WaveFormat)
                            Return sectionLevel
                        Case SoundDataUnit.linear
                            Return averageAbsoluteAmplitude
                    End Select

            End Select


        Catch ex As Exception
            AudioError(ex.ToString)
            Return Nothing
        End Try

    End Function



    ''' <summary>
    ''' Returns the RMS of the window with the highest RMS value.
    ''' </summary>
    ''' <param name="InputSound">The sound to measure.</param>
    ''' <param name="WindowSize">The windows size in samples.</param>
    ''' <param name="LoudestWindowStartSample">Upon return, holds the start sample of loudest window.</param>
    ''' <returns></returns>
    Public Shared Function GetLevelOfLoudestWindow(ByRef InputSound As Sound, ByVal Channel As Integer, ByVal WindowSize As Integer,
                                            Optional ByVal StartSample As Integer = 0, Optional ByVal SectionLength As Integer? = Nothing,
                                            Optional ByRef LoudestWindowStartSample As Integer = 0,
                                            Optional ByVal Frequencyweighting As FrequencyWeightings = FrequencyWeightings.Z,
                                            Optional ByVal ZeroPadToWindowSize As Boolean = False,
                                                Optional ByRef WindowLevels As List(Of Double) = Nothing) As Double


        CheckAndCorrectSectionLength(InputSound.WaveData.SampleData(Channel).Length, StartSample, SectionLength)

        'Stores the initial start sample value since this is changed if the sound is filtering below
        Dim InitialStartSample As Integer = StartSample

        'Resetting LoudestWindowStartSample
        LoudestWindowStartSample = 0

        'Copying the section to do measurements on to a new sound
        Dim TempSound As New Sound(New Formats.WaveFormat(InputSound.WaveFormat.SampleRate, InputSound.WaveFormat.BitDepth, 1,, InputSound.WaveFormat.Encoding))
        If Frequencyweighting = FrequencyWeightings.Z And StartSample = 0 And SectionLength = InputSound.WaveData.SampleData(Channel).Length Then
            'No need to copy, only referening the sound
            TempSound.WaveData.SampleData(1) = InputSound.WaveData.SampleData(Channel)

        Else
            'Coying samples
            Dim MeasurementArray(SectionLength - 1) As Single
            For s = StartSample To StartSample + SectionLength - 1
                MeasurementArray(s - StartSample) = InputSound.WaveData.SampleData(Channel)(s)
            Next
            TempSound.WaveData.SampleData(1) = MeasurementArray

        End If

        'Filtering the TempSound 
        If Frequencyweighting <> FrequencyWeightings.Z Then
            TempSound = DSP.IIRFilter(TempSound, Frequencyweighting, 1)
            If TempSound Is Nothing Then
                Throw New Exception("Something went wrong during IIR-filterring")
                Return Nothing 'Aborting and returning Nothing if something went wrong here
            End If
        End If


        Dim HighestSumOfSquares As Double = Double.NegativeInfinity

        'If the section to measure is shorter than the WindowSize, MeasureSectionLevel is used directly (with or without zero-padding)
        If WindowSize > TempSound.WaveData.SampleData(1).Length Then
            If ZeroPadToWindowSize = True Then

                'Zero-padding the sample array
                ReDim Preserve TempSound.WaveData.SampleData(1)(WindowSize - 1)
            End If

            'Measures and returns the level of the whole TempSound
            Return MeasureSectionLevel(TempSound, 1, ,,,, FrequencyWeightings.Z)

        End If

        Dim CurrentSumOfSquares As Double
        Dim LeftToUpdate As Integer = 0

        For s = 0 To TempSound.WaveData.SampleData(1).Length - 1 - WindowSize

            If LeftToUpdate = 0 Then

                'Re-calculating a new CurrentSumOfSquares, by iterating all samples in the current window. The reason this is done is to reduce the accumulated effect of rounding caused by adding and subtracting floating point numbers.
                CurrentSumOfSquares = 0
                For ws = s To s + WindowSize - 1
                    CurrentSumOfSquares += TempSound.WaveData.SampleData(1)(ws) ^ 2
                Next

                LeftToUpdate = WindowSize - 1

            Else

                'Subtracting the square of the s - 1 sample
                CurrentSumOfSquares -= TempSound.WaveData.SampleData(1)(s - 1) ^ 2

                'Adding the square of the new sample (s + WindowSize - 1)
                CurrentSumOfSquares += TempSound.WaveData.SampleData(1)(s + WindowSize - 1) ^ 2

                'Decreasing the samples left to update
                LeftToUpdate -= 1
            End If

            'Stroring the CurrentSumOfSquares in WindowLevels
            If WindowLevels IsNot Nothing Then
                WindowLevels.Add(CurrentSumOfSquares)
            End If

            'Updating and storing the loudest window start sample
            If CurrentSumOfSquares > HighestSumOfSquares Then
                HighestSumOfSquares = CurrentSumOfSquares
                LoudestWindowStartSample = s
            End If

        Next

        'Adding the section length removed prior to the initial start sample to LoudestWindowStartSample
        LoudestWindowStartSample += InitialStartSample

        If WindowLevels IsNot Nothing Then
            'Calulating the RMS level in dB for the WindowLevels 
            Dim dBList As New List(Of Double)
            For Each SuMOfSquareValue In WindowLevels
                If SuMOfSquareValue > 0 Then
                    dBList.Add(dBConversion(Math.Sqrt(SuMOfSquareValue / WindowSize), dBConversionDirection.to_dB, TempSound.WaveFormat))
                Else
                    dBList.Add(Double.NegativeInfinity)
                End If
            Next
            WindowLevels = dBList
        End If

        'Calulating the RMS level in dB
        If HighestSumOfSquares > 0 Then
            Return dBConversion(Math.Sqrt(HighestSumOfSquares / WindowSize), dBConversionDirection.to_dB, TempSound.WaveFormat)
        Else
            Return Double.NegativeInfinity
        End If

    End Function




    Public Class Radix2TrigonometricLookup

        Private Shared LookupDictionary As New SortedList(Of FftDirections, SortedList(Of Integer, List(Of Tuple(Of Double, Double))))
        Private Shared ArrayDictionary As New SortedList(Of FftDirections, SortedList(Of Integer, Tuple(Of Double(), Double())))

        Public Shared Function GetRadix2TrigonomerticValues(ByVal Size As Integer, ByRef Direction As FftDirections) As List(Of Tuple(Of Double, Double))

            If LookupDictionary.ContainsKey(Direction) = False Then
                LookupDictionary.Add(Direction, New SortedList(Of Integer, List(Of Tuple(Of Double, Double))))
                ArrayDictionary.Add(Direction, New SortedList(Of Integer, Tuple(Of Double(), Double())))
            End If

            If LookupDictionary(Direction).ContainsKey(Size) = True Then
                'No values need to be calculated
                Return LookupDictionary(Direction)(Size)
            End If

            Dim TrigonometricValues As New List(Of Tuple(Of Double, Double))

            Dim ExponentSign As Integer
            If Direction = FftDirections.Forward Then
                ExponentSign = -1
            ElseIf Direction = FftDirections.Backward Then
                ExponentSign = 1
            Else
                Throw New ArgumentException("Unknown FFT direction")
            End If

            Dim HalfSize As Integer = Size / 2
            For n = 0 To HalfSize - 1
                Dim LookupKey As Integer = HalfSize * (n / HalfSize)
                Dim Exponent As Double = ExponentSign * (n / HalfSize) * Math.PI
                TrigonometricValues.Add(New Tuple(Of Double, Double)(Math.Cos(Exponent), Math.Sin(Exponent)))
            Next

            'Adds the values
            LookupDictionary(Direction).Add(Size, TrigonometricValues)

            Dim PcCos As New List(Of Double)
            Dim PcSin As New List(Of Double)

            For Each kvp In TrigonometricValues
                PcCos.Add(kvp.Item1)
                PcSin.Add(kvp.Item2)
            Next

            ArrayDictionary(Direction).Add(Size, New Tuple(Of Double(), Double())(PcCos.ToArray, PcSin.ToArray))

            'And also returns them
            Return TrigonometricValues

        End Function


        Public Shared Function GetArrays(ByVal Size As Integer, ByRef Direction As FftDirections) As Tuple(Of Double(), Double())

            If ArrayDictionary.ContainsKey(Direction) = False Then
                GetRadix2TrigonomerticValues(Size, Direction)
            End If

            If ArrayDictionary(Direction).ContainsKey(Size) = False Then
                GetRadix2TrigonomerticValues(Size, Direction)
            End If

            Return ArrayDictionary(Direction)(Size)

        End Function

    End Class


    Public Enum FftDirections
        Forward
        Backward
    End Enum


    Public Enum ProcessingDomain
        TimeDomain
        FrequencyDomain
    End Enum


    Public Shared Function FIRFilter(ByVal inputSound As Sound, ByVal impulseResponse As Sound,
                  ByRef fftFormat As Formats.FftFormat, Optional ByVal inputSoundChannel As Integer? = Nothing,
                  Optional ByVal startSample As Integer = 0, Optional ByVal sectionLength As Integer? = Nothing,
                  Optional ByVal filteringDomain As ProcessingDomain = ProcessingDomain.FrequencyDomain,
                  Optional ByVal ScaleImpulseResponse As Boolean = False, Optional InActivateWarnings As Boolean = False,
                  Optional ByVal KeepInputSoundLength As Boolean = False) As Sound

        If filteringDomain = ProcessingDomain.TimeDomain Then
            Throw New ArgumentException("The filteringDomain argument value ProcessingDomain.TimeDomain is no longer supported. Use the function FIRFilterTimeDomain instead.")
        End If

        'Frequency domain convolution (complex multiplication)

        'Reference: This frequency domain calculation is based on the overlap-add method as
        'described in Bateman, A. & Paterson-Stephens, I. (2002). The DSP Handbook. Algorithms, Applications and Design Techniques.
        'chapter 7, pp 451-453.

        Try

            Dim IRChannel As Integer = 1

            Dim outputSound As New Sound(inputSound.WaveFormat)
            Dim AudioOutputConstructor As New AudioOutputConstructor(inputSound.WaveFormat, inputSoundChannel)
            outputSound.FFT = New FftData(inputSound.WaveFormat, fftFormat)
            Dim FS As Double = inputSound.WaveFormat.PositiveFullScale

            'Main section
            Select Case inputSound.WaveFormat.BitDepth
                Case 16, 32
                    For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                        'Copies the impulse response to a new array of double
                        Dim TempArray = impulseResponse.WaveData.SampleData(IRChannel)
                        Dim IRArray(TempArray.Length - 1) As Double
                        CopyToDouble(TempArray, IRArray)

                        'Scaling impulse response. 
                        If ScaleImpulseResponse = True Then

                            Dim ImpulseResponseArraySum As Double = IRArray.Sum

                            If ImpulseResponseArraySum <> 0 Then
                                MultiplyArray(IRArray, (1 / ImpulseResponseArraySum))

                                'For n = 0 To IRArray.Length - 1
                                '    IRArray(n) = IRArray(n) / ImpulseResponseArraySum
                                'Next
                            Else
                                MsgBox("The impulse response sums to 0!")
                            End If
                        End If

                        'Referencing the current channel array, and noting its original length
                        Dim OriginalInputChannelLength As Integer = inputSound.WaveData.SampleData(c).Length
                        Dim IrArrayLength As Integer = IRArray.Length
                        Dim IntendedOutputLength As Integer = OriginalInputChannelLength + IrArrayLength

                        CheckAndAdjustFFTSize(fftFormat.FftWindowSize, (IrArrayLength * 2) + 1, InActivateWarnings)

                        Dim AddOverlapSliceLength As Integer = fftFormat.FftWindowSize / 2
                        Dim NumberOfTimeWindows As Integer = Int(IntendedOutputLength / AddOverlapSliceLength)

                        'Copies the input array to a new array of Double, if needed, also extends the input array to a whole number multiple of the length of the sound data that goes into each dft
                        If IntendedOutputLength Mod AddOverlapSliceLength > 0 Then NumberOfTimeWindows += 1
                        Dim IntputChannelSoundLength As Integer = AddOverlapSliceLength * (NumberOfTimeWindows + 1)
                        Dim InputSoundChannelArrayDouble(IntputChannelSoundLength - 1) As Double
                        CopyToDouble(inputSound.WaveData.SampleData(c), InputSoundChannelArrayDouble)

                        'Creates an array in which to store the convoluted sound
                        Dim OutputChannelSampleArray(IntputChannelSoundLength - 1) As Single

                        'Creates dft bins for the IR data
                        Dim dftIR_Bin_x(fftFormat.FftWindowSize - 1) As Double
                        Dim dftIR_Bin_y(fftFormat.FftWindowSize - 1) As Double

                        'Copied the IR data into the IR bins
                        Array.Copy(IRArray, dftIR_Bin_x, IrArrayLength)

                        'Calculates forward FFT for the IR 
                        FastFourierTransform(FftDirections.Forward, dftIR_Bin_x, dftIR_Bin_y, True)

                        'Starts convolution one window at a time
                        Dim readSample As Integer = 0
                        Dim writeSample As Integer = 0

                        For windowNumber = 0 To NumberOfTimeWindows - 1 'Step 2

                            'Creates a zero-padded sound array with the length of the dft windows size ()

                            'Creates a new bins for the input sound data
                            Dim dftSoundBin_x(fftFormat.FftWindowSize - 1) As Double
                            Dim dftSoundBin_y(fftFormat.FftWindowSize - 1) As Double

                            'Copies the slice samples into the dft x-bin, and increases readSample by AddOverlapSliceLength
                            Array.Copy(InputSoundChannelArrayDouble, readSample, dftSoundBin_x, 0, AddOverlapSliceLength)
                            readSample += AddOverlapSliceLength

                            'Calculates forward FFT for the current sound window (Skipping the forward transform scaling on x instead of h, for optimization (the results of the complex multiplication should be the same))
                            FastFourierTransform(FftDirections.Forward, dftSoundBin_x, dftSoundBin_y, False)

                            'Performs complex multiplications
                            ComplexMultiplication(dftSoundBin_x, dftSoundBin_y, dftIR_Bin_x, dftIR_Bin_y)

                            'Dim tempDftSoundBin_x As Double = 0
                            'For n = 0 To fftFormat.FftWindowSize - 1
                            '    tempDftSoundBin_x = dftSoundBin_x(n) 'stores this value so that it does not get overwritten in the following line (it needs to be used also two lines below)
                            '    dftSoundBin_x(n) = tempDftSoundBin_x * dftIR_Bin_x(n) - dftSoundBin_y(n) * dftIR_Bin_y(n)
                            '    dftSoundBin_y(n) = tempDftSoundBin_x * dftIR_Bin_y(n) + dftSoundBin_y(n) * dftIR_Bin_x(n)
                            'Next

                            'Calculates inverse FFT
                            FastFourierTransform(FftDirections.Backward, dftSoundBin_x, dftSoundBin_y)

                            'Puts the convoluted sound in the output array
                            For sample = 0 To fftFormat.FftWindowSize - 1
                                OutputChannelSampleArray(writeSample) += dftSoundBin_x(sample)
                                writeSample += 1
                            Next
                            writeSample -= AddOverlapSliceLength

                            'Referencin the sound array in the output sound
                            outputSound.WaveData.SampleData(c) = OutputChannelSampleArray

                        Next

                        If KeepInputSoundLength = True Then

                            'Correcting the channel length, by copying the section needed
                            Dim ItitialTrimLength As Integer = impulseResponse.WaveData.SampleData(IRChannel).Length / 2
                            Dim NewChannelArray(OriginalInputChannelLength - 1) As Single
                            'For s = 0 To NewChannelArray.Length - 1
                            '    NewChannelArray(s) = OutputChannelSampleArray(s + ItitialTrimLength)
                            'Next
                            Array.Copy(OutputChannelSampleArray, ItitialTrimLength, NewChannelArray, 0, NewChannelArray.Length)

                            outputSound.WaveData.SampleData(c) = NewChannelArray

                            'Fading the beginning and the end of the channel array
                            DSP.Fade(outputSound, Nothing, 0, c, 0, Math.Min(NewChannelArray.Length / 10, 100))
                            DSP.Fade(outputSound, 0, Nothing, c, NewChannelArray.Length - Math.Min(NewChannelArray.Length / 10, 100))

                        Else

                            Dim NewChannelArray(IntendedOutputLength - 1) As Single
                            Array.Copy(outputSound.WaveData.SampleData(c), 0, NewChannelArray, 0, NewChannelArray.Length)
                            outputSound.WaveData.SampleData(c) = NewChannelArray

                            'ReDim Preserve outputSound.WaveData.SampleData(c)(IntendedOutputLength - 1)

                        End If

                        'Increasing impulse response channel if the impulse response has more than 1 channel.
                        If impulseResponse.WaveFormat.Channels > 1 Then IRChannel += 1

                    Next

                    Return outputSound

                Case Else
                    Throw New NotImplementedException(inputSound.WaveFormat.BitDepth & " bit depth is not yet supported.")
                    Return Nothing
            End Select


        Catch ex As Exception
            AudioError(ex.ToString)
            Return Nothing
        End Try


    End Function


    ''' <summary>
    ''' Calculates the fast fourier transform using the complex Radix-2 FFT algorithm
    ''' </summary>
    ''' <param name="Direction">Transform direction</param>
    ''' <param name="x">Real data array</param>
    ''' <param name="y">Imaginary data array</param>
    ''' <param name="ScaleForwardTransform"></param>
    ''' <param name="Reorder">Set to false to skip sample reordering</param>
    Public Shared Sub FastFourierTransform(ByVal Direction As FftDirections, ByRef x() As Double, ByRef y() As Double,
                               Optional ScaleForwardTransform As Boolean = True, Optional ByVal Reorder As Boolean = True)


        If Globals.StfBase.UseOptimizationLibraries = False Then
            FftRadix2(x, y, Direction, ScaleForwardTransform, Reorder)

        Else

            LibOstfDsp_VB.Fft_complex(x, y, x.Length, Direction, Reorder, ScaleForwardTransform)

        End If

    End Sub


    ''' <summary>
    ''' Complex Radix-2 FFT
    ''' </summary>
    ''' <param name="x">Real data array</param>
    ''' <param name="y">Imaginary data array</param>
    ''' <param name="Direction">Transform direction</param>
    ''' <param name="ScaleForwardTransform"></param>
    ''' <param name="Reorder">Set to false to skip sample reordering</param>
    Public Shared Sub FftRadix2(ByRef x() As Double, ByRef y() As Double, ByRef Direction As FftDirections, Optional ByVal ScaleForwardTransform As Boolean = True, Optional ByVal Reorder As Boolean = True)

        ' This is a modified VB translation of the MIT licensed code in Mathnet Numerics, See https://github.com/mathnet/mathnet-numerics/blob/306fb068d73f3c3d0e90f6f644b55cddfdeb9a0c/src/Numerics/Providers/FourierTransform/ManagedFourierTransformProvider.Radix2.cs

        Dim ExponentSign As Integer
        If Direction = FftDirections.Forward Then
            ExponentSign = -1
        ElseIf Direction = FftDirections.Backward Then
            ExponentSign = 1
        Else
            Throw New ArgumentException("Unknown FFT direction")
        End If

        If Reorder = True Then

            Dim TempX As Double
            Dim TempY As Double

            Dim j As Integer = 0
            For i = 0 To x.Length - 2

                If i < j Then
                    TempX = x(i)
                    x(i) = x(j)
                    x(j) = TempX

                    TempY = y(i)
                    y(i) = y(j)
                    y(j) = TempY

                End If

                Dim m As Integer = x.Length

                Do
                    m >>= 1
                    j = j Xor m
                Loop While (j And m) = 0

            Next

        End If

        'Defining some temporary variables to avoid definition inside loop
        Dim aiX As Double
        Dim aiY As Double
        Dim Real1 As Double
        Dim Imaginary1 As Double
        Dim Real2 As Double
        Dim Imaginary2 As Double
        Dim TempReal1 As Double

        Dim LevelSize As Integer = 1
        While LevelSize < x.Length

            Dim StepSize = LevelSize << 1

            For k = 0 To LevelSize - 1

                Dim exponent = (ExponentSign * k) * Math.PI / LevelSize
                Dim wX As Double = Math.Cos(exponent) ' N.B. this step of the algorithm suffers from the inexact floating point numbers returned from the trigonometric functions Cos and Sin
                Dim wY As Double = Math.Sin(exponent)

                Dim i As Integer = k
                While i < x.Length - 1

                    aiX = x(i)
                    aiY = y(i)

                    Real1 = wX
                    Imaginary1 = wY
                    Real2 = x(i + LevelSize)
                    Imaginary2 = y(i + LevelSize)

                    'Complex multiplication
                    TempReal1 = Real1
                    Real1 = TempReal1 * Real2 - Imaginary1 * Imaginary2
                    Imaginary1 = TempReal1 * Imaginary2 + Imaginary1 * Real2

                    x(i) = aiX + Real1
                    y(i) = aiY + Imaginary1

                    x(i + LevelSize) = aiX - Real1
                    y(i + LevelSize) = aiY - Imaginary1

                    i += StepSize

                End While

            Next

            LevelSize *= 2

        End While

        'Scaling
        If Direction = FftDirections.Forward And ScaleForwardTransform = True Then
            Dim scalingFactor = 1.0 / x.Length
            For i = 0 To x.Length - 1
                x(i) *= scalingFactor
                y(i) *= scalingFactor
            Next
        End If

    End Sub


    ''' <summary>
    ''' Crops the sound (removes all sound before StartSample, and after StartSample + Length).
    ''' </summary>
    ''' <param name="inputSound"></param>
    ''' <param name="startSample"></param>
    ''' <param name="sectionLength"></param>
    Public Shared Sub CropSection(ByRef InputSound As Sound,
                     Optional ByVal StartSample As Integer = 0,
                           Optional ByVal SectionLength As Integer? = Nothing)

        Try

            'Main section
            For c = 1 To InputSound.WaveFormat.Channels

                Dim InputSoundArray() As Single = InputSound.WaveData.SampleData(c)

                Dim CorrectedStartSample = StartSample
                Dim CorrectedSectionLength = SectionLength
                CheckAndCorrectSectionLength(InputSoundArray.Length, CorrectedStartSample, CorrectedSectionLength)

                'Creating a new sound array
                Dim NewSoundArray(CorrectedSectionLength - 1) As Single

                'Copies the sound
                For s = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1
                    NewSoundArray(s - CorrectedStartSample) = InputSoundArray(s)
                Next

                'Replacing the sound array with the cropped sound array
                InputSound.WaveData.SampleData(c) = NewSoundArray

            Next

        Catch ex As Exception
            AudioError(ex.ToString)
        End Try

    End Sub


    ''' <summary>
    ''' Copies a selected section of the sound to a new Sound.
    ''' </summary>
    ''' <param name="inputSound"></param>
    ''' <param name="startSample"></param>
    ''' <param name="sectionLength"></param>
    Public Shared Function CopySection(ByRef InputSound As Sound, Optional ByVal StartSample As Integer = 0,
                                Optional ByVal SectionLength As Integer? = Nothing,
                                Optional ByVal Channel As Integer? = Nothing) As Sound

        Try

            If Channel Is Nothing Then

                Dim Output As New Sound(InputSound.WaveFormat)

                'Main section
                For c = 1 To InputSound.WaveFormat.Channels

                    Dim InputSoundArray() As Single = InputSound.WaveData.SampleData(c)

                    Dim CorrectedStartSample = StartSample
                    Dim CorrectedSectionLength = SectionLength
                    CheckAndCorrectSectionLength(InputSoundArray.Length, CorrectedStartSample, CorrectedSectionLength)

                    'Creating a new sound array
                    Dim NewSoundArray(CorrectedSectionLength - 1) As Single

                    'Copies the sound
                    For s = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1
                        NewSoundArray(s - CorrectedStartSample) = InputSoundArray(s)
                    Next

                    'Replacing the sound array with the cropped sound array
                    Output.WaveData.SampleData(c) = NewSoundArray

                Next

                Return Output


            Else

                Dim Output As New Sound(New Formats.WaveFormat(InputSound.WaveFormat.SampleRate, InputSound.WaveFormat.BitDepth, 1,, InputSound.WaveFormat.Encoding))

                'Main section
                Dim InputSoundArray() As Single = InputSound.WaveData.SampleData(Channel)

                Dim CorrectedStartSample = StartSample
                Dim CorrectedSectionLength = SectionLength
                CheckAndCorrectSectionLength(InputSoundArray.Length, CorrectedStartSample, CorrectedSectionLength)

                'Creating a new sound array
                Dim NewSoundArray(CorrectedSectionLength - 1) As Single

                'Copies the sound
                For s = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1
                    NewSoundArray(s - CorrectedStartSample) = InputSoundArray(s)
                Next

                'Referencing the new sound array in the output sound
                Output.WaveData.SampleData(1) = NewSoundArray

                Return Output

            End If


        Catch ex As Exception
            AudioError(ex.ToString)
            Return Nothing
        End Try

    End Function


    ''' <summary>
    ''' Amplifys a section of the sound.
    ''' </summary>
    ''' <param name="InputSound">The sound to modify.</param>
    ''' <param name="Gain">The amount of gain applied to the specified section (gain unit is set to dB or linear with the parameter GainUnit).</param>
    ''' <param name="Channel">The channel to be modified. If left empty, all channels will be modified.</param>
    ''' <param name="StartSample">Start sample of the section to be amplified.</param>
    ''' <param name="SectionLength">Length (in samples) of the section to be amplified.</param>
    ''' <param name="GainUnit">The unit of the gain paramameter (dB or linear)</param>
    Public Shared Sub AmplifySection(ByRef InputSound As Sound, ByVal Gain As Double,
                                   Optional ByVal Channel As Integer? = Nothing,
                                   Optional ByVal StartSample As Integer = 0, Optional ByVal SectionLength As Integer? = Nothing,
                                   Optional ByVal GainUnit As SoundDataUnit = SoundDataUnit.dB)

        Try

            Dim AudioOutputConstructor As New AudioOutputConstructor(InputSound.WaveFormat, Channel)

            'Main section
            For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                Dim gainFactor As Double = 0

                Select Case GainUnit
                    Case SoundDataUnit.dB
                        gainFactor = 10 ^ (Gain / 20)

                    Case SoundDataUnit.linear
                        gainFactor = Gain

                    Case Else
                        Throw New NotImplementedException("Unsupported SoundDataUnit. Use either SoundDataUnit.dB Or .Linear.")
                End Select


                'Verkställer förstärkningen i sektionen (och varnar för distorsion)
                Dim otherErrorExeption As Exception = Nothing

                Dim SoundArray = InputSound.WaveData.SampleData(c)

                Dim CorrectedStartSample = StartSample
                Dim CorrectedSectionLength = SectionLength
                CheckAndCorrectSectionLength(SoundArray.Length, CorrectedStartSample, CorrectedSectionLength)

                If Globals.StfBase.UseOptimizationLibraries = False Then
                    MultiplyArray(SoundArray, gainFactor, CorrectedStartSample, CorrectedSectionLength)

                Else
                    ''Calling libostfdsp function multiplyFloatArraySection
                    LibOstfDsp_VB.MultiplyArraySection(SoundArray, gainFactor, CorrectedStartSample, CorrectedSectionLength)

                End If

            Next


        Catch ex As Exception
            AudioError(ex.ToString)
        End Try

    End Sub


    Public Shared Function IIRFilter(ByVal inputSound As Sound, FrequencyWeighting As FrequencyWeightings,
                              Optional ByVal channelToFilter As Integer? = Nothing) As Sound

        Select Case FrequencyWeighting
            Case FrequencyWeightings.Z
                'Returns input sound straight away if Z-Weighting is chosen
                Return inputSound

            Case FrequencyWeightings.K

                'Since K-Weighting is done with 2 subsequent filters a special code path with 2 filters is made here.

                'Filter 1
                Dim ACoefficients() As Double = {1, -1.69065929318241, 0.73248077421585}
                Dim BCoefficients() As Double = {1.53512485958697, -2.69169618940638, 1.19839281085285}
                Dim GainIn_dB As Double = 0
                Dim tempSound As Sound = IIR(inputSound, ACoefficients, BCoefficients, GainIn_dB, channelToFilter)

                'Filter 2
                ACoefficients = Nothing
                BCoefficients = Nothing
                GainIn_dB = 0
                If Set_IIR_FrequencyWeightingCoefficients(inputSound.WaveFormat, FrequencyWeightings.RLB, ACoefficients, BCoefficients, GainIn_dB) = False Then
                    Return Nothing
                End If
                Dim outputSound As Sound = IIR(tempSound, ACoefficients, BCoefficients, GainIn_dB, channelToFilter)

                If Not outputSound Is Nothing Then
                    Return outputSound
                Else
                    Return Nothing
                End If


            Case Else

                'For all other filter types

                Dim ACoefficients() As Double = Nothing
                Dim BCoefficients() As Double = Nothing

                Dim GainIn_dB As Double = 0

                If Set_IIR_FrequencyWeightingCoefficients(inputSound.WaveFormat, FrequencyWeighting, ACoefficients, BCoefficients, GainIn_dB) = False Then
                    Return Nothing
                End If

                Dim outputSound As Sound = IIR(inputSound, ACoefficients, BCoefficients, GainIn_dB, channelToFilter)
                If Not outputSound Is Nothing Then
                    Return outputSound
                Else
                    Return Nothing
                End If


        End Select


    End Function


    Public Shared Function IIR(ByVal inputSound As Sound, ByVal ACoefficients() As Double, ByVal BCoefficients() As Double, Optional ByVal Gain As Double = 0,
                     Optional ByVal channelToFilter As Integer? = Nothing) As Sound

        Try

            'Built on the standard formula for recursive filters. For Ref se e.g. Kates Digital Hearing Aids (2008), p 34.

            Select Case inputSound.WaveFormat.BitDepth
                Case 16, 32
                Case Else
                    Throw New NotImplementedException("")
            End Select

            Dim outputSound As New Sound(inputSound.WaveFormat)
            Dim AudioOutputConstructor As New AudioOutputConstructor(inputSound.WaveFormat, channelToFilter)
            Dim FS As Double = inputSound.WaveFormat.PositiveFullScale

            'Main section

            Dim Zeroes As Double = 0
            Dim Poles As Double = 0

            Dim XMemory As New List(Of Double)
            For sample = 0 To BCoefficients.Length - 2
                XMemory.Add(0)
            Next

            Dim YMemory As New List(Of Double)
            For sample = 0 To ACoefficients.Length - 2
                YMemory.Add(0)
            Next

            Dim OutputSampleValue As Double = 0

            For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                Dim outputDoubleArray(inputSound.WaveData.SampleData(c).Length - 1) As Double
                Dim CurrentInputArray = inputSound.WaveData.SampleData(c)

                For sample = 0 To CurrentInputArray.Length - 1

                    'Summing up zeroes
                    Zeroes = BCoefficients(0) * CurrentInputArray(sample)
                    For index = 1 To BCoefficients.Count - 1
                        Zeroes += BCoefficients(index) * XMemory(index - 1)
                    Next

                    'Summing poles
                    Poles = 0
                    For index = 1 To ACoefficients.Count - 1
                        'Try 'Try-block commented out on 2019-09-03, The block should definitely not be necessary!
                        Poles += ACoefficients(index) * YMemory(index - 1)
                        'Catch ex As Exception
                        'MsgBox(ex.ToString)
                        'End Try
                    Next

                    OutputSampleValue = (1 / ACoefficients(0)) * (Zeroes - Poles)
                    outputDoubleArray(sample) = OutputSampleValue

                    'Adjusting the memories
                    XMemory.Insert(0, CurrentInputArray(sample))
                    YMemory.Insert(0, OutputSampleValue)
                    XMemory.RemoveAt(XMemory.Count - 1)
                    YMemory.RemoveAt(YMemory.Count - 1)

                Next

                'Applying gain and Convering to single
                Dim OutputSingleArray(outputDoubleArray.Length - 1) As Single
                For sample = 0 To outputDoubleArray.Length - 1
                    OutputSingleArray(sample) = CSng(outputDoubleArray(sample)) '*CSng(Gain) 
                Next
                outputSound.WaveData.SampleData(c) = OutputSingleArray

                'Appyling gain
                If Not Gain = 0 Then
                    DSP.AmplifySection(outputSound, Gain, c,,,)
                End If

            Next

            Return outputSound

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function



    Private Shared Function Set_IIR_FrequencyWeightingCoefficients(ByVal soundFormat As Formats.WaveFormat, ByVal FrequencyWeighting As FrequencyWeightings, ByRef ACoefficients() As Double, ByRef BCoefficients() As Double, Optional ByRef GainIn_dB As Double = 1) As Boolean

        Select Case FrequencyWeighting
            Case FrequencyWeightings.A

                'Source: Rimell, A. N., et al. (2015). "Design of digital filters for frequency weightings (A and C) required for risk assessments of workers exposed to noise." Industrial Health 53(1): 21-27.

                Dim f1 As Double = 20.598997
                Dim f2 As Double = 107.65265
                Dim f3 As Double = 737.86223
                Dim f4 As Double = 12194.217

                Dim fs As Double = soundFormat.SampleRate

                'Dim PI As Double = 3.1415926535897932384626433832D
                'Dim PI As Double = Math.Acos(-1.0)
                Dim PI As Double = Math.PI

                'W1 represeents ω1′
                'W4 represeents ω4′

                'Dim W1 As Double = 2D * Math.Tan(PI * (f1 / fs))
                'Dim W2 As Double = 2D * Math.Tan(PI * (f2 / fs))
                'Dim W3 As Double = 2D * Math.Tan(PI * (f3 / fs))
                'Dim W4 As Double = 2D * Math.Tan(PI * (f4 / fs))

                Dim W1 As Double = 2 * Math.Tan(PI * (f1 / fs))
                Dim W2 As Double = 2 * Math.Tan(PI * (f2 / fs))
                Dim W3 As Double = 2 * Math.Tan(PI * (f3 / fs))
                Dim W4 As Double = 2 * Math.Tan(PI * (f4 / fs))


                ReDim ACoefficients(10)
                ReDim BCoefficients(10)

                ACoefficients(0) = 64 + (16 * W2 * W1 * W3) + (4 * W2 * W1 ^ 2 * W3) + (32 * W2 * W1 * W4) + (8 * W2 * W1 ^ 2 * W4) + (32 * W1 * W3 * W4) + (16 * W2 * W3 * W4) + (64 * W1) + (32 * W2) + (32 * W3) + (64 * W4) + (32 * W2 * W1) + (8 * W2 * W1 ^ 2) + (16 * W1 ^ 2) + (16 * W2 * W1 * W3 * W4) + (4 * W2 * W1 ^ 2 * W3 * W4) + (32 * W1 * W3) + (16 * W2 * W3) + (8 * W1 ^ 2 * W3) + (64 * W1 * W4) + (32 * W2 * W4) + (32 * W3 * W4) + (16 * W1 ^ 2 * W4) + (8 * W1 ^ 2 * W3 * W4) + (16 * W4 ^ 2) + (16 * W4 ^ 2 * W1) + (4 * W4 ^ 2 * W1 ^ 2) + (4 * W4 ^ 2 * W1 * W2 * W3) + (W4 ^ 2 * W1 ^ 2 * W2 * W3) + (8 * W4 ^ 2 * W1 * W2) + (2 * W4 ^ 2 * W1 ^ 2 * W2) + (8 * W4 ^ 2 * W2) + (8 * W4 ^ 2 * W3) + (8 * W4 ^ 2 * W1 * W3) + (2 * W4 ^ 2 * W1 ^ 2 * W3) + (4 * W4 ^ 2 * W2 * W3)
                ACoefficients(1) = -128 + (64 * W2 * W1 * W3) + (24 * W2 * W1 ^ 2 * W3) + (128 * W2 * W1 * W4) + (48 * W2 * W1 ^ 2 * W4) + (128 * W1 * W3 * W4) + (64 * W2 * W3 * W4) + (64 * W2 * W1) + (32 * W2 * W1 ^ 2) + (32 * W1 ^ 2) + (96 * W2 * W1 * W3 * W4) + (32 * W2 * W1 ^ 2 * W3 * W4) + (64 * W1 * W3) + (32 * W2 * W3) + (32 * W1 ^ 2 * W3) + (128 * W1 * W4) + (64 * W2 * W4) + (64 * W3 * W4) + (64 * W1 ^ 2 * W4) + (48 * W1 ^ 2 * W3 * W4) + (32 * W4 ^ 2) + (64 * W4 ^ 2 * W1) + (24 * W4 ^ 2 * W1 ^ 2) + (32 * W4 ^ 2 * W1 * W2 * W3) + (10 * W4 ^ 2 * W1 ^ 2 * W2 * W3) + (48 * W4 ^ 2 * W1 * W2) + (16 * W4 ^ 2 * W1 ^ 2 * W2) + (32 * W4 ^ 2 * W2) + (32 * W4 ^ 2 * W3) + (48 * W4 ^ 2 * W1 * W3) + (16 * W4 ^ 2 * W1 ^ 2 * W3) + (24 * W4 ^ 2 * W2 * W3)
                ACoefficients(2) = -192 + (48 * W2 * W1 * W3) + (52 * W2 * W1 ^ 2 * W3) + (96 * W2 * W1 * W4) + (104 * W2 * W1 ^ 2 * W4) + (96 * W1 * W3 * W4) + (48 * W2 * W3 * W4) - (320 * W1) - (160 * W2) - (160 * W3) - (320 * W4) - (96 * W2 * W1) + (24 * W2 * W1 ^ 2) - (48 * W1 ^ 2) + (208 * W2 * W1 * W3 * W4) + (108 * W2 * W1 ^ 2 * W3 * W4) - (96 * W1 * W3) - (48 * W2 * W3) + (24 * W1 ^ 2 * W3) - (192 * W1 * W4) - (96 * W2 * W4) - (96 * W3 * W4) + (48 * W1 ^ 2 * W4) + (104 * W1 ^ 2 * W3 * W4) - (48 * W4 ^ 2) + (48 * W4 ^ 2 * W1) + (52 * W4 ^ 2 * W1 ^ 2) + (108 * W4 ^ 2 * W1 * W2 * W3) + (45 * W4 ^ 2 * W1 ^ 2 * W2 * W3) + (104 * W4 ^ 2 * W1 * W2) + (54 * W4 ^ 2 * W1 ^ 2 * W2) + (24 * W4 ^ 2 * W2) + (24 * W4 ^ 2 * W3) + (104 * W4 ^ 2 * W1 * W3) + (54 * W4 ^ 2 * W1 ^ 2 * W3) + (52 * W4 ^ 2 * W2 * W3)
                ACoefficients(3) = 512 - (128 * W2 * W1 * W3) + (32 * W2 * W1 ^ 2 * W3) - (256 * W2 * W1 * W4) + (64 * W2 * W1 ^ 2 * W4) - (256 * W1 * W3 * W4) - (128 * W2 * W3 * W4) - (256 * W2 * W1) - (64 * W2 * W1 ^ 2) - (128 * W1 ^ 2) + (128 * W2 * W1 * W3 * W4) + (192 * W2 * W1 ^ 2 * W3 * W4) - (256 * W1 * W3) - (128 * W2 * W3) - (64 * W1 ^ 2 * W3) - (512 * W1 * W4) - (256 * W2 * W4) - (256 * W3 * W4) - (128 * W1 ^ 2 * W4) + (64 * W1 ^ 2 * W3 * W4) - (128 * W4 ^ 2) - (128 * W4 ^ 2 * W1) + (32 * W4 ^ 2 * W1 ^ 2) + (192 * W4 ^ 2 * W1 * W2 * W3) + (120 * W4 ^ 2 * W1 ^ 2 * W2 * W3) + (64 * W4 ^ 2 * W1 * W2) + (96 * W4 ^ 2 * W1 ^ 2 * W2) - (64 * W4 ^ 2 * W2) - (64 * W4 ^ 2 * W3) + (64 * W4 ^ 2 * W1 * W3) + (96 * W4 ^ 2 * W1 ^ 2 * W3) + (32 * W4 ^ 2 * W2 * W3)
                ACoefficients(4) = 128 - (224 * W2 * W1 * W3) - (56 * W2 * W1 ^ 2 * W3) - (448 * W2 * W1 * W4) - (112 * W2 * W1 ^ 2 * W4) - (448 * W1 * W3 * W4) - (224 * W2 * W3 * W4) + (640 * W1) + (320 * W2) + (320 * W3) + (640 * W4) + (64 * W2 * W1) - (112 * W2 * W1 ^ 2) + (32 * W1 ^ 2) - (224 * W2 * W1 * W3 * W4) + (168 * W2 * W1 ^ 2 * W3 * W4) + (64 * W1 * W3) + (32 * W2 * W3) - (112 * W1 ^ 2 * W3) + (128 * W1 * W4) + (64 * W2 * W4) + (64 * W3 * W4) - (224 * W1 ^ 2 * W4) - (112 * W1 ^ 2 * W3 * W4) + (32 * W4 ^ 2) - (224 * W4 ^ 2 * W1) - (56 * W4 ^ 2 * W1 ^ 2) + (168 * W4 ^ 2 * W1 * W2 * W3) + (210 * W4 ^ 2 * W1 ^ 2 * W2 * W3) - (112 * W4 ^ 2 * W1 * W2) + (84 * W4 ^ 2 * W1 ^ 2 * W2) - (112 * W4 ^ 2 * W2) - (112 * W4 ^ 2 * W3) - (112 * W4 ^ 2 * W1 * W3) + (84 * W4 ^ 2 * W1 ^ 2 * W3) - (56 * W4 ^ 2 * W2 * W3)
                ACoefficients(5) = -(448 * W2 * W1 * W3 * W4) - (224 * W1 ^ 2 * W3 * W4) + (384 * W3 * W4) - (112 * W2 * W1 ^ 2 * W3) - (112 * W4 ^ 2 * W1 ^ 2) + (384 * W1 * W3) - (224 * W4 ^ 2 * W1 * W3) + (192 * W2 * W3) - (224 * W2 * W1 ^ 2 * W4) + (192 * W1 ^ 2) + (252 * W4 ^ 2 * W1 ^ 2 * W2 * W3) + (384 * W2 * W1) - (768) - (224 * W4 ^ 2 * W1 * W2) - (112 * W4 ^ 2 * W2 * W3) + (384 * W2 * W4) + (192 * W4 ^ 2) + (768 * W1 * W4)
                ACoefficients(6) = 128 + (224 * W2 * W1 * W3) - (56 * W2 * W1 ^ 2 * W3) + (448 * W2 * W1 * W4) - (112 * W2 * W1 ^ 2 * W4) + (448 * W1 * W3 * W4) + (224 * W2 * W3 * W4) - (640 * W1) - (320 * W2) - (320 * W3) - (640 * W4) + (64 * W2 * W1) + (112 * W2 * W1 ^ 2) + (32 * W1 ^ 2) - (224 * W2 * W1 * W3 * W4) - (168 * W2 * W1 ^ 2 * W3 * W4) + (64 * W1 * W3) + (32 * W2 * W3) + (112 * W1 ^ 2 * W3) + (128 * W1 * W4) + (64 * W2 * W4) + (64 * W3 * W4) + (224 * W1 ^ 2 * W4) - (112 * W1 ^ 2 * W3 * W4) + (32 * W4 ^ 2) + (224 * W4 ^ 2 * W1) - (56 * W4 ^ 2 * W1 ^ 2) - (168 * W4 ^ 2 * W1 * W2 * W3) + (210 * W4 ^ 2 * W1 ^ 2 * W2 * W3) - (112 * W4 ^ 2 * W1 * W2) - (84 * W4 ^ 2 * W1 ^ 2 * W2) + (112 * W4 ^ 2 * W2) + (112 * W4 ^ 2 * W3) - (112 * W4 ^ 2 * W1 * W3) - (84 * W4 ^ 2 * W1 ^ 2 * W3) - (56 * W4 ^ 2 * W2 * W3)
                ACoefficients(7) = 512 + (128 * W2 * W1 * W3) + (32 * W2 * W1 ^ 2 * W3) + (256 * W2 * W1 * W4) + (64 * W2 * W1 ^ 2 * W4) + (256 * W1 * W3 * W4) + (128 * W2 * W3 * W4) - (256 * W2 * W1) + (64 * W2 * W1 ^ 2) - (128 * W1 ^ 2) + (128 * W2 * W1 * W3 * W4) - (192 * W2 * W1 ^ 2 * W3 * W4) - (256 * W1 * W3) - (128 * W2 * W3) + (64 * W1 ^ 2 * W3) - (512 * W1 * W4) - (256 * W2 * W4) - (256 * W3 * W4) + (128 * W1 ^ 2 * W4) + (64 * W1 ^ 2 * W3 * W4) - (128 * W4 ^ 2) + (128 * W4 ^ 2 * W1) + (32 * W4 ^ 2 * W1 ^ 2) - (192 * W4 ^ 2 * W1 * W2 * W3) + (120 * W4 ^ 2 * W1 ^ 2 * W2 * W3) + (64 * W4 ^ 2 * W1 * W2) - (96 * W4 ^ 2 * W1 ^ 2 * W2) + (64 * W4 ^ 2 * W2) + (64 * W4 ^ 2 * W3) + (64 * W4 ^ 2 * W1 * W3) - (96 * W4 ^ 2 * W1 ^ 2 * W3) + (32 * W4 ^ 2 * W2 * W3)
                ACoefficients(8) = -192 - (48 * W2 * W1 * W3) + (52 * W2 * W1 ^ 2 * W3) - (96 * W2 * W1 * W4) + (104 * W2 * W1 ^ 2 * W4) - (96 * W1 * W3 * W4) - (48 * W2 * W3 * W4) + (320 * W1) + (160 * W2) + (160 * W3) + (320 * W4) - (96 * W2 * W1) - (24 * W2 * W1 ^ 2) - (48 * W1 ^ 2) + (208 * W2 * W1 * W3 * W4) - (108 * W2 * W1 ^ 2 * W3 * W4) - (96 * W1 * W3) - (48 * W2 * W3) - (24 * W1 ^ 2 * W3) - (192 * W1 * W4) - (96 * W2 * W4) - (96 * W3 * W4) - (48 * W1 ^ 2 * W4) + (104 * W1 ^ 2 * W3 * W4) - (48 * W4 ^ 2) - (48 * W4 ^ 2 * W1) + (52 * W4 ^ 2 * W1 ^ 2) - (108 * W4 ^ 2 * W1 * W2 * W3) + (45 * W4 ^ 2 * W1 ^ 2 * W2 * W3) + (104 * W4 ^ 2 * W1 * W2) - (54 * W4 ^ 2 * W1 ^ 2 * W2) - (24 * W4 ^ 2 * W2) - (24 * W4 ^ 2 * W3) + (104 * W4 ^ 2 * W1 * W3) - (54 * W4 ^ 2 * W1 ^ 2 * W3) + (52 * W4 ^ 2 * W2 * W3)
                ACoefficients(9) = -128 - (64 * W2 * W1 * W3) + (24 * W2 * W1 ^ 2 * W3) - (128 * W2 * W1 * W4) + (48 * W2 * W1 ^ 2 * W4) - (128 * W1 * W3 * W4) - (64 * W2 * W3 * W4) + (64 * W2 * W1) - (32 * W2 * W1 ^ 2) + (32 * W1 ^ 2) + (96 * W2 * W1 * W3 * W4) - (32 * W2 * W1 ^ 2 * W3 * W4) + (64 * W1 * W3) + (32 * W2 * W3) - (32 * W1 ^ 2 * W3) + (128 * W1 * W4) + (64 * W2 * W4) + (64 * W3 * W4) - (64 * W1 ^ 2 * W4) + (48 * W1 ^ 2 * W3 * W4) + (32 * W4 ^ 2) - (64 * W4 ^ 2 * W1) + (24 * W4 ^ 2 * W1 ^ 2) - (32 * W4 ^ 2 * W1 * W2 * W3) + (10 * W4 ^ 2 * W1 ^ 2 * W2 * W3) + (48 * W4 ^ 2 * W1 * W2) - (16 * W4 ^ 2 * W1 ^ 2 * W2) - (32 * W4 ^ 2 * W2) - (32 * W4 ^ 2 * W3) + (48 * W4 ^ 2 * W1 * W3) - (16 * W4 ^ 2 * W1 ^ 2 * W3) + (24 * W4 ^ 2 * W2 * W3)
                ACoefficients(10) = 64 - (16 * W2 * W1 * W3) + (4 * W2 * W1 ^ 2 * W3) - (32 * W2 * W1 * W4) + (8 * W2 * W1 ^ 2 * W4) - (32 * W1 * W3 * W4) - (16 * W2 * W3 * W4) - (64 * W1) - (32 * W2) - (32 * W3) - (64 * W4) + (32 * W2 * W1) - (8 * W2 * W1 ^ 2) + (16 * W1 ^ 2) + (16 * W2 * W1 * W3 * W4) - (4 * W2 * W1 ^ 2 * W3 * W4) + (32 * W1 * W3) + (16 * W2 * W3) - (8 * W1 ^ 2 * W3) + (64 * W1 * W4) + (32 * W2 * W4) + (32 * W3 * W4) - (16 * W1 ^ 2 * W4) + (8 * W1 ^ 2 * W3 * W4) + (16 * W4 ^ 2) - (16 * W4 ^ 2 * W1) + (4 * W4 ^ 2 * W1 ^ 2) - (4 * W4 ^ 2 * W1 * W2 * W3) + (W4 ^ 2 * W1 ^ 2 * W2 * W3) + (8 * W4 ^ 2 * W1 * W2) - (2 * W4 ^ 2 * W1 ^ 2 * W2) - (8 * W4 ^ 2 * W2) - (8 * W4 ^ 2 * W3) + (8 * W4 ^ 2 * W1 * W3) - (2 * W4 ^ 2 * W1 ^ 2 * W3) + (4 * W4 ^ 2 * W2 * W3)
                BCoefficients(0) = 16 * W4 ^ 2
                BCoefficients(1) = 32 * W4 ^ 2
                BCoefficients(2) = -48 * W4 ^ 2
                BCoefficients(3) = -128 * W4 ^ 2
                BCoefficients(4) = 32 * W4 ^ 2
                BCoefficients(5) = 192 * W4 ^ 2
                BCoefficients(6) = 32 * W4 ^ 2
                BCoefficients(7) = -128 * W4 ^ 2
                BCoefficients(8) = -48 * W4 ^ 2
                BCoefficients(9) = 32 * W4 ^ 2
                BCoefficients(10) = 16 * W4 ^ 2
                'Gain = 10.0 ^ (2.0 / 20.0)
                GainIn_dB = -2

                Return True

            Case FrequencyWeightings.C

                'Source: Rimell, A. N., et al. (2015). "Design of digital filters for frequency weightings (A and C) required for risk assessments of workers exposed to noise." Industrial Health 53(1): 21-27.

                Dim f1 As Double = 20.598997
                Dim f2 As Double = 107.65265
                Dim f3 As Double = 737.86223
                Dim f4 As Double = 12194.217

                Dim fs As Double = soundFormat.SampleRate

                'Dim PI As double = 3.1415926535897932384626433832D
                'Dim PI As double = Math.Acos(-1.0)
                Dim PI As Double = Math.PI

                'W1 represeents ω1′
                'W4 represeents ω4′

                Dim W1 As Double = 2D * Math.Tan(PI * (f1 / fs))
                Dim W2 As Double = 2D * Math.Tan(PI * (f2 / fs))
                Dim W3 As Double = 2D * Math.Tan(PI * (f3 / fs))
                Dim W4 As Double = 2D * Math.Tan(PI * (f4 / fs))



                Try

                    ReDim ACoefficients(6)
                    ReDim BCoefficients(6)

                    ACoefficients(0) = (16 * W4) + (16 * W1) + (4 * W1 ^ 2) + (16 * W4 * W1) + (4 * W4 * W1 ^ 2) + (W4 ^ 2 * W1 ^ 2) + (4 * W4 ^ 2) + 16 + (4 * W4 ^ 2 * W1)
                    ACoefficients(1) = (8 * W4 ^ 2) - 32 + (6 * W4 ^ 2 * W1 ^ 2) + (32 * W4 * W1) + (16 * W4 * W1 ^ 2) + (8 * W1 ^ 2) + (16 * W4 ^ 2 * W1)
                    ACoefficients(2) = -16 + (20 * W4 * W1 ^ 2) - (48 * W4) - (4 * W4 ^ 2) - (4 * W1 ^ 2) + (15 * W4 ^ 2 * W1 ^ 2) - (48 * W1) + (20 * W4 ^ 2 * W1) - (16 * W4 * W1)
                    ACoefficients(3) = 64 - (16 * W1 ^ 2) - (64 * W4 * W1) + (20 * W4 ^ 2 * W1 ^ 2) - (16 * W4 ^ 2)
                    ACoefficients(4) = -(4 * W1 ^ 2) + (15 * W4 ^ 2 * W1 ^ 2) - (4 * W4 ^ 2) + (48 * W1) - (16 * W4 * W1) - (20 * W4 * W1 ^ 2) - 16 - (20 * W4 ^ 2 * W1) + (48 * W4)
                    ACoefficients(5) = (8 * W4 ^ 2) - (16 * W4 ^ 2 * W1) + (8 * W1 ^ 2) + (32 * W4 * W1) - (16 * W4 * W1 ^ 2) + (6 * W4 ^ 2 * W1 ^ 2) - 32
                    ACoefficients(6) = -(4 * W4 * W1 ^ 2) + (16 * W4 * W1) + (W4 ^ 2 * W1 ^ 2) - (4 * W4 ^ 2 * W1) + 16 - (16 * W1) - (16 * W4) + (4 * W4 ^ 2) + (4 * W1 ^ 2)
                    BCoefficients(0) = 4 * W4 ^ 2
                    BCoefficients(1) = 8 * W4 ^ 2
                    BCoefficients(2) = -4 * W4 ^ 2
                    BCoefficients(3) = -16 * W4 ^ 2
                    BCoefficients(4) = -4 * W4 ^ 2
                    BCoefficients(5) = 8 * W4 ^ 2
                    BCoefficients(6) = 4 * W4 ^ 2
                    'Gain = 10 ^ (-0.062 / 20)
                    GainIn_dB = 0.062

                    'case ITU

                    Return True
                Catch ex As Exception
                    Return False
                End Try

            Case FrequencyWeightings.RLB
                ReDim ACoefficients(2)
                ReDim BCoefficients(2)

                ACoefficients(0) = 1.0
                ACoefficients(1) = -1.99004745483398
                ACoefficients(2) = 0.99007225036621
                BCoefficients(0) = 1.0
                BCoefficients(1) = -2.0
                BCoefficients(2) = 1.0

                Return True

            Case Else
                Throw New NotImplementedException("Non implemented frequency weigthing.")

        End Select

    End Function

    ''' <summary>
    ''' Concatenates the input sounds. All input sounds must have the same format. They may however differ in number of channels, in which case the output file will contain the lowest number of channels among the input sounds. Data in channels higher than the lowest channel count will be ignored.
    ''' </summary>
    ''' <param name="InputSounds">A list of Sound to concatenate.</param>
    ''' <param name="EqualizeSoundLevel">If set to true, all sounds will be set to EqualizationLevel before conatenation.</param>
    ''' <param name="EqualizationLevel"></param>
    ''' <param name="EqualizationLevelFrequencyWeighting">The frequency weighting used in the sound level equalization measurement.</param>
    ''' <param name="AllowChangingInputSounds">Set to false to work on copies of the input sounds. This will require more memory use!</param>
    ''' <param name="CrossFadeLength">The length (in sample) of a cross-fade section.</param>
    ''' <param name="ThrowOnUnequalNominalLevels">If True, checks to ensure that all nominal levels stored in the SMA object of each sound are the same (except if they are Nothing, then they are ignored)</param>
    ''' <returns></returns>
    Public Shared Function ConcatenateSounds2(ByRef InputSounds As List(Of Sound),
                                           Optional ByVal CrossFadeLength As Integer? = Nothing,
                                           Optional ByVal SkewedFade As Boolean = False,
                                           Optional ByVal CosinePower As Double = 10,
                                           Optional ByVal EqualPower As Boolean = True,
                                           Optional ByVal CheckForDifferentSoundFormats As Boolean = True,
                                           Optional ByVal ThrowOnUnequalNominalLevels As Boolean = True) As Sound

        Try

            'Returning nothing if there are no input sounds
            If InputSounds.Count = 0 Then Return Nothing

            'Returning a single input sound directly
            If InputSounds.Count = 1 Then Return InputSounds(0)

            Dim DetectedNominalLevel As Double? = Nothing
            Dim NominalLevelList As New List(Of Double)
            For Each Sound In InputSounds
                If Sound.SMA IsNot Nothing Then
                    If Sound.SMA.NominalLevel.HasValue Then
                        NominalLevelList.Add(Sound.SMA.NominalLevel.Value)
                        DetectedNominalLevel = Sound.SMA.NominalLevel.Value
                    End If
                End If
            Next
            If ThrowOnUnequalNominalLevels Then
                If NominalLevelList.Count > 0 Then
                    For i = 0 To NominalLevelList.Count - 2
                        If NominalLevelList(i) <> NominalLevelList(i + 1) Then Throw New Exception("Unequal nominal levels detected in concatenated sound files! This may lead to unexpected sound levels during playback!")
                    Next
                End If
            End If

            If CheckForDifferentSoundFormats = True Then
                If InputSounds.Count > 1 Then
                    For n = 1 To InputSounds.Count - 1
                        If InputSounds(n).WaveFormat.SampleRate <> InputSounds(n - 1).WaveFormat.SampleRate Or
                                InputSounds(n).WaveFormat.BitDepth <> InputSounds(n - 1).WaveFormat.BitDepth Or
                                InputSounds(n).WaveFormat.Encoding <> InputSounds(n - 1).WaveFormat.Encoding Then
                            Throw New ArgumentException("Different formats in ConcatenateSounds input sounds. Aborting!")
                        End If
                    Next
                End If
            End If


            'Getting the length of the output sound
            Dim TotalOutputLength As Long = 0
            For Each InputSound In InputSounds
                'Getting length from channel 1 only
                If CrossFadeLength Is Nothing Then
                    TotalOutputLength += InputSound.WaveData.SampleData(1).Length
                Else
                    TotalOutputLength += InputSound.WaveData.SampleData(1).Length - CrossFadeLength
                    'Ensuring that TotalOutputLength (and later in the code below, WriteSampleIndex) is never negative
                    TotalOutputLength = System.Math.Max(0, TotalOutputLength)
                End If
            Next

            'Adjusting the length derived when crossfading, since the last sound is not faded
            If CrossFadeLength IsNot Nothing Then
                TotalOutputLength += CrossFadeLength
            End If

            'Getting the lowest number of channels
            Dim LowestNumberOfChannels As Integer = InputSounds(0).WaveFormat.Channels
            For n = 1 To InputSounds.Count - 1
                LowestNumberOfChannels = Math.Min(InputSounds(n).WaveFormat.Channels, LowestNumberOfChannels)
            Next

            'Creating an output sound format
            Dim OutputSoundWaveFormat As New Formats.WaveFormat(InputSounds(0).WaveFormat.SampleRate,
                                                                  InputSounds(0).WaveFormat.BitDepth,
                                                                  LowestNumberOfChannels,, InputSounds(0).WaveFormat.Encoding)

            'Creating an output sound
            Dim OutputSound As New Sound(OutputSoundWaveFormat)

            'Concatenating the sounds
            If CrossFadeLength Is Nothing Then

                'Creating sample arrays for each channel
                For c = 1 To OutputSound.WaveFormat.Channels
                    Dim NewChannelArray(TotalOutputLength - 1) As Single
                    OutputSound.WaveData.SampleData(c) = NewChannelArray
                Next

                Dim WriteSampleIndex As Long = 0
                For Each InputSound In InputSounds

                    For c = 1 To OutputSound.WaveFormat.Channels
                        Dim TargetArray = OutputSound.WaveData.SampleData(c)
                        Dim SourceArray = InputSound.WaveData.SampleData(c)
                        Array.Copy(SourceArray, 0, TargetArray, WriteSampleIndex, SourceArray.Length)
                    Next

                    'Increasing WriteSampleIndex after all channels have been copied, based on the channel 1 length (basicaly requires equal channel lengths in each incoming sound)
                    WriteSampleIndex += InputSound.WaveData.SampleData(1).Length

                Next

            Else

                'Fading input sounds
                For InputSoundIndex = 0 To InputSounds.Count - 1

                    Dim InputSound = InputSounds(InputSoundIndex)

                    'Fading beginnings (not of first sound)
                    If InputSoundIndex > 0 Then
                        Dim CrossInFadeSlopeType As FadeSlopeType = FadeSlopeType.Linear
                        If SkewedFade = True Then CrossInFadeSlopeType = FadeSlopeType.PowerCosine_SkewedInFadeDirection
                        DSP.Fade(InputSound, Nothing, 0,, 0, CrossFadeLength, CrossInFadeSlopeType, CosinePower, EqualPower)
                    End If

                    'Fading ends (not of last sound)
                    If InputSoundIndex < InputSounds.Count - 1 Then
                        Dim CrossOutFadeSlopeType As FadeSlopeType = FadeSlopeType.Linear
                        If SkewedFade = True Then CrossOutFadeSlopeType = FadeSlopeType.PowerCosine_SkewedFromFadeDirection
                        DSP.Fade(InputSound, 0, Nothing, , InputSound.WaveData.SampleData(1).Length - CrossFadeLength, CrossFadeLength, CrossOutFadeSlopeType, CosinePower, EqualPower)
                    End If

                Next

                'Copying the input sounds to their intended position in the output sound
                Dim WriteSampleIndex As Long = 0
                Dim OutputSoundLayers As New List(Of SortedList(Of Integer, Single()))
                For i = 0 To InputSounds.Count - 1
                    OutputSoundLayers.Add(New SortedList(Of Integer, Single()))
                    For c = 1 To OutputSound.WaveFormat.Channels
                        'Creating an empty array
                        Dim OutputSoundLayerChannelArray(TotalOutputLength - 1) As Single

                        'Getting the sound in channel c
                        Dim SourceArray = InputSounds(i).WaveData.SampleData(c)

                        'Copying the sound to the OutputSoundLayerChannelArray, from the intended position (WriteSampleIndex)
                        Array.Copy(SourceArray, 0, OutputSoundLayerChannelArray, WriteSampleIndex, SourceArray.Length)

                        'Adding the OutputSoundLayerChannelArray to OutputSoundLayers
                        OutputSoundLayers(i).Add(c, OutputSoundLayerChannelArray)

                    Next

                    'Increasing WriteSampleIndex after all channels have been copied, based on the channel 1 length (basicaly requires equal channel lengths in each incoming sound)
                    WriteSampleIndex += InputSounds(i).WaveData.SampleData(1).Length

                    'Moving the write position backwards
                    WriteSampleIndex -= CrossFadeLength

                    'Ensuring that WriteSampleIndex is non-negative
                    WriteSampleIndex = System.Math.Max(0, WriteSampleIndex)

                Next

                'Superpositioning the channel-layer arrays
                For c = 1 To OutputSound.WaveFormat.Channels
                    Dim OutputSoundChannelArray = OutputSoundLayers(0)(c)
                    For i = 1 To InputSounds.Count - 1

                        AddTwoArrays(OutputSoundChannelArray, OutputSoundLayers(i)(c))

                    Next
                    OutputSound.WaveData.SampleData(c) = OutputSoundChannelArray
                Next

            End If

            If DetectedNominalLevel.HasValue Then
                OutputSound.SMA.NominalLevel = DetectedNominalLevel.Value
            End If

            Return OutputSound

        Catch ex As Exception
            AudioError(ex.ToString)
            Return Nothing
        End Try

    End Function


    ''' <summary>
    ''' Concatenates the input sounds. All input sounds must have the same format. They may however differ in number of channels, in which case the output file will contain the lowest number of channels among the input sounds. Data in channels higher than the lowest channel count will be ignored.
    ''' </summary>
    ''' <param name="InputSounds">A list of Sound to concatenate.</param>
    ''' <param name="EqualizeSoundLevel">If set to true, all sounds will be set to EqualizationLevel before conatenation.</param>
    ''' <param name="EqualizationLevel"></param>
    ''' <param name="EqualizationLevelFrequencyWeighting">The frequency weighting used in the sound level equalization measurement.</param>
    ''' <param name="AllowChangingInputSounds">Set to false to work on copies of the input sounds. This will require more memory use!</param>
    ''' <param name="CrossFadeLength">The length (in sample) of a cross-fade section.</param>
    ''' <param name="ThrowOnUnequalNominalLevels">If True, checks to ensure that all nominal levels stored in the SMA object of each sound are the same (except if they are Nothing, then they are ignored)</param>
    ''' <returns></returns>
    Public Shared Function ConcatenateSounds(ByRef InputSounds As List(Of Sound),
                                      Optional ByVal EqualizeSoundLevel As Boolean = False,
                                      Optional ByVal EqualizationLevel As Double? = -40,
                                      Optional ByVal EqualizationLevelFrequencyWeighting As FrequencyWeightings = FrequencyWeightings.Z,
                                      Optional ByVal AllowChangingInputSounds As Boolean = True,
                                      Optional ByVal CheckForDifferentSoundFormats As Boolean = True,
                                      Optional ByVal CrossFadeLength As Integer? = Nothing,
                                      Optional ByVal SkewedFade As Boolean = False,
                                      Optional ByVal CosinePower As Double = 10,
                                      Optional ByVal EqualPower As Boolean = True,
                                          Optional ByVal ThrowOnUnequalNominalLevels As Boolean = True) As Sound

        Try

            Dim DetectedNominalLevel As Double? = Nothing
            Dim NominalLevelList As New List(Of Double)
            For Each Sound In InputSounds
                If Sound.SMA IsNot Nothing Then
                    If Sound.SMA.NominalLevel.HasValue Then
                        NominalLevelList.Add(Sound.SMA.NominalLevel.Value)
                        DetectedNominalLevel = Sound.SMA.NominalLevel.Value
                    End If
                End If
            Next
            If ThrowOnUnequalNominalLevels Then
                If NominalLevelList.Count > 0 Then
                    For i = 0 To NominalLevelList.Count - 2
                        If NominalLevelList(i) <> NominalLevelList(i + 1) Then Throw New Exception("Unequal nominal levels detected in concatenated sound files! This may lead to unexpected sound levels during playback!")
                    Next
                End If
            End If

            'Returning nothing if there are no input sounds
            If InputSounds.Count = 0 Then Return Nothing

            If CheckForDifferentSoundFormats = True Then
                If InputSounds.Count > 1 Then
                    For n = 1 To InputSounds.Count - 1
                        If InputSounds(n).WaveFormat.SampleRate <> InputSounds(n - 1).WaveFormat.SampleRate Or
                                InputSounds(n).WaveFormat.BitDepth <> InputSounds(n - 1).WaveFormat.BitDepth Or
                                InputSounds(n).WaveFormat.Encoding <> InputSounds(n - 1).WaveFormat.Encoding Then
                            Throw New ArgumentException("Different formats in ConcatenateSounds input sounds. Aborting!")
                        End If
                    Next
                End If
            End If

            'Getting the length of the output sound
            Dim TotalLength As Long = 0
            For Each InputSound In InputSounds
                'Getting length from channel 1 only
                If CrossFadeLength Is Nothing Then
                    TotalLength += InputSound.WaveData.SampleData(1).Length
                Else
                    TotalLength += InputSound.WaveData.SampleData(1).Length - CrossFadeLength
                End If
            Next

            'Adjusting the length derived when crossfading, since the last sound is not faded
            If CrossFadeLength IsNot Nothing Then
                TotalLength += CrossFadeLength
            End If


            'Getting the lowest number of channels
            Dim LowestNumberOfChannels As Integer = InputSounds(0).WaveFormat.Channels
            For n = 1 To InputSounds.Count - 1
                LowestNumberOfChannels = Math.Min(InputSounds(n).WaveFormat.Channels, LowestNumberOfChannels)
            Next

            'Creates copies of the input sounds if needed
            Dim SoundsToUse As List(Of Sound)
            If EqualizeSoundLevel = True And AllowChangingInputSounds = False Then
                SoundsToUse = New List(Of Sound)
                For Each InputSound In InputSounds
                    SoundsToUse.Add(InputSound.CreateCopy)
                Next
            Else
                SoundsToUse = InputSounds
            End If

            'Setting the sound level of each sound to TargetLevel
            If EqualizeSoundLevel = True Then
                'Setting all input sounds to EqualizationLevel
                For Each InputSound In SoundsToUse
                    DSP.MeasureAndAdjustSectionLevel(InputSound, EqualizationLevel,,,, EqualizationLevelFrequencyWeighting)
                Next
            End If


            'Creating an output sound
            Dim OutputSoundWaveFormat As New Formats.WaveFormat(SoundsToUse(0).WaveFormat.SampleRate,
                                                                  SoundsToUse(0).WaveFormat.BitDepth,
                                                                  LowestNumberOfChannels,, SoundsToUse(0).WaveFormat.Encoding)
            Dim OutputSound As New Sound(OutputSoundWaveFormat)
            'Creating sample arrays for each channel
            For c = 1 To OutputSound.WaveFormat.Channels
                Dim NewChannelArray(TotalLength - 1) As Single
                OutputSound.WaveData.SampleData(c) = NewChannelArray
            Next

            'Concatenating the sounds
            If CrossFadeLength Is Nothing Then

                Dim WriteSampleIndex As Long = 0
                For Each InputSound In SoundsToUse
                    For s = 0 To InputSound.WaveData.SampleData(1).Length - 1 'Using channel 1 lengths only
                        For c = 1 To OutputSound.WaveFormat.Channels
                            OutputSound.WaveData.SampleData(c)(WriteSampleIndex) = InputSound.WaveData.SampleData(c)(s)
                        Next
                        'Increasing WriteSampleIndex after all channels have been copied
                        WriteSampleIndex += 1
                    Next
                Next
            Else

                Dim WriteSampleIndex As Long = 0
                For InputSoundIndex = 0 To SoundsToUse.Count - 1

                    Dim InputSound = SoundsToUse(InputSoundIndex)

                    'Fading beginnings (not of first sound)
                    If InputSoundIndex > 0 Then
                        Dim CrossInFadeSlopeType As FadeSlopeType = FadeSlopeType.Linear
                        If SkewedFade = True Then CrossInFadeSlopeType = FadeSlopeType.PowerCosine_SkewedInFadeDirection
                        DSP.Fade(InputSound, Nothing, 0,, 0, CrossFadeLength, CrossInFadeSlopeType, CosinePower, EqualPower)
                    End If

                    'Fading ends (not of last sound)
                    If InputSoundIndex < SoundsToUse.Count - 1 Then
                        Dim CrossOutFadeSlopeType As FadeSlopeType = FadeSlopeType.Linear
                        If SkewedFade = True Then CrossOutFadeSlopeType = FadeSlopeType.PowerCosine_SkewedFromFadeDirection
                        DSP.Fade(InputSound, 0, Nothing, , InputSound.WaveData.SampleData(1).Length - CrossFadeLength, CrossFadeLength, CrossOutFadeSlopeType, CosinePower, EqualPower)
                    End If

                    'Adding the overlapping sounds into the output arrays
                    For s = 0 To InputSound.WaveData.SampleData(1).Length - 1 'Using channel 1 lengths only

                        'Skipping to next sample if WriteSampleIndex is below 0. This happens when concatenating sounds that are shorter than the CrossFadeLength
                        If WriteSampleIndex < 0 Then
                            'Increasing WriteSampleIndex after all channels have been copied
                            WriteSampleIndex += 1
                            Continue For
                        End If

                        For c = 1 To OutputSound.WaveFormat.Channels
                            OutputSound.WaveData.SampleData(c)(WriteSampleIndex) += InputSound.WaveData.SampleData(c)(s)
                        Next
                        'Increasing WriteSampleIndex after all channels have been copied
                        WriteSampleIndex += 1
                    Next

                    'Moving the write position backwards
                    WriteSampleIndex -= CrossFadeLength
                Next


            End If

            If DetectedNominalLevel.HasValue Then
                OutputSound.SMA.NominalLevel = DetectedNominalLevel.Value
            End If

            Return OutputSound

        Catch ex As Exception
            AudioError(ex.ToString)
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Superpositions the input sounds to a new sound. The lengths and wave format (including channel count) are required to be the same in all input sounds.
    ''' </summary>
    ''' <param name="InputSounds"></param>
    ''' <returns></returns>
    Public Shared Function SuperpositionEqualLengthSounds(ByRef InputSounds As List(Of Sound)) As Sound

        If InputSounds.Count = 0 Then Return Nothing

        Dim WaveFormat = InputSounds(0).WaveFormat
        For i = 1 To InputSounds.Count - 1
            If WaveFormat.IsEqual(InputSounds(i).WaveFormat) = False Then Throw New ArgumentException("All wave formats need to be the same!")
        Next

        Dim OutputSound As New Sound(WaveFormat)

        'Checking length equality of each channel, compared to channel 1 of the first sound
        Dim ChannelLength As Integer = InputSounds(0).WaveData.SampleData(1).Length
        For i = 0 To InputSounds.Count - 1
            For c = 1 To InputSounds(i).WaveFormat.Channels
                If InputSounds(i).WaveData.SampleData(c).Length <> ChannelLength Then Throw New ArgumentException("All sounds need to have the same sample array lengths (in corresponding channels)!")
            Next
        Next

        For Channel = 1 To WaveFormat.Channels

            Dim NewChannelArray(ChannelLength - 1) As Single
            'Copies the first sound
            Array.Copy(InputSounds(0).WaveData.SampleData(Channel), NewChannelArray, NewChannelArray.Length)

            'Superpositions the remaining sounds
            For i = 1 To InputSounds.Count - 1
                Dim CurrentChannelArray = InputSounds(i).WaveData.SampleData(Channel)
                For s = 0 To ChannelLength - 1
                    NewChannelArray(s) += CurrentChannelArray(s)
                Next
            Next

            OutputSound.WaveData.SampleData(Channel) = NewChannelArray

        Next

        Return OutputSound

    End Function


    ''' <summary>
    ''' Sets the sound level of the indicated section of the indicated sound to a target level.
    ''' </summary>
    ''' <param name="InputSound"></param>
    ''' <param name="targetLevel"></param>
    ''' <param name="channel"></param>
    ''' <param name="startSample"></param>
    ''' <param name="sectionLength"></param>
    ''' <param name="FrequencyWeighting"></param>
    Public Shared Sub MeasureAndAdjustSectionLevel(ByRef InputSound As Sound, ByVal targetLevel As Decimal, Optional ByVal channel As Integer? = Nothing,
                                    Optional ByVal startSample As Integer? = Nothing, Optional ByVal sectionLength As Integer? = Nothing,
                                             Optional ByVal FrequencyWeighting As FrequencyWeightings = FrequencyWeightings.Z)


        'Setting default start sample and length values
        If startSample Is Nothing Then startSample = 0
        If sectionLength Is Nothing Then sectionLength = InputSound.WaveData.ShortestChannelSampleCount

        Dim AudioOutputConstructor As New AudioOutputConstructor(InputSound.WaveFormat, channel)

        'Main section
        For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

            'Measures the level of the channel section
            Dim SectionLevel As Double = MeasureSectionLevel(InputSound, c, startSample, sectionLength,,, FrequencyWeighting)

            'Calculating the needed gain
            Dim Gain As Double = targetLevel - SectionLevel

            'Amplifies the section
            AmplifySection(InputSound, Gain, c, startSample, sectionLength,)

        Next

    End Sub


    Public Enum FadeSlopeType
        Smooth
        Linear
        PowerCosine_SkewedInFadeDirection
        PowerCosine_SkewedFromFadeDirection
    End Enum

    Public Class FadeSpecifications

        Public ReadOnly StartAttenuation As Double?
        Public ReadOnly EndAttenuation As Double?
        Public ReadOnly StartSample As Integer
        Public ReadOnly SectionLength As Integer?
        Public ReadOnly SlopeType As FadeSlopeType
        Public ReadOnly CosinePower As Double
        Public ReadOnly EqualPower As Boolean

        ''' <summary>
        ''' Creates a new instance of FadeSpecifications
        ''' </summary>
        ''' <param name="StartAttenuation">The attenuation (in dB) applied in the start of the fade period. I left empty, fade will start from silence.</param>
        ''' <param name="EndAttenuation">The attenuation (in dB) applied in the end of the fade period. I left empty, the fade period will end in silence.</param>
        ''' <param name="startSample">The start sample of the section to fade.</param>
        ''' <param name="sectionLength">The length (in samples) of the section to fade. If left to Nothing, fading will go to the end of the shortest sound channel.</param>
        ''' <param name="slopeType">Specifies the curvature of the fade section. Linear creates a linear fade, and Smooth fades using a cosine function to smoothen out the fade section.</param>
        Public Sub New(Optional ByVal StartAttenuation As Double? = Nothing, Optional ByVal EndAttenuation As Double? = Nothing,
                        Optional ByVal StartSample As Integer = 0, Optional ByVal SectionLength As Integer? = Nothing,
                        Optional ByVal SlopeType As FadeSlopeType = FadeSlopeType.Linear, Optional CosinePower As Double = 10, Optional ByVal EqualPower As Boolean = False)

            Me.StartAttenuation = StartAttenuation
            Me.EndAttenuation = EndAttenuation
            Me.StartSample = StartSample
            Me.SectionLength = SectionLength
            Me.SlopeType = SlopeType
            Me.CosinePower = CosinePower
            Me.EqualPower = EqualPower

        End Sub

    End Class

    ''' <summary>
    ''' Fading the indicated section of the indicated sound using the specified fading type.
    ''' </summary>
    ''' <param name="InputSound">The sound to be modified.</param>
    ''' <param name="FadeSpecifications"></param>
    ''' <param name="Channel">The channel to be modified. If left empty all channels will be modified.</param>
    Public Shared Sub Fade(ByRef InputSound As Sound, ByVal FadeSpecifications As FadeSpecifications, Optional ByVal Channel As Integer? = Nothing)

        Fade(InputSound, FadeSpecifications.StartAttenuation, FadeSpecifications.EndAttenuation, Channel, FadeSpecifications.StartSample, FadeSpecifications.SectionLength, FadeSpecifications.SlopeType, FadeSpecifications.CosinePower, FadeSpecifications.EqualPower)

    End Sub



    ''' <summary>
    ''' Fading the indicated section of the indicated sound using the specified fading type.
    ''' </summary>
    ''' <param name="input">The sound to be modified.</param>
    ''' <param name="StartAttenuation">The attenuation (in dB) applied in the start of the fade period. I left empty, fade will start from silence.</param>
    ''' <param name="EndAttenuation">The attenuation (in dB) applied in the end of the fade period. I left empty, the fade period will end in silence.</param>
    ''' <param name="channel">The channel to be modified. If left empty all channels will be modified.</param>
    ''' <param name="startSample">The start sample of the section to fade.</param>
    ''' <param name="sectionLength">The length (in samples) of the section to fade. If left to Nothing, fading will go to the end of the shortest sound channel.</param>
    ''' <param name="slopeType">Specifies the curvature of the fade section. Linear creates a linear fade, and Smooth fades using a cosine function to smoothen out the fade section.</param>
    Public Shared Sub Fade(ByRef input As Sound, Optional ByVal StartAttenuation As Double? = Nothing, Optional ByVal EndAttenuation As Double? = Nothing,
                        Optional ByVal channel As Integer? = Nothing,
                        Optional ByVal startSample As Integer = 0, Optional ByVal sectionLength As Integer? = Nothing,
                        Optional ByVal slopeType As FadeSlopeType = FadeSlopeType.Smooth,
                        Optional CosinePower As Double = 10,
                    Optional ByVal EqualPower As Boolean = False)

        'Ska uppdateras med kurvtyp, linjär ska läggas till som val
        'Även "micro255 law" ska läggas till, Equation 22-1 i dspguide.com: y=(ln(1+0.000001*x))/(ln(1+0.000001)) ' Skapar en naturlig fadening genom "compounding"
        'även en alternativ ekvation finns "A Law" (Stämmer detta. Borde nog byta input och output, jämfört med grafen i boken)

        Try

            Dim AudioOutputConstructor As New AudioOutputConstructor(input.WaveFormat, channel)

            'Main section
            For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                Dim inputArray() As Single = input.WaveData.SampleData(c)
                'CheckAndCorrectSectionLength(inputArray.Length, startSample, sectionLength)

                'Main section
                Dim startFactor As Double
                Dim endFactor As Double

                'Setting start factor
                If StartAttenuation Is Nothing Then
                    startFactor = 0
                Else
                    startFactor = 10 ^ (-StartAttenuation / 20) 'the minus sign converts attenuation to gain 

                    'Modifying the startFactor to account for the sqrt applied later
                    If EqualPower = True Then
                        startFactor = startFactor ^ 2
                    End If

                End If

                'Setting end factor
                If EndAttenuation Is Nothing Then
                    endFactor = 0
                Else
                    endFactor = 10 ^ (-EndAttenuation / 20) 'the minus sign converts attenuation to gain

                    'Modifying the startFactor to account for the sqrt applied later
                    If EqualPower = True Then
                        endFactor = endFactor ^ 2
                    End If

                End If


                Dim fadeSampleCount As Integer = 0
                Dim fadeProgress As Double = 0

                Dim IsFadeIn As Boolean = False
                If startFactor < endFactor Then IsFadeIn = True

                Dim CorrectedStartSample = startSample
                Dim CorrectedSectionLength = sectionLength
                CheckAndCorrectSectionLength(inputArray.Length, CorrectedStartSample, CorrectedSectionLength)

                If CorrectedSectionLength > 1 Then ' To avoid division by zero some lines below

                    'Modifies currentFadeFactor according to a cosine finction, whereby currentModFactor starts on 1 and end at 0
                    Dim currentModFactor As Double
                    Dim currentFadeFactor As Double
                    Select Case slopeType
                        Case FadeSlopeType.Smooth

                            If EqualPower = True Then

                                For currentSample = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1

                                    'fadeProgress goes from 0 to 1 during the fade section
                                    fadeProgress = fadeSampleCount / (CorrectedSectionLength - 1)

                                    currentModFactor = ((Math.Cos(twopi * (fadeProgress / 2)) + 1) / 2)
                                    'currentFadeFactor = Math.Sqrt(startFactor * currentModFactor + endFactor * (1 - currentModFactor))

                                    'Fading the section
                                    inputArray(currentSample) = inputArray(currentSample) * Math.Sqrt(startFactor * currentModFactor + endFactor * (1 - currentModFactor))
                                    'inputArray(currentSample) = (inputArray(currentSample) * currentFadeFactor)
                                    fadeSampleCount += 1
                                Next

                            Else

                                For currentSample = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1

                                    'fadeProgress goes from 0 to 1 during the fade section
                                    fadeProgress = fadeSampleCount / (CorrectedSectionLength - 1)

                                    currentModFactor = ((Math.Cos(twopi * (fadeProgress / 2)) + 1) / 2)
                                    'currentFadeFactor = startFactor * currentModFactor + endFactor * (1 - currentModFactor)

                                    'Fading the section
                                    inputArray(currentSample) = inputArray(currentSample) * (startFactor * currentModFactor + endFactor * (1 - currentModFactor))
                                    'inputArray(currentSample) = (inputArray(currentSample) * currentFadeFactor)
                                    fadeSampleCount += 1

                                Next
                            End If


                        Case FadeSlopeType.PowerCosine_SkewedFromFadeDirection

                            For currentSample = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1

                                'fadeProgress goes from 0 to 1 during the fade section
                                fadeProgress = fadeSampleCount / (CorrectedSectionLength - 1)

                                If IsFadeIn = True Then
                                    currentModFactor = ((Math.Cos(twopi * ((1 - fadeProgress) / 2)) + 1) / 2) ^ CosinePower
                                    currentFadeFactor = (startFactor * (1 - currentModFactor) + endFactor * currentModFactor)
                                Else
                                    currentModFactor = ((Math.Cos(twopi * (fadeProgress / 2)) + 1) / 2) ^ CosinePower
                                    currentFadeFactor = (startFactor * currentModFactor) + endFactor * (1 - currentModFactor)
                                End If

                                If EqualPower = True Then
                                    currentFadeFactor = Math.Sqrt(currentFadeFactor)
                                End If

                                'Fading the section
                                inputArray(currentSample) = (inputArray(currentSample) * currentFadeFactor)
                                fadeSampleCount += 1
                            Next

                        Case FadeSlopeType.PowerCosine_SkewedInFadeDirection

                            For currentSample = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1

                                'fadeProgress goes from 0 to 1 during the fade section
                                fadeProgress = fadeSampleCount / (CorrectedSectionLength - 1)

                                If IsFadeIn = True Then
                                    currentModFactor = ((Math.Cos(twopi * ((fadeProgress) / 2)) + 1) / 2) ^ CosinePower
                                    currentFadeFactor = (startFactor * (currentModFactor) + endFactor * (1 - currentModFactor))
                                Else
                                    currentModFactor = ((Math.Cos(twopi * ((1 - fadeProgress) / 2)) + 1) / 2) ^ CosinePower
                                    currentFadeFactor = (startFactor * (1 - currentModFactor)) + endFactor * (currentModFactor)
                                End If

                                If EqualPower = True Then
                                    currentFadeFactor = Math.Sqrt(currentFadeFactor)
                                End If

                                'Fading the section
                                inputArray(currentSample) = (inputArray(currentSample) * currentFadeFactor)
                                fadeSampleCount += 1
                            Next

                        Case Else 'I.e. Linear!

                            If EqualPower = True Then

                                For currentSample = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1

                                    'fadeProgress goes from 0 to 1 during the fade section
                                    fadeProgress = fadeSampleCount / (CorrectedSectionLength - 1)
                                    'currentFadeFactor = Math.Sqrt(startFactor * (1 - fadeProgress) + endFactor * fadeProgress)

                                    'Fading the section
                                    inputArray(currentSample) = inputArray(currentSample) * Math.Sqrt(startFactor * (1 - fadeProgress) + endFactor * fadeProgress)
                                    'inputArray(currentSample) = (inputArray(currentSample) * currentFadeFactor)
                                    fadeSampleCount += 1

                                Next

                            Else

                                For currentSample = CorrectedStartSample To CorrectedStartSample + CorrectedSectionLength - 1

                                    'fadeProgress goes from 0 to 1 during the fade section
                                    fadeProgress = fadeSampleCount / (CorrectedSectionLength - 1)
                                    'currentFadeFactor = startFactor * (1 - fadeProgress) + endFactor * fadeProgress

                                    'Fading the section
                                    inputArray(currentSample) = inputArray(currentSample) * (startFactor * (1 - fadeProgress) + endFactor * fadeProgress)
                                    'inputArray(currentSample) = (inputArray(currentSample) * currentFadeFactor)
                                    fadeSampleCount += 1

                                Next

                            End If

                    End Select
                End If
            Next

        Catch ex As Exception
            AudioError(ex.ToString)
        End Try

    End Sub





    Public Shared Sub InsertSound(ByRef SoundToInsert As Sound, ByVal SourceChannel As Integer, ByRef TargetSound As Sound, ByVal TargetChannel As Integer, ByVal StartInsertSample As Integer)

        Try

            'This check allows different channel counts
            If SoundToInsert.WaveFormat.IsEqual(TargetSound.WaveFormat, False) = False Then Throw New ArgumentException("Different formats in InsertSound sounds. Aborting!")

            Dim SourceChannelArray = SoundToInsert.WaveData.SampleData(SourceChannel)
            Dim TargetChannelArray = TargetSound.WaveData.SampleData(TargetChannel)
            Dim CopyLength As Integer = SourceChannelArray.Length

            If Globals.StfBase.UseOptimizationLibraries = True Then

                'Placing the sound to insert in an array with the same length as TargetSound (as same length is required by AddTwoFloatArrays below)
                Dim ExtendedSourceArray(TargetChannelArray.Length - 1) As Single
                Array.Copy(SourceChannelArray, 0, ExtendedSourceArray, StartInsertSample, CopyLength)

                'Dim TestSound As New Audio.Sound(New Formats.WaveFormat(TargetSound.WaveFormat.SampleRate, TargetSound.WaveFormat.BitDepth, 1,, TargetSound.WaveFormat.Encoding))
                'TestSound.WaveData.SampleData(1) = ExtendedSourceArray
                'TestSound.WriteWaveFile(IO.Path.Combine(Utils.logFilePath, "AdaptiveSipSounds", "TestSound_Source"))

                'Dim TestSound2 As New Audio.Sound(New Formats.WaveFormat(TargetSound.WaveFormat.SampleRate, TargetSound.WaveFormat.BitDepth, 1,, TargetSound.WaveFormat.Encoding))
                'TestSound2.WaveData.SampleData(1) = TargetChannelArray
                'TestSound2.WriteWaveFile(IO.Path.Combine(Utils.logFilePath, "AdaptiveSipSounds", "TestSound_TargetChannelArray"))

                'Summing the arrays using AddTwoFloatArrays
                LibOstfDsp_VB.AddTwoFloatArrays(TargetChannelArray, ExtendedSourceArray)

                'Dim TestSound3 As New Audio.Sound(New Formats.WaveFormat(TargetSound.WaveFormat.SampleRate, TargetSound.WaveFormat.BitDepth, 1,, TargetSound.WaveFormat.Encoding))
                'TestSound3.WaveData.SampleData(1) = TargetChannelArray
                'TestSound3.WriteWaveFile(IO.Path.Combine(Utils.logFilePath, "AdaptiveSipSounds", "TestSound_AfterMix"))

            Else

                Dim RequiredTargetLength As Integer = StartInsertSample + SourceChannelArray.Length
                If RequiredTargetLength > TargetChannelArray.Length Then
                    CopyLength = TargetChannelArray.Length - StartInsertSample
                End If

                For s = 0 To CopyLength - 1
                    TargetChannelArray(StartInsertSample + s) += SourceChannelArray(s)
                Next

            End If

        Catch ex As Exception
            AudioError(ex.ToString)
        End Try

    End Sub


    ''' <summary>
    ''' Soft limits the indicated section of the input sound to the specified RMS level (in dBFS).
    ''' </summary>
    ''' <param name="inputSound"></param>
    ''' <param name="startSample"></param>
    ''' <param name="sectionLength"></param>
    ''' <param name="WindowDuration">The time in seconds of windows within which to average the RMS level.</param>
    ''' <param name="Channel">The channel to hard limit. If left to Nothing, all channels will be limited.</param>
    ''' <param name="ReturnReport">If ReturnReport is set to True, the function returns a string containing a report of the applied limiting.</param>
    ''' <returns>If ReturnReport is set to True, the function returns a string containing a report of the applied limiting.</returns>
    Public Shared Function SoftLimitSection(ByRef InputSound As Sound,
                                         ByVal Channel As Integer,
                                         ByVal ThresholdLevel As Double,
                                         Optional ByVal StartSample As Integer = 0,
                                         Optional ByVal SectionLength As Integer? = Nothing,
                                         Optional ByVal WindowDuration As Double = 0.2,
                                         Optional ByVal FrequencyWeighting As FrequencyWeightings = FrequencyWeightings.Z,
                                         Optional ByVal ReturnReport As Boolean = True,
                                         Optional ByVal slopeType As FadeSlopeType = FadeSlopeType.Linear,
                                         Optional CosinePower As Double = 10,
                                         Optional ByVal EqualPower As Boolean = False) As String

        Dim Report As String = ""

        Try

            'Declares a list used for reporting of limiter data
            Dim ReportList As List(Of String) = Nothing
            If ReturnReport = True Then
                ReportList = New List(Of String)
            End If

            'Calculating the averaging window length in samples
            Dim WindowLength As Integer = WindowDuration * InputSound.WaveFormat.SampleRate

            'Performs limiting by calling LimitChannelSection
            'Limiting the sound in the specified channel
            SoftLimitChannelSection(InputSound, Channel, ThresholdLevel, StartSample, SectionLength, WindowLength,
                                        FrequencyWeighting, ReportList, slopeType, CosinePower, EqualPower)

            'Creating the report
            If ReturnReport = True Then
                Report = String.Join(vbCrLf, ReportList)
            End If

        Catch ex As Exception
            AudioError(ex.ToString)
            'Storing the exception info in Report
            If ReturnReport = True Then
                Report &= "The following error occured in SoftLimitSection: " & ex.ToString
            End If
        End Try

        Return Report

    End Function


    ''' <summary>
    ''' Applies limiting to the specified channel, in the specified sound.
    ''' </summary>
    ''' <param name="InputSound"></param>
    ''' <param name="Channel"></param>
    ''' <param name="ThresholdLevel"></param>
    ''' <param name="StartSample"></param>
    ''' <param name="SectionLength">The length of the measurement windows. Must be an even integer.</param>
    ''' <param name="WindowLength"></param>
    ''' <param name="ReportList">A list of string that can be used to log limiter data.</param>
    Private Shared Sub SoftLimitChannelSection(ByRef InputSound As Sound,
                                        ByVal Channel As Integer,
                                        ByVal ThresholdLevel As Double,
                                        ByVal StartSample As Integer,
                                        ByVal SectionLength As Integer?,
                                        ByVal WindowLength As Integer,
                                        ByVal FrequencyWeighting As FrequencyWeightings,
                                        Optional ByRef ReportList As List(Of String) = Nothing,
                                        Optional ByVal SlopeType As FadeSlopeType = FadeSlopeType.Linear,
                                        Optional ByVal CosinePower As Double = 10,
                                        Optional ByVal EqualPower As Boolean = False)

        Try

            'TODO: This function ignores the last window, if not full!

            CheckAndCorrectSectionLength(InputSound.WaveData.SampleData(Channel).Length, StartSample, SectionLength)

            'Calculating the number of sections
            Dim SectionCount As Integer = Int(SectionLength / WindowLength)

            'Measures each section of the sound, and stores the results in SoundLevelList
            Dim SoundLevelList As New List(Of Double)
            For SectionIndex = 0 To SectionCount - 1
                SoundLevelList.Add(MeasureSectionLevel(InputSound, Channel, StartSample + SectionIndex * WindowLength, WindowLength,,, FrequencyWeighting))
            Next

            'Going though each section and calculates needed attenuation 
            Dim AttenuationList As New List(Of Double)
            For SectionIndex = 0 To SectionCount - 1
                AttenuationList.Add(Math.Max(0, SoundLevelList(SectionIndex) - ThresholdLevel))
            Next

            Dim LastAttenuation As Double = AttenuationList(0)
            For SectionIndex = 0 To SectionCount - 1

                'Calculating attenuation
                Dim StartAttenuation As Double
                Dim EndAttenuation As Double
                If SectionIndex = SectionCount - 1 Then
                    'The last window uses only the attentation level of the last index in AttenuationList
                    StartAttenuation = Math.Max(AttenuationList(SectionIndex), AttenuationList(SectionIndex))
                    EndAttenuation = StartAttenuation
                Else
                    StartAttenuation = LastAttenuation
                    EndAttenuation = Math.Max(AttenuationList(SectionIndex), AttenuationList(SectionIndex + 1))
                    LastAttenuation = EndAttenuation
                End If

                'Using Fade to perform attenuation
                Fade(InputSound, StartAttenuation, EndAttenuation, Channel, StartSample + SectionIndex * WindowLength, WindowLength, SlopeType, CosinePower, EqualPower)

                'Reporting only limited sections
                If ReportList IsNot Nothing Then
                    If StartAttenuation > 0 Or EndAttenuation > 0 Then
                        ReportList.Add(SectionIndex & vbTab & StartAttenuation & vbTab & EndAttenuation)
                    End If
                End If
            Next

        Catch ex As Exception
            AudioError(ex.ToString)
        End Try

    End Sub



    Public Shared Function CreateSilence(ByRef format As Formats.WaveFormat, Optional ByVal channel As Integer? = Nothing,
                                 Optional ByVal duration As Double = 1,
                                     Optional durationTimeUnit As TimeUnits = TimeUnits.seconds) As Sound

        Try

            Dim outputSound As New Sound(format)
            Dim AudioOutputConstructor As New AudioOutputConstructor(format, channel)

            Dim posFS As Double = format.PositiveFullScale
            Dim negFS As Double = format.NegativeFullScale

            Dim dataLength As Long = 0
            Select Case durationTimeUnit
                Case TimeUnits.seconds
                    dataLength = duration * format.SampleRate
                Case TimeUnits.samples
                    dataLength = duration
            End Select

            Dim rnd As New Random

            'Main section

            For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                Dim channelArray(dataLength - 1) As Single
                For n = 1 To channelArray.Length - 1
                    channelArray(n) = 0
                Next
                outputSound.WaveData.SampleData(c) = channelArray

            Next

            Return outputSound

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function


    Public Shared Function CreateWhiteNoise(ByRef format As Formats.WaveFormat, Optional ByVal channel As Integer? = Nothing,
                                Optional ByVal level As Double = 1, Optional ByVal duration As Double = 1,
                                     Optional durationTimeUnit As TimeUnits = TimeUnits.seconds,
                                     Optional ByRef RandomSource As Random = Nothing) As Sound

        Try

            If level > 1 Then
                level = 1
                MsgBox("Level was outside allowed value (-1 through 1)" & vbCr & vbCr & "The level was adjusted To 1", "Waring from createSineWave")
            End If

            If level < -1 Then
                level = -1
                MsgBox("Level was outside allowed value (-1 through 1)" & vbCr & vbCr & "The level was adjusted To -1", "Waring from createSineWave")
            End If

            Dim outputSound As New Sound(format)
            Dim AudioOutputConstructor As New AudioOutputConstructor(format, channel)

            Dim posFS As Double = format.PositiveFullScale
            Dim negFS As Double = format.NegativeFullScale

            Dim dataLength As Long = 0
            Select Case durationTimeUnit
                Case TimeUnits.seconds
                    dataLength = duration * format.SampleRate
                Case TimeUnits.samples
                    dataLength = duration
            End Select

            If RandomSource Is Nothing Then
                Thread.Sleep(20) ' The reason for sleeping the thread is to avoid the same random seed on multiple close calls. (Due to the limited resolution of the system clock.)
                RandomSource = New Random()
            End If

            'Main section
            Select Case format.Encoding
                Case Formats.WaveFormat.WaveFormatEncodings.IeeeFloatingPoints

                    For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                        Dim channelArray(dataLength - 1) As Single

                        Select Case format.BitDepth
                            Case 32

                                For n = 0 To channelArray.Length - 1
                                    channelArray(n) = (level * (2 * (RandomSource.NextDouble() - 0.5)))
                                Next

                            Case Else
                                Throw New NotImplementedException

                        End Select

                        outputSound.WaveData.SampleData(c) = channelArray

                    Next


                Case Formats.WaveFormat.WaveFormatEncodings.PCM

                    For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                        Dim channelArray(dataLength - 1) As Single

                        For n = 0 To channelArray.Length - 1
                            channelArray(n) = level * RandomSource.Next(negFS, posFS)
                        Next

                        outputSound.WaveData.SampleData(c) = channelArray

                    Next

            End Select

            Return outputSound

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function


    ''' <summary>
    ''' Creates a new sound containing a sine wave.
    ''' </summary>
    ''' <param name="format"></param>
    ''' <param name="channel"></param>
    ''' <param name="freq">Specified in Hz</param>
    ''' <param name="intensity"></param>
    ''' <param name="intensityUnit"></param>
    ''' <param name="duration"></param>
    ''' <param name="durationTimeUnit"></param>
    ''' <param name="Phase">Specified in radians.</param>
    ''' <returns></returns>
    Public Shared Function CreateSineWave(ByRef format As Formats.WaveFormat, Optional ByVal channel As Integer? = Nothing,
                                Optional ByVal freq As Double = 1000, Optional ByVal intensity As Decimal = 1,
                                   Optional intensityUnit As SoundDataUnit = SoundDataUnit.unity,
                                   Optional ByVal duration As Double = 1, Optional durationTimeUnit As TimeUnits = TimeUnits.seconds,
                                   Optional ByVal Phase As Double = 0, Optional ByVal OverrideLevelCheck As Boolean = False) As Sound
        Try




            Dim outputSound As New Sound(format)
            Dim AudioOutputConstructor As New AudioOutputConstructor(format, channel)

            'Checking valid input intensity values
            If OverrideLevelCheck = False Then
                Select Case intensityUnit
                    Case SoundDataUnit.unity
                        If intensity > 1 Then
                            intensity = 1
                            MsgBox("Level was outside allowed value (0 through 1)" & vbCr & vbCr & "The level was adjusted to 1",, "CreateSineWave")
                        End If

                        If intensity < 0 Then
                            intensity = 1
                            MsgBox("Level was outside allowed value (0 through 1)" & vbCr & vbCr & "The level was adjusted to 1",, "CreateSineWave")
                        End If
                    Case SoundDataUnit.dB
                        If intensity > 0 Then
                            intensity = 0
                            MsgBox("Level was above allowed max value (0 dBFS)" & vbCr & vbCr & "The level was adjusted To 0 dBFS",, "CreateSineWave")
                        End If
                    Case SoundDataUnit.linear
                        If intensity > format.PositiveFullScale Then
                            intensity = format.PositiveFullScale
                            MsgBox("Level was outside allowed value (0 through " & format.PositiveFullScale & ")" & vbCr & vbCr & "The level was adjusted To " & format.PositiveFullScale,, "Waring from createSineWave")
                        End If

                        If intensity < 0 Then
                            intensity = format.PositiveFullScale
                            MsgBox("Level was outside allowed value (0 through " & format.PositiveFullScale & ")" & vbCr & vbCr & "The level was adjusted To " & format.PositiveFullScale,, "CreateSineWave")
                        End If
                End Select
            End If

            'Converting intensity values
            Select Case intensityUnit
                Case SoundDataUnit.unity
                            'no conversion is needed (since the signal generation words in unity scale)

                Case SoundDataUnit.dB
                    'convert to linear
                    intensity = dBConversion(intensity, dBConversionDirection.from_dB, outputSound.WaveFormat)

                    'convering to unity
                    intensity = intensity / format.PositiveFullScale

                Case SoundDataUnit.linear
                    'convering to unity
                    intensity = intensity / format.PositiveFullScale

            End Select

            Dim dataLength As Long = 0
            Select Case durationTimeUnit
                Case TimeUnits.seconds
                    dataLength = duration * format.SampleRate
                Case TimeUnits.samples
                    dataLength = duration
            End Select

            'Main section
            For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                Dim channelArray(dataLength - 1) As Single

                Select Case format.BitDepth
                    Case 8 'Actually used for formats that use unsigned datatypes
                        For n = 0 To channelArray.Length - 1
                            channelArray(n) = (intensity * (format.PositiveFullScale / 2)) * Math.Sin(Phase + twopi * (freq / format.SampleRate) * n) + format.PositiveFullScale / 2 ' - _
                        Next

                    Case 16, 32, 64 'Actually used for formats that use signed datatypes
                        For n = 0 To channelArray.Length - 1
                            channelArray(n) = (intensity * format.PositiveFullScale) * Math.Sin(Phase + twopi * (freq / format.SampleRate) * n) ' - _
                        Next
                    Case Else
                        Throw New NotImplementedException(format.BitDepth & " bit depth Is Not yet supported.")

                End Select

                outputSound.WaveData.SampleData(c) = channelArray

            Next

            Return outputSound

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function



    ''' <summary>
    ''' Creates sound containing a frequency modulated (or more correctly phase moduated) sine wave.
    ''' </summary>
    ''' <param name="format"></param>
    ''' <param name="channel"></param>
    ''' <param name="CarrierFrequency"></param>
    ''' <param name="intensity"></param>
    ''' <param name="ModulationFrequency"></param>
    ''' <param name="ModulationDepth">Modulating depth as a proportion of the carrier frequency.</param>
    ''' <param name="intensityUnit"></param>
    ''' <param name="duration"></param>
    ''' <param name="durationTimeUnit"></param>
    ''' <returns></returns>
    Public Shared Function CreateFrequencyModulatedSineWave(ByRef format As Formats.WaveFormat, Optional ByVal channel As Integer? = Nothing,
                                                     Optional ByVal CarrierFrequency As Double = 1000, Optional ByVal intensity As Decimal = 1,
                                                     Optional ByVal ModulationFrequency As Double = 10, Optional ByVal ModulationDepth As Double = 0.1,
                                                     Optional intensityUnit As SoundDataUnit = SoundDataUnit.unity,
                                                     Optional ByVal duration As Double = 1, Optional durationTimeUnit As TimeUnits = TimeUnits.seconds) As Sound

        'FM modulation 
        '(Source: Sinusoidal Modulation of Sinusoids, JULIUS O. SMITH III. https://ccrma.stanford.edu/~jos/rbeats/Sinusoidal_Frequency_Modulation_FM.html)
        'x(t) = Ac *Cos [wc*t + fc + Am*Sin(wmt +fm)]
        '
        't=time
        'Ac=Amplitude of carrier
        'wc=frequency of carrier
        'fc= phase of carrier
        '
        'Am=Amplitude of modulator
        'wm=frequency of modulator
        'fm= phase of modulator

        'And here's another useful link: http://www.rfcafe.com/references/electrical/frequency-modulation.htm

        Try


            Dim outputSound As New Sound(format)
            Dim AudioOutputConstructor As New AudioOutputConstructor(format, channel)


            'Checking valid input intensity values
            Select Case intensityUnit
                Case SoundDataUnit.unity
                    If intensity > 1 Then
                        intensity = 1
                        MsgBox("Level was outside allowed value (0 through 1)" & vbCr & vbCr & "The level was adjusted to 1",, "CreateSineWave")
                    End If

                    If intensity < 0 Then
                        intensity = 1
                        MsgBox("Level was outside allowed value (0 through 1)" & vbCr & vbCr & "The level was adjusted to 1",, "CreateSineWave")
                    End If
                Case SoundDataUnit.dB
                    If intensity > 0 Then
                        intensity = 0
                        MsgBox("Level was above allowed max value (0 dBFS)" & vbCr & vbCr & "The level was adjusted To 0 dBFS",, "CreateSineWave")
                    End If
                Case SoundDataUnit.linear
                    If intensity > format.PositiveFullScale Then
                        intensity = format.PositiveFullScale
                        MsgBox("Level was outside allowed value (0 through " & format.PositiveFullScale & ")" & vbCr & vbCr & "The level was adjusted To " & format.PositiveFullScale,, "Waring from createSineWave")
                    End If

                    If intensity < 0 Then
                        intensity = format.PositiveFullScale
                        MsgBox("Level was outside allowed value (0 through " & format.PositiveFullScale & ")" & vbCr & vbCr & "The level was adjusted To " & format.PositiveFullScale,, "CreateSineWave")
                    End If
            End Select

            'Converting intensity values
            Select Case intensityUnit
                Case SoundDataUnit.unity
                            'no conversion is needed (since the signal generation words in unity scale)

                Case SoundDataUnit.dB
                    'convert to linear
                    intensity = dBConversion(intensity, dBConversionDirection.from_dB, outputSound.WaveFormat)

                    'convering to unity
                    intensity = intensity / format.PositiveFullScale

                Case SoundDataUnit.linear
                    'convering to unity
                    intensity = intensity / format.PositiveFullScale

            End Select

            Dim dataLength As Long = 0
            Select Case durationTimeUnit
                Case TimeUnits.seconds
                    dataLength = duration * format.SampleRate
                Case TimeUnits.samples
                    dataLength = duration
            End Select

            'Calculating the maximum frequency deviation
            Dim MaximumFrequencyDeviation As Double = CarrierFrequency * ModulationDepth

            'Main section
            Dim CurrentFrequency As Double = 0
            For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                Dim channelArray(dataLength - 1) As Single

                Select Case format.BitDepth
                    Case 8 'Actually used for formats that use unsigned datatypes
                        For n = 0 To channelArray.Length - 1

                            channelArray(n) = (intensity * (format.PositiveFullScale / 2)) * Math.Sin(twopi * (CurrentFrequency / format.SampleRate) * n) + format.PositiveFullScale / 2 ' - _

                        Next

                    Case 16, 32, 64 'Actually used for formats that use signed datatypes
                        For n = 0 To channelArray.Length - 1

                            'Intuitively for carrier frequency:
                            'channelArray(n) = (intensity * format.PositiveFullScale) * Math.Sin("One lap" * "number of laps in a second" * "the current time (in seconds)")
                            'channelArray(n) = (intensity * format.PositiveFullScale) * Math.Sin(twopi * CurrentFrequency * (n / format.SampleRate)) ' - _


                            'Applying modulation by varying the phase (Actually I did this by trial and error, I'm not sure why the factor (twopi*Pi*0.1) is needed (+/-0.1 is the modulator range))
                            'channelArray(n) = (intensity * format.PositiveFullScale) * Math.Sin(twopi * CarrierFrequency * (n / format.SampleRate) - (twopi * Math.PI * ModulationFrequencyDeviation * ModulationWaveForm.WaveData.SampleData(c)(n)))

                            'x(t) = Ac *Cos [wc*t + fc + Am*Sin(wmt +fm)]
                            channelArray(n) = (intensity * format.PositiveFullScale) * Math.Cos(twopi * CarrierFrequency * (n / format.SampleRate) + 0 +
                                                                                                (MaximumFrequencyDeviation / ModulationFrequency) * Math.Sin(twopi * ModulationFrequency * (n / format.SampleRate) + 0))


                        Next
                    Case Else
                        Throw New NotImplementedException(format.BitDepth & " bit depth Is Not yet supported.")

                End Select

                outputSound.WaveData.SampleData(c) = channelArray

            Next

            Return outputSound

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function




    ''' <summary>
    ''' Creates a log sine sweep.
    ''' </summary>
    ''' <param name="format"></param>
    ''' <param name="channel"></param>
    ''' <param name="StartFrequency"></param>
    ''' <param name="ArrivalFrequency"></param>
    ''' <param name="intensity"></param>
    ''' <param name="intensityUnit"></param>
    ''' <param name="TotalDuration"></param>
    ''' <param name="durationTimeUnit"></param>
    ''' <returns></returns>
    Public Shared Function CreateLogSineSweep(ByRef format As Formats.WaveFormat, Optional ByVal channel As Integer? = Nothing,
                                       Optional ByVal StartFrequency As Double = 20, Optional ByVal ArrivalFrequency As Double = 20000,
                                       Optional ByVal FlatSpectrum As Boolean = False,
                                       Optional ByVal intensity As Decimal = 1, Optional intensityUnit As SoundDataUnit = SoundDataUnit.unity,
                                       Optional ByVal TotalDuration As Double = 1, Optional durationTimeUnit As TimeUnits = TimeUnits.seconds) As Sound

        Try

            Dim outputSound As New Sound(format)
            Dim AudioOutputConstructor As New AudioOutputConstructor(format, channel)

            'Converting intensity values
            Select Case intensityUnit
                Case SoundDataUnit.unity
                            'no conversion is needed (since the signal generation words in unity scale)

                Case SoundDataUnit.dB
                    'convert to linear
                    intensity = dBConversion(intensity, dBConversionDirection.from_dB, outputSound.WaveFormat)

                    'convering to unity
                    intensity = intensity / format.PositiveFullScale

                Case SoundDataUnit.linear
                    'convering to unity
                    intensity = intensity / format.PositiveFullScale

            End Select

            Dim dataLength As Long = 0
            Select Case durationTimeUnit
                Case TimeUnits.seconds
                    dataLength = TotalDuration * format.SampleRate
                Case TimeUnits.samples
                    dataLength = TotalDuration
            End Select

            'Main section
            For c = AudioOutputConstructor.FirstChannelIndex To AudioOutputConstructor.LastChannelIndex

                Dim channelArray(dataLength - 1) As Single

                Select Case format.BitDepth
                    Case 16, 32, 64 'Should be used for formats that use signed datatypes

                        'Formula from Picinali, Lorenzo. "Techniques for the extraction of the impulse response of a linear and time-invariant system." 
                        'x(t) = sin(((TwoPi*StartFreq * T)*(Ln((TwoPi*ArrivalFreq)/(TwoPi*StartFreq)) *(e ^((t/T)* ln((TwoPi*ArrivalFreq)/(TwoPi*StartFreq)))-1))
                        Dim CurrentTime As Double
                        Dim LnExpression As Double = Math.Log(ArrivalFrequency / StartFrequency) 'Simplification of: Math.Log((twopi * ArrivalFrequency) / (twopi * StartFrequency))
                        Dim Factor1 As Double = ((StartFrequency * TotalDuration) / LnExpression)
                        Dim LocalIntensity As Double = (intensity * format.PositiveFullScale)
                        Dim Exponent1 As Double

                        If FlatSpectrum = True Then

                            Dim Normalizer As Double = 1 / Math.Exp(LnExpression / 2)

                            For s = 0 To channelArray.Length - 1

                                CurrentTime = s / format.SampleRate
                                Exponent1 = (CurrentTime / TotalDuration) * LnExpression
                                channelArray(s) = LocalIntensity * Math.Sin(twopi * Factor1 * (Math.Exp(Exponent1) - 1)) * (Math.Exp(Exponent1 / 2) * Normalizer) 'The flat spectrum level factor in this equation was created (by EW) much by trial and error...

                            Next

                        Else
                            For s = 0 To channelArray.Length - 1

                                CurrentTime = s / format.SampleRate
                                Exponent1 = (CurrentTime / TotalDuration) * LnExpression
                                channelArray(s) = LocalIntensity * Math.Sin(twopi * Factor1 * (Math.Exp(Exponent1) - 1))
                            Next

                        End If

                    Case Else
                        Throw New NotImplementedException(format.BitDepth & " bit depth Is Not yet supported.")

                End Select

                outputSound.WaveData.SampleData(c) = channelArray

            Next

            Return outputSound

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function


    Public Enum dBTypes
        SoundPressure
        SoundPower
    End Enum

    Public Shared Function dBConversion(ByVal inputValue As Double, ByVal conversionDirection As dBConversionDirection,
                                 ByVal soundFormat As Formats.WaveFormat,
                                 Optional ByVal dBConversionType As dBTypes = dBTypes.SoundPressure) As Double

        Try

            Dim posFS As Double = soundFormat.PositiveFullScale

            Select Case dBConversionType
                Case dBTypes.SoundPressure
                    Select Case conversionDirection
                        Case dBConversionDirection.to_dB
                            Dim dBFS = 20 * Math.Log10(inputValue / posFS)
                            Return dBFS
                        Case dBConversionDirection.from_dB
                            Dim RMS As Double = posFS * 10 ^ (inputValue / 20)
                            Return RMS
                        Case Else
                            Throw New ArgumentException("Invalid conversionDirection")
                    End Select

                Case dBTypes.SoundPower
                    Select Case conversionDirection
                        Case dBConversionDirection.to_dB
                            Dim dBFS = 10 * Math.Log10(inputValue / posFS)
                            Return dBFS
                        Case dBConversionDirection.from_dB
                            Dim RMS As Double = posFS * 10 ^ (inputValue / 10)
                            Return RMS
                        Case Else
                            Throw New ArgumentException("Invalid conversionDirection")
                    End Select
                Case Else
                    Throw New ArgumentException("Invalid dBConversionType")
            End Select

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function

    Public Enum SoundMeasurementType
        RMS 'Measuring the RMS value
        AbsolutePeakAmplitude 'Measuring the absolute peak amplidude
        averageAbsoluteAmplitude 'Measuring the average absolute valued amplidude
    End Enum

    Public Enum SoundDataUnit
        dB
        linear
        unity
    End Enum

    Public Enum dBConversionDirection
        to_dB
        from_dB
    End Enum

    ''' <summary>
    ''' Holds the simulated sound field output level of a 1 kHz sine wave at an (hypothetical) RMS level of 0 dBFS. 
    ''' </summary>
    Public Const Standard_dBFS_dBSPL_Difference As Double = 100

    ''' <summary>
    ''' Converts the sound pressure level given by InputSPL to a value in dB FS using the conversion value given by Standard_dBFS_dBSPL_Difference
    ''' </summary>
    ''' <param name="InputSPL"></param>
    ''' <returns></returns>
    Public Shared Function Standard_dBSPL_To_dBFS(ByVal InputSPL As Double) As Double
        Return InputSPL - Standard_dBFS_dBSPL_Difference
    End Function

    ''' <summary>
    ''' Converts the full scale sound level given by InputFS to a sound pressure level value using the conversion value given by Standard_dBFS_dBSPL_Difference
    ''' </summary>
    ''' <param name="InputFS"></param>
    ''' <returns></returns>
    Public Shared Function Standard_dBFS_To_dBSPL(ByVal InputFS As Double) As Double
        Return Standard_dBFS_dBSPL_Difference + InputFS
    End Function


    ''' <summary>
    ''' Performs variuos checks on a chosen fft size. If input fftSize is set to nothing, fft size is set to the lowest valid 
    ''' value equal to or higher than the lower inclusive limit. If fft size is lower than the lower inclusive limit, fft size is set 
    ''' to a value equal to or higher than the lower inclusive limit. Is a non valid fft size (I.E. not a base to integer exponent), the
    ''' fft size is set to a value equal to or higher than the input fft size.
    ''' </summary>
    ''' <param name="fftSize"></param>
    ''' <param name="lowerInclusiveLimit"></param>
    Public Shared Sub CheckAndAdjustFFTSize(ByRef fftSize As Double, ByVal lowerInclusiveLimit As Double, Optional ByRef inActivateWarnings As Boolean = False)

        'limit is an inclusive value, ie fftSize can be set to limit

        'Finding the smallest fftSize => lower inclusive limit
        If fftSize = Nothing Then
            'sets the fftsize
            Dim exponent As Integer = 1
            fftSize = 0
            Do Until fftSize >= lowerInclusiveLimit
                fftSize = 2 ^ exponent
                exponent += 1
            Loop
            If inActivateWarnings = False Then
                MsgBox("No fft size is chosen. Setting the fftSize automatically to: " & fftSize)

                'Inactivates the warning after the first time it's been shown
                inActivateWarnings = True
            End If
        Else

            'If the manually set FFT size is too small, it is increased automatically.
            If fftSize < lowerInclusiveLimit Then
                Dim exponent As Integer = 1
                fftSize = 0
                Do Until fftSize >= lowerInclusiveLimit
                    fftSize = 2 ^ exponent
                    exponent += 1
                Loop
                If inActivateWarnings = False Then
                    MsgBox("The chosen FFT size is to small. Increasing it to " & fftSize & " points.")

                    'Inactivates the warning after the first time it's been shown
                    inActivateWarnings = True
                End If
            End If

            'Checks that the set fft windows size value is valid
            Dim validValues As Integer = 0
            For exponent = 0 To 1023
                If 2 ^ exponent = fftSize Then
                    validValues += 1
                    Exit For
                End If
            Next
            If Not validValues > 0 Then

                'Corrects the erraneous fft size to the nearest valid fft size that is larger than the input fft size
                Dim exponent As Integer = 1
                Dim newFftSize As Integer
                newFftSize = 0
                Do Until newFftSize >= fftSize
                    newFftSize = 2 ^ exponent
                    exponent += 1
                Loop
                fftSize = newFftSize
                If inActivateWarnings = False Then
                    MsgBox("The specified fft size is not valid. Setting it automatically to: " & fftSize)

                    'Inactivates the warning after the first time it's been shown
                    inActivateWarnings = True
                End If
            End If

        End If

    End Sub



    Public Enum WindowingType
        Rectangular
        Hamming
        Hanning
        Sine
        Blackman
        Triangular
        Tukey
    End Enum




#Region "Math"

    Public Const twopi As Double = 2 * System.Math.PI 'Or 2 * Math.Acos(-1)

    ''' <summary>
    ''' Returns a vector of length n, with random integers sampled in from the range of min (includive) to max (exclusive).
    ''' </summary>
    ''' <returns></returns>
    Public Shared Function SampleWithoutReplacement(ByVal n As Integer, ByVal min As Integer, ByVal max As Integer,
                                         Optional randomSource As Random = Nothing) As Integer()

        If randomSource Is Nothing Then randomSource = New Random()

        If n > max - min Then Throw New ArgumentException("max minus min must be equal to or greater than n")

        Dim SampleData As New HashSet(Of Integer)
        Dim NewSample As Integer

        'Sampling data until the length of SampleData equals n
        Do Until SampleData.Count >= n

            'Getting a random sample
            NewSample = randomSource.Next(min, max)

            'Adding the sample only if it is not already present in SampleData 
            If Not SampleData.Contains(NewSample) Then SampleData.Add(NewSample)
        Loop

        Return SampleData.ToArray

    End Function

    Public Enum RoundingMethods
        GetClosestValue
        AlwaysDown
        AlwaysUp
        DoNotRound
    End Enum

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="inputValue"></param>
    ''' <param name="roundingMethod"></param>
    ''' <param name="DecimalsInReturnsString"></param>
    ''' <param name="SkipRounding">If set to true, the rounding function is inactivated and the input value is returned unalterred.</param>
    ''' <returns></returns>
    Public Shared Function Rounding(ByVal inputValue As Object, Optional ByVal roundingMethod As RoundingMethods = RoundingMethods.GetClosestValue,
                         Optional DecimalsInReturnsString As Integer? = Nothing, Optional ByVal SkipRounding As Boolean = False,
                         Optional MinimumNonDecimalsInReturnString As Integer? = Nothing)

        'Returns the input value, if SkipRounding is true
        If SkipRounding = True Then Return inputValue

        Try

            Dim ReturnValue As Double = inputValue

            Select Case roundingMethod
                Case RoundingMethods.AlwaysDown
                    ReturnValue = Int(ReturnValue)

                Case RoundingMethods.AlwaysUp
                    If Not inputValue - Int(inputValue) = 0 Then
                        ReturnValue = Int(ReturnValue) + 1
                    Else
                        ReturnValue = ReturnValue
                    End If

                Case RoundingMethods.DoNotRound
                    ReturnValue = ReturnValue

                Case RoundingMethods.GetClosestValue

                    If DecimalsInReturnsString Is Nothing Then
                        ReturnValue = (System.Math.Round(ReturnValue))
                        'If not midpoint rounding is done below
                    End If

                Case Else
                    Throw New Exception("The " & roundingMethod & " rounding method enumerator is not valid.")
                    Return Nothing
            End Select

            Dim RetString As String = ""
            If DecimalsInReturnsString IsNot Nothing Or MinimumNonDecimalsInReturnString IsNot Nothing Then

                If DecimalsInReturnsString < 0 Then Throw New ArgumentException("DecimalsInReturnsString cannot be lower than 0.")
                If MinimumNonDecimalsInReturnString < 0 Then Throw New ArgumentException("MinimumNonDecimalsInReturnString cannot be lower than 0.")

                'Adding decimals to format
                Dim NumberFormat As String = "0"
                If DecimalsInReturnsString IsNot Nothing Then
                    For n = 0 To DecimalsInReturnsString - 1
                        If n = 0 Then NumberFormat &= "."
                        NumberFormat &= "0"
                    Next
                End If

                'Adding non-decimals to format
                If MinimumNonDecimalsInReturnString IsNot Nothing Then
                    For n = 0 To MinimumNonDecimalsInReturnString - 2 ' -2 as one 0 has already been added above
                        NumberFormat = "0" & NumberFormat
                    Next
                End If

                RetString = ReturnValue.ToString(NumberFormat).TrimEnd("0").Trim(".").Trim(",")
                If RetString = "" Then RetString = "0"

                Return RetString

            End If

            Return ReturnValue

        Catch ex As Exception
            MsgBox(ex.ToString)
            Return Nothing
        End Try

    End Function



    Public Shared Function Shuffle(ByVal Input As List(Of SpeechTestResponseAlternative), ByRef Randomizer As Random) As List(Of SpeechTestResponseAlternative)
        Dim SampleOrder = SampleWithoutReplacement(Input.Count, 0, Input.Count, Randomizer)
        Dim Output As New List(Of SpeechTestResponseAlternative)
        For Each RandomIndex In SampleOrder
            Output.Add(Input(RandomIndex))
        Next
        Return Output
    End Function



    Public Shared Sub ComplexMultiplication(Real1 As Double(), Imag1 As Double(), Real2 As Double(), Imag2 As Double())

        If Real1.Length <> Imag1.Length Or Real1.Length <> Imag1.Length Or Real1.Length <> Real2.Length Or Real1.Length <> Imag2.Length Then
            Throw New ArgumentException("Unequal length of input arrays")
        End If

        If Globals.StfBase.UseOptimizationLibraries = False Then

            'Performs complex multiplications
            Dim TempValue As Double = 0
            For n = 0 To Real1.Length - 1
                TempValue = Real1(n) 'stores this value so that it does not get overwritten in the following line (it needs to be used also two lines below)
                Real1(n) = TempValue * Real2(n) - Imag1(n) * Imag2(n)
                Imag1(n) = TempValue * Imag2(n) + Imag1(n) * Real2(n)
            Next

        Else

            LibOstfDsp_VB.ComplexMultiplication(Real1, Imag1, Real2, Imag2)

        End If

    End Sub


    Public Shared Sub CopyToDouble(SourceArray As Single(), TargetArray As Double())

        If TargetArray.Length < SourceArray.Length Then Throw New ArgumentException("TargetArray cannot be shorter than SourceArray")

        If Globals.StfBase.UseOptimizationLibraries = False Then
            For i = 0 To SourceArray.Length - 1
                TargetArray(i) = SourceArray(i)
            Next
        Else
            LibOstfDsp_VB.CopyToDouble(SourceArray, TargetArray)
        End If

    End Sub

    ''' <summary>
    ''' Multiplies each element in the Array1 array with the Factor using fast SIMD (Single Instruction, Multiple Data) operations
    ''' </summary>
    ''' <param name="Values"></param>
    ''' <returns></returns>
    Public Shared Sub MultiplyArray(Values() As Double, Factor As Double)

        If Globals.StfBase.UseOptimizationLibraries = False Then

            Dim VectorSize As Integer = System.Numerics.Vector(Of Double).Count
            Dim FactorVector = New System.Numerics.Vector(Of Double)(Factor)

            Dim i As Integer
            For i = 0 To Values.Length - VectorSize Step VectorSize
                Dim v As New System.Numerics.Vector(Of Double)(Values, i)
                v = v * FactorVector
                v.CopyTo(Values, i)
            Next

            ' Handle any remaining elements at the end that don't fit into a full vector.
            For i = i To Values.Length - 1
                Values(i) *= Factor
            Next

        Else

            LibOstfDsp_VB.MultiplyArray(Values, Factor)

        End If

    End Sub



    ''' <summary>
    ''' Multiplies each element in a section of the Array1 array with the Factor using fast SIMD (Single Instruction, Multiple Data) operations
    ''' </summary>
    ''' <param name="Values">The input array</param>
    ''' <param name="StartIndex">The start index of the section</param>
    ''' <param name="SectionLength">The length of the section</param>
    ''' <returns>The sum of squares of the specified section</returns>
    Public Shared Sub MultiplyArray(Values() As Single, Factor As Single, StartIndex As Integer, SectionLength As Integer)

        If Globals.StfBase.UseOptimizationLibraries = False Then

            Dim VectorSize As Integer = System.Numerics.Vector(Of Single).Count
            Dim FactorVector = New System.Numerics.Vector(Of Single)(Factor)

            ' Ensure we do not exceed the array bounds
            Dim endIndex As Integer = System.Math.Min(StartIndex + SectionLength, Values.Length)

            Dim i As Integer = StartIndex
            While i < endIndex AndAlso i + VectorSize <= endIndex ' Make sure there's enough room for a full vector
                Dim v As New System.Numerics.Vector(Of Single)(Values, i)
                v = v * FactorVector
                v.CopyTo(Values, i)
                i += VectorSize
            End While

            ' Handle any remaining elements at the end that don't fit into a full vector.
            While i < endIndex
                Values(i) *= Factor
                i += 1
            End While

        Else

            LibOstfDsp_VB.MultiplyArraySection(Values, Factor, StartIndex, SectionLength)

        End If

    End Sub

    ''' <summary>
    ''' Adds the two arrays. Either fast SIMD (Single Instruction, Multiple Data) operations are used, or if StfBase.UseOptimizationLibraries is True using the LibOstfDsp.
    ''' Arrays need to be the same lengths, otherwise an exception is thrown.
    ''' </summary>
    ''' <param name="Array1">The first input/output data array. Upon return this corresponding data array contains the sum of the values in array1 And array2</param>
    ''' <param name="Array2">The the input data array containing the values which should be added to array1</param>
    ''' <returns></returns>
    Public Shared Sub AddTwoArrays(Array1() As Single, Array2() As Single)

        If Array1.Length <> Array2.Length Then Throw New ArgumentException("Arrays 1 and 2 need to have the same lengths.")

        If Globals.StfBase.UseOptimizationLibraries = False Then

            Dim VectorSize As Integer = System.Numerics.Vector(Of Single).Count

            Dim i As Integer
            For i = 0 To Array1.Length - VectorSize Step VectorSize
                Dim v1 As New System.Numerics.Vector(Of Single)(Array1, i)
                Dim v2 As New System.Numerics.Vector(Of Single)(Array2, i)
                v1 += v2
                v1.CopyTo(Array1, i)
            Next

            ' Handle any remaining elements at the end that don't fit into a full vector.
            For i = i To Array1.Length - 1
                Array1(i) += Array2(i)
            Next


            'Untested paralell processing alternative
            'Parallel.For(0, Array1.Length \ VectorSize, Sub(i)
            '                                                Dim offset = i * VectorSize
            '                                                Dim v1 As New System.Numerics.Vector(Of Single)(Array1, offset)
            '                                                Dim v2 As New System.Numerics.Vector(Of Single)(Array2, offset)
            '                                                v1 += v2
            '                                                v1.CopyTo(Array1, offset)
            '                                            End Sub)

            '' Handle any remaining elements
            'For i = (Array1.Length \ VectorSize) * VectorSize To Array1.Length - 1
            '    Array1(i) += Array2(i)
            'Next

        Else
            LibOstfDsp_VB.AddTwoFloatArrays(Array1, Array2)
        End If

    End Sub

    ''' <summary>
    ''' Calculates the sum-of-square value of a section of an array using fast SIMD (Single Instruction, Multiple Data) operations
    ''' </summary>
    ''' <param name="Values">The input array</param>
    ''' <param name="startIndex">The start index of the section</param>
    ''' <param name="sectionLength">The length of the section</param>
    ''' <returns>The sum of squares of the specified section</returns>
    Public Shared Function CalculateSumOfSquare(Values() As Single, startIndex As Integer, sectionLength As Integer) As Single

        If Globals.StfBase.UseOptimizationLibraries = False Then

            Dim VectorSize As Integer = System.Numerics.Vector(Of Single).Count
            Dim SumOfSquaresVector As System.Numerics.Vector(Of Single) = System.Numerics.Vector(Of Single).Zero

            ' Ensure we do not exceed the array bounds
            Dim endIndex As Integer = System.Math.Min(startIndex + sectionLength, Values.Length)

            Dim i As Integer = startIndex
            While i < endIndex AndAlso i + VectorSize <= endIndex ' Make sure there's enough room for a full vector
                Dim v As New System.Numerics.Vector(Of Single)(Values, i)
                SumOfSquaresVector += v * v
                i += VectorSize
            End While

            Dim SumOfSquares As Single = 0

            For j As Integer = 0 To VectorSize - 1
                SumOfSquares += SumOfSquaresVector(j)
            Next

            ' Handle any remaining elements at the end that don't fit into a full vector.
            While i < endIndex
                SumOfSquares += Values(i) ^ 2
                i += 1
            End While

            Return SumOfSquares

        Else

            'Calculating the sum of sqares in libostfdsp
            Return LibOstfDsp_VB.CalculateSumOfSquare(Values, startIndex, sectionLength)

        End If

    End Function



    Public Shared Sub DeinterleaveSoundArray(interleavedArray As Single(), channelCount As Integer, channelLength As Integer, concatenatedArrays As Single())

        If Globals.StfBase.UseOptimizationLibraries = False Or Global.STFN.Core.Globals.StfBase.CurrentPlatForm = Global.STFN.Core.Platforms.WinUI Then 'TODO: Change this when windows opimization dlls implement this function

            ' Takes a flattened matrix in which each channel Is put after each other, And interleaves the channels values
            Dim targetIndex As Integer = 0
            For s = 0 To channelLength - 1
                For c = 0 To channelCount - 1
                    concatenatedArrays(c * channelLength + s) = interleavedArray(targetIndex)
                    targetIndex += 1
                Next
            Next

        Else
            LibOstfDsp_VB.DeinterleaveSoundArray(interleavedArray, channelCount, channelLength, concatenatedArrays)
        End If

    End Sub

    ''' <summary>
    ''' Unwraps the indicated angle into the range -180 (is lower than) Azimuth (which is equal to or lower than) 180 degrees.
    ''' </summary>
    ''' <param name="Angle">The angle in degrees</param>
    ''' <returns></returns>
    Public Shared Function UnwrapAngle(ByVal Angle As Double) As Double

        'Gets the remainder when dividing by 360
        Dim UnwrappedAngle As Double = Angle Mod 360

        'Sets the Azimuth in the following range: -180 < Azimuth <= 180
        If UnwrappedAngle > 180 Then UnwrappedAngle -= 360
        If UnwrappedAngle <= -180 Then UnwrappedAngle += 360

        Return UnwrappedAngle
    End Function

    Public Shared Function Degrees2Radians(ByVal Degrees As Double) As Double
        Return Degrees * System.Math.PI / 180
    End Function

    Public Shared Function Radians2Degrees(ByVal Radians As Double) As Double
        Return Radians * 180 / System.Math.PI
    End Function


#End Region

End Class




