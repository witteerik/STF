' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte


Namespace Audio

    Namespace Formats

        Public Enum SupportedWaveFormatsEncodings
            PCM = 1
            IeeeFloatingPoints = 3
        End Enum

        <Serializable>
        Public Class WaveFormat

            Public ReadOnly Property SampleRate As UInteger
            Public ReadOnly Property BitDepth As UShort
            Public ReadOnly Property Channels As Short
            Public ReadOnly Property Encoding As WaveFormatEncodings 'This property is acctually redundant since it hold the same numeric value as FmtCode. It is retained for now since it clearly specifies the names of the formats, and FmtCode need to be in UShort

            Public Property Name As String
            Public ReadOnly Property MainChunkID As String = "RIFF"
            Public ReadOnly Property RiffType As String = "WAVE"
            Public ReadOnly Property FmtID As String = "fmt "
            Public ReadOnly Property FmtSize As UInteger = 16
            Public ReadOnly Property FmtCode As UShort
            Public ReadOnly Property FmtAvgBPS As UInteger
            Public ReadOnly Property FmtBlockAlign As UShort
            Public ReadOnly Property DataID As String = "data"
            Public ReadOnly Property PositiveFullScale As Double
            Public ReadOnly Property NegativeFullScale As Double

            Public Enum WaveFormatEncodings
                PCM = 1
                IeeeFloatingPoints = 3
                'WAVE_FORMAT_EXTENSIBLE = 65534
            End Enum

            Public Sub New(ByVal SampleRate As Integer, ByVal BitDepth As Integer, ByVal Channels As Integer,
                       Optional Name As String = "",
                       Optional Encoding As WaveFormatEncodings = WaveFormatEncodings.IeeeFloatingPoints)

                Me.SampleRate = SampleRate
                Me.BitDepth = BitDepth
                Me.Name = Name

                'Setting bit depth, and checks that the bit depth is supported
                If Not CheckIfBitDepthIsSupported(Encoding, Me.BitDepth) = True Then Throw New NotImplementedException("BitDepth " & Me.BitDepth & " Is Not yet implemented.")

                'Setting format code (automatically converting the numeric value of the enumerator to UShort)
                FmtCode = Encoding
                Me.Encoding = Encoding

                'Set full scale
                Select Case Me.Encoding
                    Case WaveFormatEncodings.PCM
                        Select Case Me.BitDepth
                            Case 8  'Byte
                                'Note that 8 bit wave has the numeric range of 0-255,
                                PositiveFullScale = Byte.MaxValue
                                NegativeFullScale = Byte.MinValue + 1
                            Case 16 'Short
                                PositiveFullScale = Short.MaxValue
                                NegativeFullScale = Short.MinValue + 1
                            Case 32 'Integer
                                PositiveFullScale = Integer.MaxValue
                                NegativeFullScale = Integer.MinValue + 1
                            Case Else
                                Throw New NotImplementedException(Me.BitDepth & " bit depth is not yet supported for the PCM wave format.")
                        End Select
                    Case WaveFormatEncodings.IeeeFloatingPoints
                        Select Case Me.BitDepth
                            Case 32 'Single
                                PositiveFullScale = 1
                                NegativeFullScale = -1
                            Case 64 'Double
                                PositiveFullScale = 1
                                NegativeFullScale = -1
                            Case Else
                                Throw New NotImplementedException(Me.BitDepth & " bit depth is not yet supported for Ieee Floating Points format.")
                        End Select
                End Select

                'Setting (and correcting) number of channels
                If Channels < 1 Then Channels = 1
                Me.Channels = Channels

                'Calculating and setting FmtAvgBPS and FmtBlockAlign 
                FmtAvgBPS = Me.SampleRate * Me.Channels * (Me.BitDepth / 8)
                FmtBlockAlign = Me.Channels * (Me.BitDepth / 8)

            End Sub

            ''' <summary>
            ''' Determines whether values are the same between the current WaveFormat and ComparisonFormat
            ''' </summary>
            ''' <param name="ComparisonFormat"></param>
            ''' <returns></returns>
            Public Function IsEqual(ByRef ComparisonFormat As WaveFormat,
                                    Optional ByVal CompareChannels As Boolean = True,
                                    Optional ByVal CompareSampleRate As Boolean = True,
                                    Optional ByVal CompareBitDepth As Boolean = True,
                                    Optional ByVal CompareEncoding As Boolean = True) As Boolean

                If CompareChannels = True Then If ComparisonFormat.Channels <> Me.Channels Then Return False
                If CompareSampleRate = True Then If ComparisonFormat.SampleRate <> Me.SampleRate Then Return False
                If CompareBitDepth = True Then If ComparisonFormat.BitDepth <> Me.BitDepth Then Return False
                If CompareEncoding = True Then If ComparisonFormat.Encoding <> Me.Encoding Then Return False

                Return True

            End Function

            Public Overrides Function ToString() As String

                Dim Output As New List(Of String)

                Output.Add("Samplerate: " & SampleRate.ToString & " Hz")
                Output.Add("Bit depth: " & BitDepth.ToString)
                Output.Add("Channels: " & Channels.ToString)
                Output.Add("Encoding: " & Encoding.ToString)

                Return String.Join(vbCrLf, Output)

            End Function

        End Class


        <Serializable>
        Public Class FftFormat

            Public Property FftWindowSize As Integer
            Public Property AnalysisWindowSize As Integer
            Public Property OverlapSize As Integer
            Public Property WindowingType As DSP.WindowingType = WindowingType.Rectangular
            Public Property Tukey_r As Double

            ''' <summary>
            ''' Creates a new fft format.
            ''' </summary>
            ''' <param name="setAnalysisWindowSize">Determines the length in samples of each part of the wave form that will be analysed in fft.</param>
            ''' <param name="setFftWindowSize">Determines the FFT length (frequency resolution) that will be used when calculating fft (is automatically set to the length of the analysis window if left emtpy).</param>
            ''' <param name="setoverlapSize">Determines the overlap between two analysis windows (in samples.)</param>
            ''' <param name="setWindowing">Determines which windowing function of the analysis window that will be used the before the fft calculation.</param>
            ''' <param name="Tukey_r">The ratio between windowing size and the cosine sections in a Tukey window. Only needed if Tukey windowing is used.</param>
            Public Sub New(Optional ByRef setAnalysisWindowSize As Integer = 1024, Optional ByRef setFftWindowSize As Integer = -1, Optional ByRef setoverlapSize As Integer = 0,
                   Optional setWindowing As DSP.WindowingType = DSP.WindowingType.Rectangular, Optional ByRef InActivateWarnings As Boolean = False,
                   Optional Tukey_r As Double = 0.5)

                'Adjusting setAnalysisWindowSize
                If setAnalysisWindowSize < 0 Then setAnalysisWindowSize = 1024
                If setAnalysisWindowSize Mod 2 = 1 Then setAnalysisWindowSize += 1
                AnalysisWindowSize = setAnalysisWindowSize

                'Adding default value for setFftWindowSize, if left empty
                If setFftWindowSize < 0 Then setFftWindowSize = AnalysisWindowSize

                'Checking fft size
                DSP.CheckAndAdjustFFTSize(setFftWindowSize, AnalysisWindowSize, InActivateWarnings)
                FftWindowSize = setFftWindowSize

                'Checking overlap size
                If Not setoverlapSize < AnalysisWindowSize Then setoverlapSize = AnalysisWindowSize - 1
                OverlapSize = setoverlapSize

                WindowingType = setWindowing
                Me.Tukey_r = Tukey_r
            End Sub


        End Class


        <Serializable>
        Public Class SoundLevelFormat

            Private _SoundMeasurementType As SoundMeasurementTypes
            Public Property SoundMeasurementType As SoundMeasurementTypes
                Get
                    Return _SoundMeasurementType
                End Get
                Set(value As SoundMeasurementTypes)
                    _SoundMeasurementType = value
                    Select Case value
                        Case SoundMeasurementTypes.LoudestSection_Z_Weighted
                            LoudestSectionMeasurement = True
                            FrequencyWeighting = FrequencyWeightings.Z

                        Case SoundMeasurementTypes.LoudestSection_C_Weighted
                            LoudestSectionMeasurement = True
                            FrequencyWeighting = FrequencyWeightings.C

                        Case SoundMeasurementTypes.LoudestSection_RLB_Weighted
                            LoudestSectionMeasurement = True
                            FrequencyWeighting = FrequencyWeightings.RLB

                        Case SoundMeasurementTypes.LoudestSection_K_Weighted
                            LoudestSectionMeasurement = True
                            FrequencyWeighting = FrequencyWeightings.K

                        Case SoundMeasurementTypes.Average_C_Weighted
                            LoudestSectionMeasurement = False
                            FrequencyWeighting = FrequencyWeightings.C

                        Case SoundMeasurementTypes.Average_RLB_Weighted
                            LoudestSectionMeasurement = False
                            FrequencyWeighting = FrequencyWeightings.RLB

                        Case SoundMeasurementTypes.Average_K_Weighted
                            LoudestSectionMeasurement = False
                            FrequencyWeighting = FrequencyWeightings.K

                        Case SoundMeasurementTypes.Average_Z_Weighted
                            LoudestSectionMeasurement = False
                            FrequencyWeighting = FrequencyWeightings.Z

                        Case Else
                            Throw New NotImplementedException("Unsupported frequency weighting.")
                    End Select
                End Set
            End Property

            Public Enum SoundMeasurementTypes
                Average_Z_Weighted = 0
                LoudestSection_Z_Weighted = 100
                Average_A_Weighted = 1
                LoudestSection_A_Weighted = 101
                Average_C_Weighted = 2
                LoudestSection_C_Weighted = 102
                Average_RLB_Weighted = 3
                LoudestSection_RLB_Weighted = 103
                Average_K_Weighted = 4
                LoudestSection_K_Weighted = 104
            End Enum

            Public Property LoudestSectionMeasurement As Boolean
            ''' <summary>
            ''' The lengths of the sound level measurement sections. (Should appropriately be set to the auditory temporal integration time (appr 0.1-0.2 s))
            ''' </summary>
            ''' <returns></returns>
            Public Property TemporalIntegrationDuration As Decimal
            Public Property FrequencyWeighting As FrequencyWeightings

            ''' <summary>
            ''' Creates a new instance of the SoundLevelFormat class
            ''' </summary>
            ''' <param name="DefaultSoundMeasurementType"></param>
            ''' <param name="DefaultTemporalIntegrationDuration"></param>
            Public Sub New(DefaultSoundMeasurementType As SoundMeasurementTypes,
                       Optional DefaultTemporalIntegrationDuration As Decimal = 0.05)

                SoundMeasurementType = DefaultSoundMeasurementType
                TemporalIntegrationDuration = DefaultTemporalIntegrationDuration

            End Sub

            ''' <summary>
            ''' Compares the instance of SoundLevelFormat to another instance.
            ''' </summary>
            ''' <param name="SoundLevelFormat"></param>
            ''' <returns>Returns False if the compared instances differ in SoundMeasurementType, and if time weighting is used, also if the temporal integration time differs. Otherwise returns True</returns>
            Public Function IsEqual(ByVal SoundLevelFormat As SoundLevelFormat) As Boolean

                If SoundMeasurementType <> SoundLevelFormat.SoundMeasurementType Then Return False

                If LoudestSectionMeasurement Then
                    If TemporalIntegrationDuration <> SoundLevelFormat.TemporalIntegrationDuration Then Return False
                End If

                Return True

            End Function

        End Class

    End Namespace


End Namespace