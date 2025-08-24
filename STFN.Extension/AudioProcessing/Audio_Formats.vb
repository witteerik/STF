' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core.Audio.Formats

Namespace Audio


    Namespace Formats

        Public Enum SoundFileFormats
            wav
            ptwf
        End Enum


        <Serializable>
        Public Class SpectrogramFormat

            Public Property SpectrogramFftFormat As FftFormat
            Public Property SpectrogramPreFilterKernelFftFormat As FftFormat
            Public Property SpectrogramPreFirFilterFftFormat As FftFormat

            Public Property SpectrogramCutFrequency As Single
            Public Property UsePreFftFiltering As Boolean
            Public Property PreFftFilteringAttenuationRate As Single
            Public Property PreFftFilteringKernelSize As Integer
            Public Property PreFftFilteringKernelCreationAnalysisWindowSize As Integer
            Public Property PreFftFilteringAnalysisWindowSize As Integer
            Public Property InActivateWarnings As Boolean

            Public Sub New(Optional ByVal setSpectrogramCutFrequency As Single = 8000, Optional spectrogramAnalysisWindowSize As Integer = 1024, Optional spectrogramFftSize As Integer = 1024,
        Optional spectrogramAnalysisWindowOverlapSize As Integer = 512,
                   Optional spectrogramWindowingMethod As DSP.WindowingType = DSP.WindowingType.Hamming,
                   Optional ByVal setUsePreFftFiltering As Boolean = True,
                   Optional ByVal setPreFftFilteringAttenuationRate As Single = 6, Optional ByVal setPreFftFilteringKernelSize As Integer = 2000,
                   Optional ByVal setPreFftFilteringAnalysisWindowSize As Integer = 1024,
                   Optional setInActivateWarnings As Boolean = False)

                InActivateWarnings = setInActivateWarnings
                UsePreFftFiltering = setUsePreFftFiltering
                PreFftFilteringAttenuationRate = setPreFftFilteringAttenuationRate
                PreFftFilteringKernelSize = setPreFftFilteringKernelSize
                PreFftFilteringAnalysisWindowSize = setPreFftFilteringAnalysisWindowSize
                PreFftFilteringKernelCreationAnalysisWindowSize = setPreFftFilteringAnalysisWindowSize

                SpectrogramCutFrequency = setSpectrogramCutFrequency
                SpectrogramFftFormat = New FftFormat(spectrogramAnalysisWindowSize, spectrogramFftSize, spectrogramAnalysisWindowOverlapSize, spectrogramWindowingMethod, InActivateWarnings)
                SpectrogramPreFilterKernelFftFormat = New FftFormat(PreFftFilteringKernelCreationAnalysisWindowSize,,, DSP.WindowingType.Hamming, InActivateWarnings) 'TODO should there be a windowing function specified here?
                SpectrogramPreFirFilterFftFormat = New FftFormat(PreFftFilteringAnalysisWindowSize,,, DSP.WindowingType.Hamming, InActivateWarnings) 'TODO should there be a windowing function specified here?

            End Sub

        End Class



    End Namespace



End Namespace
