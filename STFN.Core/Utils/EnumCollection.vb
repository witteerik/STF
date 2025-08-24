' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte


Namespace Utils

    Public Module EnumCollection

        Public Enum Sides
            Left
            Right
        End Enum

        Public Enum ElevationChange
            Ascending
            Unchanged
            Descendning
        End Enum

        Public Enum SidesWithBoth
            Left
            Right
            Both
        End Enum

        Public Enum Languages
            English
            Swedish
        End Enum

        Public Enum UserTypes
            Research
            Clinical
        End Enum

        Public Enum SoundPropagationTypes
            PointSpeakers
            SimulatedSoundField
            Ambisonics
        End Enum




    End Module

End Namespace


'Enums directly available under STFN.Core
Public Enum Platforms ' These reflects the platform names currently specified in NET MAUI. Not all will work with STF.
    iOS
    WinUI
    UWP
    Tizen
    tvOS
    MacCatalyst
    macOS
    watchOS
    Unknown
    Android
End Enum

Public Enum TimeUnits
    samples
    seconds
    pixels
End Enum

Public Enum FrequencyWeightings
    A
    C
    Z
    RLB
    K
End Enum

Public Module AudiologyTypes

    Public Enum BmldModes
        RightOnly
        LeftOnly
        BinauralSamePhase
        BinauralPhaseInverted
        BinauralUncorrelated 'Noise only
    End Enum

End Module