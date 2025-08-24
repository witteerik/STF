' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports System.Runtime.CompilerServices
Imports STFN.Core
Imports STFN.Core.Messager

'This file contains extension methods for STFN.Core.StfBase

Public Class SftBaseExtension
    Inherits STFN.Core.StfBase

    Protected _PortAudioIsInitialized As Boolean = False
    ''' <summary>
    ''' Returns True if the PortAudio library has been successfylly initialized by call the the StfBase function InitializeSTF.
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property PortAudioIsInitialized As Boolean
        Get
            Return _PortAudioIsInitialized
        End Get
    End Property


    Public Overrides Sub InitializeSTF(ByVal PlatForm As Platforms, ByVal MediaPlayerType As AudioSystemSpecification.MediaPlayerTypes, Optional ByVal MediaRootDirectory As String = "")

        'Exits the sub to avoid multiple calls (which should be avoided, especially to the Audio.PortAudio.Pa_Initialize function). (N.B. The call to MyBase.InitializeSTF sets StfIsInitialized to True so it does not have to be done here.)
        If StfIsInitialized = True Then Exit Sub

        MyBase.InitializeSTF(PlatForm, MediaPlayerType, MediaRootDirectory)

        'Initiating PortAudio if PaBased MediaPlayerType is used
        If Global.STFN.Core.Globals.StfBase.CurrentMediaPlayerType = Global.STFN.Core.AudioSystemSpecification.MediaPlayerTypes.PaBased Then

            'Initializing the port audio library
            If Audio.PortAudio.Pa_GetDeviceCount = Audio.PortAudio.PaError.paNotInitialized Then
                Dim Pa_Initialize_ReturnValue = Audio.PortAudio.Pa_Initialize
                If Pa_Initialize_ReturnValue = Global.STFN.Extension.Audio.PortAudio.PaError.paNoError Then
                    _PortAudioIsInitialized = True
                Else
                    Throw New Exception("Unable to initialize PortAudio library for audio processing." & vbCrLf & vbCrLf &
                                    "The following error occurred: " & vbCrLf & vbCrLf &
                                    Audio.PortAudio.Pa_GetErrorText(Pa_Initialize_ReturnValue))
                    ' if Pa_Initialize() returns an error code, 
                    ' Pa_Terminate() should NOT be called.
                End If
            Else
                'If the we end up here, PortAudio will have already bee initialized, which it should not. (Then there will be missing calls to Pa_Terminate.)
            End If
        End If


        'Initializing the sound player
        If Global.STFN.Core.Globals.StfBase.CurrentMediaPlayerType = Global.STFN.Core.AudioSystemSpecification.MediaPlayerTypes.PaBased Then
            Dim SoundPlayer As STFN.Core.Audio.SoundPlayers.iSoundPlayer = New Audio.PortAudioVB.PortAudioBasedSoundPlayer(False, False, False, False)
            Globals.StfBase.InitializeSoundPlayer(SoundPlayer)
        End If


    End Sub


    ''' <summary>
    ''' Loading the audio systems specifications.
    ''' </summary>
    ''' <param name="InputLines">An array of (text) lines in the audio system specification (.txt) file. </param>
    ''' <param name="AudioSystemSpecificationFilePath">The file path to the audio systems specifications file (used only for error reporting)</param>
    Public Overrides Sub LoadAudioSystemSpecifications(ByVal InputLines() As String, ByVal AudioSystemSpecificationFilePath As String)


        'Reads first all lines and sort them into a player-MediaPlayerType specific dictionary
        Dim PlayerTypeDictionary As New SortedList(Of AudioSystemSpecification.MediaPlayerTypes, String())
        Dim CurrentPlayerType As AudioSystemSpecification.MediaPlayerTypes? = Nothing

        Dim CurrentSoundPlayerList As List(Of String) = Nothing
        For i = 0 To InputLines.Length - 1
            Dim Line As String = InputLines(i).Trim
            If Line.StartsWith("<New media player>") Then
                If CurrentSoundPlayerList IsNot Nothing Then
                    'Storing the loaded sound player data in PlayerTypeDictionary
                    If PlayerTypeDictionary.ContainsKey(CurrentPlayerType) Then
                        Throw New Exception("In the file " & AudioSystemSpecificationFilePath & " each MediaPlayerType (PaBased, AudioTrackBased, etc.) can only be specified once. It seems as the MediaPlayerType " & CurrentPlayerType.ToString & " occurres multiple times.")
                    End If
                    PlayerTypeDictionary.Add(CurrentPlayerType, CurrentSoundPlayerList.ToArray)
                End If
                'Creating a new CurrentSoundPlayerList 
                CurrentSoundPlayerList = New List(Of String)
            ElseIf Line.StartsWith("MediaPlayerType") Then
                CurrentPlayerType = InputFileSupport.InputFileEnumValueParsing(Line, GetType(AudioSystemSpecification.MediaPlayerTypes), AudioSystemSpecificationFilePath, True)
            ElseIf Line.Trim = "" Then
                'Just ignores empty lines
            Else
                If CurrentSoundPlayerList IsNot Nothing Then
                    CurrentSoundPlayerList.Add(Line)
                Else
                    Throw New Exception("The file " & AudioSystemSpecificationFilePath & " must begin with by specfiying a <New media MediaPlayerType> line followed by a MediaPlayerType line (e.g. MediaPlayerType = PaBased)")
                End If
            End If
        Next

        If CurrentSoundPlayerList IsNot Nothing Then
            'Also storing the last loaded sound player data in PlayerTypeDictionary
            PlayerTypeDictionary.Add(CurrentPlayerType, CurrentSoundPlayerList.ToArray)
        End If

        Dim PlayerWasLoaded As Boolean = False

        'Looking for and loading the settings for the first player of the intended MediaPlayerType
        For Each PlayerType In PlayerTypeDictionary

            'Skipping the player MediaPlayerType if it is not the CurrentMediaPlayerType 
            If PlayerType.Key <> Globals.StfBase.CurrentMediaPlayerType Then Continue For

            Try

                Dim MediaPlayerInputLines = PlayerType.Value

                'Reads the API settings, and tries to select the API and device if available, otherwise lets the user select a device manually

                Dim ApiName As String = ""
                Dim OutputDeviceName As String = ""
                Dim OutputDeviceNames As List(Of String) = Nothing ' Used for MME multiple device support
                Dim InputDeviceName As String = ""
                Dim InputDeviceNames As List(Of String) = Nothing ' Used for MME multiple device support
                Dim BufferSize As Integer = 2048
                Dim AllowDefaultOutputDevice As Boolean? = Nothing
                Dim AllowDefaultInputDevice As Boolean? = Nothing

                Dim LinesRead As Integer = 0
                For i = 0 To MediaPlayerInputLines.Length - 1
                    LinesRead += 1
                    Dim Line As String = MediaPlayerInputLines(i).Trim

                    'Skips empty and outcommented lines
                    If Line = "" Then Continue For
                    If Line.StartsWith("//") Then Continue For

                    'If Line.StartsWith("<AudioDevices>") Then
                    '    'No need to do anything?
                    '    Continue For
                    'End If
                    If Line.StartsWith("<New transducer>") Then Exit For

                    If Global.STFN.Core.Globals.StfBase.CurrentMediaPlayerType = Global.STFN.Core.AudioSystemSpecification.MediaPlayerTypes.PaBased Then
                        If Line.StartsWith("ApiName") Then ApiName = InputFileSupport.GetInputFileValue(Line, True)
                    End If

                    If Line.Replace(" ", "").StartsWith("OutputDevice=") Then OutputDeviceName = InputFileSupport.GetInputFileValue(Line, True)
                    If Line.Replace(" ", "").StartsWith("OutputDevices=") Then OutputDeviceNames = InputFileSupport.InputFileListOfStringParsing(Line, False, True)
                    If Line.Replace(" ", "").StartsWith("InputDevice=") Then InputDeviceName = InputFileSupport.GetInputFileValue(Line, True)
                    If Line.Replace(" ", "").StartsWith("InputDevices=") Then InputDeviceNames = InputFileSupport.InputFileListOfStringParsing(Line, False, True)
                    If Line.StartsWith("BufferSize") Then BufferSize = InputFileSupport.InputFileIntegerValueParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("AllowDefaultOutputDevice") Then AllowDefaultOutputDevice = InputFileSupport.InputFileBooleanValueParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("AllowDefaultInputDevice") Then AllowDefaultInputDevice = InputFileSupport.InputFileBooleanValueParsing(Line, True, AudioSystemSpecificationFilePath)

                Next

                Dim AudioSettings As STFN.Core.Audio.AudioSettings = Nothing

                If Global.STFN.Core.Globals.StfBase.CurrentMediaPlayerType = Global.STFN.Core.AudioSystemSpecification.MediaPlayerTypes.PaBased Then

                    Dim DeviceLoadSuccess As Boolean = True
                    If OutputDeviceName = "" And InputDeviceName = "" And OutputDeviceNames Is Nothing And InputDeviceNames Is Nothing Then
                        'No device names have been specified
                        DeviceLoadSuccess = False
                    End If

                    If ApiName <> "MME" Then
                        If OutputDeviceNames IsNot Nothing Or InputDeviceNames IsNot Nothing Then
                            DeviceLoadSuccess = False
                            MsgBox("When specifying multiple sound (input or output) devices in the file " & AudioSystemSpecificationFilePath & ", the sound API must be MME ( not " & ApiName & ")!", MsgBoxStyle.Exclamation, "Sound device specification error!")
                        End If
                    End If

                    If OutputDeviceNames IsNot Nothing And OutputDeviceName <> "" Then
                        DeviceLoadSuccess = False
                        MsgBox("Either (not both) of single or multiple sound output devices must be specified in the file " & AudioSystemSpecificationFilePath & "!", MsgBoxStyle.Exclamation, "Sound device specification error!")
                    End If

                    If InputDeviceNames IsNot Nothing And InputDeviceName <> "" Then
                        DeviceLoadSuccess = False
                        MsgBox("Either (not both) of single or multiple sound input devices must be specified in the file " & AudioSystemSpecificationFilePath & "!", MsgBoxStyle.Exclamation, "Sound device specification error!")
                    End If

                    'Tries to setup the PortAudioApiSettings using the loaded data
                    AudioSettings = New Audio.PortAudioApiSettings

                    If AllowDefaultOutputDevice Is Nothing Then
                        DeviceLoadSuccess = False
                        MsgBox("The AllowDefaultOutputDevice behaviour must be specified in the file " & AudioSystemSpecificationFilePath & "!" & vbCrLf & vbCrLf &
                               "Use either:" & vbCrLf & "AllowDefaultOutputDevice = True" & vbCrLf & "or" & vbCrLf & "AllowDefaultOutputDevice = False" & vbCrLf & vbCrLf &
                               "Press OK to close the application.", MsgBoxStyle.Exclamation, "Sound device specification error!")
                        Messager.RequestCloseApp()
                    End If

                    AudioSettings.AllowDefaultOutputDevice = AllowDefaultOutputDevice

                    If AllowDefaultInputDevice Is Nothing Then
                        DeviceLoadSuccess = False
                        MsgBox("The AllowDefaultInputDevice behaviour must be specified in the file " & AudioSystemSpecificationFilePath & "!" & vbCrLf & vbCrLf &
                               "Use either:" & vbCrLf & "AllowDefaultInputDevice = True" & vbCrLf & "or" & vbCrLf & "AllowDefaultInputDevice = False" & vbCrLf & vbCrLf &
                               "Press OK to close the application.", MsgBoxStyle.Exclamation, "Sound device specification error!")
                        Messager.RequestCloseApp()
                    End If

                    AudioSettings.AllowDefaultInputDevice = AllowDefaultInputDevice

                    If DeviceLoadSuccess = True Then
                        If ApiName = "ASIO" Then
                            DeviceLoadSuccess = DirectCast(AudioSettings, Audio.PortAudioApiSettings).SetAsioSoundDevice(OutputDeviceName, BufferSize)
                        Else
                            If OutputDeviceNames Is Nothing And InputDeviceNames Is Nothing Then
                                DeviceLoadSuccess = DirectCast(AudioSettings, Audio.PortAudioApiSettings).SetNonAsioSoundDevice(ApiName, OutputDeviceName, InputDeviceName, BufferSize)
                            Else
                                DeviceLoadSuccess = DirectCast(AudioSettings, Audio.PortAudioApiSettings).SetMmeMultipleDevices(InputDeviceNames, OutputDeviceNames, BufferSize)
                            End If
                        End If
                    End If

                    If DeviceLoadSuccess = False Then

                        If OutputDeviceNames Is Nothing Then OutputDeviceNames = New List(Of String)
                        If InputDeviceNames Is Nothing Then InputDeviceNames = New List(Of String)

                        If AllowDefaultOutputDevice = True Then
                            MsgBox("The Speech Test Framework (STF) was unable to load the sound API (" & ApiName & ") and device/s indicated in the file " & AudioSystemSpecificationFilePath & vbCrLf & vbCrLf &
                        "Output device: " & OutputDeviceName & vbCrLf &
                        "Output devices: " & String.Join(", ", OutputDeviceNames) & vbCrLf &
                        "Input device: " & InputDeviceName & vbCrLf &
                        "Input devices: " & String.Join(", ", InputDeviceNames) & vbCrLf & vbCrLf &
                        "Click OK to use the default input/output devices." & vbCrLf & vbCrLf &
                        "IMPORTANT: Sound tranducer calibration and/or routing may not be correct!", MsgBoxStyle.Exclamation, "STF sound device not found!")

                            DirectCast(AudioSettings, Audio.PortAudioApiSettings).SelectDefaultAudioDevice()
                        Else
                            MsgBox("Selecting default device for PaBased sound players has been disabled in the audio system specifications file " & AudioSystemSpecificationFilePath & vbCrLf & vbCrLf & "Press OK to close the application.")
                            Messager.RequestCloseApp()
                        End If
                        'End If
                    End If

                ElseIf Global.STFN.Core.Globals.StfBase.CurrentMediaPlayerType = Global.STFN.Core.AudioSystemSpecification.MediaPlayerTypes.AudioTrackBased Then

                    Dim DeviceLoadSuccess As Boolean = True
                    If OutputDeviceName = "" Then
                        'No device names have been specified
                        MsgBox("An output device must be specified for the android AudioTrack based player in the file " & AudioSystemSpecificationFilePath & "!", MsgBoxStyle.Exclamation, "Sound device specification error!")
                        DeviceLoadSuccess = False
                    End If

                    'No input device is required
                    'If InputDeviceName = "" Then
                    '    DeviceLoadSuccess = False
                    'MsgBox("An input device must be specified for the android AudioTrack based player in the file " & AudioSystemSpecificationFilePath & "!", MsgBoxStyle.Exclamation, "Sound device specification error!")
                    'End If

                    'Setting up the player must be done in STFM as there is no access to Android AudioTrack in STFN, however the object holding the AudioSettings must be created here as it needs to be referenced in the Transducers below
                    AudioSettings = New STFN.Core.Audio.AndroidAudioTrackPlayerSettings
                    AudioSettings.AllowDefaultOutputDevice = AllowDefaultOutputDevice
                    AudioSettings.AllowDefaultInputDevice = AllowDefaultInputDevice
                    DirectCast(AudioSettings, STFN.Core.Audio.AndroidAudioTrackPlayerSettings).SelectedOutputDeviceName = OutputDeviceName
                    DirectCast(AudioSettings, STFN.Core.Audio.AndroidAudioTrackPlayerSettings).SelectedInputDeviceName = InputDeviceName
                    AudioSettings.FramesPerBuffer = BufferSize

                Else
                    Throw New NotImplementedException("Unknown media player MediaPlayerType specified in the file " & AudioSystemSpecificationFilePath)
                End If


                'Reads the remains of the file
                _AvaliableTransducers = New List(Of AudioSystemSpecification)

                'Backs up one line
                LinesRead = Math.Max(0, LinesRead - 1)

                Dim CurrentTransducer As AudioSystemSpecification = Nothing
                For i = LinesRead To MediaPlayerInputLines.Length - 1
                    Dim Line As String = MediaPlayerInputLines(i).Trim

                    'Skips empty and outcommented lines
                    If Line = "" Then Continue For
                    If Line.StartsWith("//") Then Continue For

                    If Line.StartsWith("<New transducer>") Then
                        If CurrentTransducer Is Nothing Then
                            'Creates the first transducer
                            CurrentTransducer = New AudioSystemSpecification(CurrentMediaPlayerType, AudioSettings)
                        Else
                            'Stores the transducer
                            _AvaliableTransducers.Add(CurrentTransducer)
                            'Creates a new one
                            CurrentTransducer = New AudioSystemSpecification(CurrentMediaPlayerType, AudioSettings)
                        End If
                    End If

                    If Line.StartsWith("Name") Then CurrentTransducer.Name = InputFileSupport.GetInputFileValue(Line, True)
                    If Line.StartsWith("LoudspeakerAzimuths") Then CurrentTransducer.LoudspeakerAzimuths = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("LoudspeakerElevations") Then CurrentTransducer.LoudspeakerElevations = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("LoudspeakerDistances") Then CurrentTransducer.LoudspeakerDistances = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("HardwareOutputChannels") Then CurrentTransducer.HardwareOutputChannels = InputFileSupport.InputFileListOfIntegerParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("CalibrationGain") Then CurrentTransducer.CalibrationGain = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("HostVolumeOutputLevel") Then CurrentTransducer.HostVolumeOutputLevel = InputFileSupport.InputFileDoubleValueParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("PtaCalibrationGainFrequencies") Then CurrentTransducer.PtaCalibrationGainFrequencies = InputFileSupport.InputFileListOfIntegerParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("PtaCalibrationGainValues") Then CurrentTransducer.PtaCalibrationGainValues = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("RETSPL_Speech") Then CurrentTransducer.RETSPL_Speech = InputFileSupport.InputFileDoubleValueParsing(Line, True, AudioSystemSpecificationFilePath)
                    If Line.StartsWith("LimiterThreshold") Then CurrentTransducer.LimiterThreshold = InputFileSupport.InputFileDoubleValueParsing(Line, True, AudioSystemSpecificationFilePath)

                Next

                'Stores the last transducer
                If CurrentTransducer IsNot Nothing Then _AvaliableTransducers.Add(CurrentTransducer)

                'Adding a default transducer if none were sucessfully read
                If _AvaliableTransducers.Count = 0 Then _AvaliableTransducers.Add(New AudioSystemSpecification(CurrentMediaPlayerType, AudioSettings))

                For Each Transducer In _AvaliableTransducers
                    Transducer.SetupMixer()
                Next

                'Checking calibration gain values and issues warnings if calibration gain is above 30 dB
                For Each Transducer In _AvaliableTransducers
                    For i = 0 To Transducer.CalibrationGain.Count - 1
                        If Transducer.CalibrationGain(i) > 30 Then
                            MsgBox("Calibration gain number " & i & " for the audio transducer '" & Transducer.Name & "' exceeds 30 dB. " & vbCrLf & vbCrLf &
                                   "Make sure that this is really correct before you continue and be cautios not to inflict personal injuries or damage your equipment if you continue! " & vbCrLf & vbCrLf &
                                   "This calibration value is set in the audio system specifications file: " & AudioSystemSpecificationFilePath, MsgBoxStyle.Exclamation, "Warning - High calibration gain value!")
                        End If
                    Next

                    If Transducer.PtaCalibrationGainFrequencies IsNot Nothing And Transducer.PtaCalibrationGainValues IsNot Nothing Then
                        If Transducer.PtaCalibrationGainFrequencies.Count = Transducer.PtaCalibrationGainValues.Count Then
                            Transducer.PtaCalibrationGain = New SortedList(Of Integer, Double)
                            For i = 0 To Transducer.PtaCalibrationGainFrequencies.Count - 1
                                Transducer.PtaCalibrationGain.Add(Transducer.PtaCalibrationGainFrequencies(i), Transducer.PtaCalibrationGainValues(i))
                            Next
                        Else

                            MsgBox("The number of specified PTA calibration gain values and frequencies do not match for the audio transducer '" & Transducer.Name & "'" & vbCrLf & vbCrLf &
                                   "Press OK to closing the application! (These calibration values are set in the audio system specifications file: " & AudioSystemSpecificationFilePath, MsgBoxStyle.Exclamation, "Warning - error in PTA calibration values!")
                            Messager.RequestCloseApp()
                        End If

                    End If


                Next

                PlayerWasLoaded = True

                'Exits the loop after the first MediaPlayerType successfully read
                Exit For

            Catch ex As Exception
                MsgBox("An error occurred trying to parse the file: " & AudioSystemSpecificationFilePath & vbCrLf & " The application may not work as intended!")
            End Try

        Next

        If PlayerWasLoaded = False Then
            MsgBox("Unable to load any media player with the MediaPlayerType " & Globals.StfBase.CurrentMediaPlayerType.ToString & vbCrLf & " The application may not work as intended!")
        End If




    End Sub



    ''' <summary>
    ''' This sub needs to be called when closing the last STFN application.
    ''' </summary>
    Public Overrides Sub TerminateSTF()

        MyBase.TerminateSTF()

        If Global.STFN.Core.Globals.StfBase.CurrentMediaPlayerType = Global.STFN.Core.AudioSystemSpecification.MediaPlayerTypes.PaBased Then
            'Terminating Port Audio
            If PortAudioIsInitialized Then Audio.PortAudio.Pa_Terminate()
        End If

    End Sub


    'Protected Shared Sub LoadAudioSystemSpecificationFile_OLD( )

    '    Dim AudioSystemSpecificationFilePath = IO.Path.Combine(StfBase.MediaRootDirectory, StfBase.AudioSystemSettingsFile)

    '    'Reads the API settings, and tries to select the API and device if available, otherwise lets the user select a device manually

    '    ''Getting calibration file descriptions from the text file SignalDescriptions.txt
    '    Dim LinesRead As Integer = 0
    '    Dim InputLines() As String = System.IO.File.ReadAllLines(AudioSystemSpecificationFilePath, System.Text.Encoding.UTF8)

    '    Dim ApiName As String = "MME"
    '    Dim OutputDeviceName As String = "Högtalare (2- Realtek(R) Audio)"
    '    Dim OutputDeviceNames As New List(Of String) ' Used for MME multiple device support
    '    Dim InputDeviceName As String = ""
    '    Dim InputDeviceNames As New List(Of String) ' Used for MME multiple device support
    '    Dim BufferSize As Integer = 2048


    '    For i = 0 To InputLines.Length - 1
    '        LinesRead += 1
    '        Dim Line As String = InputLines(i).Trim

    '        'Skips empty and outcommented lines
    '        If Line = "" Then Continue For
    '        If Line.StartsWith("//") Then Continue For

    '        If Line = "<AudioDevices>" Then
    '            'No need to do anything?
    '            Continue For
    '        End If
    '        If Line = "<New transducer>" Then Exit For

    '        If Line.StartsWith("ApiName") Then ApiName = InputFileSupport.GetInputFileValue(Line, True)
    '        If Line.Replace(" ", "").StartsWith("OutputDevice=") Then OutputDeviceName = InputFileSupport.GetInputFileValue(Line, True)
    '        If Line.Replace(" ", "").StartsWith("OutputDevices=") Then OutputDeviceNames = InputFileSupport.InputFileListOfStringParsing(Line, False, True)
    '        If Line.Replace(" ", "").StartsWith("InputDevice=") Then InputDeviceName = InputFileSupport.GetInputFileValue(Line, True)
    '        If Line.Replace(" ", "").StartsWith("InputDevices=") Then InputDeviceNames = InputFileSupport.InputFileListOfStringParsing(Line, False, True)
    '        If Line.StartsWith("BufferSize") Then BufferSize = InputFileSupport.InputFileIntegerValueParsing(Line, True, AudioSystemSpecificationFilePath)

    '    Next

    '    Dim DeviceLoadSuccess As Boolean = True
    '    If OutputDeviceName = "" And InputDeviceName = "" And OutputDeviceNames Is Nothing And InputDeviceNames Is Nothing Then
    '        'No device names have been specified
    '        DeviceLoadSuccess = False
    '    End If

    '    If ApiName <> "MME" Then
    '        If OutputDeviceNames IsNot Nothing Or InputDeviceNames IsNot Nothing Then
    '            DeviceLoadSuccess = False
    '            MsgBox("When specifying multiple sound (input or output) devices in the file " & AudioSystemSpecificationFilePath & ", the sound API must be MME ( not " & ApiName & ")!", MsgBoxStyle.Exclamation, "Sound device specification error!")
    '        End If
    '    End If

    '    If OutputDeviceNames IsNot Nothing And OutputDeviceName <> "" Then
    '        DeviceLoadSuccess = False
    '        MsgBox("Either (not both) of single or multiple sound output devices must be specified in the file " & AudioSystemSpecificationFilePath & "!", MsgBoxStyle.Exclamation, "Sound device specification error!")
    '    End If

    '    If InputDeviceNames IsNot Nothing And InputDeviceName <> "" Then
    '        DeviceLoadSuccess = False
    '        MsgBox("Either (not both) of single or multiple sound input devices must be specified in the file " & AudioSystemSpecificationFilePath & "!", MsgBoxStyle.Exclamation, "Sound device specification error!")
    '    End If

    '    'Tries to setup the PortAudioApiSettings using the loaded data
    '    Dim AudioApiSettings As New Audio.PortAudioApiSettings
    '    If DeviceLoadSuccess = True Then
    '        If ApiName = "ASIO" Then
    '            DeviceLoadSuccess = AudioApiSettings.SetAsioSoundDevice(OutputDeviceName, BufferSize)
    '        Else
    '            If OutputDeviceNames Is Nothing And InputDeviceNames Is Nothing Then
    '                DeviceLoadSuccess = AudioApiSettings.SetNonAsioSoundDevice(ApiName, OutputDeviceName, InputDeviceName, BufferSize)
    '            Else
    '                DeviceLoadSuccess = AudioApiSettings.SetMmeMultipleDevices(InputDeviceNames, OutputDeviceNames, BufferSize)
    '            End If
    '        End If
    '    End If

    '    If DeviceLoadSuccess = False Then

    '        If OutputDeviceNames Is Nothing Then OutputDeviceNames = New List(Of String)
    '        If InputDeviceNames Is Nothing Then InputDeviceNames = New List(Of String)

    '        MsgBox("The Open Speech Test Framework (OSTF) was unable to load the sound API (" & ApiName & ") and device/s indicated in the file " & AudioSystemSpecificationFilePath & vbCrLf & vbCrLf &
    '            "Output device: " & OutputDeviceName & vbCrLf &
    '            "Output devices: " & String.Join(", ", OutputDeviceNames) & vbCrLf &
    '            "Input device: " & InputDeviceName & vbCrLf &
    '            "Input devices: " & String.Join(", ", InputDeviceNames) & vbCrLf & vbCrLf &
    '            "Click OK to manually select audio input/output devices." & vbCrLf & vbCrLf &
    '            "IMPORTANT: Sound tranducer calibration and/or routing may not be correct when manually selected sound devices are used!", MsgBoxStyle.Exclamation, "OSTF sound device not found!")

    '        'Using default settings, as their is not yet any GUI for selecting settings such as the NET FrameWork AudioSettingsDialog 

    '        'Dim NewAudioSettingsDialog As New AudioSettingsDialog()
    '        'Dim AudioSettingsDialogResult = NewAudioSettingsDialog.ShowDialog()
    '        'If AudioSettingsDialogResult = Windows.Forms.DialogResult.OK Then
    '        '    PortAudioApiSettings = NewAudioSettingsDialog.CurrentAudioApiSettings
    '        'Else
    '        '    MsgBox("You pressed cancel. Default sound settings will be used", MsgBoxStyle.Exclamation, "Select sound device!")
    '        AudioApiSettings.SelectDefaultAudioDevice()
    '        'End If
    '    End If

    '    'Reads the remains of the file
    '    _AvaliableTransducers = New List(Of AudioSystemSpecification)

    '    'Backs up one line
    '    LinesRead = Math.Max(0, LinesRead - 1)

    '    Dim CurrentTransducer As AudioSystemSpecification = Nothing
    '    For i = LinesRead To InputLines.Length - 1
    '        Dim Line As String = InputLines(i).Trim

    '        'Skips empty and outcommented lines
    '        If Line = "" Then Continue For
    '        If Line.StartsWith("//") Then Continue For

    '        If Line = "<New transducer>" Then
    '            If CurrentTransducer Is Nothing Then
    '                'Creates the first transducer
    '                CurrentTransducer = New AudioSystemSpecification(CurrentMediaPlayerType, AudioApiSettings)
    '            Else
    '                'Stores the transducer
    '                _AvaliableTransducers.Add(CurrentTransducer)
    '                'Creates a new one
    '                CurrentTransducer = New AudioSystemSpecification(CurrentMediaPlayerType, AudioApiSettings)
    '            End If
    '        End If

    '        If Line.StartsWith("Name") Then CurrentTransducer.Name = InputFileSupport.GetInputFileValue(Line, True)
    '        If Line.StartsWith("LoudspeakerAzimuths") Then CurrentTransducer.LoudspeakerAzimuths = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
    '        If Line.StartsWith("LoudspeakerElevations") Then CurrentTransducer.LoudspeakerElevations = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
    '        If Line.StartsWith("LoudspeakerDistances") Then CurrentTransducer.LoudspeakerDistances = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
    '        If Line.StartsWith("HardwareOutputChannels") Then CurrentTransducer.HardwareOutputChannels = InputFileSupport.InputFileListOfIntegerParsing(Line, True, AudioSystemSpecificationFilePath)
    '        If Line.StartsWith("CalibrationGain") Then CurrentTransducer.CalibrationGain = InputFileSupport.InputFileListOfDoubleParsing(Line, True, AudioSystemSpecificationFilePath)
    '        If Line.StartsWith("LimiterThreshold") Then CurrentTransducer.LimiterThreshold = InputFileSupport.InputFileDoubleValueParsing(Line, True, AudioSystemSpecificationFilePath)

    '    Next

    '    'Stores the last transducer
    '    If CurrentTransducer IsNot Nothing Then _AvaliableTransducers.Add(CurrentTransducer)

    '    'Adding a default transducer if none were sucessfully read
    '    If _AvaliableTransducers.Count = 0 Then _AvaliableTransducers.Add(New AudioSystemSpecification(CurrentMediaPlayerType, AudioApiSettings))

    '    If StfBase.CurrentMediaPlayerType <> AudioSystemSpecification.MediaPlayerTypes.AudioTrackBased Then
    '        'This has to be made later in STFM for the AudioTrackBased player
    '        For Each Transducer In _AvaliableTransducers
    '            Transducer.SetupMixer()
    '        Next
    '    End If

    '    'Checking calibration gain values and issues warnings if calibration gain is above 30 dB
    '    For Each Transducer In _AvaliableTransducers
    '        For i = 0 To Transducer.CalibrationGain.Count - 1
    '            If Transducer.CalibrationGain(i) > 30 Then
    '                MsgBox("Calibration gain number " & i & " for the audio transducer '" & Transducer.Name & "' exceeds 30 dB. " & vbCrLf & vbCrLf &
    '                       "Make sure that this is really correct before you continue and be cautios not to inflict personal injuries or damage your equipment if you continue! " & vbCrLf & vbCrLf &
    '                       "This calibration value is set in the audio system specifications file: " & AudioSystemSpecificationFilePath, MsgBoxStyle.Exclamation, "Warning - High calibration gain value!")
    '            End If
    '        Next
    '    Next

    'End Sub





End Class



