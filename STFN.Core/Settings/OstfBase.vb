' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core.Audio
Imports STFN.Core.Messager


Public Module Globals

    ''' <summary>
    ''' The launching app is responsible to create a single instance of StfBase and keep it alive throughout the apps lifetime.
    ''' </summary>
    Public StfBase As StfBase

End Module


Public Class StfBase

    Public Property CurrentPlatForm As Platforms

    Public CurrentMediaPlayerType As AudioSystemSpecification.MediaPlayerTypes

    'Optimization libraries
    Public Property UseOptimizationLibraries As Boolean = False ' This can be used to determine if C++ libraries should be called, such as the libostfdsp, instead of calling equivalent OSTF functions implemented in the managed (.NET) code.


    'Other basic settings
    Public Property AllowDirectionalSimulation As Boolean = True

    Public Property UseExtraWindows As Boolean = True


    ''' <summary>
    ''' Determines if a calibration check control should be displayed in the test settings.
    ''' </summary>
    ''' <returns></returns>
    Public Property ShowCalibrationCheck As Boolean = False

    Public Property LogAllPlayedSoundFiles As Boolean = False

    ' Program location
    Public Property MediaRootDirectory As String = "" '= IO.Path.Combine("C:\", "OSTFMedia") 'Indicates the root path. Other paths given in the project setting files are relative (subpaths) to this path only if they begin with .\ otherwise they are taken as absolute paths.
    Public Property DefaultMediaRootFolderName As String = "OSTFMedia"

    Public Property WasStartedFromVisualStudio As Boolean = False

    Public Property AvailableSpeechMaterialsSubFolder As String = "AvailableSpeechMaterials"
    Public Property CalibrationSignalSubDirectory As String = "CalibrationSignals"
    Public Property AudioSystemSubDirectory As String = "AudioSystem"
    Public Property AudioSystemSettingsFile As String = IO.Path.Combine(AudioSystemSubDirectory, "AudioSystemSpecification.txt")
    Public Property RoomImpulsesSubDirectory As String = "RoomImpulses"
    Public Property AvailableImpulseResponseSetsFile As String = IO.Path.Combine(RoomImpulsesSubDirectory, "AvailableImpulseResponseSets.txt")

    Public Property AvailableSpeechMaterials As New List(Of SpeechMaterialSpecification)

    Public Property AvailableTestsSubDirectory As String = "AvailableTests"

    Protected Property _AvailableTests As List(Of String) = Nothing
    Public ReadOnly Property AvailableTests As List(Of String)
        Get
            If _AvailableTests Is Nothing Then
                'Loading available tests, on first call
                If LoadAvailableTests() = False Then
                    Return Nothing
                End If
            End If

            Return _AvailableTests
        End Get
    End Property


    ''' <summary>
    ''' The SoundPlayer shared between all STFN applications. Each application that uses it, is responsible of initiating it, with the settings required by the specific application. As well as disposing it when the application is closed.
    ''' </summary>
    Public WithEvents SoundPlayer As Audio.SoundPlayers.iSoundPlayer

    Public Function InitializeSoundPlayer(ByRef SoundPlayer As Audio.SoundPlayers.iSoundPlayer) As Boolean
        Globals.StfBase.SoundPlayer = SoundPlayer
        Return SoundPlayerIsInitialized()
    End Function

    Public Function SoundPlayerIsInitialized() As Boolean
        Return _SoundPlayer IsNot Nothing
    End Function


    Protected StfIsInitialized As Boolean = False

    Protected _DirectionalSimulator As DirectionalSimulation = Nothing

    Public ReadOnly Property DirectionalSimulator As DirectionalSimulation
        Get
            If AllowDirectionalSimulation = True Then
                If _DirectionalSimulator Is Nothing Then _DirectionalSimulator = New DirectionalSimulation()
                Return _DirectionalSimulator
            Else
                Return Nothing
            End If
        End Get
    End Property

    ''' <summary>
    ''' This sub needs to be called upon startup of all STF applications. (Any subsequent calls to this Sub will be ignored.)
    ''' </summary>
    ''' <param name="PlatForm"></param>
    ''' <param name="MediaRootDirectory"></param>
    ''' <param name="MediaPlayerType"></param>
    Public Overridable Sub InitializeSTF(ByVal PlatForm As Platforms, ByVal MediaPlayerType As AudioSystemSpecification.MediaPlayerTypes, Optional ByVal MediaRootDirectory As String = "")

        'Exits the sub to avoid multiple calls (which should be avoided, especially to the Audio.PortAudio.Pa_Initialize function).
        If StfIsInitialized = True Then Exit Sub
        StfIsInitialized = True

        'Using optimization libraries if available (currently only on Windows)
        Select Case PlatForm
            Case Platforms.WinUI, Platforms.UWP
                UseOptimizationLibraries = False
            Case Platforms.Android
                UseOptimizationLibraries = True
        End Select

        'Storing the platform instruction and the MediaPlayerType specified by the calling code 
        CurrentPlatForm = PlatForm
        CurrentMediaPlayerType = MediaPlayerType

        Try

            'Determining which media player MediaPlayerType to use
            Select Case CurrentMediaPlayerType
                Case AudioSystemSpecification.MediaPlayerTypes.Default

                    'Selecting default media player MediaPlayerType depending on the current platform
                    Select Case CurrentPlatForm
                        Case Platforms.WinUI, Platforms.UWP
                            'Selects Port Audio based sound player as media player MediaPlayerType
                            CurrentMediaPlayerType = AudioSystemSpecification.MediaPlayerTypes.PaBased
                        Case Platforms.Android
                            'Selects MAUI community toolkit media element as media player MediaPlayerType
                            CurrentMediaPlayerType = AudioSystemSpecification.MediaPlayerTypes.AudioTrackBased
                        Case Platforms.Unknown
                            Throw New Exception("Unable to initialize media player for " & CurrentPlatForm.ToString & " platform MediaPlayerType! The application may not work as intended!")
                        Case Else
                            Throw New Exception("There is no supported STF sound player for the " & CurrentPlatForm.ToString & " platform! The application may not work as intended!")
                    End Select

                Case AudioSystemSpecification.MediaPlayerTypes.PaBased

                    'Checks that the current platform supports MediaPlayerType PaBased
                    Select Case CurrentPlatForm
                        Case Platforms.WinUI, Platforms.UWP
                            'WinUI and UWP should work with the Port Audio based sound player
                        Case Else
                            'Everything else will currently not work
                            Throw New Exception("Unable to initialize media player for " & CurrentPlatForm.ToString & " platform MediaPlayerType! The application may not work as intended!")
                    End Select

                Case AudioSystemSpecification.MediaPlayerTypes.AudioTrackBased

                    'Checks that the current platform supports the selected MediaPlayerType
                    Select Case CurrentPlatForm
                        Case Platforms.Android
                            'Android should work with the AudioTrack based sound player
                        Case Else
                            'Everything else will currently not work
                            Throw New Exception("Unable to initialize media player for " & CurrentPlatForm.ToString & " platform MediaPlayerType! The application may not work as intended!")
                    End Select

            End Select

            'If the code runs to this point, a platform along with a supported MediaPlayerType will have been selected.


            'Sets the MediaRootDirectory 
            Globals.StfBase.MediaRootDirectory = MediaRootDirectory

            'Checks that the folder exists
            If IO.Directory.Exists(MediaRootDirectory) = False Then

                Throw New Exception("Unable to locate the media root directory! Cannot start the application.")

            End If

            'Checks that it seems to be the right (media) folder
            If IO.Directory.Exists(IO.Path.Combine(MediaRootDirectory, AudioSystemSubDirectory)) = False Then
                MsgBox("It seems like you have selected an incorrect OSTF media folder. The OSTF media folder should for example contain the folder " & AudioSystemSubDirectory & vbCrLf &
                                    "Please try again.", MsgBoxStyle.Exclamation, "Unable to find the OSTF media folder!")

                Throw New Exception("Unable to locate the OSTF media folder! Cannot start the application.")

            End If

        Catch ex As Exception
            Throw New Exception("The following error occurred when trying to initialize STF:" & vbCrLf & vbCrLf & ex.ToString)
        End Try

    End Sub

    ''' <summary>
    ''' This sub needs to be called when closing the last STF application.
    ''' </summary>
    Public Overridable Sub TerminateSTF()

        'Disposing the sound player. 
        If SoundPlayerIsInitialized() = True Then SoundPlayer.Dispose()

    End Sub

    ''' <summary>
    ''' Returns the logotype path if a logotype file exists, otherwise returns an empty string.
    ''' </summary>
    ''' <returns></returns>
    Public Function GetLogotypePath() As String

        Dim LogotypePath As String = IO.Path.Combine(MediaRootDirectory, "Logo", "Logotype.jpg")

        If IO.File.Exists(LogotypePath) Then
            Return LogotypePath
        Else
            Return ""
        End If

    End Function

    Public Sub LoadAvailableSpeechMaterialSpecifications()

        Dim TestSpecificationFolder As String = IO.Path.Combine(MediaRootDirectory, AvailableSpeechMaterialsSubFolder)

        'Getting .txt files in that folder
        Dim ExistingFiles = IO.Directory.GetFiles(TestSpecificationFolder)
        Dim TextFileNames As New List(Of String)
        For Each FullFilePath In ExistingFiles
            If FullFilePath.EndsWith(".txt") Then
                'Adds only the file name
                TextFileNames.Add(IO.Path.GetFileName(FullFilePath))
            End If
        Next

        'Clears any tests previously loaded before adding new tests
        Globals.StfBase.AvailableSpeechMaterials.Clear()

        For Each TextFileName In TextFileNames
            'Ties to use the text file in order to create a new test specification object, and just skipps it if unsuccessful
            Dim NewTestSpecification = SpeechMaterialSpecification.LoadSpecificationFile(TextFileName)
            If NewTestSpecification IsNot Nothing Then
                Globals.StfBase.AvailableSpeechMaterials.Add(NewTestSpecification)
            End If
        Next

    End Sub

    Protected Function LoadAvailableTests() As Boolean

        Dim TestSpecificationFile As String = IO.Path.Combine(MediaRootDirectory, AvailableTestsSubDirectory, "AvailableTests.txt")

        If System.IO.File.Exists(TestSpecificationFile) = False Then
            Return False
        End If

        'Clears any tests previously loaded tests before adding new tests
        Globals.StfBase._AvailableTests = New List(Of String)

        Dim Input = IO.File.ReadAllLines(TestSpecificationFile, Text.Encoding.UTF8)

        For line = 0 To Input.Length - 1

            If Input(line).Trim = "" Then Continue For
            If Input(line).Trim.StartsWith("//") Then Continue For

            _AvailableTests.Add(Input(line).Trim)

        Next

        Return True

    End Function


    Protected _AvaliableTransducers As List(Of AudioSystemSpecification) = Nothing
    Public ReadOnly Property AvaliableTransducers As List(Of AudioSystemSpecification)
        Get
            If _AvaliableTransducers Is Nothing Then LoadAudioSystemSpecificationFile()
            Return _AvaliableTransducers
        End Get
    End Property

    Public Sub ClearAudioSpecifications()

        _AvaliableTransducers.Clear()

    End Sub

    Protected Sub LoadAudioSystemSpecificationFile()

        Dim AudioSystemSpecificationFilePath = Utils.GeneralIO.NormalizeCrossPlatformPath(IO.Path.Combine(Globals.StfBase.MediaRootDirectory, Globals.StfBase.AudioSystemSettingsFile))
        ''Getting calibration file descriptions from the text file
        Dim InputLines() As String = System.IO.File.ReadAllLines(AudioSystemSpecificationFilePath, System.Text.Encoding.UTF8)

        LoadAudioSystemSpecifications(InputLines, AudioSystemSpecificationFilePath)

    End Sub


    ''' <summary>
    ''' Loading the audio systems specifications.
    ''' </summary>
    ''' <param name="InputLines">An array of (text) lines in the audio system specification (.txt) file. </param>
    ''' <param name="AudioSystemSpecificationFilePath">The file path to the audio systems specifications file (used only for error reporting)</param>
    Public Overridable Sub LoadAudioSystemSpecifications(ByVal InputLines() As String, ByVal AudioSystemSpecificationFilePath As String)


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

                Dim AudioSettings As Audio.AudioSettings = Nothing

                If Global.STFN.Core.Globals.StfBase.CurrentMediaPlayerType = Global.STFN.Core.AudioSystemSpecification.MediaPlayerTypes.AudioTrackBased Then

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
                    AudioSettings = New Audio.AndroidAudioTrackPlayerSettings
                    AudioSettings.AllowDefaultOutputDevice = AllowDefaultOutputDevice
                    AudioSettings.AllowDefaultInputDevice = AllowDefaultInputDevice
                    DirectCast(AudioSettings, AndroidAudioTrackPlayerSettings).SelectedOutputDeviceName = OutputDeviceName
                    DirectCast(AudioSettings, AndroidAudioTrackPlayerSettings).SelectedInputDeviceName = InputDeviceName
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
                            CurrentTransducer = New AudioSystemSpecification(Globals.StfBase.CurrentMediaPlayerType, AudioSettings)
                        Else
                            'Stores the transducer
                            _AvaliableTransducers.Add(CurrentTransducer)
                            'Creates a new one
                            CurrentTransducer = New AudioSystemSpecification(Globals.StfBase.CurrentMediaPlayerType, AudioSettings)
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
                If _AvaliableTransducers.Count = 0 Then _AvaliableTransducers.Add(New AudioSystemSpecification(Globals.StfBase.CurrentMediaPlayerType, AudioSettings))

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

    Public Function GetAudioSystemSpecificationFieldsDescriptions() As List(Of String)

        Dim Output As New List(Of String)

        Output.Add("<New media player>, a tag to define a new media player" + vbCrLf)
        Output.Add("MediaPlayerType, Use PaBased on Windows or AudioTrackBased on Android devices. (PaBased = PortAudio based STF sound player, AudioTrackBased = Android AudioTrack based sound player)" + vbCrLf)
        Output.Add("ApiName, Only relevant on Windows and is ignored on Android. Values can be MME, Windows WASAPI, ASIO, Windows DirectSound, etc." + vbCrLf)
        Output.Add("OutputDevices, audio output devices to use if more than one. OutputDevices is only relevant with the MME API on Windows. For Android, use OutputDevice instead" + vbCrLf)
        Output.Add("OutputDevice, the audio output device to use" + vbCrLf)
        Output.Add("InputDevices, audio input devices to use if more than one. InputDevices is only relevant with the MME API on Windows. For Android, use InputDevice instead" + vbCrLf)
        Output.Add("InputDevice, the audio input device to use" + vbCrLf)
        Output.Add("Please note: On windows output and input devices should be defined using their name (use the 'View devices' button to get the right one). On andorid devices, ouput and input devices should be defined using the [ProductName]+[DeviceType] (e.g. 'USB-Audio - AudioQuest DragonFly Black v1.5+UsbHeadset')" + vbCrLf)
        Output.Add("BufferSize, the size of the audiobuffer in each package to the sound device. Must be powers of 2. Only relevant on Windows, ignored on android devices." + vbCrLf)
        Output.Add("AllowDefaultOutputDevice, If True, STF is allowed to select the default audio output device/s. If set to False, STF will close the application if the the intended audio output device/s is not present on the system." + vbCrLf)
        Output.Add("AllowDefaultInputDevice, If True, STF is allowed to select the default audio input device/s. If set to False, STF will close the application if the the intended audio output device/s is not present on the system." + vbCrLf + vbCrLf)

        Output.Add("<New transducer>, a tag to define a new transducer (there can be several transducers under each media player)" + vbCrLf)
        Output.Add("Name, the name of the transducer as shown in the software" + vbCrLf)
        Output.Add("LoudspeakerAzimuths, a comma delimited vector indicating the actual physical azimuth angle between each loudspeaker and the frontal angle at the listener's position (in degrees)" + vbCrLf)
        Output.Add("LoudspeakerElevations, a comma delimited vector indicating the actual physical elevation angle between each loudspeaker and the horizon at the listener's position (in degrees)" + vbCrLf)
        Output.Add("LoudspeakerDistances, a comma delimited vector indicating the actual physical distance from the loudspeakers to the listener (in meters)" + vbCrLf)
        Output.Add("HardwareOutputChannels, a comma delimited vector with hardware output channels to use (as enumerated by the sound device, often 1,2,3 etc)" + vbCrLf)
        Output.Add("CalibrationGain, a comma delimited vector holding the calibration gain applied to the speaker connected to each of the indicated HardwareOutputChannels during playback (so that a signal of 0 dBFS = 100 dBPSL). Note that values should be comma separated, and dots (.) should be used as decimal mark." + vbCrLf)
        Output.Add("HostVolumeOutputLevel, The host volume level (for the selected API) in percentages (0-100). If possible, this value will be used to set and maintain the volume of the selected output sound unit during speech tests. Currently only supported on Android devices." + vbCrLf)
        Output.Add("PtaCalibrationGainFrequencies, a comma delimited vector of frequencies (in Hz) for which calibration gain values are given by PtaCalibrationGainValues" + vbCrLf)
        Output.Add("PtaCalibrationGainValues, a comma delimited vector of pure tone calibration values for the specified transducer. These are the values to modify during calibration of pure tone audiometry stimuli. Please note that this calibration depends on the CalibrationGain values for each channel, so that if CalibrationGain is changed, PtaCalibrationGainValues also has to be changed (but not the other way round)" + vbCrLf)
        Output.Add("RETSPL_Speech, the RETSPL value assumed when displaying speech audiometry levels in dB HL." + vbCrLf)
        Output.Add("LimiterThreshold, a limiter threshold (in dB SPL) that can be set to limit the output level in each channel of the transducer." + vbCrLf + vbCrLf)

        Output.Add("Note that you can enter comments in the audio system specifications using double slashes //")
        Output.Add("For example:")
        Output.Add("LoudspeakerAzimuths = 0, 180 // This is my front-back loudspeaker setup")

        Return Output

    End Function





End Class


