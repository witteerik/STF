﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports System.IO
Imports System.Threading

Namespace Audio
    Public Module AudioLog

        Public ShowAudioErrors As Boolean = False
        Public LogAudioErrors As Boolean = False
        Public AudioLogIsInMultiThreadApplication As Boolean = True
        Public AudioLogSpinLock As New Threading.SpinLock

        Public Sub SendInfoToAudioLog(ByVal message As String,
                         Optional ByVal LogFileNameWithoutExtension As String = "",
                         Optional LogFileTemporaryPath As String = "",
                         Optional ByVal OmitDateInsideLog As Boolean = False,
                         Optional ByVal OmitDateInFileName As Boolean = False)

            Dim SpinLockTaken As Boolean = False

            Try

                'Getting the application path
                Dim logFilePath As String = IO.Path.Combine(IO.Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).FullName, "AudioLog")

                'Attempts to enter a spin lock to avoid multiple thread conflicts when saving to the same file
                AudioLogSpinLock.Enter(SpinLockTaken)

                If LogFileTemporaryPath = "" Then LogFileTemporaryPath = Logging.LogFileDirectory

                Dim FileNameToUse As String = ""

                If OmitDateInFileName = False Then
                    If LogFileNameWithoutExtension = "" Then
                        FileNameToUse = "log-" & DateTime.Now.ToShortDateString.Replace("/", "-") & ".txt"
                    Else
                        FileNameToUse = LogFileNameWithoutExtension & "-" & DateTime.Now.ToShortDateString.Replace("/", "-") & ".txt"
                    End If
                Else
                    If LogFileNameWithoutExtension = "" Then
                        FileNameToUse = "log.txt"
                    Else
                        FileNameToUse = LogFileNameWithoutExtension & ".txt"
                    End If

                End If

                Dim OutputFilePath As String = Path.Combine(LogFileTemporaryPath, FileNameToUse)

                'Adds a thread ID if in multi thread app
                If AudioLogIsInMultiThreadApplication = True Then
                    Dim TreadName As String = Thread.CurrentThread.ManagedThreadId
                    OutputFilePath &= "ThreadID_" & TreadName
                End If

                Try
                    If Not Directory.Exists(LogFileTemporaryPath) Then Directory.CreateDirectory(LogFileTemporaryPath)
                    Dim samplewriter As New StreamWriter(OutputFilePath, FileMode.Append)
                    If OmitDateInsideLog = False Then
                        samplewriter.WriteLine(DateTime.Now.ToString & vbCrLf & message)
                    Else
                        samplewriter.WriteLine(message)
                    End If
                    samplewriter.Close()

                Catch ex As Exception
                    'Just ignores any errors here...
                End Try
            Finally

                'Releases any spinlock
                If SpinLockTaken = True Then AudioLogSpinLock.Exit()
            End Try

        End Sub

        Public Sub AudioError(ByVal errorText As String, Optional ByVal errorTitle As String = "Error")

            If ShowAudioErrors = True Then
                MsgBox(errorText, MsgBoxStyle.Critical, errorTitle)
            End If

            If LogAudioErrors = True Then
                SendInfoToAudioLog("The following error occurred: " & vbCrLf & errorTitle & errorText, "Errors")
            End If

        End Sub

    End Module

End Namespace