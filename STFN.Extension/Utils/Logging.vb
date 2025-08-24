' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte



Imports System.IO


Public Class Logging
    Inherits STFN.Core.Logging

    ''' <summary>
    ''' Copies files at the indicated file paths to the indicated Output folder. 
    ''' </summary>
    ''' <param name="OutputFolder"></param>
    ''' <param name="NewFilesToCopy">A list containing the file paths to copy. If left to nothing the files stored by using AddFileToCopy will be used instead.</param>
    '''<return>Returns True is all files were sucessfully copied, or if no output files existed, and False if one or more files were failed to be copied.</return>
    Public Function CopyFilesToFolder(ByRef OutputFolder As String, Optional ByRef NewFilesToCopy As SortedSet(Of String) = Nothing) As Boolean

        If NewFilesToCopy Is Nothing Then NewFilesToCopy = FilesToCopy
        If NewFilesToCopy Is Nothing Then Return True

        Dim FailedFiles As Integer = 0

        For Each CurrentFile In NewFilesToCopy

            Try
                'Makes sure that the output folder exists
                If Not Directory.Exists(OutputFolder) Then Directory.CreateDirectory(OutputFolder)

                Dim OutputFilePath As String = Path.Combine(OutputFolder, Path.GetFileName(CurrentFile))

                'Ensures that an old file with the same filename is not overwritten by adding a number to existing files
                OutputFilePath = STFN.Core.Utils.GeneralIO.CheckFileNameConflict(OutputFilePath)

                'Copies the file
                File.Copy(CurrentFile, OutputFilePath)

            Catch ex As Exception
                SendInfoToLog("Could not copy the file " & CurrentFile & " to the folder: " & OutputFolder & vbCrLf & ex.ToString)
                FailedFiles += 1
            End Try
        Next

        If FailedFiles = 0 Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Sub SetLogFolder()

        Dim SuggestedLogFolder As String = LogFileDirectory

0:

        Dim SelectedFolder = STFN.Core.Messager.GetFolder(SuggestedLogFolder, "Select the log / export folder you want to use!").Result
        If SelectedFolder = "" Then
            MsgBox("No folder selected! Click OK to try again!", MsgBoxStyle.Exclamation, "Unable to set the log folder")
            GoTo 0
        End If

        'Checking that the folder is valid
        If IO.Directory.Exists(SelectedFolder) Then
            SuggestedLogFolder = SelectedFolder
            MsgBox("Log / export folder was successfully set to: " & vbCrLf & vbCrLf & SuggestedLogFolder, MsgBoxStyle.Information, "STF")
        Else
            MsgBox("Could not locate the selected folder! Click OK to try again!", MsgBoxStyle.Exclamation, "Unable to set the log folder")
            GoTo 0
        End If

        'Actually storing the log folder
        LogFileDirectory = SuggestedLogFolder

    End Sub

End Class

