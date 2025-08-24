' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte





Namespace Utils

    Public Class StringManipulation
        Inherits STFN.Core.Utils.StringManipulation

        ''' <summary>
        ''' Compares two string arrays and returns true only if they have the same lengths and all corresponding items are equal. Items are compared in the given order. 
        ''' </summary>
        ''' <param name="Strings1"></param>
        ''' <param name="Strings2"></param>
        ''' <returns></returns>
        Public Shared Function AllStringsEqual(ByVal Strings1() As String, ByVal Strings2() As String) As Boolean
            If Strings1.Length <> Strings2.Length Then Return False
            For n = 0 To Strings1.Length - 1
                If Strings1(n) <> Strings2(n) Then Return False
            Next
            Return True
        End Function


        Public Shared Sub ExcludeStringArrayMembersDueToLength(ByRef input() As String, ByVal upperStringLengthInclusionLimit As Integer)

            Dim inclusionCount As Integer = 0

            For n = 0 To input.Length - 1

                Dim split = input(n).Split(vbTab)

                input(n) = split(0)

                If split(0).Length <= upperStringLengthInclusionLimit Then inclusionCount += 1
            Next

            Dim output(inclusionCount - 1) As String

            Dim outputCounter As Integer = 0
            For n = 0 To input.Length - 1
                If input(n).Length <= upperStringLengthInclusionLimit Then
                    output(outputCounter) = input(n)
                    outputCounter += 1
                End If
            Next

            input = output

        End Sub


        ''' <summary>
        ''' Removes all exact duplicates of a string in an array of String.
        ''' </summary>
        ''' <param name="inputArray"></param>
        Public Shared Sub RemoveStringArrayDuplicates(ByRef inputArray() As String, Optional ByVal LogActive As Boolean = False, Optional TrimBlankSpaces As Boolean = True, Optional UseToLower As Boolean = True)

            Dim startTime As DateTime = DateTime.Now

            Dim originalArrayLength As Integer = inputArray.Length

            If LogActive = True Then
                Logging.SendInfoToLog("Initializing removal of string array duplicates.")
            End If

            'Putting unique strings in a temporary SortedSet (and selectively applying trimming and/or to lower)
            Dim TempSortedSet As New SortedSet(Of String)
            If UseToLower = True Then
                If TrimBlankSpaces = True Then
                    For inputArrayIndex = 0 To inputArray.Length - 1
                        If Not TempSortedSet.Contains(inputArray(inputArrayIndex).Trim.ToLower) Then
                            TempSortedSet.Add(inputArray(inputArrayIndex).Trim.ToLower)
                        End If
                    Next
                Else
                    For inputArrayIndex = 0 To inputArray.Length - 1
                        If Not TempSortedSet.Contains(inputArray(inputArrayIndex).ToLower) Then
                            TempSortedSet.Add(inputArray(inputArrayIndex).ToLower)
                        End If
                    Next
                End If
            Else
                If TrimBlankSpaces = True Then
                    For inputArrayIndex = 0 To inputArray.Length - 1
                        If Not TempSortedSet.Contains(inputArray(inputArrayIndex).Trim) Then
                            TempSortedSet.Add(inputArray(inputArrayIndex).Trim)
                        End If
                    Next
                Else
                    For inputArrayIndex = 0 To inputArray.Length - 1
                        If Not TempSortedSet.Contains(inputArray(inputArrayIndex)) Then
                            TempSortedSet.Add(inputArray(inputArrayIndex))
                        End If
                    Next
                End If
            End If


            'Putting all unique strings in a string array
            Dim outputArray(TempSortedSet.Count - 1) As String
            Dim outputArrayIndex As Integer = 0
            For Each item As String In TempSortedSet
                outputArray(outputArrayIndex) = item
                outputArrayIndex += 1
            Next

            inputArray = outputArray

            If LogActive = True Then
                Logging.SendInfoToLog("     " & originalArrayLength - outputArray.Length & " duplicate strings were removed. " & outputArray.Length & " strings remain in the array." & " Processing time: " & (DateTime.Now - startTime).TotalSeconds & " seconds.")
            End If

        End Sub


        ''' <summary>
        ''' Removes all duplicates strings in an array of string, as long as the duplicates come in a straight order after each other. 
        ''' </summary>
        ''' <param name="inputArray"></param>
        Public Shared Sub RemoveSortedStringArrayDuplicates(ByRef inputArray() As String)

            Dim originalArrayLength As Long = inputArray.Length

            Dim tempArray(inputArray.Length - 1) As String

            tempArray(0) = inputArray(0)
            Dim tempArrayIndex As Long = 1
            For inputArrayIndex = 1 To inputArray.Length - 1
                If Not tempArray(tempArrayIndex - 1) = (inputArray(inputArrayIndex)) Then
                    tempArray(tempArrayIndex) = inputArray(inputArrayIndex)
                    tempArrayIndex += 1
                End If
            Next
            ReDim Preserve tempArray(tempArrayIndex - 1)

            inputArray = tempArray

            'Log originalArrayLength
            'Log tempArray.Length
            'MsgBox(originalArrayLength & " " & tempArray.Length)

        End Sub

        ''' <summary>
        ''' Returns a string array with all strings in the InputArray that also exist in the ListArray.
        ''' </summary>
        ''' <param name="InputArray"></param>
        ''' <param name="ListArray"></param>
        Public Shared Function RemoveStringsNotInList(ByVal InputArray() As String, ByVal ListArray As String()) As String()

            Dim SortedListArraySet As New SortedSet(Of String)
            For Each Item In ListArray
                If Not SortedListArraySet.Contains(Item) Then
                    SortedListArraySet.Add(Item)
                End If
            Next

            Dim OutputList As New List(Of String)

            For n = 0 To InputArray.Length - 1
                If SortedListArraySet.Contains(InputArray(n)) Then
                    OutputList.Add(InputArray(n))
                End If
            Next

            Return OutputList.ToArray

        End Function

    End Class

End Namespace