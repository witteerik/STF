﻿' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports System.Runtime.InteropServices

Public Class LibOstfDsp_VB

    <DllImport("libostfdsp_x64.dll", EntryPoint:="copyToDouble", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub copyToDouble_win_64(sourceArray As Single(), size As Integer, targetArray As Double())
    End Sub
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="copyToDouble", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub copyToDouble_win_32(sourceArray As Single(), size As Integer, targetArray As Double())
    End Sub

    <DllImport("libostfdspandroid", EntryPoint:="copyToDouble", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub copyToDouble_android(sourceArray As Single(), size As Integer, targetArray As Double())
    End Sub

    Public Shared Sub CopyToDouble(SourceArray As Single(), TargetArray As Double())

        Dim Size As Integer = SourceArray.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI
                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    copyToDouble_win_32(SourceArray, Size, TargetArray)
                Else
                    copyToDouble_win_64(SourceArray, Size, TargetArray)
                End If

            Case Platforms.Android
                copyToDouble_android(SourceArray, Size, TargetArray)
            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")

        End Select

    End Sub

    <DllImport("libostfdsp_x64.dll", EntryPoint:="copyToFloat", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub copyToFloat_win_64(sourceArray As Double(), size As Integer, targetArray As Single())
    End Sub
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="copyToFloat", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub copyToFloat_win_32(sourceArray As Double(), size As Integer, targetArray As Single())
    End Sub

    <DllImport("libostfdspandroid", EntryPoint:="copyToFloat", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub copyToFloat_android(sourceArray As Double(), size As Integer, targetArray As Single())
    End Sub

    Public Shared Sub CopyToSingle(SourceArray As Double(), TargetArray As Single())

        Dim Size As Integer = SourceArray.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI
                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    copyToFloat_win_32(SourceArray, Size, TargetArray)
                Else
                    copyToFloat_win_64(SourceArray, Size, TargetArray)
                End If

            Case Platforms.Android
                copyToFloat_android(SourceArray, Size, TargetArray)
            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")

        End Select

    End Sub


    <DllImport("libostfdsp_x64.dll", EntryPoint:="multiplyDoubleArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyDoubleArray_64(values As Double(), size As Integer, factor As Double) As Integer
    End Function
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="multiplyDoubleArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyDoubleArray_32(values As Double(), size As Integer, factor As Double) As Integer
    End Function

    <DllImport("libostfdspandroid", EntryPoint:="multiplyDoubleArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyDoubleArray_android(values As Double(), size As Integer, factor As Double) As Integer
    End Function

    ''' <summary>
    ''' Multiplies each value in Array by Factor.
    ''' </summary>
    ''' <param name="Array"></param>
    ''' <param name="Factor"></param>
    ''' <returns>Returns the number of samples that were trimmed due to exceeding the Singel range.</returns>
    Public Shared Function MultiplyArray(Array As Double(), Factor As Double) As Integer

        Dim Size As Integer = Array.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    Return multiplyDoubleArray_32(Array, Size, Factor)
                Else
                    Return multiplyDoubleArray_64(Array, Size, Factor)
                End If

            Case Platforms.Android

                Return multiplyDoubleArray_android(Array, Size, Factor)

            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")

        End Select

    End Function


    <DllImport("libostfdsp_x64.dll", EntryPoint:="multiplyFloatArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyFloatArray_64(values As Single(), size As Integer, factor As Single) As Integer
    End Function
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="multiplyFloatArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyFloatArray_32(values As Single(), size As Integer, factor As Single) As Integer
    End Function

    <DllImport("libostfdspandroid", EntryPoint:="multiplyFloatArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyFloatArray_android(values As Single(), size As Integer, factor As Single) As Integer
    End Function

    ''' <summary>
    ''' Multiplies each value in Array by Factor.
    ''' </summary>
    ''' <param name="Array"></param>
    ''' <param name="Factor"></param>
    ''' <returns>Returns the number of samples that were trimmed due to exceeding the Singel range.</returns>
    Public Shared Function MultiplyArray(Array As Single(), Factor As Single) As Integer

        Dim Size As Integer = Array.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    Return multiplyFloatArray_32(Array, Size, Factor)
                Else
                    Return multiplyFloatArray_64(Array, Size, Factor)
                End If

            Case Platforms.Android

                Return multiplyFloatArray_android(Array, Size, Factor)

            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")

        End Select

    End Function

    <DllImport("libostfdsp_x64.dll", EntryPoint:="multiplyFloatArraySection", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyFloatArraySection_64(values As Single(), arraySize As Integer, factor As Single, startIndex As Integer, length As Integer) As Integer
    End Function
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="multiplyFloatArraySection", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyFloatArraySection_32(values As Single(), arraySize As Integer, factor As Single, startIndex As Integer, length As Integer) As Integer
    End Function

    <DllImport("libostfdspandroid", EntryPoint:="multiplyFloatArraySection", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function multiplyFloatArraySection_android(values As Single(), arraySize As Integer, factor As Single, startIndex As Integer, length As Integer) As Integer
    End Function

    ''' <summary>
    ''' Multiplies each value in a specified section of Array by Factor.
    ''' </summary>
    ''' <param name="Array"></param>
    ''' <param name="Factor"></param>
    ''' <param name="StartIndex"></param>
    ''' <param name="Length"></param>
    ''' <returns>Returns the number of samples that were trimmed due to exceeding the Singel range.</returns>
    Public Shared Function MultiplyArraySection(Array As Single(), Factor As Single, StartIndex As Integer, Length As Integer) As Integer

        Dim ArraySize As Integer = Array.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    Return multiplyFloatArraySection_32(Array, ArraySize, Factor, StartIndex, Length)
                Else
                    Return multiplyFloatArraySection_64(Array, ArraySize, Factor, StartIndex, Length)
                End If

            Case Platforms.Android

                Return multiplyFloatArraySection_android(Array, ArraySize, Factor, StartIndex, Length)

            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")
        End Select


    End Function


    <DllImport("libostfdsp_x64.dll", EntryPoint:="calculateFloatSumOfSquare", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function calculateFloatSumOfSquare_64(values As Single(), arraySize As Integer, startIndex As Integer, sectionLength As Integer) As Double
    End Function
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="calculateFloatSumOfSquare", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function calculateFloatSumOfSquare_32(values As Single(), arraySize As Integer, startIndex As Integer, sectionLength As Integer) As Double
    End Function

    <DllImport("libostfdspandroid", EntryPoint:="calculateFloatSumOfSquare", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Function calculateFloatSumOfSquare_android(values As Single(), arraySize As Integer, startIndex As Integer, sectionLength As Integer) As Double
    End Function

    ''' <summary>
    ''' Calculates the sum-of-square for the indicated section of the Values array
    ''' </summary>
    ''' <param name="Values"></param>
    ''' <param name="StartIndex"></param>
    ''' <param name="SectionLength"></param>
    ''' <returns></returns>
    Public Shared Function CalculateSumOfSquare(Values As Single(), StartIndex As Integer, SectionLength As Integer) As Double

        Dim ArraySize As Integer = Values.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    Return calculateFloatSumOfSquare_32(Values, ArraySize, StartIndex, SectionLength)
                Else
                    Return calculateFloatSumOfSquare_64(Values, ArraySize, StartIndex, SectionLength)
                End If

            Case Platforms.Android

                Return calculateFloatSumOfSquare_android(Values, ArraySize, StartIndex, SectionLength)

            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")
        End Select

    End Function


    <DllImport("libostfdsp_x64.dll", EntryPoint:="addTwoFloatArrays", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub addTwoFloatArrays_64(array1 As Single(), array2 As Single(), size As Integer)
    End Sub

    <DllImport("libostfdsp_Win32.dll", EntryPoint:="addTwoFloatArrays", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub addTwoFloatArrays_32(array1 As Single(), array2 As Single(), size As Integer)
    End Sub

    <DllImport("libostfdspandroid", EntryPoint:="addTwoFloatArrays", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub addTwoFloatArrays_android(array1 As Single(), array2 As Single(), size As Integer)
    End Sub

    ''' <summary>
    ''' Adds the values of Array2 to the corresponding indices of Array1. Equal Array lengths are required.
    ''' </summary>
    ''' <param name="Array1">The first input/output data array. Upon return this corresponding data array contains the sum of the values in Array1 and Array2.</param>
    ''' <param name="Array2">The the input data array containing the values which should be added to Array1. </param>
    Public Shared Sub AddTwoFloatArrays(Array1 As Single(), Array2 As Single())

        If Array1.Length <> Array2.Length Then Throw New ArgumentException("Input arrays must have the same lengths.")

        Dim size As Integer = Array1.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    addTwoFloatArrays_32(Array1, Array2, size)
                Else
                    addTwoFloatArrays_64(Array1, Array2, size)
                End If

            Case Platforms.Android

                addTwoFloatArrays_android(Array1, Array2, size)

            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")
        End Select

    End Sub


    <DllImport("libostfdsp_x64.dll", EntryPoint:="fft_complex", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub fft_complex_64_win(real As Double(), imag As Double(), size As Integer, Optional direction As Integer = 1, Optional reorder As Boolean = True, Optional scaleForwardTransform As Boolean = True)
    End Sub

    <DllImport("libostfdsp_Win32.dll", EntryPoint:="fft_complex", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub fft_complex_32_win(real As Double(), imag As Double(), size As Integer, Optional direction As Integer = 1, Optional reorder As Boolean = True, Optional scaleForwardTransform As Boolean = True)
    End Sub

    <DllImport("libostfdspandroid", EntryPoint:="fft_complex", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub fft_complex_android(real As Double(), imag As Double(), size As Integer, pccos As Double(), pcsin As Double(), Optional direction As Integer = 1, Optional reorder As Boolean = True, Optional scaleForwardTransform As Boolean = True)
    End Sub


    Public Shared Sub Fft_complex(real As Double(), imag As Double(), size As Integer, Optional Direction As DSP.FftDirections = DSP.FftDirections.Forward, Optional reorder As Boolean = True, Optional scaleForwardTransform As Boolean = True)

        'Translating Direction to the corresponding integer Array used in the optimization libraries
        Dim dir As Integer
        Select Case Direction
            Case DSP.FftDirections.Forward
                dir = 1
            Case DSP.FftDirections.Backward
                dir = -1
            Case Else
                Throw New ArgumentException("Unknown value for Direction!")
        End Select

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    fft_complex_32_win(real, imag, size, dir, reorder, scaleForwardTransform)
                Else
                    fft_complex_64_win(real, imag, size, dir, reorder, scaleForwardTransform)
                End If

            Case Platforms.Android

                'Ensures that x and y have the same Length
                If real.Length <> imag.Length Then
                    Throw New ArgumentException("The x and y arrays need to have the same Length in FastFourierTransform!")
                End If

                Dim TrigonArrays = DSP.Radix2TrigonometricLookup.GetArrays(size, Direction) ' Note that the Enum (not the translated value for Direction is used here)

                fft_complex_android(real, imag, size, TrigonArrays.Item1, TrigonArrays.Item2, dir, reorder, scaleForwardTransform)

            Case Else
                Throw New NotImplementedException("Optimization library FFT is not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")
        End Select

    End Sub

    <DllImport("libostfdsp_x64.dll", EntryPoint:="createInterleavedArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub createInterleavedArray_win_64(concatenatedArrays As Single(), channelCount As Integer, channelLength As Integer, interleavedArray As Single(), applyGain As Boolean, channelGainFactor As Single())
    End Sub
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="createInterleavedArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub createInterleavedArray_win_32(concatenatedArrays As Single(), channelCount As Integer, channelLength As Integer, interleavedArray As Single(), applyGain As Boolean, channelGainFactor As Single())
    End Sub

    <DllImport("libostfdspandroid", EntryPoint:="createInterleavedArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub createInterleavedArray_android(concatenatedArrays As Single(), channelCount As Integer, channelLength As Integer, interleavedArray As Single(), applyGain As Boolean, channelGainFactor As Single())
    End Sub

    Public Shared Sub CreateInterleavedArray(concatenatedArrays As Single(), channelCount As Integer, channelLength As Integer, interleavedArray As Single(), applyGain As Boolean, channelGainFactor As Single())

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    createInterleavedArray_win_32(concatenatedArrays, channelCount, channelLength, interleavedArray, applyGain, channelGainFactor)
                Else
                    createInterleavedArray_win_64(concatenatedArrays, channelCount, channelLength, interleavedArray, applyGain, channelGainFactor)
                End If

            Case Platforms.Android

                createInterleavedArray_android(concatenatedArrays, channelCount, channelLength, interleavedArray, applyGain, channelGainFactor)

            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")
        End Select

    End Sub

    <DllImport("libostfdsp_x64.dll", EntryPoint:="deinterleaveArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub deinterleaveArray_win_64(interleavedArray As Single(), channelCount As Integer, channelLength As Integer, concatenatedArrays As Single())
    End Sub
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="deinterleaveArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub deinterleaveArray_win_32(interleavedArray As Single(), channelCount As Integer, channelLength As Integer, concatenatedArrays As Single())
    End Sub


    <DllImport("libostfdspandroid", EntryPoint:="deinterleaveArray", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub deinterleaveArray_android(interleavedArray As Single(), channelCount As Integer, channelLength As Integer, concatenatedArrays As Single())
    End Sub

    Public Shared Sub DeinterleaveSoundArray(interleavedArray As Single(), channelCount As Integer, channelLength As Integer, concatenatedArrays As Single())

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI

                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    deinterleaveArray_win_32(interleavedArray, channelCount, channelLength, concatenatedArrays)
                Else
                    deinterleaveArray_win_64(interleavedArray, channelCount, channelLength, concatenatedArrays)
                End If

            Case Platforms.Android
                Try
                    deinterleaveArray_android(interleavedArray, channelCount, channelLength, concatenatedArrays)
                Catch ex As Exception
                    Messager.MsgBox(ex.ToString)
                End Try

            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")
        End Select

    End Sub


    <DllImport("libostfdsp_x64.dll", EntryPoint:="complexMultiplication", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub complexMultiplication_win_64(Real1 As Double(), Imag1 As Double(), Real2 As Double(), Imag2 As Double(), size As Integer)
    End Sub
    <DllImport("libostfdsp_Win32.dll", EntryPoint:="complexMultiplication", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub complexMultiplication_win_32(Real1 As Double(), Imag1 As Double(), Real2 As Double(), Imag2 As Double(), size As Integer)
    End Sub

    <DllImport("libostfdspandroid", EntryPoint:="complexMultiplication", CallingConvention:=CallingConvention.Cdecl)>
    Private Shared Sub complexMultiplication_android(Real1 As Double(), Imag1 As Double(), Real2 As Double(), Imag2 As Double(), size As Integer)
    End Sub

    Public Shared Sub ComplexMultiplication(Real1 As Double(), Imag1 As Double(), Real2 As Double(), Imag2 As Double())

        Dim Size As Integer = Real1.Length

        Select Case Globals.StfBase.CurrentPlatForm
            Case Platforms.WinUI
                'Checking whether a 32-bit or 64-bit environment is running
                If IntPtr.Size = 4 Then
                    complexMultiplication_win_32(Real1, Imag1, Real2, Imag2, Size)
                Else
                    complexMultiplication_win_64(Real1, Imag1, Real2, Imag2, Size)
                End If

            Case Platforms.Android
                complexMultiplication_android(Real1, Imag1, Real2, Imag2, Size)
            Case Else
                Throw New NotImplementedException("Optimization library not implemented for the " & Globals.StfBase.CurrentPlatForm.ToString & " platform.")

        End Select

    End Sub



End Class