' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Namespace Utils

    Public Class StringManipulation

        Public Shared Function SplitStringByLines(input As String) As String()

            Return input.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)

            'As an alternative using regex, this expression should match any combination of \r and \n
            'Return Regex.Split(input, "\r\n|\n\r|\r|\n")
        End Function

    End Class

End Namespace