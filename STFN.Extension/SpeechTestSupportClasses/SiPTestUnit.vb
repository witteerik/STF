' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte

Imports STFN.Core.SipTest

Public Class SiPTestUnit
    Inherits STFN.Core.SipTest.SiPTestUnit

    Public Sub New(ByRef ParentMeasurement As SipMeasurement, Optional Description As String = "")
        MyBase.New(ParentMeasurement, Description)
    End Sub
End Class
