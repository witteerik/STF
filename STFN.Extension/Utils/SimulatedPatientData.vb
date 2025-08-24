' License
' This project Is available free for noncommercial use under the PolyForm Noncommercial License 1.0.0.
' Commercial use requires a separate commercial license.See the `LICENSE` file for details.
' 
' Copyright (c) 2025 Erik Witte



Namespace Utils

    Public Structure SimulatedPatientData

        Property TrueInnerEarThresholds_Right As SortedList(Of Integer, Double)
        Property TrueInnerEarThresholds_Left As SortedList(Of Integer, Double)
        Property ConductiveComponents_Right As SortedList(Of Integer, Double)
        Property ConductiveComponents_Left As SortedList(Of Integer, Double)
        Property CollapsingEarCanalDegree_Right As CollapsingEarCanalDegrees
        Property CollapsingEarCanalDegree_Left As CollapsingEarCanalDegrees
        Property LowestResponseTime As Double
        Property HighestResponseTime As Double
        Property Default_FalsePositiveRate As Double
        Property Default_FalseNegativRate As Double
        Property Right_FalsePositiveRate As SortedList(Of Integer, Double)
        Property Right_FalseNegativeRate As SortedList(Of Integer, Double)
        Property Left_FalsePositiveRate As SortedList(Of Integer, Double)
        Property Left_FalseNegativeRate As SortedList(Of Integer, Double)

    End Structure

    Public Enum CollapsingEarCanalDegrees
        NotCollapsing
        SlightlyCollapsing
        PartiallyCollapsing
        CompletelyCollapsed
    End Enum

End Namespace