Public Class PerformanceAnalysis

    Public Function DoIT(ByVal TotalCounts As Integer) As Integer

        Return ((TotalCounts / (TotalCounts + 10)) * 10)

    End Function

End Class
