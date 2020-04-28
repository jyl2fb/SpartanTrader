Module Recommender
    Public Sub ResetAllRecommendations()
        Dim ticker As String
        For i = 0 To 11
            ticker = myDataSet.Tabes("TickersTbl").Rows(i)("Ticker")
            Recommendations(i) = New Transaction
            Recommendations(i).familyTicker = ticker.Trim()
            Globals.dashboard.Range("I4").Offset(i, 0).Value = ticker
            Globals.Dashboard.Range("J4").Offset(i, 0).Value = ticker
        Next

        CandidateRecList = New List(Of Transaction)
    End Sub

    Public Sub CalcAllRecommendations(targetDate As Date)

    End Sub

End Module
