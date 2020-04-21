
Public Class Dashboard

    Private Sub Sheet2_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet2_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub FillCBoxes()
        TickersCBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            TickersCBox.Items.Add(myRow("Ticker").ToString.Trim())
        Next
        TickersCBox.Text = "Select Ticker"
        SymbolsCBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            SymbolsCBox.Items.Add(myRow("Symbol").ToString.Trim())
        Next
        SymbolsCBox.Text = "Select Symbol"
    End Sub

    Private Sub BuyStockBtn_Click(sender As Object, e As EventArgs) Handles BuyStockBtn.Click
        CT.Clear()
        CT.type = "Buy"
        If IsStockInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub SellStockBtn_Click(sender As Object, e As EventArgs) Handles SellStockBtn.Click
        CT.Clear()
        CT.type = "Sell"
        If IsStockInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub SellShortBtn_Click(sender As Object, e As EventArgs) Handles SellShortBtn.Click
        CT.Clear()
        CT.type = "SellShort"
        If IsStockInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub CashDivBtn_Click(sender As Object, e As EventArgs) Handles CashDivBtn.Click
        CT.Clear()
        CT.type = "CashDiv"
        If IsStockInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub ExecuteStockTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecuteStockTransactionBtn.Click
        If IsStockInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            If IsValid(CT) Then
                Execute(CT)
                CT.Highlight()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                CalcAllRecommendations(currentDate)
                DisplayAllRecommendations()
            End If
        End If
    End Sub

    Private Sub ExecuteOptionTransactionBtn_Click(sender As Object, e As EventArgs) Handles ExecuteOptionTransactionBtn.Click
        If IsOptionInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            If IsValid(CT) Then
                Execute(CT)
                CT.Highlight()
                CalcFinancialMetrics(currentDate)
                DisplayFinancialMetrics(currentDate)
                CalcAllRecommendations(currentDate)
                DisplayAllRecommendations()
            End If
        End If
    End Sub

    Private Sub BuyOptionBtn_Click(sender As Object, e As EventArgs) Handles BuyOptionBtn.Click
        CT.Clear()
        CT.type = "Buy"
        If IsOptionInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub SellOptionBtn_Click(sender As Object, e As EventArgs) Handles SellOptionBtn.Click
        CT.Clear()
        CT.type = "Sell"
        If IsOptionInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub SellShortOptionBtn_Click(sender As Object, e As EventArgs) Handles SellShortOptionBtn.Click
        CT.Clear()
        CT.type = "SellShort"
        If IsOptionInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub ExerciseOptionBtn_Click(sender As Object, e As EventArgs) Handles ExerciseOptionBtn.Click
        CT.Clear()
        If SymbolsCBox.SelectedItem <> Nothing Then
            If IsACall(SymbolsCBox.SelectedItem) Then
                CT.type = "X-Call"
            Else
                If IsAPut(SymbolsCBox.SelectedItem) Then
                    CT.type = "X-Put"
                End If
            End If
        End If
        If IsOptionInputValid() = True Then
            CT.CalcTransactionProperties(currentDate)
            CT.Show()
        End If
    End Sub

    Private Sub ManualExecutionLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ManualExecutionLBox.SelectedIndexChanged
        Dim selectedTrade As Integer
        selectedTrade = ManualExecutionLBox.SelectedIndex
        If selectedTrade < 0 Or selectedTrade > 11 Then
            Exit Sub
        End If
        If IsValid(Recommendations(selectedTrade)) Then
            Execute(Recommendations(selectedTrade))
            CalcFinancialMetrics(currentDate)
            DisplayFinancialMetrics(currentDate)
            CalcAllRecommendations(currentDate)
            DisplayAllRecommendations()
        End If
    End Sub

    Public Sub DisplayAllRecommendations()
        For i = 0 To 11
            DisplayRecommendation(i)
        Next
        Globals.Dashboard.ManualExecutionLBox.ClearSelected()
    End Sub

    Public Sub DisplayRecommendation(i As Integer)
        For i = 0 To 11
            Globals.Dashboard.Range("K4").Offset(i, 0).Value = Recommendations(i).familyDelta
            Globals.Dashboard.Range("L4").Offset(i, 0).Value = Recommendations(i).type
            Globals.Dashboard.Range("M4").Offset(i, 0).Value = Recommendations(i).symbol
            Globals.Dashboard.Range("N4").Offset(i, 0).Value = Recommendations(i).qty
            Globals.Dashboard.Range("O4").Offset(i, 0).Value = Recommendations(i).totValue
        Next

    End Sub
End Class
