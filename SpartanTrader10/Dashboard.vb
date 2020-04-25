
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

    Public Sub SetupTEChart()
        'create the table
        If myDataSet.Tables.Contains("TETbl") Then
            myDataSet.Tables("TETbl").Clear()
        Else ' create a new table in the dataset
            myDataSet.Tables.Add("TETbl")
            myDataSet.Tables("TETbl").Columns.Add("Date", GetType(Date))
            myDataSet.Tables("TETbl").Columns.Add("TaTPV", GetType(Double))
            myDataSet.Tables("TETbl").Columns.Add("NoHedge", GetType(Double))
            myDataSet.Tables("TETbl").Columns.Add("TPV", GetType(Double))
        End If
        TELO.DataSource = myDataSet.Tables("TETbl")

        'format the chart - feel free to change any formatting
        TEChart.ChartType = Excel.XlChartType.xlLine
        'TEChart.ChartStyle = 6
        'TEChart.ApplyLayout(3)
        TEChart.HasTitle = False

        'format the y axis as $
        Dim y As Excel.Axis = TEChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$#,###"
        y.MinimumScaleIsAuto = False
        y.MaximumScaleIsAuto = True

        'format the x axis as dates
        Dim x As Excel.Axis = TEChart.Axes(Excel.XlAxisType.xlCategory)
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "[$-409]d-mmm;@"

        Dim s As Excel.SeriesCollection = TEChart.SeriesCollection
        s(0).Format.Line.Weight = 2
        s(0).Format.Line.ForeColor.RGB = System.Drawing.Color.SteelBlue
        s(1).Format.Line.Weight = 2
        s(1).Format.Line.ForeColor.RGB = System.Drawing.Color.Gray
        s(2).Format.Line.Weight = 2
        s(2).Format.Line.ForeColor.RGB = System.Drawing.Color.Orange
    End Sub

    Public Sub UpdateTEChart(targetDate As Date)
        Dim interestOnInitialCA As Double = 0
        Dim ts As TimeSpan = targetDate.Date - startDate.Date
        Dim t As Double = ts.Days / 365.25

        interestOnInitialCA = initialCAccount * (Math.Exp(riskFreeRate * t) - 1)

        Dim tempRow As DataRow
        tempRow = myDataSet.Tables("TETbl").Rows.Add()
        tempRow("Date") = targetDate.ToShortDateString
        tempRow("TPV") = TPV
        tempRow("TaTPV") = TaTPV
        tempRow("NoHedge") = IPvalue + initialCAccount + interestOnInitialCA

        TEChart.SetSourceData(TELO.Range)

        ' these lines set the scale of the chart for better viewing
        Dim y As Excel.Axis = TEChart.Axes(Excel.XlAxisType.xlValue)
        y.MinimumScale = Math.Truncate((FindMinInTPVTrackingTable() / 10000000)) * 10000000
    End Sub

    Public Function FindMinInTPVTrackingTable() As Double
        Dim tempMin As Double = 100000000
        For Each myRow As DataRow In myDataSet.Tables("TETbl").Rows
            tempMin = Math.Min(myRow("TPV"), tempMin)
            tempMin = Math.Min(myRow("TaTPV"), tempMin)
            tempMin = Math.Min(myRow("NoHedge"), tempMin)
        Next
        Return tempMin
    End Function
End Class
