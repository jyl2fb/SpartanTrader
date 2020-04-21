
Public Class FinCharts

    Private Sub Sheet5_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet5_Shutdown() Handles Me.Shutdown

    End Sub

    Public Sub FillLBoxes()
        TickersLBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            TickersLBox.Items.Add(myRow("Ticker").ToString.Trim())
        Next

        SymbolsLBox.Items.Clear()
        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            SymbolsLBox.Items.Add(myRow("Symbol").ToString.Trim())
        Next
    End Sub


    Public Sub SetupFinCharts()
        'format the chart - feel free to change any formatting
        StockChart.ChartType = Excel.XlChartType.xlLine
        StockDataToChartLO.AutoSetDataBoundColumnHeaders = True

        'format the y axis as $
        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        y.HasTitle = False
        y.HasMinorGridlines = True
        y.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y.TickLabels.NumberFormat = "$###.00"

        'format the x axis as dates
        Dim x As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlCategory)
        x.HasTitle = False
        x.CategoryType = Excel.XlCategoryType.xlTimeScale
        x.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x.BaseUnit = Excel.XlTimeUnit.xlDays
        x.TickLabels.NumberFormat = "d-mmm"

        'format the chart - feel free to change any formatting
        OptionChart.ChartType = Excel.XlChartType.xlLine
        OptionDataToChartLO.AutoSetDataBoundColumnHeaders = True

        'format the y axis as $
        Dim y2 As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlValue)
        y2.HasTitle = False
        y2.HasMinorGridlines = True
        y2.MinorTickMark = Excel.XlTickMark.xlTickMarkOutside
        y2.TickLabels.NumberFormat = "$###.00"

        'format the x axis as dates
        Dim x2 As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlCategory)
        x2.HasTitle = False
        x2.CategoryType = Excel.XlCategoryType.xlTimeScale
        x2.MajorTickMark = Excel.XlTickMark.xlTickMarkCross
        x2.BaseUnit = Excel.XlTimeUnit.xlDays
        x2.TickLabels.NumberFormat = "d-mmm"
    End Sub

    Private Sub SymbolsLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SymbolsLBox.SelectedIndexChanged
        Dim t As String = ""
        Dim sql As String = ""
        t = SymbolsLBox.SelectedItem.Trim()
        sql = "Select date, bid, ask from optionmarket where symbol = '" + t + "'"
        DownloadDataTableFromDB(sql, "OptionDataToChartTbl")
        OptionDataToChartLO.DataSource = myDataSet.Tables("OptionDataToChartTbl")
        OptionChart.SetSourceData(OptionDataToChartLO.Range)
        OptionChart.ChartTitle.Text = "Daily Closings for " + t
        Dim y As Excel.Axis = OptionChart.Axes(Excel.XlAxisType.xlValue)
        ' this line sets the scale of the chart for better viewing 
        y.MinimumScale = Math.Truncate(FindMinBid("OptionDataToChartTbl") / 10) * 10
    End Sub

    Private Sub TickersLBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TickersLBox.SelectedIndexChanged
        Dim t As String = ""
        Dim sql As String = ""
        t = TickersLBox.SelectedItem.Trim()
        sql = "Select date, bid, ask from stockmarket where ticker = '" + t + "'"
        DownloadDataTableFromDB(sql, "StockDataToChartTbl")
        StockDataToChartLO.DataSource = myDataSet.Tables("StockDataToChartTbl")
        StockChart.SetSourceData(StockDataToChartLO.Range)
        StockChart.ChartTitle.Text = "Daily Closings for " + t
        Dim y As Excel.Axis = StockChart.Axes(Excel.XlAxisType.xlValue)
        ' this line sets the scale of the chart for better viewing 
        y.MinimumScale = Math.Truncate(FindMinBid("StockDataToChartTbl") / 10) * 10
    End Sub

    Public Function FindMinBid(tablename As String) As Double
        Dim tempMin As Double = 1000000
        For Each myRow As DataRow In myDataSet.Tables(tablename).Rows
            tempMin = Math.Min(myRow("Bid"), tempMin)
        Next
        Return tempMin
    End Function
End Class
