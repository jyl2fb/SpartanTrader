Module DataSetProcedures
    Public Function GetMaxMargin() As Double
        Dim value As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTbl").Rows
            If myRow("Name").Trim() = "MaxMargins" Then
                value = myRow("Value")
                Return Double.Parse(value)
            End If
        Next
        MessageBox.Show("Holy BatMouse! Could not find 'MaxMargins'. Returned 0.")
        Return 0
    End Function

    Public Function GetRiskFreeRate() As Double
        Dim value As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTbl").Rows
            If myRow("Name").Trim() = "RiskFreeRate" Then
                value = myRow("Value")
                Return Double.Parse(value)
            End If
        Next
        MessageBox.Show("Holy BatMouse! Could not find 'RiskFreeRate'. Returned 0.")
        Return 0
    End Function

    Public Function GetInitialCAccount() As Double
        Dim value As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTbl").Rows
            If myRow("Name").Trim() = "CAccount" Then
                value = myRow("Value")
                Return Double.Parse(value)
            End If
        Next
        MessageBox.Show("Holy BatMouse! Could not find 'InitialCAccount'. Returned 0.")
        Return 0
    End Function

    Public Function GetStartDate() As Date
        Dim value As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTbl").Rows
            If myRow("Name").Trim() = "StartDate" Then
                value = myRow("Value")
                Return Date.Parse(value)
            End If
        Next
        MessageBox.Show("Holy BatScreen! Could not find 'StartDate'. Returned 1/1/1.")
        Return "1/1/1"
    End Function

    Public Function GetCAccount() As Double
        Dim value As String
        For Each myrow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            If myrow("symbol").trim() = "CAccount" Then
                value = myrow("units")
                Return Double.Parse(value)
            End If
        Next
        MessageBox.Show("holy Batkeyboard! could not find 'CAccount'. I reset the portfolio!",
                        "Reset Portfolio", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        UploadPosition("CAccount", GetInitialCAccount())
        Globals.Ribbons.RibbonST.AcquiredPositionsButton_Click(Nothing, Nothing)
        Return GetInitialCAccount()
    End Function

    Public Function IsAStock(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("TickersTbl").Rows
            If myRow("Ticker").trim() = symbol Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function GetAsk(symbol As String, targetDate As Date) As Double
        symbol = symbol.Trim()
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        DownloadPricesForOneDay(targetDate)

        If IsAStock(symbol) Then
            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTbl").Rows
                If myRow("Ticker").trim() = symbol Then
                    Return myRow("Ask")
                End If
            Next
        Else ' is an option
            For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
                If myRow("Symbol").trim() = symbol Then
                    Return myRow("Ask")
                End If
            Next
        End If
        MessageBox.Show("Holy Batsandals! Could not find the ask for " + symbol + ". Returned 0.")
        Return 0
    End Function

    Public Function GetBid(symbol As String, targetDate As Date) As Double
        symbol = symbol.Trim()
        If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
            targetDate = targetDate.AddDays(-1)
        End If
        If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
            targetDate = targetDate.AddDays(-2)
        End If

        DownloadPricesForOneDay(targetDate)

        If IsAStock(symbol) Then
            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTbl").Rows
                If myRow("Ticker").trim() = symbol Then
                    Return myRow("Bid")
                End If
            Next
        Else ' is an option
            For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
                If myRow("Symbol").trim() = symbol Then
                    Return myRow("Bid")
                End If
            Next
        End If
        MessageBox.Show("Holy Batsandals! Could not find the bid for " + symbol + ". Returned 0.")
        Return 0
    End Function

    Public Function GetDividend(ticker As String, targetDate As Date) As Double
        If IsAStock(ticker) Then
            If (targetDate.DayOfWeek = DayOfWeek.Saturday) Then
                targetDate = targetDate.AddDays(-1)
            End If
            If (targetDate.DayOfWeek = DayOfWeek.Sunday) Then
                targetDate = targetDate.AddDays(-2)
            End If

            DownloadPricesForOneDay(targetDate)

            For Each myRow As DataRow In myDataSet.Tables("StockMarketOneDayTbl").Rows
                If myRow("Ticker").trim() = ticker Then
                    Return Double.Parse(myRow("Dividend"))
                End If
            Next
        End If
        MessageBox.Show("Holy Batshoelace! I could not find the dividend for " + ticker + ". Returned 0.")
        Return 0
    End Function

    Public Function GetStrike(symbol As String) As Double
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol Then
                Return Double.Parse(myRow("Strike"))
            End If
        Next
        MessageBox.Show("Holy Batshoewax! Could not find the strike for " + symbol + ". Returned 0")
        Return 0
    End Function

    Public Function GetTrCostCoefficient(secType As String, trType As String) As Double
        For Each myRow As DataRow In myDataSet.Tables("TransactionCostTbl").Rows
            If myRow("SecurityType").Trim() = secType And myRow("TransactionType").Trim() = trType Then
                Return Double.Parse(myRow("CostCoeff"))
            End If
        Next
        MessageBox.Show("Holy Batsputs! I could not find the transaction cost. Returned 0")
        Return 0
    End Function

    Public Function GetCurrentPositionInAP(symbol) As Double
        For Each myRow As DataRow In myDataSet.Tables("AcquiredPositionsTbl").Rows
            If myRow("Symbol").ToString().Trim() = symbol Then
                Return Double.Parse(myRow("Units"))
            End If
        Next
        Return 0
    End Function

    Public Function GetUnderlier(symbol As String) As String
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOnedayTbl").Rows
            If myRow("Symbol").Trim() = symbol Then
                Return myRow("Underlier").Trim()
            End If
        Next
        MessageBox.Show("Holy BatCucumber! Could not find the underlier for " + symbol + ". Returned ???")
        Return "???"
    End Function

    Public Function IsACall(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol And myRow("Type").Trim() = "Call" Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function IsAPut(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").trim() = symbol And myRow("Type").Trim() = "Put" Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetExpiration(symbol As String) As Date
        symbol = symbol.Trim
        For Each myRow As DataRow In myDataSet.Tables("OptionMarketOneDayTbl").Rows
            If myRow("Symbol").Trim() = symbol Then
                Return myRow("Expiration")
            End If
        Next
        MessageBox.Show("Holy BatVBA! Could not find the expiration for " + symbol + ". Returned 1/1/1")
        Return "1/1/1"
    End Function

    Public Function IsAnOption(symbol As String) As Boolean
        symbol = symbol.Trim()
        For Each myRow As DataRow In myDataSet.Tables("SymbolsTbl").Rows
            If myRow("Symbol").trim() = symbol Then
                Return True
            End If
        Next
        Return False
    End Function

    Public Function GetEndDate() As Date
        Dim value As String
        For Each myRow As DataRow In myDataSet.Tables("EnvironmentVariableTbl").Rows
            If myRow("Name").Trim() = "EndDate" Then
                value = myRow("Value")
                Return Date.Parse(value)
            End If
        Next
        MessageBox.Show("Holy BatDoor! Could not find 'EndDate'. Returned 1/1/1.")
        Return "1/1/1"
    End Function
End Module
