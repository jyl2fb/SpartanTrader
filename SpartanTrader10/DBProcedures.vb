Module ConnectToDatabase
    ' DB Connection-related objects
    Public myConnection As SqlClient.SqlConnection
    Public myCommand As SqlClient.SqlCommand
    Public myDataAdapter As SqlClient.SqlDataAdapter
    Public myDataSet As DataSet

    Public Sub ConnectToDB(myConnString As String)          ' this uses ADO technology 
        Try
            myConnection = New SqlClient.SqlConnection      ' create connection
            myConnection.ConnectionString = myConnString    ' sets connection string
            myCommand = New SqlClient.SqlCommand            ' the command represents the SQL query
            myCommand.Connection = myConnection             ' links command and connection
            myDataAdapter = New SqlClient.SqlDataAdapter    ' adapter
            myDataAdapter.SelectCommand = myCommand         ' links adapter and command
            myDataSet = New DataSet
            myConnection.Open()
        Catch ex As Exception
            MessageBox.Show("Holy Batmobile! I could not connect. This is what you gave me: " + myConnString,
                            "Connection problem!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub DisconnectFromDB()
        myConnection.Close()
    End Sub

    Public Sub DownloadDataTableFromDB(mySQLQuery As String, NameOfOutputTable As String)
        If myDataSet.Tables.Contains(NameOfOutputTable) Then
            myDataSet.Tables(NameOfOutputTable).Clear()
        End If
        Try
            myCommand.CommandText = mySQLQuery
            myDataAdapter.Fill(myDataSet, NameOfOutputTable)
        Catch ex As Exception
            MessageBox.Show("Holy Batchopper! I could not run this query. This is the query: " + mySQLQuery, "Query Problem!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub RunNonQuery(SQLNonQuery As String)
        Try
            myCommand.CommandText = SQLNonQuery
            myCommand.ExecuteNonQuery()

        Catch ex As Exception
            MessageBox.Show("Holy Batplane! I could not run this nonquery. This is the query: " + SQLNonQuery,
                            "Query problem!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub ConnectToActiveDB()
        Select Case ActiveDB
            Case "Alpha"
                ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;" +
                            "Initial Catalog=HedgeTournamentALPHA;Integrated Security=True")
            Case "Beta"
                ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;" +
                            "Initial Catalog=HedgeTournamentBETA;Integrated Security=True")
            Case "Gamma"
                ConnectToDB("Data Source=f-sg6m-s4.comm.virginia.edu;" +
                            "Initial Catalog=HedgeTournamentGAMMA;Integrated Security=True")
            Case Else
                MessageBox.Show("Holy Batmobile! No active database selected.",
                    "Spartan Trader", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Select

    End Sub

    Public Function DownloadCurrentDate() As Date
        Try
            Dim temp As String = ""
            myCommand.CommandText = "Select Value from EnvironmentVariable where Name = 'CurrentDate'"
            temp = myCommand.ExecuteScalar()
            Return Date.Parse(temp)
        Catch ex As Exception
            MessageBox.Show("Holy Batfrog! I could not get you the current date.", "Spartan Trader", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Date.Parse("1/1/1")
        End Try
    End Function

    Public Sub DownloadStaticData()
        DownloadDataTableFromDB("Select * from InitialPosition order by symbol", "InitialPositionsTbl")
        DownloadDataTableFromDB("Select distinct ticker from StockMarket order by ticker", "TickersTbl")
        DownloadDataTableFromDB("Select * from EnvironmentVariable", "EnvironmentVariableTbl")
        DownloadDataTableFromDB("Select distinct symbol, underlier from OptionMarket order by symbol", "SymbolsTbl")
        DownloadDataTableFromDB("Select * from TransactionCost", "TransactionCostTbl")
    End Sub

    Public Sub DownloadTeamData(targetDate As Date)
        DownloadDataTableFromDB("Select * from portfolioteam" + teamID + " order by symbol", "AcquiredPositionsTbl")
        lastTransactionDate = DownloadLastTransactionDate(targetDate)
        lastTEUpDate = startDate
        TEpercent = 0
        sumTE = 0
    End Sub

    Public Function DownloadLastTransactionDate(targetDate As Date) As Date
        Dim temp As String = ""
        myCommand.CommandText = String.Format("Select max(date) from TransactionQueue where teamid = {0} and date <= '{1}'",
                                            teamID, targetDate.ToShortDateString())
        Try
            temp = myCommand.ExecuteScalar()
            Return Date.Parse(temp)
        Catch myException As Exception
            Return startDate

        End Try
    End Function

    Public Sub DownloadPricesForOneDay(targetDate As Date)
        If targetDate.Date = lastPriceDownloadDate.Date Then
            Return
        Else
            DownloadDataTableFromDB("Select * from StockMarket where Date = '" + targetDate.ToShortDateString() + "'", "StockMarketOneDayTbl")
            DownloadDataTableFromDB("Select * from OptionMarket where Date = '" + targetDate.ToShortDateString() + "'", "OptionMarketOneDayTbl")
            lastPriceDownloadDate = targetDate

        End If
    End Sub

    Public Sub UploadPosition(sym As String, newUnits As Double)
        Try
            Dim sql As String
            newUnits = Math.Round(newUnits, 2)
            sym = sym.Trim()

            sql = "Delete from portfolioteam" + teamID + " where Symbol = '" + sym + "';"
            RunNonQuery(sql)

            If (newUnits = 0) And (sym <> "CAccount") Then
                'skip 
            Else
                sql = String.Format("Insert into portfolioTeam{0} Values ('{1}', {2}, 0)",
                teamID,
                sym,
                newUnits)
                RunNonQuery(sql)
            End If
        Catch myException As Exception
            MessageBox.Show("I could not upload " + sym + " to AP, Dave. " +
                        "Maybe this will help: " + myException.Message)
        End Try
    End Sub

    Public Sub ClearTeamPortfolioOnDB()
        Dim sql As String = "Delete from PortfolioTeam" + teamID
        RunNonQuery(sql)
    End Sub

    Public Sub UploadScreenPortfolioToDb()
        Dim tempSymbol, tempUnits As String
        If Globals.ThisWorkbook.ActiveSheet.Name <> "Dashboard" Then
            MessageBox.Show("Are you looking at the Portfolio that you want me to upload, Dave?",
                            "Portfolio Not Active",
                            MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return
        End If

        If Globals.Dashboard.AcquiredLO.IsSelected Then
            MessageBox.Show("Click outside the ListObject to confirm data entry, Dave.", "Edit In Progress",
                            MessageBoxButtons.OK, MessageBoxIcon.Hand)
            Return
        End If
        ClearTeamPortfolioOnDB()
        For i As Integer = 1 To Globals.Dashboard.AcquiredLO.DataBodyRange.Rows.Count()
            tempSymbol = Globals.Dashboard.AcquiredLO.DataBodyRange.Cells(i, 1).Value
            tempUnits = Globals.Dashboard.AcquiredLO.DataBodyRange.Cells(i, 2).Value
            If IsAPEntryValid(tempSymbol, tempUnits) Then
                UploadPosition(tempSymbol, tempUnits)
            End If
        Next
    End Sub

    Public Function IsThereData() As Boolean
        Try
            myCommand.CommandText = "Select count(*) from StockMarket"
            If myCommand.ExecuteScalar() > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

End Module

