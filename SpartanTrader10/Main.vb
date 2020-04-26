Module Main
    'testing my change

    Public Sub Initialization()
        Globals.ThisWorkbook.Application.DisplayFormulaBar = False
        Globals.Dashboard.TeamIDCell.Value = "TeamID: " + teamID
        Globals.Ribbons.RibbonST.AlphaTBtn_Click(Nothing, Nothing)
    End Sub

    Public Sub MainProgram()
        Globals.Dashboard.Activate()
        ClearAllLO()
        lastPriceDownloadDate = "1/1/1"
        ConnectToActiveDB()
        Globals.Dashboard.SetupTEChart()
        CreateCurrentTransaction()
        If IsThereData() = True Then
            currentDate = DownloadCurrentDate()
            DownloadStaticData()
            DownloadTeamData(currentDate)
            SetFinancialConstants()
            CreateCurrentTransaction()
            ResetAllRecommendations()
            Select Case traderMode
                Case "Manual"
                    ResetToBeginningOfTournament()
                    StopTimers()
                    RunDailyRoutine(currentDate)
                Case "Simulation", "StepSim"
                    StopTimers()
                    ResetToBeginningOfTournament()
                    Try
                        Do
                            RunDailyRoutine(currentDate)
                            currentDate = currentDate.AddDays(1)
                        Loop While currentDate <= GetEndDate() And traderMode <> "Manual"
                    Catch ex As Exception
                        Exit Sub
                    End Try
                Case "Sync", "RoboTrader"
                    DisplayFinancialMetrics(currentDate)
                    StartTimers()
            End Select
        Else 'no data
            WaitForData()
        End If
    End Sub

    Sub ClearAllLO()
        Globals.Markets.StockMarketLO.DataSource = Nothing
        Globals.Markets.OptionsLO.DataSource = Nothing
        Globals.Markets.SP500LO.DataSource = Nothing
        Globals.Parameters.EnvVariablesLO.DataSource = Nothing
        Globals.Parameters.TransactionCostsLO.DataSource = Nothing
        Globals.Transactions.TransactionsLO.DataSource = Nothing
        Globals.Transactions.ConfirmationTicketsLO.DataSource = Nothing
        Globals.Dashboard.AcquiredLO.DataSource = Nothing
        Globals.Dashboard.InitialLO.DataSource = Nothing


    End Sub

    Public Sub SetFinancialConstants()
        Globals.Dashboard.FillCBoxes()
        maxMargin = GetMaxMargin()
        startDate = GetStartDate()
        riskFreeRate = GetRiskFreeRate()
        initialCAccount = GetInitialCAccount()
        TPVatStart = CalcTPVAtStart()
    End Sub

    Public Sub CalcFinancialMetrics(targetDate As Date)
        CAccount = GetCAccount()
        IPvalue = CalcIPValue(targetDate)
        APvalue = CalcAPValue(targetDate)
        margin = CalcMargin(targetDate)
        TPVatStart = CalcTPVAtStart()
        TaTPV = CalcTaTPV(targetDate)
        TPV = CalcTPV(targetDate)
        TE = CalcTE()
        TEpercent = TE / TaTPV
        sumTE = sumTE + UpdateSumTE(targetDate)

    End Sub

    Public Sub CreateCurrentTransaction()
        CT = New Transaction()
        CT.Clear()
    End Sub

    Public Sub DisplayFinancialMetrics(targetDate As Date)

        CT.Show()

        Globals.Dashboard.Range("G04").Value = CAccount
        Globals.Dashboard.Range("G05").Value = margin * 0.3
        Globals.Dashboard.Range("G06").Value = margin
        Globals.Dashboard.Range("G07").Value = maxMargin
        Globals.Dashboard.Range("G09").Value = IPvalue
        Globals.Dashboard.Range("G10").Value = APvalue
        Globals.Dashboard.Range("G11").Value = TPVatStart
        Globals.Dashboard.Range("G12").Value = TPV
        Globals.Dashboard.Range("G13").Value = TaTPV
        Globals.Dashboard.Range("G15").Value = TE
        Globals.Dashboard.Range("G16").Value = TEpercent
        Globals.Dashboard.Range("G17").Value = sumTE

        Globals.Dashboard.InitialLO.DataBodyRange.Interior.Color = System.Drawing.Color.Black
        Globals.Dashboard.InitialLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.LightSkyBlue
        Globals.Dashboard.InitialLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0;[White]-"
        Globals.Dashboard.InitialLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,###;[Red]$ -#,##0;[Green]-"
        Globals.Dashboard.InitialLO.DataSource = myDataSet.Tables("InitialPositionsTbl")
        Globals.Dashboard.InitialLO.Range.Columns.AutoFit()


        Globals.Dashboard.AcquiredLO.DataBodyRange.Interior.Color = System.Drawing.Color.Black
        Globals.Dashboard.AcquiredLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.LightSkyBlue
        Globals.Dashboard.AcquiredLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0;[White]-"
        Globals.Dashboard.AcquiredLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,###;[Red]$ -#,##0;[Green]-"
        Globals.Dashboard.AcquiredLO.DataSource = myDataSet.Tables("AcquiredPositionsTbl")
        Globals.Dashboard.AcquiredLO.Range.Columns.AutoFit()
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

    Public Sub RunDailyRoutine(tdate As Date)
        Globals.Dashboard.DateLine.Value = tdate.ToLongDateString()
        CalcFinancialMetrics(tdate)
        DisplayFinancialMetrics(tdate)
        Globals.Dashboard.UpdateTEChart(tdate)
        Select Case traderMode
            Case "Manual"
                CalcAllRecommendations(currentDate)
                Globals.Dashboard.DisplayAllRecommendations()
            Case "Simulation"
                RoboExecuteAll(tdate)
            Case "StepSim"
                RoboExecuteStepByStep(tdate)
            Case "Sync"
                CalcAllRecommendations(currentDate)
                Globals.Dashboard.DisplayAllRecommendations()
            Case "RoboTrader"
                RoboExecuteAll(tdate)
        End Select
    End Sub

    Public Sub ResetToBeginningOfTournament()
        currentDate = GetStartDate()
        ClearTeamPortfolioOnDB()
        initialCAccount = GetInitialCAccount()
        UploadPosition("CAccount", initialCAccount)
        DownloadTeamData(currentDate)
        lastTransactionDate = GetStartDate()
        CalcAllRecommendations(currentDate)
        Globals.Dashboard.DisplayAllRecommendations()
    End Sub

    Public Sub WaitForData()
        If MessageBox.Show("There is no data in the database, Dave... Would you like to wait for it?", "Hal",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) = DialogResult.Yes Then
            Globals.Ribbons.RibbonST.ManualTBtn.Checked = False
            Globals.Ribbons.RibbonST.SyncTBtn.Checked = False
            Globals.Ribbons.RibbonST.SimulationTBtn.Checked = False
            Globals.Ribbons.RibbonST.StepSimTBtn.Checked = False
            Globals.Ribbons.RibbonST.RoboTraderTBtn.Checked = False
            waitingForData = True
            traderMode = "RoboTrader"
            StartTimers()
        Else
            Select Case ActiveDB
                Case "Alpha"
                    Globals.Ribbons.RibbonST.BetaTBtn_Click(Nothing, Nothing)
                Case "Beta"
                    Globals.Ribbons.RibbonST.GammaTBtn_Click(Nothing, Nothing)
                Case "Gamma"
                    Globals.Ribbons.RibbonST.AlphaTBtn_Click(Nothing, Nothing)
            End Select
        End If
    End Sub
End Module
