Imports Microsoft.Office.Tools.Ribbon

Public Class RibbonST

    Public Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        'Program starts here
        'Activates Ribbon
        RibbonUI.ActivateTabMso("TabAddIns")
        Initialization()

    End Sub

    Public Sub AlphaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles AlphaTBtn.Click
        AlphaTBtn.Checked = True
        BetaTBtn.Checked = False
        GammaTBtn.Checked = False
        ActiveDB = "Alpha"
        MainProgram()
    End Sub

    Public Sub BetaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles BetaTBtn.Click
        AlphaTBtn.Checked = False
        BetaTBtn.Checked = True
        GammaTBtn.Checked = False
        ActiveDB = "Beta"
        MainProgram()
    End Sub

    Public Sub GammaTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles GammaTBtn.Click
        AlphaTBtn.Checked = False
        BetaTBtn.Checked = False
        GammaTBtn.Checked = True
        ActiveDB = "Gamma"
        MainProgram()
    End Sub

    Public Sub DashboardBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles DashboardBtn.Click
        Globals.Dashboard.Activate()
        Globals.Dashboard.Range("G1").Select()
    End Sub
    Public Sub TransactionsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TransactionsBtn.Click
        Globals.Transactions.Activate()
        Globals.Transactions.TransactionsLO.DataSource = Nothing
        Application.DoEvents()
        DownloadDataTableFromDB("Select * from TransactionQueue where teamid = " + teamID + " order by rowid desc", "TransactionsTbl")
        Globals.Transactions.TransactionsLO.DataSource = myDataSet.Tables("TransactionsTbl")

        Globals.Transactions.Range("B3").NumberFormat = "mm/dd/yyyy"
        Globals.Transactions.Range("M3").NumberFormat = "mm/dd/yyyy"

        Globals.Transactions.TransactionsLO.Range.Columns.AutoFit()
        Globals.Transactions.Range("A1").Select()
    End Sub
    Public Sub StockMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StockMktBtn.Click
        Globals.Markets.Activate()
        Globals.Markets.StockMarketLO.DataSource = Nothing
        DownloadDataTableFromDB("Select * from StockMarket order by date desc", "StockMarketTbl")
        Application.DoEvents()
        Globals.Markets.StockMarketLO.DataSource = myDataSet.Tables("StockMarketTbl")
        Globals.Markets.Range("B3").NumberFormat = "mm/dd/yyyy"
        Globals.Markets.Range("F3").NumberFormat = "mm/dd/yyyy"
        Globals.Markets.StockMarketLO.Range.Columns.AutoFit()
        Globals.Markets.Range("A1").Select()
    End Sub

    Public Sub OptionsMktBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles OptionsMktBtn.Click
        Globals.Markets.Activate()
        Globals.Markets.OptionsLO.DataSource = Nothing
        DownloadDataTableFromDB("Select * from OptionMarket order by date desc", "OptionMarketTbl")
        Application.DoEvents()
        Globals.Markets.OptionsLO.DataSource = myDataSet.Tables("OptionMarketTbl")
        Globals.Markets.Range("I3").NumberFormat = "mm/dd/yyyy"
        Globals.Markets.Range("N3").NumberFormat = "mm/dd/yyyy"
        Globals.Markets.OptionsLO.Range.Columns.AutoFit()
        Globals.Markets.Range("A1").Select()
    End Sub

    Public Sub SP500Btn_Click(sender As Object, e As RibbonControlEventArgs) Handles SP500Btn.Click
        Globals.Markets.Activate()
        Globals.Markets.SP500LO.DataSource = Nothing
        DownloadDataTableFromDB("Select * from StockIndex order by date desc", "IndexTbl")
        Application.DoEvents()
        Globals.Markets.SP500LO.DataSource = myDataSet.Tables("IndexTbl")
        Globals.Markets.Range("R3").NumberFormat = "mm/dd/yyyy"
        Globals.Markets.SP500LO.Range.Columns.AutoFit()
        Globals.Markets.Range("A1").Select()
    End Sub



    Public Sub QuitBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles QuitBtn.Click
        Globals.ThisWorkbook.Application.DisplayAlerts = False
        Globals.ThisWorkbook.Application.DisplayFormulaBar = False
        If traderMode = "Simulation" Then
            currentDate = GetEndDate()
        End If
        DisconnectFromDB()
        Globals.ThisWorkbook.Application.Quit()
    End Sub

    Public Sub InitialPositionsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles InitialPositionsButton.Click
        Globals.Dashboard.Activate()
        Globals.Dashboard.InitialLO.DataSource = Nothing
        Application.DoEvents()
        DownloadDataTableFromDB("Select * from InitialPosition order by symbol", "InitialPositionTbl")
        Globals.Dashboard.InitialLO.DataSource = myDataSet.Tables("InitialPositionTbl")
        Globals.Dashboard.InitialLO.DataBodyRange.Interior.Color = System.Drawing.Color.Black
        Globals.Dashboard.InitialLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.LightSkyBlue
        Globals.Dashboard.InitialLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0;-"
        Globals.Dashboard.InitialLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,###;[Red]$ -#,##0;-"
        Globals.Dashboard.InitialLO.Range.ColumnWidth = 12
        Globals.Dashboard.Range("G1").Select()
        CalcFinancialMetrics(currentDate)
        DisplayFinancialMetrics(currentDate)

    End Sub

    Public Sub AcquiredPositionsButton_Click(sender As Object, e As RibbonControlEventArgs) Handles AcquiredPositionsButton.Click
        Globals.Dashboard.Activate()
        Globals.Dashboard.AcquiredLO.DataSource = Nothing
        Application.DoEvents()
        DownloadDataTableFromDB("Select * from PortfolioTeam" + teamID + " order by symbol", "AcquiredPositionTbl")
        Globals.Dashboard.AcquiredLO.DataBodyRange.Interior.Color = System.Drawing.Color.Black
        Globals.Dashboard.AcquiredLO.ListColumns(1).Range.Font.Color = System.Drawing.Color.LightSkyBlue
        Globals.Dashboard.AcquiredLO.ListColumns(2).Range.NumberFormat = "[White]#,##0;[Red]-#,##0;[White]-"
        Globals.Dashboard.AcquiredLO.ListColumns(3).Range.NumberFormat = "[Green]$ #,###;[Red]$ -#,##0;[Green]-"
        Globals.Dashboard.AcquiredLO.Range.ColumnWidth = 12
        Globals.Dashboard.Range("G1").Select()
        CalcFinancialMetrics(currentDate)
        DisplayFinancialMetrics(currentDate)
    End Sub

    Public Sub ConfirmationTicketButton_Click(sender As Object, e As RibbonControlEventArgs) Handles ConfirmationTicketButton.Click
        Globals.Transactions.Activate()
        Globals.Transactions.ConfirmationTicketsLO.DataSource = Nothing
        Application.DoEvents()
        DownloadDataTableFromDB("Select * from ConfirmationTicketTeam" + teamID + " order by rowid desc", "ConfirmationTicketTbl")
        Globals.Transactions.ConfirmationTicketsLO.DataSource = myDataSet.Tables("ConfirmationTicketTbl")
        Globals.Transactions.ConfirmationTicketsLO.Range.Columns.AutoFit()
        Globals.Transactions.Range("A1").Select()
    End Sub

    Public Sub TCostsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles TCostsBtn.Click
        Globals.Parameters.Activate()
        Globals.Parameters.TransactionCostsLO.DataSource = Nothing
        DownloadDataTableFromDB("Select * from TransactionCost", "TcostTbl")
        Application.DoEvents()
        Globals.Parameters.TransactionCostsLO.DataSource = myDataSet.Tables("TcostTbl")
        Globals.Parameters.TransactionCostsLO.Range.Columns.AutoFit()
        Globals.Parameters.Range("A1").Select()
    End Sub

    Public Sub ParametersBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ParametersBtn.Click
        Globals.Parameters.Activate()
        Globals.Parameters.EnvVariablesLO.DataSource = Nothing
        DownloadDataTableFromDB("Select * from EnvironmentVariable", "EnvVarTbl")
        Application.DoEvents()
        Globals.Parameters.EnvVariablesLO.DataSource = myDataSet.Tables("EnvVarTbl")
        Globals.Parameters.EnvVariablesLO.Range.Columns.AutoFit()
        Globals.Parameters.Range("A1").Select()
    End Sub

    Public Sub ResetPortfolioBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ResetPortfolioBtn.Click
        If MessageBox.Show("Are you sure, Dave? There is no undo...", "Reset AP?",
        MessageBoxButtons.YesNo, MessageBoxIcon.Hand) = DialogResult.Yes Then

            ClearTeamPortfolioOnDB()
            initialCAccount = GetInitialCAccount()
            UploadPosition("CAccount", initialCAccount)
            MainProgram()
        End If
    End Sub

    Public Sub EditControlBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles EditControlBtn.Click
        AcquiredPositionsButton_Click(Nothing, Nothing)
        For i = 1 To 5
            myDataSet.Tables("AcquiredPOsitionsTbl").Rows.Add()
        Next
    End Sub

    Public Sub UploadPortfolioBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles UploadPortfolioBtn.Click
        UploadScreenPortfolioToDb()
        MainProgram()
    End Sub

    Public Sub FinChartsBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles FinChartsBtn.Click
        Globals.FinCharts.FillLBoxes()
        Globals.FinCharts.SetupFinCharts()
        Globals.FinCharts.Activate()
        Globals.FinCharts.Range("A1").Select()

    End Sub

    Public Sub ManualTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles ManualTBtn.Click
        ManualTBtn.Checked = True
        SyncTBtn.Checked = False
        SimulationTBtn.Checked = False
        StepSimTBtn.Checked = False
        RoboTraderTBtn.Checked = False
        traderMode = "Manual"
        MainProgram()
    End Sub

    Public Sub SyncTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SyncTBtn.Click
        ManualTBtn.Checked = False
        SyncTBtn.Checked = True
        SimulationTBtn.Checked = False
        StepSimTBtn.Checked = False
        RoboTraderTBtn.Checked = False
        If traderMode = "RoboTrader" Then
            traderMode = "Sync"
        Else
            traderMode = "Sync"
            MainProgram()
        End If
    End Sub

    Public Sub RoboTraderTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles RoboTraderTBtn.Click
        ManualTBtn.Checked = False
        SyncTBtn.Checked = False
        SimulationTBtn.Checked = False
        StepSimTBtn.Checked = False
        RoboTraderTBtn.Checked = True
        If traderMode = "Sync" Then
            traderMode = "RoboTrader"

        Else
            traderMode = "RoboTrader"
            MainProgram()
        End If
    End Sub

    Public Sub SimulationTBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles SimulationTBtn.Click
        ManualTBtn.Checked = False
        SyncTBtn.Checked = False
        SimulationTBtn.Checked = True
        StepSimTBtn.Checked = False
        RoboTraderTBtn.Checked = False
        traderMode = "Simulation"
        MainProgram()
    End Sub

    Public Sub StepSimBtn_Click(sender As Object, e As RibbonControlEventArgs) Handles StepSimTBtn.Click
        ManualTBtn.Checked = False
        SyncTBtn.Checked = False
        SimulationTBtn.Checked = False
        StepSimTBtn.Checked = True
        RoboTraderTBtn.Checked = False
        traderMode = "StepSim"
        MainProgram()
    End Sub
End Class
