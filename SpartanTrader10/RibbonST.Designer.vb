Partial Class RibbonST
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Database = Me.Factory.CreateRibbonGroup
        Me.AlphaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.BetaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.GammaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.Dashboard = Me.Factory.CreateRibbonGroup
        Me.DashboardBtn = Me.Factory.CreateRibbonButton
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.InitialPositionsButton = Me.Factory.CreateRibbonButton
        Me.AcquiredPositionsButton = Me.Factory.CreateRibbonButton
        Me.TransactionsBtn = Me.Factory.CreateRibbonButton
        Me.ResetPortfolioBtn = Me.Factory.CreateRibbonButton
        Me.EditControlBtn = Me.Factory.CreateRibbonButton
        Me.UploadPortfolioBtn = Me.Factory.CreateRibbonButton
        Me.ConfirmationTicketButton = Me.Factory.CreateRibbonButton
        Me.BI = Me.Factory.CreateRibbonGroup
        Me.StockMktBtn = Me.Factory.CreateRibbonButton
        Me.OptionsMktBtn = Me.Factory.CreateRibbonButton
        Me.SP500Btn = Me.Factory.CreateRibbonButton
        Me.TCostsBtn = Me.Factory.CreateRibbonButton
        Me.ParametersBtn = Me.Factory.CreateRibbonButton
        Me.FinChartsBtn = Me.Factory.CreateRibbonButton
        Me.Control = Me.Factory.CreateRibbonGroup
        Me.QuitBtn = Me.Factory.CreateRibbonButton
        Me.ModeGroup = Me.Factory.CreateRibbonGroup
        Me.ManualTBtn = Me.Factory.CreateRibbonToggleButton
        Me.SyncTBtn = Me.Factory.CreateRibbonToggleButton
        Me.RoboTraderTBtn = Me.Factory.CreateRibbonToggleButton
        Me.SimulationTBtn = Me.Factory.CreateRibbonToggleButton
        Me.StepSimBtn = Me.Factory.CreateRibbonToggleButton
        Me.Tab1.SuspendLayout()
        Me.Database.SuspendLayout()
        Me.Dashboard.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.BI.SuspendLayout()
        Me.Control.SuspendLayout()
        Me.ModeGroup.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Database)
        Me.Tab1.Groups.Add(Me.ModeGroup)
        Me.Tab1.Groups.Add(Me.Dashboard)
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.BI)
        Me.Tab1.Groups.Add(Me.Control)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Database
        '
        Me.Database.Items.Add(Me.AlphaTBtn)
        Me.Database.Items.Add(Me.BetaTBtn)
        Me.Database.Items.Add(Me.GammaTBtn)
        Me.Database.Label = "Database"
        Me.Database.Name = "Database"
        '
        'AlphaTBtn
        '
        Me.AlphaTBtn.Label = "Alpha DB"
        Me.AlphaTBtn.Name = "AlphaTBtn"
        '
        'BetaTBtn
        '
        Me.BetaTBtn.Label = "Beta DB"
        Me.BetaTBtn.Name = "BetaTBtn"
        '
        'GammaTBtn
        '
        Me.GammaTBtn.Label = "Gamma DB"
        Me.GammaTBtn.Name = "GammaTBtn"
        '
        'Dashboard
        '
        Me.Dashboard.Items.Add(Me.DashboardBtn)
        Me.Dashboard.Label = "Dashboard"
        Me.Dashboard.Name = "Dashboard"
        '
        'DashboardBtn
        '
        Me.DashboardBtn.Label = "Dashboard"
        Me.DashboardBtn.Name = "DashboardBtn"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.InitialPositionsButton)
        Me.Group1.Items.Add(Me.AcquiredPositionsButton)
        Me.Group1.Items.Add(Me.TransactionsBtn)
        Me.Group1.Items.Add(Me.ResetPortfolioBtn)
        Me.Group1.Items.Add(Me.EditControlBtn)
        Me.Group1.Items.Add(Me.UploadPortfolioBtn)
        Me.Group1.Items.Add(Me.ConfirmationTicketButton)
        Me.Group1.Label = "Portfolio Management"
        Me.Group1.Name = "Group1"
        '
        'InitialPositionsButton
        '
        Me.InitialPositionsButton.Label = "Initial Positions"
        Me.InitialPositionsButton.Name = "InitialPositionsButton"
        '
        'AcquiredPositionsButton
        '
        Me.AcquiredPositionsButton.Label = "Acquired Positions"
        Me.AcquiredPositionsButton.Name = "AcquiredPositionsButton"
        '
        'TransactionsBtn
        '
        Me.TransactionsBtn.Label = "Transactions"
        Me.TransactionsBtn.Name = "TransactionsBtn"
        '
        'ResetPortfolioBtn
        '
        Me.ResetPortfolioBtn.Label = "Reset"
        Me.ResetPortfolioBtn.Name = "ResetPortfolioBtn"
        '
        'EditControlBtn
        '
        Me.EditControlBtn.Label = "Edit"
        Me.EditControlBtn.Name = "EditControlBtn"
        '
        'UploadPortfolioBtn
        '
        Me.UploadPortfolioBtn.Label = "Upload"
        Me.UploadPortfolioBtn.Name = "UploadPortfolioBtn"
        '
        'ConfirmationTicketButton
        '
        Me.ConfirmationTicketButton.Label = "Confirmation Tickets"
        Me.ConfirmationTicketButton.Name = "ConfirmationTicketButton"
        '
        'BI
        '
        Me.BI.Items.Add(Me.StockMktBtn)
        Me.BI.Items.Add(Me.OptionsMktBtn)
        Me.BI.Items.Add(Me.SP500Btn)
        Me.BI.Items.Add(Me.TCostsBtn)
        Me.BI.Items.Add(Me.ParametersBtn)
        Me.BI.Items.Add(Me.FinChartsBtn)
        Me.BI.Label = "Business Intelligence"
        Me.BI.Name = "BI"
        '
        'StockMktBtn
        '
        Me.StockMktBtn.Label = "Stock Mkt."
        Me.StockMktBtn.Name = "StockMktBtn"
        '
        'OptionsMktBtn
        '
        Me.OptionsMktBtn.Label = "Options Mkt"
        Me.OptionsMktBtn.Name = "OptionsMktBtn"
        '
        'SP500Btn
        '
        Me.SP500Btn.Label = "SP500"
        Me.SP500Btn.Name = "SP500Btn"
        '
        'TCostsBtn
        '
        Me.TCostsBtn.Label = "T-Costs"
        Me.TCostsBtn.Name = "TCostsBtn"
        '
        'ParametersBtn
        '
        Me.ParametersBtn.Label = "Parameters"
        Me.ParametersBtn.Name = "ParametersBtn"
        '
        'FinChartsBtn
        '
        Me.FinChartsBtn.Label = "Fin Charts"
        Me.FinChartsBtn.Name = "FinChartsBtn"
        '
        'Control
        '
        Me.Control.Items.Add(Me.QuitBtn)
        Me.Control.Label = "Control"
        Me.Control.Name = "Control"
        '
        'QuitBtn
        '
        Me.QuitBtn.Label = "Quit"
        Me.QuitBtn.Name = "QuitBtn"
        '
        'ModeGroup
        '
        Me.ModeGroup.Items.Add(Me.ManualTBtn)
        Me.ModeGroup.Items.Add(Me.SyncTBtn)
        Me.ModeGroup.Items.Add(Me.RoboTraderTBtn)
        Me.ModeGroup.Items.Add(Me.SimulationTBtn)
        Me.ModeGroup.Items.Add(Me.StepSimBtn)
        Me.ModeGroup.Label = "Mode"
        Me.ModeGroup.Name = "ModeGroup"
        '
        'ManualTBtn
        '
        Me.ManualTBtn.Label = "Manual"
        Me.ManualTBtn.Name = "ManualTBtn"
        '
        'SyncTBtn
        '
        Me.SyncTBtn.Label = "Sync"
        Me.SyncTBtn.Name = "SyncTBtn"
        '
        'RoboTraderTBtn
        '
        Me.RoboTraderTBtn.Label = "RoboTrader"
        Me.RoboTraderTBtn.Name = "RoboTraderTBtn"
        '
        'SimulationTBtn
        '
        Me.SimulationTBtn.Label = "Simulation"
        Me.SimulationTBtn.Name = "SimulationTBtn"
        '
        'StepSimBtn
        '
        Me.StepSimBtn.Label = "StepSim"
        Me.StepSimBtn.Name = "StepSimBtn"
        '
        'RibbonST
        '
        Me.Name = "RibbonST"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Database.ResumeLayout(False)
        Me.Database.PerformLayout()
        Me.Dashboard.ResumeLayout(False)
        Me.Dashboard.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.BI.ResumeLayout(False)
        Me.BI.PerformLayout()
        Me.Control.ResumeLayout(False)
        Me.Control.PerformLayout()
        Me.ModeGroup.ResumeLayout(False)
        Me.ModeGroup.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Database As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AlphaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents BetaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents GammaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Dashboard As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents BI As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DashboardBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents StockMktBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OptionsMktBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SP500Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Control As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents QuitBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents InitialPositionsButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AcquiredPositionsButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TransactionsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ConfirmationTicketButton As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TCostsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ParametersBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ResetPortfolioBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditControlBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UploadPortfolioBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FinChartsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ModeGroup As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ManualTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents SyncTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents RoboTraderTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents SimulationTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents StepSimBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As RibbonST
        Get
            Return Me.GetRibbon(Of RibbonST)()
        End Get
    End Property
End Class
