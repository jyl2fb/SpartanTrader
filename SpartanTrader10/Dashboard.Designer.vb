﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On



'''
<Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(1),  _
 Global.System.Security.Permissions.PermissionSetAttribute(Global.System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")>  _
Partial Public NotInheritable Class Dashboard
    Inherits Microsoft.Office.Tools.Excel.WorksheetBase
    
    Friend WithEvents DateLine As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents AcquiredLO As Microsoft.Office.Tools.Excel.ListObject
    
    Friend WithEvents InitialLO As Microsoft.Office.Tools.Excel.ListObject
    
    Friend WithEvents TeamIDCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TEChart As Microsoft.Office.Tools.Excel.Chart
    
    Friend WithEvents SecondsCell As Microsoft.Office.Tools.Excel.NamedRange
    
    Friend WithEvents TELO As Microsoft.Office.Tools.Excel.ListObject
    
    Friend WithEvents TickersCBox As Microsoft.Office.Tools.Excel.Controls.ComboBox
    
    Friend WithEvents StockQtyBox As Microsoft.Office.Tools.Excel.Controls.TextBox
    
    Friend WithEvents BuyStockBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellStockBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellShortBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents CashDivBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ExecuteStockTransactionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SymbolsCBox As Microsoft.Office.Tools.Excel.Controls.ComboBox
    
    Friend WithEvents OptionQtyBox As Microsoft.Office.Tools.Excel.Controls.TextBox
    
    Friend WithEvents BuyOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents SellShortOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ExerciseOptionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ExecuteOptionTransactionBtn As Microsoft.Office.Tools.Excel.Controls.Button
    
    Friend WithEvents ManualExecutionLBox As Microsoft.Office.Tools.Excel.Controls.ListBox
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Public Sub New(ByVal factory As Global.Microsoft.Office.Tools.Excel.Factory, ByVal serviceProvider As Global.System.IServiceProvider)
        MyBase.New(factory, serviceProvider, "Sheet2", "Sheet2")
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub Initialize()
        MyBase.Initialize
        Globals.Dashboard = Me
        Global.System.Windows.Forms.Application.EnableVisualStyles
        Me.InitializeCachedData
        Me.InitializeControls
        Me.InitializeComponents
        Me.InitializeData
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub FinishInitialization()
        Me.OnStartup
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub InitializeDataBindings()
        Me.BeginInitialization
        Me.BindToData
        Me.EndInitialization
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeCachedData()
        If (Me.DataHost Is Nothing) Then
            Return
        End If
        If Me.DataHost.IsCacheInitialized Then
            Me.DataHost.FillCachedData(Me)
        End If
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeData()
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub BindToData()
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Sub StartCaching(ByVal MemberName As String)
        Me.DataHost.StartCaching(Me, MemberName)
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Sub StopCaching(ByVal MemberName As String)
        Me.DataHost.StopCaching(Me, MemberName)
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Function IsCached(ByVal MemberName As String) As Boolean
        Return Me.DataHost.IsCached(Me, MemberName)
    End Function
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub BeginInitialization()
        Me.BeginInit
        Me.DateLine.BeginInit
        Me.AcquiredLO.BeginInit
        Me.InitialLO.BeginInit
        Me.TeamIDCell.BeginInit
        Me.TEChart.BeginInit
        Me.SecondsCell.BeginInit
        Me.TELO.BeginInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub EndInitialization()
        Me.TELO.EndInit
        Me.SecondsCell.EndInit
        Me.TEChart.EndInit
        Me.TeamIDCell.EndInit
        Me.InitialLO.EndInit
        Me.AcquiredLO.EndInit
        Me.DateLine.EndInit
        Me.EndInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeControls()
        Me.DateLine = Globals.Factory.CreateNamedRange(Nothing, Nothing, "DateLine", "DateLine", Me)
        Me.AcquiredLO = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet2:AcquiredLO", "AcquiredLO", Me)
        Me.InitialLO = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet2:InitialLO", "InitialLO", Me)
        Me.TeamIDCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "TeamIDCell", "TeamIDCell", Me)
        Me.TEChart = Globals.Factory.CreateChart(Nothing, Nothing, "Sheet2:Chart 1", "TEChart", Me)
        Me.SecondsCell = Globals.Factory.CreateNamedRange(Nothing, Nothing, "SecondsCell", "SecondsCell", Me)
        Me.TELO = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet2:TELO", "TELO", Me)
        Me.TickersCBox = New Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "1B9CFAB0817D4614E371BA261C76ECA0C6DA41", "1B9CFAB0817D4614E371BA261C76ECA0C6DA41", Me, "TickersCBox")
        Me.StockQtyBox = New Microsoft.Office.Tools.Excel.Controls.TextBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "22AE137DF231F5242212B9B62458710E26B702", "22AE137DF231F5242212B9B62458710E26B702", Me, "StockQtyBox")
        Me.BuyStockBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "3A737A9013AE31342143AAC332FEB1D70166A3", "3A737A9013AE31342143AAC332FEB1D70166A3", Me, "BuyStockBtn")
        Me.SellStockBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "475CE45464FC5C44EC1491CD403B214C11B264", "475CE45464FC5C44EC1491CD403B214C11B264", Me, "SellStockBtn")
        Me.SellShortBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "84E0CECCD8B11684AA8882828679A4649A6238", "84E0CECCD8B11684AA8882828679A4649A6238", Me, "SellShortBtn")
        Me.CashDivBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "9DF148E8B9C34E949E39BB769715A2AAD5B239", "9DF148E8B9C34E949E39BB769715A2AAD5B239", Me, "CashDivBtn")
        Me.ExecuteStockTransactionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "13F3441AA143AE14CCF181E016606E72B3BC61", "13F3441AA143AE14CCF181E016606E72B3BC61", Me, "ExecuteStockTransactionBtn")
        Me.SymbolsCBox = New Microsoft.Office.Tools.Excel.Controls.ComboBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "1B41DF30A103CB1433018E231B7416FB5ED861", "1B41DF30A103CB1433018E231B7416FB5ED861", Me, "SymbolsCBox")
        Me.OptionQtyBox = New Microsoft.Office.Tools.Excel.Controls.TextBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "1818A3A8A1BA671426E1B206184EE4BFD9C941", "1818A3A8A1BA671426E1B206184EE4BFD9C941", Me, "OptionQtyBox")
        Me.BuyOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1454BFB5E1D645141091B73211A51887DA7C71", "1454BFB5E1D645141091B73211A51887DA7C71", Me, "BuyOptionBtn")
        Me.SellOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1BBCC80AD153C4144421923510F58B2602C731", "1BBCC80AD153C4144421923510F58B2602C731", Me, "SellOptionBtn")
        Me.SellShortOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1A18C20641A72C1401B18C5C1A90A56B754661", "1A18C20641A72C1401B18C5C1A90A56B754661", Me, "SellShortOptionBtn")
        Me.ExerciseOptionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "12486F748194FB14E891ADBE158DFF48F524B1", "12486F748194FB14E891ADBE158DFF48F524B1", Me, "ExerciseOptionBtn")
        Me.ExecuteOptionTransactionBtn = New Microsoft.Office.Tools.Excel.Controls.Button(Globals.Factory, Me.ItemProvider, Me.HostContext, "1E2A52455132D914AEE1AEF2144BD19AEA9CE1", "1E2A52455132D914AEE1AEF2144BD19AEA9CE1", Me, "ExecuteOptionTransactionBtn")
        Me.ManualExecutionLBox = New Microsoft.Office.Tools.Excel.Controls.ListBox(Globals.Factory, Me.ItemProvider, Me.HostContext, "13A3AB2DF190D214B5019698171B1D64DD8681", "13A3AB2DF190D214B5019698171B1D64DD8681", Me, "ManualExecutionLBox")
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeComponents()
        '
        'TickersCBox
        '
        Me.TickersCBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(255,Byte),Integer), CType(CType(192,Byte),Integer), CType(CType(128,Byte),Integer))
        Me.TickersCBox.Name = "TickersCBox"
        Me.TickersCBox.Text = "Choose a Ticker"
        '
        'StockQtyBox
        '
        Me.StockQtyBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(255,Byte),Integer), CType(CType(192,Byte),Integer), CType(CType(128,Byte),Integer))
        Me.StockQtyBox.Name = "StockQtyBox"
        Me.StockQtyBox.Text = "0"
        '
        'BuyStockBtn
        '
        Me.BuyStockBtn.BackColor = System.Drawing.SystemColors.Control
        Me.BuyStockBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BuyStockBtn.Name = "BuyStockBtn"
        Me.BuyStockBtn.Text = "Buy"
        Me.BuyStockBtn.UseVisualStyleBackColor = false
        '
        'SellStockBtn
        '
        Me.SellStockBtn.BackColor = System.Drawing.SystemColors.Control
        Me.SellStockBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SellStockBtn.Name = "SellStockBtn"
        Me.SellStockBtn.Text = "Sell"
        Me.SellStockBtn.UseVisualStyleBackColor = false
        '
        'SellShortBtn
        '
        Me.SellShortBtn.BackColor = System.Drawing.SystemColors.Control
        Me.SellShortBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SellShortBtn.Name = "SellShortBtn"
        Me.SellShortBtn.Text = "Sell Short"
        Me.SellShortBtn.UseVisualStyleBackColor = false
        '
        'CashDivBtn
        '
        Me.CashDivBtn.BackColor = System.Drawing.SystemColors.Control
        Me.CashDivBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.CashDivBtn.Name = "CashDivBtn"
        Me.CashDivBtn.Text = "CashDiv"
        Me.CashDivBtn.UseVisualStyleBackColor = false
        '
        'ExecuteStockTransactionBtn
        '
        Me.ExecuteStockTransactionBtn.BackColor = System.Drawing.SystemColors.Control
        Me.ExecuteStockTransactionBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExecuteStockTransactionBtn.Name = "ExecuteStockTransactionBtn"
        Me.ExecuteStockTransactionBtn.Text = "Execute Transaction"
        Me.ExecuteStockTransactionBtn.UseVisualStyleBackColor = false
        '
        'SymbolsCBox
        '
        Me.SymbolsCBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(255,Byte),Integer), CType(CType(192,Byte),Integer), CType(CType(128,Byte),Integer))
        Me.SymbolsCBox.Name = "SymbolsCBox"
        Me.SymbolsCBox.Text = "Choose a Symbol"
        '
        'OptionQtyBox
        '
        Me.OptionQtyBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(255,Byte),Integer), CType(CType(192,Byte),Integer), CType(CType(128,Byte),Integer))
        Me.OptionQtyBox.Name = "OptionQtyBox"
        Me.OptionQtyBox.Text = "0"
        '
        'BuyOptionBtn
        '
        Me.BuyOptionBtn.BackColor = System.Drawing.SystemColors.Control
        Me.BuyOptionBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.BuyOptionBtn.Name = "BuyOptionBtn"
        Me.BuyOptionBtn.Text = "Buy"
        Me.BuyOptionBtn.UseVisualStyleBackColor = false
        '
        'SellOptionBtn
        '
        Me.SellOptionBtn.BackColor = System.Drawing.SystemColors.Control
        Me.SellOptionBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SellOptionBtn.Name = "SellOptionBtn"
        Me.SellOptionBtn.Text = "Sell"
        Me.SellOptionBtn.UseVisualStyleBackColor = false
        '
        'SellShortOptionBtn
        '
        Me.SellShortOptionBtn.BackColor = System.Drawing.SystemColors.Control
        Me.SellShortOptionBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.SellShortOptionBtn.Name = "SellShortOptionBtn"
        Me.SellShortOptionBtn.Text = "Sell Short"
        Me.SellShortOptionBtn.UseVisualStyleBackColor = false
        '
        'ExerciseOptionBtn
        '
        Me.ExerciseOptionBtn.BackColor = System.Drawing.SystemColors.Control
        Me.ExerciseOptionBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExerciseOptionBtn.Name = "ExerciseOptionBtn"
        Me.ExerciseOptionBtn.Text = "Exercise"
        Me.ExerciseOptionBtn.UseVisualStyleBackColor = false
        '
        'ExecuteOptionTransactionBtn
        '
        Me.ExecuteOptionTransactionBtn.BackColor = System.Drawing.SystemColors.Control
        Me.ExecuteOptionTransactionBtn.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExecuteOptionTransactionBtn.Name = "ExecuteOptionTransactionBtn"
        Me.ExecuteOptionTransactionBtn.Text = "Execute Transaction"
        Me.ExecuteOptionTransactionBtn.UseVisualStyleBackColor = false
        '
        'ManualExecutionLBox
        '
        Me.ManualExecutionLBox.BackColor = System.Drawing.Color.Black
        Me.ManualExecutionLBox.Font = New System.Drawing.Font("Arial", 14!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0,Byte))
        Me.ManualExecutionLBox.ForeColor = System.Drawing.Color.FromArgb(CType(CType(255,Byte),Integer), CType(CType(128,Byte),Integer), CType(CType(0,Byte),Integer))
        Me.ManualExecutionLBox.ItemHeight = 22
        Me.ManualExecutionLBox.Items.AddRange(New Object() {"Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade", "Trade"})
        Me.ManualExecutionLBox.Name = "ManualExecutionLBox"
        '
        'DateLine
        '
        Me.DateLine.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'AcquiredLO
        '
        Me.AcquiredLO.AutoSetDataBoundColumnHeaders = true
        Me.AcquiredLO.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'InitialLO
        '
        Me.InitialLO.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TeamIDCell
        '
        Me.TeamIDCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TEChart
        '
        Me.TEChart.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'SecondsCell
        '
        Me.SecondsCell.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'TELO
        '
        Me.TELO.AutoSetDataBoundColumnHeaders = true
        Me.TELO.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'Dashboard
        '
        Me.TickersCBox.BindingContext = Me.BindingContext
        Me.StockQtyBox.BindingContext = Me.BindingContext
        Me.BuyStockBtn.BindingContext = Me.BindingContext
        Me.SellStockBtn.BindingContext = Me.BindingContext
        Me.SellShortBtn.BindingContext = Me.BindingContext
        Me.CashDivBtn.BindingContext = Me.BindingContext
        Me.ExecuteStockTransactionBtn.BindingContext = Me.BindingContext
        Me.SymbolsCBox.BindingContext = Me.BindingContext
        Me.OptionQtyBox.BindingContext = Me.BindingContext
        Me.BuyOptionBtn.BindingContext = Me.BindingContext
        Me.SellOptionBtn.BindingContext = Me.BindingContext
        Me.SellShortOptionBtn.BindingContext = Me.BindingContext
        Me.ExerciseOptionBtn.BindingContext = Me.BindingContext
        Me.ExecuteOptionTransactionBtn.BindingContext = Me.BindingContext
        Me.ManualExecutionLBox.BindingContext = Me.BindingContext
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Private Function NeedsFill(ByVal MemberName As String) As Boolean
        Return Me.DataHost.NeedsFill(Me, MemberName)
    End Function
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub OnShutdown()
        Me.TELO.Dispose
        Me.SecondsCell.Dispose
        Me.TEChart.Dispose
        Me.TeamIDCell.Dispose
        Me.InitialLO.Dispose
        Me.AcquiredLO.Dispose
        Me.DateLine.Dispose
        MyBase.OnShutdown
    End Sub
End Class

Partial Friend NotInheritable Class Globals
    
    Private Shared _Dashboard As Dashboard
    
    Friend Shared Property Dashboard() As Dashboard
        Get
            Return _Dashboard
        End Get
        Set
            If (_Dashboard Is Nothing) Then
                _Dashboard = value
            Else
                Throw New System.NotSupportedException()
            End If
        End Set
    End Property
End Class
