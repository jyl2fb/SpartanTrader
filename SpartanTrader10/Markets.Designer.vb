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
<Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(5),  _
 Global.System.Security.Permissions.PermissionSetAttribute(Global.System.Security.Permissions.SecurityAction.Demand, Name:="FullTrust")>  _
Partial Public NotInheritable Class Markets
    Inherits Microsoft.Office.Tools.Excel.WorksheetBase
    
    Friend WithEvents OptionsLO As Microsoft.Office.Tools.Excel.ListObject
    
    Friend WithEvents SP500LO As Microsoft.Office.Tools.Excel.ListObject
    
    Friend WithEvents StockMarketLO As Microsoft.Office.Tools.Excel.ListObject
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Public Sub New(ByVal factory As Global.Microsoft.Office.Tools.Excel.Factory, ByVal serviceProvider As Global.System.IServiceProvider)
        MyBase.New(factory, serviceProvider, "Sheet1", "Sheet1")
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Protected Overrides Sub Initialize()
        MyBase.Initialize
        Globals.Markets = Me
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
        Me.OptionsLO.BeginInit
        Me.SP500LO.BeginInit
        Me.StockMarketLO.BeginInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub EndInitialization()
        Me.StockMarketLO.EndInit
        Me.SP500LO.EndInit
        Me.OptionsLO.EndInit
        Me.EndInit
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeControls()
        Me.OptionsLO = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet1:OptionsLO", "OptionsLO", Me)
        Me.SP500LO = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet1:SP500LO", "SP500LO", Me)
        Me.StockMarketLO = Globals.Factory.CreateListObject(Nothing, Nothing, "Sheet1:StockMarketLO", "StockMarketLO", Me)
    End Sub
    
    '''
    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "16.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Never)>  _
    Private Sub InitializeComponents()
        '
        'OptionsLO
        '
        Me.OptionsLO.AutoSetDataBoundColumnHeaders = true
        Me.OptionsLO.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'SP500LO
        '
        Me.SP500LO.AutoSetDataBoundColumnHeaders = true
        Me.SP500LO.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'StockMarketLO
        '
        Me.StockMarketLO.AutoSetDataBoundColumnHeaders = true
        Me.StockMarketLO.DefaultDataSourceUpdateMode = System.Windows.Forms.DataSourceUpdateMode.Never
        '
        'Markets
        '
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
        Me.StockMarketLO.Dispose
        Me.SP500LO.Dispose
        Me.OptionsLO.Dispose
        MyBase.OnShutdown
    End Sub
End Class

Partial Friend NotInheritable Class Globals
    
    Private Shared _Markets As Markets
    
    Friend Shared Property Markets() As Markets
        Get
            Return _Markets
        End Get
        Set
            If (_Markets Is Nothing) Then
                _Markets = value
            Else
                Throw New System.NotSupportedException()
            End If
        End Set
    End Property
End Class
