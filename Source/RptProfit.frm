VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form RptProfit 
   Caption         =   ":: PROFIT DETAILS :."
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   Icon            =   "RptProfit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "RptProfit.frx":0ECA
   ScaleHeight     =   6585
   ScaleWidth      =   8415
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer 
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      lastProp        =   500
      _cx             =   14843
      _cy             =   11668
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "RptProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim rs As ADODB.Recordset
    Dim crystal As CRAXDRT.Application
    Dim report As CRAXDRT.report

Private Sub Form_Load()
    Connect
    
    CRViewer.DisplayBorder = False
    CRViewer.DisplayTabs = False
    CRViewer.EnableDrillDown = False
    CRViewer.EnableRefreshButton = False
    
    Set rs = New ADODB.Recordset
    rs.Open RptSql, conn, adOpenStatic, adLockReadOnly
    
    Set crystal = New CRAXDRT.Application
    Set report = crystal.OpenReport(App.Path & RptPathIs)
    
    report.DiscardSavedData
    report.ReportTitle = "Sales"
    report.Database.SetDataSource rs
    CRViewer.ReportSource = report
    CRViewer.ViewReport
    
    Do While CRViewer.IsBusy
        DoEvents
    Loop
    CRViewer.Zoom 100
    rs.Close
    Set rs = Nothing
    Set crystal = Nothing
    Set report = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer.Top = 0
    CRViewer.Left = 0
    CRViewer.Height = ScaleHeight
    CRViewer.Width = ScaleWidth
End Sub
