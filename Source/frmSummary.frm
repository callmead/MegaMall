VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSummary 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: DATABASE SUMMARY :."
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   Icon            =   "frmSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmSummary.frx":0ECA
   ScaleHeight     =   7605
   ScaleWidth      =   12510
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   480
      Top             =   6240
   End
   Begin VB.Timer TmPB 
      Interval        =   140
      Left            =   480
      Top             =   5760
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   6960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   4680
      X2              =   4680
      Y1              =   1560
      Y2              =   4920
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   5040
      X2              =   8760
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Today's Profit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblTP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6840
      TabIndex        =   22
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Statistics..."
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   44.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1185
      Left            =   360
      TabIndex        =   21
      Top             =   120
      Width           =   8685
   End
   Begin VB.Label lblpb 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblSDMS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      ToolTipText     =   "Shows the number of products for which the quantity in stock is in Minus!"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblSDRP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      ToolTipText     =   "Shows the number of products which needs to be Re-Ordered!"
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Minus Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      ToolTipText     =   "Shows the number of products for which the quantity in stock is in Minus!"
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "ReORder Products"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      ToolTipText     =   "Shows the number of products which needs to be Re-Ordered!"
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   4080
      Width           =   2415
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00808080&
      X1              =   600
      X2              =   4320
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   12240
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label lblTSA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      ToolTipText     =   "Shows the Total Amount of Sale that has been made so far"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label lblTEA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6840
      TabIndex        =   12
      ToolTipText     =   "Shows the Total Amount of expenses that has been made so far"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblNP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      ToolTipText     =   "Shows the total profit that has been made so far, excluding expenses"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label lblTSD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      ToolTipText     =   "Shows the total amount that is involved in supplier accounts"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblTSBA 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   6960
      TabIndex        =   9
      ToolTipText     =   "Shows the total amount that is involved in supplier accounts"
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sale"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   720
      TabIndex        =   8
      ToolTipText     =   "Shows the Total Amount of Sale that has been made so far"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   600
      X2              =   4320
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Expense"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5160
      TabIndex        =   6
      ToolTipText     =   "Shows the Total Amount of expenses that has been made so far"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   5040
      X2              =   8760
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Profit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Shows the total profit that has been made so far, excluding expenses"
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Profit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   600
      X2              =   4320
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Due"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   "Shows the total amount that is involved in supplier accounts"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Bills Amount"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5160
      TabIndex        =   1
      ToolTipText     =   "Shows the total amount that is involved in supplier accounts"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Accounts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   5040
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   5040
      X2              =   8760
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   240
      X2              =   240
      Y1              =   120
      Y2              =   7440
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    lblTSBA.Caption = "Loading..."
    lblTSD.Caption = "Loading..."
    lblNP.Caption = "Loading..."
    lblTEA.Caption = "Loading..."
    lblTSA.Caption = "Loading..."
    lblSDMS.Caption = "Loading..."
    lblSDRP.Caption = "Loading..."
    lblTP.Caption = "Loading..."
End Sub

Private Sub GetSupplierAccountInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Total_Bills_Amount) as 'TBA',SUM(Total_Due) as 'TD' FROM Supplier;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If rs!TBA <> "" Then
        lblTSBA.Caption = rs!TBA
        lblTSD.Caption = rs!TD
    Else
        lblTSBA.Caption = "0"
        lblTSD.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTotalExpenseInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Amount) as 'Sum' FROM Expenditure;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    
    If rs!Sum <> "" Then
        lblTEA.Caption = rs!Sum
    Else
        lblTEA.Caption = "0"
    End If
    
    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTotalSaleInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT SUM(Grand_Total) as 'Sum' FROM Sales;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!Sum <> "" Then
        lblTSA.Caption = rs!Sum
    Else
        lblTSA.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetProfitInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic

    sql = "SELECT SUM(Profit) as 'Sum' FROM Sales;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!Sum <> "" Then
        lblNP.Caption = rs!Sum
    Else
        lblNP.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetTodayProfitInfo()
    GetDate
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic

    sql = "SELECT SUM(Profit) AS 'Sum' FROM Sales WHERE Date='" & DateToday & "';"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!Sum <> "" Then
        lblTP.Caption = rs!Sum
    Else
        lblTP.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetStockROLInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT COUNT(*) as 'No' FROM Stock WHERE Stock_In_Hand<ReOrder_Level;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!No <> "" Then
        lblSDRP.Caption = rs!No
    Else
        lblSDRP.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub GetStockMinusInfo()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT COUNT(*) as No FROM Stock WHERE Stock_In_Hand<=0;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    If rs!No <> "" Then
        lblSDMS.Caption = rs!No
    Else
        lblSDMS.Caption = "0"
    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Timer2_Timer()
    lblpb.Caption = "Data Loaded"
    PB1.Visible = False
    
    GetSupplierAccountInfo
    GetTotalSaleInfo
    GetTotalExpenseInfo
    GetProfitInfo
    GetTodayProfitInfo
    GetStockROLInfo
    GetStockMinusInfo
    
End Sub

Private Sub TmPB_Timer()
    PB1.Value = PB1.Value + 5
    If (PB1.Value = PB1.Max) Then
        TmPB.Enabled = False
    End If
End Sub
