VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomerAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: CUSTOMER ACCOUNTS :."
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmCustomerAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCustomerAccount.frx":0ECA
   ScaleHeight     =   8130
   ScaleWidth      =   9165
   Begin VB.TextBox txtDA 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   6600
      TabIndex        =   10
      Text            =   "txtDA"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtPA 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   2280
      TabIndex        =   9
      Text            =   "txtPA"
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtInv 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   6600
      TabIndex        =   5
      Text            =   "txtInv"
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdMF 
      Caption         =   "Move &First"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4320
      TabIndex        =   16
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdP 
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdN 
      Caption         =   "Ne&xt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6720
      TabIndex        =   18
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtR 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   795
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmCustomerAccount.frx":4B4D7
      Top             =   2880
      Width           =   6615
   End
   Begin VB.TextBox txtTA 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   2280
      TabIndex        =   7
      Text            =   "txtTA"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox txtTID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   2280
      TabIndex        =   1
      Text            =   "txtTID"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   6600
      TabIndex        =   2
      Text            =   "txtDate"
      ToolTipText     =   "Date Format yyyy-MM-dd"
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdML 
      Caption         =   "Move &Last"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   19
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   24
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdRDB 
      Caption         =   "Re&fresh DB"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3600
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox txtCID 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   2280
      TabIndex        =   3
      Text            =   "txtCID"
      Top             =   720
      Width           =   1815
   End
   Begin VB.ComboBox ST 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      ItemData        =   "frmCustomerAccount.frx":4B4DC
      Left            =   3720
      List            =   "frmCustomerAccount.frx":4B4E9
      Sorted          =   -1  'True
      TabIndex        =   21
      Text            =   "Total_Amount"
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5760
      TabIndex        =   22
      Top             =   5040
      Width           =   1575
   End
   Begin VB.TextBox txtSearch 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      Left            =   240
      TabIndex        =   20
      Text            =   "txtSearch"
      Top             =   5040
      Width           =   3375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7440
      TabIndex        =   23
      Top             =   5040
      Width           =   1455
   End
   Begin VB.ComboBox PM 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   360
      ItemData        =   "frmCustomerAccount.frx":4B514
      Left            =   6600
      List            =   "frmCustomerAccount.frx":4B521
      Sorted          =   -1  'True
      TabIndex        =   8
      Text            =   "PM"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.CommandButton cmdSelectCid 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdSelectIn 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   25
      Top             =   5520
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16744576
      DefColWidth     =   80
      Enabled         =   -1  'True
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   26
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   8880
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblDA 
      BackStyle       =   0  'Transparent
      Caption         =   "Due Amount"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   35
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblPA 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Amount"
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
      Height          =   375
      Left            =   360
      TabIndex        =   34
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblPM 
      BackStyle       =   0  'Transparent
      Caption         =   "Payment Mode"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   33
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblInv 
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   32
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   8520
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Height          =   375
      Left            =   4920
      TabIndex        =   31
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Height          =   375
      Left            =   360
      TabIndex        =   30
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblCID 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer ID"
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
      Height          =   375
      Left            =   360
      TabIndex        =   29
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblOffice 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblTID 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction ID"
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
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmCustomerAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cus_ID, sql As String
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock, AddAmount, ChangeAmounts, MinusAmounts As Boolean
Dim ExistingCustomerAmount, NewCustomerAmount, CurrentAmount, CurrentDue, ExistingDueAmount, NewDueAmount As Double

Private Sub Form_Load()

    Connect
    SQLString = "SELECT * FROM Customer_Account ORDER BY TID"
    ShowCustomerAccountData (SQLString)
    ShowCustomerAccountGrid (SQLString)
    
    ClearFields
    
    AddAmount = False
    ChangeAmounts = False
    MinusAmounts = False
    
    Normalize
    txtSearch.Text = ""
    
'For Int TextBoxes
    Dim tmp1, tmp2, tmp3 As Long
    tmp1 = SetWindowLong(txtTA.hwnd, GWL_STYLE, GetWindowLong(txtTA.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp2 = SetWindowLong(txtPA.hwnd, GWL_STYLE, GetWindowLong(txtPA.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp3 = SetWindowLong(txtDA.hwnd, GWL_STYLE, GetWindowLong(txtDA.hwnd, GWL_STYLE) Or ES_NUMBER)
End Sub

Private Sub cmdSelectCid_Click()
    ParentForm = "frmCustomerAccountCID"
    GridSQLString = "Select Customer_ID,Name,Occupation,Phone_No from Customer ORDER BY Customer_ID"
    SelectedField = 0
    frmDataSelect.Show vbModal
End Sub

Private Sub cmdSelectIn_Click()
    If txtCID.Text = "" Then
        MsgBox "Please Select a Customer First", vbOKOnly, "Information Required"
        Exit Sub
    Else
        ParentForm = "frmCustomerAccountInv"
        GridSQLString = "Select Invoice_No,Date,Grand_Total,Amount_Paid,Amount_Due FROM Sales WHERE Customer_ID='" & txtCID.Text & "' ORDER BY Date"
        SelectedField = 0
        frmDataSelect.Show vbModal
    End If
End Sub

Private Sub cmdNew_Click()
    EnterNewCustomerAccount
    txtTA.Text = "0"
    txtPA.Text = "0"
    txtDA.Text = "0"
End Sub

Private Sub cmdAdd_Click()
    CheckFieldsData
    'Updating Database

    sql = "INSERT INTO Customer_Account VALUES('" & txtTID & "','" & txtCID & "','" & txtDate & "','" & txtInv & "'," & txtTA & ",'" & PM & "'," & txtPA & "," & txtDA & ",'" & txtR & "')"
    'MsgBox SQL
    conn.Execute sql
    
    AddAmount = True
    UpdateCustomerAccounts
        
    Normalize
    cmdRDB_Click
    'cmdNew.SetFocus
    Exit Sub

End Sub

Private Sub UpdateCustomerAccounts()
    
    NewCustomerAmount = 0
    NewDueAmount = 0

    Set rsTmp = New ADODB.Recordset
    Query = "SELECT Total_Bills_Amount,Total_Due FROM Customer WHERE Customer_ID='" & txtCID.Text & "'"

    rsTmp.CursorLocation = adUseClient
    rsTmp.CursorType = adOpenStatic
    rsTmp.LockType = adLockReadOnly
    rsTmp.Open Query, conn
        If rsTmp.EOF = True Then
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Sub
        End If
    xCount = rsTmp.RecordCount
        If Rx > rsTmp.RecordCount - 1 Then
            Rx = 0
        End If
        If Rx < 0 Then
            Rx = rsTmp.RecordCount - 1
        End If
    rsTmp.Move Rx

    ExistingCustomerAmount = Val(rsTmp!Total_Bills_Amount)
    ExistingDueAmount = Val(rsTmp!Total_Due)
    
    If AddAmount = True Then
        NewCustomerAmount = Val(txtTA.Text) + ExistingCustomerAmount
        NewDueAmount = Val(txtDA.Text) + ExistingDueAmount
    
    ElseIf MinusAmounts = True Then
        NewCustomerAmount = ExistingCustomerAmount - Val(txtTA.Text)
        NewDueAmount = ExistingDueAmount - Val(txtDA.Text)
    
    ElseIf ChangeAmounts = True Then
        NewCustomerAmount = (ExistingCustomerAmount - CurrentAmount) + Val(txtTA.Text)
        NewDueAmount = (ExistingDueAmount - CurrentDue) + Val(txtDA.Text)
    End If
    
    Query = "UPDATE Customer SET Total_Bills_Amount=" & NewCustomerAmount & ",Total_Due=" & NewDueAmount & " WHERE Customer_ID='" & txtCID.Text & "'"
    'MsgBox Query
    conn.Execute Query
    
    AddAmount = False
    ChangeAmounts = False
    MinusAmounts = False
    
    rsTmp.Close
    Set rsTmp = Nothing

End Sub

Private Sub cmdEdit_Click()
    SetFields (True)
    
    CurrentAmount = Val(txtTA.Text)
    CurrentDue = Val(txtDA.Text)
    ChangeAmounts = True
    
    txtTA.SetFocus
    SetButtons (True)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdNew.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Customer_Account SET Date='" & txtDate.Text & "',Customer_ID='" & txtCID.Text & "',Invoice_No='" & txtInv.Text & "',Total_Amount=" & txtTA.Text & ",Payment_Mode='" & PM.Text & "',Amount_paid=" & txtPA.Text & ",Amount_due=" & txtDA.Text & ",Remarks='" & txtR.Text & "' Where TID='" & txtTID.Text & "'"
    conn.Execute sql
    
    UpdateCustomerAccounts
    
    ShowCustomerAccountData (SQLString)
    Set DataGrid1.DataSource = RsCustomerGrid
    ShowCustomerAccountGrid ("SELECT * FROM Customer_Account ORDER BY TID")
    DataGrid1.Row = Rx

    ClearFields
    Normalize
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
    
    sql = "DELETE FROM Customer_Account Where TID='" & txtTID.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    conn.Execute sql
    
    AddAmount = False
    MinusAmounts = True
    UpdateCustomerAccounts
    
    ClearFields
    Normalize
    cmdRDB_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    ClearFields
    SQLString = "SELECT * FROM Customer_Account ORDER BY TID"
    Rx = 0
    ShowCustomerAccountData (SQLString)
    ShowCustomerAccountGrid (SQLString)
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowCustomerAccountData ("SELECT * FROM Customer_Account ORDER BY TID")
    ShowCustomerAccountGrid ("SELECT * FROM Customer_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowCustomerAccountData ("SELECT * FROM Customer_Account ORDER BY TID")
    ShowCustomerAccountGrid ("SELECT * FROM Customer_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowCustomerAccountData ("SELECT * FROM Customer_Account ORDER BY TID")
    ShowCustomerAccountGrid ("SELECT * FROM Customer_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowCustomerAccountData ("SELECT * FROM Customer_Account ORDER BY TID")
    ShowCustomerAccountGrid ("SELECT * FROM Customer_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdSearch_Click()
If (txtSearch.Text = "" Or txtSearch.Text = " ") Then
    MsgBox "Search what?", vbExclamation, "General Error"
    txtSearch.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer

    SQLString = "SELECT * FROM Customer_Account WHERE " + ST.Text + "=" & txtSearch
    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText

    Set RsSuppAccountGrid = New ADODB.Recordset
    RsSuppAccountGrid.CursorLocation = adUseClient
    RsSuppAccountGrid.CursorType = adOpenStatic
    RsSuppAccountGrid.LockType = adLockReadOnly
    RsSuppAccountGrid.Open SQLString, conn
    Set DataGrid1.DataSource = RsSuppAccountGrid

    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing

        MsgBox "Record Not Found !!!", vbInformation, ""
        txtSearch.SetFocus
        SendKeys "{Home}+{End}"
        cmdRDB_Click
        Exit Sub
    End If
    If IsNull(rs!SuppID) Then
        ClearFields
    Else

    txtTID.Text = rs!TID
    txtCID.Text = rs!Customer_ID
    txtDate.Text = Format(rs!Dated, "YYYY-MM-DD")
    txtInv.Text = rs!PO_No
    txtTA.Text = rs!Total_Amount
    PM.Text = rs!Payment_Mode
    txtPA.Text = rs!Paid_Amount
    txtDA.Text = rs!Due_Amount
    txtR.Text = rs!Remarks

    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub PM_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtPA_KeyDown(KeyCode As Integer, Shift As Integer)
    txtDA.Text = Val(txtTA.Text) - Val(txtPA.Text)
End Sub

Public Sub EnterNewCustomerAccount()

    txtTA.Text = "0"
    txtPA.Text = "0"
    txtDA.Text = "0"
    SetButtons (False)
    SetFields (True)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdNew.Visible = False
    cmdCancel.Enabled = True
    cmdSelectCid.Enabled = True
    cmdSelectIn.Enabled = True
    cmdAdd.Enabled = True
    ClearFields
    GenerateID
    GetDate
    txtDate.Text = DateToday
    cmdSelectCid.SetFocus
    
End Sub

Private Sub SetButtons(ButtonLock As Boolean)
    cmdNew.Enabled = ButtonLock
    cmdAdd.Enabled = ButtonLock
    cmdEdit.Enabled = ButtonLock
    cmdSave.Enabled = ButtonLock
    cmdCancel.Enabled = ButtonLock
    cmdDelete.Enabled = ButtonLock
    cmdRDB.Enabled = ButtonLock
    cmdMF.Enabled = ButtonLock
    cmdN.Enabled = ButtonLock
    cmdP.Enabled = ButtonLock
    cmdML.Enabled = ButtonLock
    cmdSearch.Enabled = ButtonLock
    cmdClose.Enabled = ButtonLock
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    txtTA.Enabled = TextFieldLock
    PM.Enabled = TextFieldLock
    txtPA.Enabled = TextFieldLock
    txtDA.Enabled = TextFieldLock
    txtR.Enabled = TextFieldLock
End Sub

Private Sub ClearFields()
    txtTID.Text = ""
    txtDate.Text = ""
    txtCID.Text = ""
    txtInv.Text = ""
    txtTA.Text = ""
    PM.Text = ""
    txtPA.Text = ""
    txtDA.Text = ""
    txtR.Text = ""
End Sub

Private Sub Normalize()
    ClearFields
    SetFields (False)
    SetButtons (True)
    cmdNew.Enabled = True
    cmdNew.Visible = True
    cmdEdit.Enabled = True
    cmdEdit.Visible = True
    cmdDelete.Enabled = True

    cmdSelectCid.Enabled = False
    cmdSelectIn.Enabled = False
    
    txtSearch.Enabled = True
    ST.Enabled = True
    cmdRDB_Click
End Sub

Private Sub GenerateID()
    txtTID.Text = "T" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Sub ST_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub ST_LostFocus()
    If ST.Text = "Date" Then
        txtSearch.ToolTipText = "Date Format YYYY-MM-DD"
        txtSearch.Text = "2007-12-30"
    Else
        Exit Sub
    End If
End Sub

Private Sub txtDA_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPA_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtTA_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtTA_KeyUp(KeyCode As Integer, Shift As Integer)
txtDA.Text = Val(txtTA.Text) - Val(txtPA.Text)
End Sub

Private Sub txtPA_KeyUp(KeyCode As Integer, Shift As Integer)
txtDA.Text = Val(txtTA.Text) - Val(txtPA.Text)
End Sub

Private Sub CheckFieldsData()
    'Checking Fields for Records
    If (txtTID.Text = "" Or txtTID.Text = " ") Then
        'MsgBox "Please provide a Transaction ID !!!", vbOKOnly, "Information Required"
        txtTID.SetFocus
        Exit Sub
    End If
    If (txtCID.Text = "" Or txtCID.Text = " ") Then
        MsgBox "Please Select a Customer !!!", vbOKOnly, "Information Required"
        txtCID.SetFocus
        Exit Sub
    End If
    If (txtInv.Text = "" Or txtInv.Text = " ") Then
        MsgBox "Please Select an Invoice !!!", vbOKOnly, "Information Required"
        Exit Sub
    End If
    If (txtTA.Text = "" Or txtTA.Text = " ") Then
        MsgBox "Please Provide Total Amount for the Purchase Order " + txtInv.Text + " !!!", vbOKOnly, "Information Required"
        txtInv.SetFocus
        Exit Sub
    End If
    If (PM.Text = "PM" Or PM.Text = "") Then
        MsgBox "Please select Payment Mode !!!", vbOKOnly, "Information Required"
        PM.SetFocus
        Exit Sub
    End If
    If (txtPA.Text = "" Or txtPA.Text = " ") Then
        MsgBox "Please Provide Paid Amount for the Purchase Order " + txtInv.Text + " !!!", vbOKOnly, "Information Required"
        txtPA.SetFocus
        Exit Sub
    End If

    If (txtDA.Text = "" Or txtDA.Text = " ") Then
        'MsgBox "Please Provide Total Amount for the Purchase Order " + txtInv.Text + " !!!", vbOKOnly, "Information Required"
        'txtInv.SetFocus
        txtDA.Text = "0"
        Exit Sub
    End If
    If (txtR.Text = "") Then txtR.Text = "-"
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSearch_Click
    End If
End Sub

