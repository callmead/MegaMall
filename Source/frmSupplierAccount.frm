VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSupplierAccount 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: SUPPLIER ACCOUNTS :."
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9150
   Icon            =   "frmSupplierAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmSupplierAccount.frx":0ECA
   ScaleHeight     =   7380
   ScaleWidth      =   9150
   Begin VB.CommandButton cmdSelectPO 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      TabIndex        =   6
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdSelectSup 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   720
      Width           =   375
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
      ItemData        =   "frmSupplierAccount.frx":4B4D7
      Left            =   6600
      List            =   "frmSupplierAccount.frx":4B4E4
      Sorted          =   -1  'True
      TabIndex        =   8
      Text            =   "PM"
      Top             =   1440
      Width           =   2295
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
      TabIndex        =   22
      Top             =   4320
      Width           =   1455
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
      TabIndex        =   19
      Text            =   "txtSearch"
      Top             =   4320
      Width           =   3375
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
      TabIndex        =   21
      Top             =   4320
      Width           =   1575
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
      ItemData        =   "frmSupplierAccount.frx":4B503
      Left            =   3720
      List            =   "frmSupplierAccount.frx":4B516
      Sorted          =   -1  'True
      TabIndex        =   20
      Text            =   "Total_Amount"
      Top             =   4320
      Width           =   1935
   End
   Begin VB.TextBox txtSID 
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
      Text            =   "txtSID"
      Top             =   720
      Width           =   1815
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
      Left            =   3840
      TabIndex        =   14
      Top             =   3480
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
      Left            =   2280
      TabIndex        =   24
      Top             =   3480
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
      Left            =   6120
      TabIndex        =   18
      Top             =   3840
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
      Left            =   3000
      TabIndex        =   23
      Top             =   3840
      Width           =   1455
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
      Left            =   720
      TabIndex        =   0
      Top             =   3480
      Width           =   1455
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
      Top             =   1440
      Width           =   2295
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
      Height          =   675
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmSupplierAccount.frx":4B555
      Top             =   2520
      Width           =   6615
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
      Left            =   6960
      TabIndex        =   17
      Top             =   3480
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
      Left            =   5400
      TabIndex        =   16
      Top             =   3480
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
      Left            =   4560
      TabIndex        =   15
      Top             =   3840
      Width           =   1455
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
      Left            =   1440
      TabIndex        =   13
      Top             =   3840
      Width           =   1455
   End
   Begin VB.TextBox txtPO 
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
      Text            =   "txtPO"
      Top             =   720
      Width           =   1815
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
      Top             =   1920
      Width           =   2295
   End
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
      Top             =   1920
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   25
      Top             =   4800
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
      Left            =   720
      TabIndex        =   12
      Top             =   3480
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
      Left            =   2280
      TabIndex        =   26
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   7440
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
      TabIndex        =   35
      Top             =   240
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
      TabIndex        =   34
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label lblSupplierID 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID"
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
      TabIndex        =   33
      Top             =   720
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
      TabIndex        =   32
      Top             =   2520
      Width           =   1815
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
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblPO 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
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
      TabIndex        =   30
      Top             =   720
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
      TabIndex        =   29
      Top             =   1440
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
      TabIndex        =   28
      Top             =   1920
      Width           =   1815
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
      TabIndex        =   27
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   2400
      Y2              =   2400
   End
End
Attribute VB_Name = "frmSupplierAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sup_ID, sql As String
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock, AddAmount, ChangeAmounts, MinusAmounts As Boolean
Dim ExistingSupplierAmount, NewSupplierAmount, CurrentAmount, CurrentDue, ExistingDueAmount, NewDueAmount As Double

Private Sub Form_Load()

    Connect
    SQLString = "SELECT * FROM Supplier_Account ORDER BY TID"
    ShowSupplierAccountData (SQLString)
    ShowSupplierAccountGrid (SQLString)
    
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

Private Sub cmdSelectPO_Click()
    If txtSID.Text = "" Then
        MsgBox "Please Select a Supplier First", vbOKOnly, "Information Required"
        Exit Sub
    Else
        ParentForm = "frmSupplierAccountPO"
        GridSQLString = "Select Purchase_Order.PO_No,Purchase_Order.Date,Supplier.Company,Purchase_Order.Delivery_Date from Purchase_Order,Supplier WHERE Purchase_Order.Supplier_ID='" & txtSID.Text & "' AND Supplier.Supplier_ID='" & txtSID.Text & "' ORDER BY Supplier.Name"
        SelectedField = 0
        frmDataSelect.Show vbModal
    End If
End Sub

Private Sub cmdSelectSup_Click()
    ParentForm = "frmSupplierAccountSID"
    GridSQLString = "Select Supplier_ID,Name,Company from Supplier ORDER BY Supplier_ID"
    SelectedField = 0
    frmDataSelect.Show vbModal
End Sub

Private Sub cmdNew_Click()
    EnterNewSupplierAccount
    txtTA.Text = "0"
    txtPA.Text = "0"
    txtDA.Text = "0"
End Sub

Private Sub cmdAdd_Click()
    
    'Checking Fields for Records
    If (txtTID.Text = "" Or txtTID.Text = " ") Then
        'MsgBox "Please provide a Transaction ID !!!", vbOKOnly, "Information Required"
        txtTID.SetFocus
        Exit Sub
    End If
    If (txtSID.Text = "" Or txtSID.Text = " ") Then
        MsgBox "Please Select a Supplier !!!", vbOKOnly, "Information Required"
        txtSID.SetFocus
        Exit Sub
    End If
    If (txtPO.Text = "" Or txtPO.Text = " ") Then
        MsgBox "Please Select a Purchase Order !!!", vbOKOnly, "Information Required"
        'txtPO.SetFocus
        Exit Sub
    End If
    If (txtTA.Text = "" Or txtTA.Text = " ") Then
        MsgBox "Please Provide Total Amount for the Purchase Order " + txtPO.Text + " !!!", vbOKOnly, "Information Required"
        txtPO.SetFocus
        Exit Sub
    End If
    If (PM.Text = "PM" Or PM.Text = "") Then
        MsgBox "Please select Payment Mode !!!", vbOKOnly, "Information Required"
        PM.SetFocus
        Exit Sub
    End If
    If (txtPA.Text = "" Or txtPA.Text = " ") Then
        MsgBox "Please Provide Paid Amount for the Purchase Order " + txtPO.Text + " !!!", vbOKOnly, "Information Required"
        txtPA.SetFocus
        Exit Sub
    End If

    If (txtDA.Text = "" Or txtDA.Text = " ") Then
        'MsgBox "Please Provide Total Amount for the Purchase Order " + txtPO.Text + " !!!", vbOKOnly, "Information Required"
        'txtPO.SetFocus
        txtDA.Text = "0"
        Exit Sub
    End If
    If (txtR.Text = "") Then txtR.Text = "-"
    
    'Updating Database
    'Updating Database
    If DupCheck("SELECT * from Supplier_Account WHERE Supplier_ID='" & txtSID.Text & "' AND PO_No='" & txtPO.Text & "'") = True Then
        MsgBox "Purchase Order Already Exists !!! ", vbInformation, "Information"
        Normalize
        Exit Sub
    Else
    
        sql = "INSERT INTO Supplier_Account VALUES('" & txtTID & "','" & txtSID & "','" & txtDate & "','" & txtPO & "'," & txtTA & ",'" & PM & "'," & txtPA & "," & txtDA & ",'" & txtR & "')"
        'MsgBox SQL
        conn.Execute sql
        
        AddAmount = True
        UpdateSupplierAccounts
            
        Normalize
        cmdRDB_Click
        cmdNew.SetFocus
        Exit Sub
        
    End If
    
End Sub

Private Sub UpdateSupplierAccounts()
    
    NewSupplierAmount = 0
    NewDueAmount = 0

    Set rsTmp = New ADODB.Recordset
    Query = "SELECT Total_Bills_Amount,Total_Due FROM Supplier WHERE Supplier_ID='" & txtSID.Text & "'"

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

    ExistingSupplierAmount = Val(rsTmp!Total_Bills_Amount)
    ExistingDueAmount = Val(rsTmp!Total_Due)
    
    If AddAmount = True Then
        NewSupplierAmount = Val(txtTA.Text) + ExistingSupplierAmount
        NewDueAmount = Val(txtDA.Text) + ExistingDueAmount
    
    ElseIf MinusAmounts = True Then
        NewSupplierAmount = ExistingSupplierAmount - Val(txtTA.Text)
        NewDueAmount = ExistingDueAmount - Val(txtDA.Text)
    
'    ElseIf (AddAmount = False And ChangeAmounts = False) Then
'        NewSupplierAmount = ExistingSupplierAmount - Val(txtTA.Text)
'        NewDueAmount = ExistingDueAmount - Val(txtDA.Text)
    
    ElseIf ChangeAmounts = True Then
        NewSupplierAmount = (ExistingSupplierAmount - CurrentAmount) + Val(txtTA.Text)
        NewDueAmount = (ExistingDueAmount - CurrentDue) + Val(txtDA.Text)
    End If
    
    Query = "UPDATE Supplier SET Total_Bills_Amount=" & NewSupplierAmount & ",Total_Due=" & NewDueAmount & " WHERE Supplier_ID='" & txtSID.Text & "'"
    'MsgBox Query
    conn.Execute Query
    
    AddAmount = False
    ChangeAmounts = False
    
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
    SetButtons (False)
    cmdClose.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Supplier_Account SET Date='" & txtDate.Text & "',Supplier_ID='" & txtSID.Text & "',PO_No='" & txtPO.Text & "',Total_Amount=" & txtTA.Text & ",Payment_Mode='" & PM.Text & "',Paid_Amount=" & txtPA.Text & ",Due_Amount=" & txtDA.Text & ",Remarks='" & txtR.Text & "' Where TID='" & txtTID.Text & "'"
    conn.Execute sql
    
    UpdateSupplierAccounts
    
    ShowSupplierAccountData (SQLString)
    Set DataGrid1.DataSource = RsSuppGrid
    ShowSupplierAccountGrid ("SELECT * FROM Supplier_Account ORDER BY TID")
    DataGrid1.Row = Rx

    Normalize
    
    cmdRDB_Click
    
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
    
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorType = adOpenDynamic
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorLocation = adUseServer
    sql = "DELETE FROM Supplier_Account Where TID='" & txtTID.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Set rsTemp = Nothing
        Exit Sub
    End If
    rsTemp.Open sql, conn, adOpenKeyset
    Set rsTemp = Nothing
    
    AddAmount = False
    MinusAmounts = True
    UpdateSupplierAccounts
    
    Rx = Rx - 1
    ClearFields
    Normalize
    cmdRDB_Click
    Set DataGrid1.DataSource = RsSuppGrid
    ShowSupplierAccountGrid ("SELECT * FROM Supplier_Account ORDER BY TID")
    If (Rx <> 0) Then DataGrid1.Row = Rx
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    ClearFields
    SQLString = "SELECT * FROM Supplier_Account ORDER BY TID"
    Rx = 0
    ShowSupplierAccountData (SQLString)
    ShowSupplierAccountGrid (SQLString)
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowSupplierAccountData ("SELECT * FROM Supplier_Account ORDER BY TID")
    ShowSupplierAccountGrid ("SELECT * FROM Supplier_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowSupplierAccountData ("SELECT * FROM Supplier_Account ORDER BY TID")
    ShowSupplierAccountGrid ("SELECT * FROM Supplier_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowSupplierAccountData ("SELECT * FROM Supplier_Account ORDER BY TID")
    ShowSupplierAccountGrid ("SELECT * FROM Supplier_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowSupplierAccountData ("SELECT * FROM Supplier_Account ORDER BY TID")
    ShowSupplierAccountGrid ("SELECT * FROM Supplier_Account ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdSearch_Click()
On Error GoTo Err
If (txtSearch.Text = "" Or txtSearch.Text = " ") Then
    MsgBox "Search what?", vbExclamation, "General Error"
    txtSearch.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    
    If ST.Text = "PO_No" Or ST.Text = "Supplier_ID" Then
        SQLString = "SELECT * FROM Supplier_Account WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    Else
        SQLString = "SELECT * FROM Supplier_Account WHERE " + ST.Text + "=" & txtSearch
    End If
    'MsgBox SQLString
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
    If IsNull(rs!TID) Then
        ClearFields
    Else

    txtTID.Text = rs!TID
    txtSID.Text = rs!Supplier_ID
    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
    txtPO.Text = rs!PO_No
    txtTA.Text = rs!Total_Amount
    PM.Text = rs!Payment_Mode
    txtPA.Text = rs!Paid_Amount
    txtDA.Text = rs!Due_Amount
    txtR.Text = rs!Remarks

    End If
    rs.Close
    Set rs = Nothing
Err:
    MsgBox "Invalid Search", vbInformation
    txtSearch.SetFocus
End Sub

Private Sub PM_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub txtDA_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPA_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPA_KeyDown(KeyCode As Integer, Shift As Integer)
    txtDA.Text = Val(txtTA.Text) - Val(txtPA.Text)
End Sub

Public Sub EnterNewSupplierAccount()

    txtTA.Text = "0"
    txtPA.Text = "0"
    txtDA.Text = "0"
    SetButtons (False)
    SetFields (True)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdNew.Visible = False
    cmdCancel.Enabled = True
    cmdSelectSup.Enabled = True
    cmdSelectPO.Enabled = True
    cmdAdd.Enabled = True
    ClearFields
    GenerateID
    GetDate
    txtDate.Text = DateToday
    cmdSelectSup.SetFocus
    
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
    txtSID.Text = ""
    txtPO.Text = ""
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

    cmdSelectSup.Enabled = False
    cmdSelectPO.Enabled = False
    
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

Private Sub txtPA_KeyUp(KeyCode As Integer, Shift As Integer)
txtDA.Text = Val(txtTA.Text) - Val(txtPA.Text)
End Sub

Private Sub txtR_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtTA_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtTA_KeyUp(KeyCode As Integer, Shift As Integer)
txtDA.Text = Val(txtTA.Text) - Val(txtPA.Text)
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSearch_Click
    End If
End Sub

Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    'SQL = "SELECT * from Supplier WHERE Supplier_ID='" & chkID & "'"
    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtPO.Text = rs!PO_No Then
        DupCheck = True
    Else
        DupCheck = False
    End If
    rs.Close
    Set rs = Nothing
End Function

