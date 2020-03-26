VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: CUSTOMER DETAILS :."
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
   Icon            =   "frmCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCustomer.frx":0ECA
   ScaleHeight     =   9030
   ScaleWidth      =   9180
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
      TabIndex        =   24
      Top             =   6000
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
      TabIndex        =   21
      Text            =   "txtSearch"
      Top             =   6000
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
      TabIndex        =   23
      Top             =   6000
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
      ItemData        =   "frmCustomer.frx":4B4D7
      Left            =   3720
      List            =   "frmCustomer.frx":4B4F9
      Sorted          =   -1  'True
      TabIndex        =   22
      Text            =   "Name"
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtTBM 
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
      TabIndex        =   10
      Text            =   "txtTBM"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtName 
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
      Text            =   "txtName"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtAddress 
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
      TabIndex        =   5
      Text            =   "txtAddress"
      Top             =   1200
      Width           =   6615
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
      TabIndex        =   16
      Top             =   3960
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
      TabIndex        =   25
      Top             =   3960
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
      Left            =   6240
      TabIndex        =   20
      Top             =   4320
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
      Left            =   3120
      TabIndex        =   15
      Top             =   4320
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
      Top             =   3960
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
      TabIndex        =   1
      Text            =   "txtCID"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtPhone 
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
      Text            =   "txtPhone"
      Top             =   2280
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
      Height          =   795
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "frmCustomer.frx":4B56B
      Top             =   4920
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
      TabIndex        =   19
      Top             =   3960
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
      TabIndex        =   18
      Top             =   3960
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
      Left            =   4680
      TabIndex        =   17
      Top             =   4320
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
      Left            =   1560
      TabIndex        =   14
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtCNIC 
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
      TabIndex        =   4
      Text            =   "txtCNIC"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtOCP 
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
      TabIndex        =   6
      Text            =   "txtOCP"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtMobileNo 
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
      TabIndex        =   8
      Text            =   "txtMobileNo"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox txtOtherNo 
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
      Text            =   "txtOtherNo"
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtDue 
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
      TabIndex        =   11
      Text            =   "txtDue"
      Top             =   3360
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   26
      Top             =   6480
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16744576
      DefColWidth     =   93
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
      TabIndex        =   13
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
      Left            =   2280
      TabIndex        =   27
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   8880
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label lblTBM 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   39
      Top             =   3360
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
      TabIndex        =   38
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblPhone 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone #"
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
      TabIndex        =   37
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      TabIndex        =   36
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
      TabIndex        =   35
      Top             =   4920
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
      TabIndex        =   34
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblAddress 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   9120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   8880
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label lblCNIC 
      BackStyle       =   0  'Transparent
      Caption         =   "CNIC #"
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
   Begin VB.Label lblOcp 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
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
      TabIndex        =   31
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblMobile 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile #"
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
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblOtherNo 
      BackStyle       =   0  'Transparent
      Caption         =   "Other #"
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
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblDue 
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
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   360
      X2              =   8880
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00404040&
      X1              =   360
      X2              =   8880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00404040&
      X1              =   360
      X2              =   8880
      Y1              =   3840
      Y2              =   3840
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sup_ID, sql, returnString As String
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock As Boolean

Private Sub Form_Load()
    Connect
    GetDate
    
    SQLString = "SELECT * FROM Customer ORDER BY Customer_ID"
    ShowCustomerData (SQLString)
    ShowCustomerGrid (SQLString)
    
    ClearFields
    
    Normalize
    txtSearch.Text = ""

'For Int TextBoxes
    Dim tmp1, tmp2, tmp3, tmp4 As Long
    tmp1 = SetWindowLong(txtTBM.hwnd, GWL_STYLE, GetWindowLong(txtTBM.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp2 = SetWindowLong(txtDue.hwnd, GWL_STYLE, GetWindowLong(txtDue.hwnd, GWL_STYLE) Or ES_NUMBER)
End Sub

Private Sub cmdNew_Click()
    EnterNewCustomer
End Sub

Private Sub cmdAdd_Click()
    'Checking Fields for Records
    If (txtCID.Text = "" Or txtCID.Text = " ") Then
        MsgBox "Please provide a Customer ID !!!", vbOKOnly, "Information Required"
        txtCID.SetFocus
        Exit Sub
    End If
    If (txtName.Text = "" Or txtName.Text = " ") Then
        MsgBox "Please provide Customer Name !!!", vbOKOnly, "Information Required"
        txtName.SetFocus
        Exit Sub
    End If
    If (txtCNIC.Text = "" Or txtCNIC.Text = " ") Then
        MsgBox "Please provide CNIC# of the Customer !!!", vbOKOnly, "Information Required"
        txtCNIC.SetFocus
        Exit Sub
    End If
    If (txtAddress.Text = "" Or txtAddress.Text = " ") Then
        'MsgBox "Please provide Address for " + txtName.Text + " !!!", vbOKOnly, "Information Required"
        'txtAddress.SetFocus
        txtAddress.Text = "-"
        Exit Sub
    End If
    If (txtPhone.Text = "" Or txtPhone.Text = " ") Then
        MsgBox "Please provide a Phone Number for " + txtName.Text + " !!!", vbOKOnly, "Information Required"
        txtPhone.SetFocus
        Exit Sub
    End If
    If (txtOCP.Text = "" Or txtOCP.Text = " ") Then txtOCP.Text = "-"
    If (txtMobileNo.Text = "" Or txtMobileNo.Text = " ") Then txtMobileNo.Text = "-"
    If (txtOtherNo.Text = "" Or txtOtherNo.Text = " ") Then txtOtherNo.Text = "-"
    If (txtTBM.Text = "" Or txtTBM.Text = " ") Then txtTBM.Text = "0"
    If (txtDue.Text = "" Or txtDue.Text = " ") Then txtDue.Text = "0"
    If (txtR.Text = "") Then txtR.Text = "-"

    'Updating Database
    If DupCheck("SELECT * from Customer WHERE Customer_ID='" & txtCID.Text & "'") = True Then
        MsgBox "Customer ID Already Exists !!! ", , "General Error"
    Else
        sql = "INSERT INTO Customer VALUES('" & txtCID & "','" & txtDate & "','" & txtName & "','" & txtCNIC & "','" & txtAddress & "','" & txtOCP & "','" & txtPhone & "','" & txtMobileNo & "','" & txtOtherNo & "'," & txtTBM & "," & txtDue & ",'" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
    End If
        
    Normalize
    cmdNew.SetFocus
End Sub

Private Sub cmdEdit_Click()
    SetFields (True)
    txtName.SetFocus
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Customer SET Name='" & txtName.Text & "',CNIC_No='" & txtCNIC.Text & "',Address='" & txtAddress.Text & "',Occupation='" & txtOCP.Text & "',Phone_No='" & txtPhone.Text & "',Mobile_No='" & txtMobileNo.Text & "',Other_No='" & txtOtherNo.Text & "',Total_Bills_Amount=" & txtTBM.Text & ",Total_Due=" & txtDue.Text & ",Remarks='" & txtR.Text & "' Where Customer_ID='" & txtCID.Text & "'"
    conn.Execute sql
    ShowCustomerData (SQLString)
    Set DataGrid1.DataSource = RsCustomerGrid
    ShowCustomerGrid ("SELECT * FROM Customer ORDER BY Name")
    DataGrid1.Row = Rx

    Normalize
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
    Dim sqlInvDel, sqlSaleDel, sqlSale, sqlCA, sqlC, InNo As String
        
    If MsgBox("This will DELETE Complete Data of the current Customer from Database[Invoices & Cusotmer Accounts]. ARE YOU SURE?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    'Getting & Deleting Sale Details info
    GetInfo ("SELECT Invoice_No as 'Info' FROM Sales WHERE Customer_ID='" & txtCID.Text & "'")
    InNo = returnString
    sqlInvDel = "DELETE FROM Invoice WHERE Invoice_No='" & InNo & "'"
'    MsgBox "SQL IS " & sqlPODet
    conn.Execute sqlInvDel
    
    'Deleting Customer Account info
    sqlCA = "DELETE FROM Customer_Account WHERE Customer_ID='" & txtCID.Text & "'"
'    MsgBox "SQL IS " & sqlSA
    conn.Execute sqlCA
    
    sqlSaleDel = "DELETE FROM Sales WHERE Invoice_No='" & InNo & "'"
'    MsgBox "SQL IS " & sqlPODet
    conn.Execute sqlSaleDel
        
    'Deleting Receivings info
    sqlSale = "DELETE FROM Sales WHERE Customer_ID='" & txtCID.Text & "'"
'    MsgBox "SQL IS " & sqlR
    conn.Execute sqlSale
               
    'Deleting Customer info
    sqlC = "DELETE FROM Customer Where Customer_ID='" & txtCID.Text & "'"
'    MsgBox "SQL IS " & sqlS
    conn.Execute sqlC
       
    ClearFields
    Normalize
    cmdRDB_Click
End Sub
Private Function GetInfo(SQLStr As String)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    rs.Open SQLStr, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    
        returnString = rs!Info
'        MsgBox returnString

    rs.Close
    Set rs = Nothing
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    ClearFields
    SQLString = "SELECT * FROM Customer ORDER BY Customer_ID"
    Rx = 0
    ShowCustomerData (SQLString)
    ShowCustomerGrid (SQLString)
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowCustomerData ("SELECT * FROM Customer ORDER BY Customer_ID")
    ShowCustomerGrid ("SELECT * FROM Customer ORDER BY Customer_ID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowCustomerData ("SELECT * FROM Customer ORDER BY Customer_ID")
    ShowCustomerGrid ("SELECT * FROM Customer ORDER BY Customer_ID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowCustomerData ("SELECT * FROM Customer ORDER BY Customer_ID")
    ShowCustomerGrid ("SELECT * FROM Customer ORDER BY Customer_ID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowCustomerData ("SELECT * FROM Customer ORDER BY Customer_ID")
    ShowCustomerGrid ("SELECT * FROM Customer ORDER BY Customer_ID")
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
    
    If (ST.Text = "Total_Bills_Amount" Or ST.Text = "Total_Due" Or ST.Text = "Amount_Paid") Then
        SQLString = "SELECT * FROM Customer WHERE " + ST.Text + "=" & txtSearch
    Else
        SQLString = "SELECT * FROM Customer WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    End If
    
    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RsCustomerGrid = New ADODB.Recordset
    RsCustomerGrid.CursorLocation = adUseClient
    RsCustomerGrid.CursorType = adOpenStatic
    RsCustomerGrid.LockType = adLockReadOnly
    RsCustomerGrid.Open SQLString, conn
    Set DataGrid1.DataSource = RsCustomerGrid
      
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        
        MsgBox "Record Not Found !!!", vbInformation, ""
        txtSearch.SetFocus
        SendKeys "{Home}+{End}"
        cmdRDB_Click
        Exit Sub
    End If
    If IsNull(rs!Customer_ID) Then
        ClearFields
    Else
       
    txtCID.Text = rs!Customer_ID
    txtDate.Text = Format(rs!Dated, "YYYY-MM-DD")
    txtName.Text = rs!Name
    txtCNIC.Text = rs!CNIC_No
    txtOCP.Text = rs!Occupation
    txtAddress.Text = rs!Address
    txtPhone.Text = rs!Phone_No
    txtMobileNo.Text = rs!Mobile_No
    txtOtherNo.Text = rs!Other_No
    txtTBM.Text = rs!Total_Bills_Amount
    txtDue.Text = rs!Due
    txtR.Text = rs!Remarks
    
    End If
    rs.Close
    Set rs = Nothing

End Sub

Private Sub ClearFields()
    txtCID.Text = ""
    txtDate.Text = ""
    txtName.Text = ""
    txtCNIC.Text = ""
    txtAddress.Text = ""
    txtOCP.Text = ""
    txtPhone.Text = ""
    txtMobileNo.Text = ""
    txtOtherNo.Text = ""
    txtTBM.Text = ""
    txtDue.Text = ""
    txtR.Text = ""
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    txtName.Enabled = TextFieldLock
    txtCNIC.Enabled = TextFieldLock
    txtAddress.Enabled = TextFieldLock
    txtOCP.Enabled = TextFieldLock
    txtPhone.Enabled = TextFieldLock
    txtMobileNo.Enabled = TextFieldLock
    txtOtherNo.Enabled = TextFieldLock
    txtTBM.Enabled = TextFieldLock
    txtDue.Enabled = TextFieldLock
    txtR.Enabled = TextFieldLock
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

Private Sub Normalize()
    SetFields (False)
    SetButtons (True)
    cmdNew.Visible = True
    cmdEdit.Visible = True
    cmdDelete.Enabled = True
    
    txtSearch.Enabled = True
    ST.Enabled = True
    cmdRDB_Click
End Sub

Public Sub EnterNewCustomer()
    
    SetButtons (False)
    SetFields (True)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdNew.Visible = False
    cmdCancel.Enabled = True
    cmdAdd.Enabled = True
    ClearFields
    GenerateID
    GetDate
    txtDate.Text = DateToday
    txtName.SetFocus
    txtTBM.Text = "0"
    txtDue.Text = "0"
    
End Sub
Private Sub GenerateID()
    txtCID.Text = "C" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Sub ST_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic

    'SQL = "SELECT * FROM Customer WHERE Customer_ID='" & chkID & "'"
    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtCID = rs!Customer_ID Then
        DupCheck = True
    Else
        DupCheck = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Private Sub ST_LostFocus()
    If ST.Text = "Date" Then
        txtSearch.ToolTipText = "Date Format YYYY-MM-DD"
        txtSearch.Text = "2007-12-30"
    Else
        Exit Sub
    End If
End Sub

Private Sub txtAddress_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtCNIC_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtDue_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtMobileNo_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtOCP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtOtherNo_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtPhone_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtName_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtTBM_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSearch_Click
    End If
End Sub

