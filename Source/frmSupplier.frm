VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSupplier 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: SUPPLIER DETAILS :."
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmSupplier.frx":0ECA
   ScaleHeight     =   8940
   ScaleWidth      =   9165
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
      TabIndex        =   12
      Text            =   "txtDue"
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox txtFaxNo 
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
      Text            =   "txtFaxNo"
      Top             =   2760
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
   Begin VB.TextBox txtCP 
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
      Text            =   "txtCP"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtCompany 
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
      Text            =   "txtCompany"
      Top             =   720
      Width           =   2295
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
      TabIndex        =   35
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
      Left            =   4680
      TabIndex        =   21
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
      Left            =   5520
      TabIndex        =   22
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
      Left            =   7080
      TabIndex        =   23
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
      Height          =   675
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "frmSupplier.frx":4B4D7
      Top             =   4800
      Width           =   6615
   End
   Begin VB.TextBox txtOfficeNo 
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
      Text            =   "txtOfficeNo"
      Top             =   2280
      Width           =   2295
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
      TabIndex        =   1
      Text            =   "txtSID"
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
      Left            =   840
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
      Left            =   3120
      TabIndex        =   26
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
      Left            =   6240
      TabIndex        =   24
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
      Left            =   2400
      TabIndex        =   25
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
      Left            =   3960
      TabIndex        =   20
      Top             =   3960
      Width           =   1455
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
      TabIndex        =   6
      Text            =   "txtAddress"
      Top             =   1680
      Width           =   6615
   End
   Begin VB.TextBox txtSName 
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
      Text            =   "txtSName"
      Top             =   720
      Width           =   2295
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
      TabIndex        =   11
      Text            =   "txtTBM"
      Top             =   3360
      Width           =   2295
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
      ItemData        =   "frmSupplier.frx":4B4DC
      Left            =   3720
      List            =   "frmSupplier.frx":4B4F2
      Sorted          =   -1  'True
      TabIndex        =   16
      Text            =   "Name"
      Top             =   5880
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
      TabIndex        =   17
      Top             =   5880
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
      TabIndex        =   15
      Text            =   "txtSearch"
      Top             =   5880
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
      TabIndex        =   18
      Top             =   5880
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   19
      Top             =   6360
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
      Left            =   840
      TabIndex        =   14
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
      Left            =   2400
      TabIndex        =   34
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   2160
      Y2              =   2160
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
      TabIndex        =   41
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblFax 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax #"
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
      TabIndex        =   40
      Top             =   2760
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
      TabIndex        =   39
      Top             =   2760
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
      TabIndex        =   38
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblCP 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person"
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
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblCompany 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
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
      TabIndex        =   36
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   9120
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
      Top             =   1680
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
      TabIndex        =   32
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
      TabIndex        =   31
      Top             =   4800
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
      TabIndex        =   30
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblOffice 
      BackStyle       =   0  'Transparent
      Caption         =   "Office #"
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
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblSID 
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
      TabIndex        =   28
      Top             =   240
      Width           =   1815
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
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   360
      TabIndex        =   27
      Top             =   3360
      Width           =   1815
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sup_ID, sql As String
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock As Boolean
Private returnString As String

Private Sub Form_Load()

    Connect
    GetDate
    
    SQLString = "SELECT * FROM Supplier ORDER BY Company"
    ShowSupplierData (SQLString)
    ShowSupplierGrid (SQLString)
    
    ClearFields
    
    Normalize
    txtSearch.Text = ""
    
    'For Int TextBoxes
    Dim tmp1, tmp2 As Long
    tmp1 = SetWindowLong(txtTBM.hwnd, GWL_STYLE, GetWindowLong(txtTBM.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp1 = SetWindowLong(txtDue.hwnd, GWL_STYLE, GetWindowLong(txtDue.hwnd, GWL_STYLE) Or ES_NUMBER)

End Sub

Private Sub cmdNew_Click()
    EnterNewSupplier
    txtTBM.Text = "0"
    txtDue.Text = "0"
End Sub

Private Sub cmdAdd_Click()
    
    'Checking Fields for Records
    If (txtSID.Text = "" Or txtSID.Text = " ") Then
        MsgBox "Please provide a Supplier ID !!!", vbOKOnly, "Information Required"
        txtSID.SetFocus
        Exit Sub
    End If
    If (txtSName.Text = "" Or txtSName.Text = " ") Then
        MsgBox "Please provide Supplier Name !!!", vbOKOnly, "Information Required"
        txtSName.SetFocus
        Exit Sub
    End If
    If (txtAddress.Text = "" Or txtAddress.Text = " ") Then
        'MsgBox "Please provide Address for " + txtSName.Text + " !!!", vbOKOnly, "Information Required"
        'txtAddress.SetFocus
        txtAddress.Text = "-"
        Exit Sub
    End If
    If (txtOfficeNo.Text = "" Or txtOfficeNo.Text = " ") Then
        MsgBox "Please provide Office Number for " + txtSName.Text + " !!!", vbOKOnly, "Information Required"
        txtOfficeNo.SetFocus
        Exit Sub
    End If
    If (txtMobileNo.Text = "" Or txtMobileNo.Text = " ") Then txtMobileNo.Text = "-"
    If (txtOtherNo.Text = "" Or txtOtherNo.Text = " ") Then txtOtherNo.Text = "-"
    If (txtFaxNo.Text = "" Or txtFaxNo.Text = " ") Then txtFaxNo.Text = "-"
    If (txtTBM.Text = "" Or txtTBM.Text = " ") Then txtTBM.Text = "0"
    If (txtDue.Text = "" Or txtDue.Text = " ") Then txtDue.Text = "0"
    If (txtR.Text = "") Then txtR.Text = "-"
    
    'Updating Database
    If DupCheck("SELECT * from Supplier WHERE Supplier_ID='" & txtSID.Text & "'") = True Then
        MsgBox "Supplier ID Already Exists !!! ", , "General Error"
    Else
        sql = "INSERT INTO Supplier(Supplier_ID,Date,Name,Company,Contact_Person,Address,Office_No,Mobile_No,Other_No,Fax_No,Total_Bills_Amount,Total_Due,Remarks) values('" & txtSID & "','" & txtDate & "','" & txtSName & "','" & txtCompany & "','" & txtCP & "','" & txtAddress & "','" & txtOfficeNo & "','" & txtMobileNo & "','" & txtOtherNo & "','" & txtFaxNo & "'," & txtTBM & "," & txtDue & ",'" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
    End If
        
    Normalize
    cmdRDB_Click
    cmdNew.SetFocus
    Exit Sub
    
End Sub

Private Sub cmdEdit_Click()
    SetFields (True)
    txtSName.SetFocus
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Supplier SET Date='" & txtDate.Text & "',Name='" & txtSName.Text & "',Contact_Person='" & txtCP.Text & "',Address='" & txtAddress.Text & "',Office_No='" & txtOfficeNo.Text & "',Mobile_No='" & txtMobileNo.Text & "',Other_No='" & txtOtherNo.Text & "',Fax_No='" & txtFaxNo.Text & "',Total_Bills_Amount=" & txtTBM.Text & ",Total_Due=" & txtDue.Text & ",Remarks='" & txtR.Text & "' Where Supplier_ID='" & txtSID.Text & "'"
    conn.Execute sql
    ShowSupplierData (SQLString)
    Set DataGrid1.DataSource = RsSuppGrid
    ShowSupplierGrid ("SELECT * FROM Supplier ORDER BY Supplier_ID")
    DataGrid1.Row = Rx

    Normalize
    cmdRDB_Click
    
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
    Dim sqlPODet, sqlR, sqlPO, sqlSA, sqlS, PONo As String
        
    If MsgBox("This will DELETE Complete Data of the current Supplier from Database[Receivings, Purchase Orders & Supplier Accounts]. ARE YOU SURE?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        'Set rsTemp = Nothing
        Exit Sub
    End If
    
    'Getting & Deleting PODetails info
    GetInfo ("SELECT PO_No as 'Info' FROM Purchase_Order WHERE Supplier_ID='" & txtSID.Text & "'")
    PONo = returnString
    sqlPODet = "DELETE FROM PO_Details WHERE PO_No='" & PONo & "'"
'    MsgBox "SQL IS " & sqlPODet
    conn.Execute sqlPODet
    
    'Deleting Receivings info
    sqlR = "DELETE FROM Receivings WHERE PO_No='" & PONo & "'"
'    MsgBox "SQL IS " & sqlR
    conn.Execute sqlR
    
    'Deleting Supplier Account info
    sqlSA = "DELETE FROM Supplier_Account WHERE Supplier_ID='" & txtSID.Text & "'"
'    MsgBox "SQL IS " & sqlSA
    conn.Execute sqlSA
    
    'Deleting Purchase_Order info
    sqlPO = "DELETE FROM Purchase_Order WHERE Supplier_ID='" & txtSID.Text & "'"
'    MsgBox "SQL IS " & sqlPO
    conn.Execute sqlPO
        
    'Deleting Supplier Account info
    sqlS = "DELETE FROM Supplier Where Supplier_ID='" & txtSID.Text & "'"
'    MsgBox "SQL IS " & sqlS
    conn.Execute sqlS
       
    Rx = Rx - 1
    ClearFields
    Normalize
    cmdRDB_Click
    Set DataGrid1.DataSource = RsSuppGrid
    ShowSupplierGrid ("SELECT * FROM Supplier ORDER BY Supplier_ID")
    If (Rx <> 0) Then DataGrid1.Row = Rx
    
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
    SQLString = "SELECT * FROM Supplier ORDER BY Company"
    Rx = 0
    ShowSupplierData (SQLString)
    ShowSupplierGrid (SQLString)
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowSupplierData ("SELECT * FROM Supplier ORDER BY Company")
    ShowSupplierGrid ("SELECT * FROM Supplier ORDER BY Company")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowSupplierData ("SELECT * FROM Supplier ORDER BY Company")
    ShowSupplierGrid ("SELECT * FROM Supplier ORDER BY Company")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowSupplierData ("SELECT * FROM Supplier ORDER BY Company")
    ShowSupplierGrid ("SELECT * FROM Supplier ORDER BY Company")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowSupplierData ("SELECT * FROM Supplier ORDER BY Company")
    ShowSupplierGrid ("SELECT * FROM Supplier ORDER BY Company")
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
    
    If (ST.Text = "Total_Bills_Amount" Or ST.Text = "Total_Due") Then
        SQLString = "SELECT * FROM Supplier WHERE " + ST.Text + "=" & txtSearch
    Else
        SQLString = "SELECT * FROM Supplier WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    End If
    'MsgBox SQLString
    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    Set RsSuppGrid = New ADODB.Recordset
    RsSuppGrid.CursorLocation = adUseClient
    RsSuppGrid.CursorType = adOpenStatic
    RsSuppGrid.LockType = adLockReadOnly
    RsSuppGrid.Open SQLString, conn
    Set DataGrid1.DataSource = RsSuppGrid
      
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        
        MsgBox "Record Not Found !!!", vbInformation, ""
        txtSearch.SetFocus
        SendKeys "{Home}+{End}"
        cmdRDB_Click
        Exit Sub
    End If
    If IsNull(rs!Supplier_ID) Then
        ClearFields
    Else
       
    txtSID.Text = rs!Supplier_ID
    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
    txtSName.Text = rs!Name
    txtCompany.Text = rs!Company
    txtCP.Text = rs!Contact_Person
    txtAddress.Text = rs!Address
    txtOfficeNo.Text = rs!Office_No
    txtMobileNo.Text = rs!Mobile_No
    txtOtherNo.Text = rs!Other_No
    txtFaxNo.Text = rs!Fax_No
    txtTBM.Text = rs!Total_Bills_Amount
    txtDue.Text = rs!Total_Due
    txtR.Text = rs!Remarks
    
    End If
    rs.Close
    Set rs = Nothing
Err:
    MsgBox "Invalid Search", vbInformation
    txtSearch.SetFocus
End Sub

Private Sub ClearFields()
    txtSID.Text = ""
    txtDate.Text = ""
    txtSName.Text = ""
    txtCompany.Text = ""
    txtCP.Text = ""
    txtAddress.Text = ""
    txtOfficeNo.Text = ""
    txtMobileNo.Text = ""
    txtOtherNo.Text = ""
    txtFaxNo.Text = ""
    txtTBM.Text = ""
    txtDue.Text = ""
    txtR.Text = ""
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    txtSName.Enabled = TextFieldLock
    txtCompany.Enabled = TextFieldLock
    txtCP.Enabled = TextFieldLock
    txtAddress.Enabled = TextFieldLock
    txtOfficeNo.Enabled = TextFieldLock
    txtMobileNo.Enabled = TextFieldLock
    txtOtherNo.Enabled = TextFieldLock
    txtFaxNo.Enabled = TextFieldLock
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

Public Sub EnterNewSupplier()
    
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
    txtSName.SetFocus
    
End Sub
Private Sub GenerateID()
    txtSID.Text = "S" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
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

    'SQL = "SELECT * FROM Supplier WHERE Supplier_ID='" & chkID & "'"
    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtSID.Text = rs!Supplier_ID Then
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

Private Sub txtCompany_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtCP_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDue_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtFaxNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtMobileNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtOfficeNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtOtherNo_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSName_GotFocus()
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

