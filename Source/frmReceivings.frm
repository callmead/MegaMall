VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmReceivings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: RECEIVINGS :."
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10755
   Icon            =   "frmReceivings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmReceivings.frx":0ECA
   ScaleHeight     =   6990
   ScaleWidth      =   10755
   Begin VB.CommandButton cmdAd 
      Enabled         =   0   'False
      Height          =   375
      Left            =   8520
      Picture         =   "frmReceivings.frx":4B4D7
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Add to Cart"
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdSelProd 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox txtPID 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
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
      Left            =   240
      TabIndex        =   6
      Text            =   "txtPID"
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtPricePU 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
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
      Left            =   6720
      TabIndex        =   10
      Text            =   "txtPricePU"
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   4920
      TabIndex        =   9
      Text            =   "txtPrice"
      Top             =   3360
      Width           =   1695
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
      Left            =   2160
      TabIndex        =   3
      Text            =   "txtPO"
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
      Left            =   9000
      TabIndex        =   13
      Top             =   960
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
      Left            =   9000
      TabIndex        =   16
      Top             =   2040
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
      Left            =   9000
      TabIndex        =   17
      Top             =   2400
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
      Left            =   9000
      TabIndex        =   18
      Top             =   2760
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
      Height          =   1035
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmReceivings.frx":4C1A1
      Top             =   1560
      Width           =   6615
   End
   Begin VB.TextBox txtQty 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0"
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
      Left            =   3240
      TabIndex        =   8
      Text            =   "txtQty"
      Top             =   3360
      Width           =   1575
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
      Left            =   2160
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
      Left            =   9000
      TabIndex        =   14
      Top             =   1320
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
      Left            =   9000
      TabIndex        =   19
      Top             =   3120
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
      Left            =   9000
      TabIndex        =   25
      Top             =   600
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
      Left            =   9000
      TabIndex        =   15
      Top             =   1680
      Width           =   1455
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
      ItemData        =   "frmReceivings.frx":4C1A6
      Left            =   4440
      List            =   "frmReceivings.frx":4C1B6
      Sorted          =   -1  'True
      TabIndex        =   21
      Text            =   "Product_ID"
      Top             =   6360
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
      Left            =   6480
      TabIndex        =   22
      Top             =   6360
      Width           =   1935
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
      Top             =   6360
      Width           =   4095
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
      Left            =   8520
      TabIndex        =   23
      Top             =   6360
      Width           =   1935
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
      Left            =   9000
      TabIndex        =   24
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelectPO 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   375
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
      Left            =   9000
      TabIndex        =   0
      Top             =   240
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   26
      Top             =   3840
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16744576
      DefColWidth     =   73
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
   Begin MSFlexGridLib.MSFlexGrid PrdGrid 
      Height          =   2325
      Left            =   240
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3840
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   4101
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColor       =   16744576
      ForeColor       =   16777215
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorSel    =   8421631
      BackColorBkg    =   9081241
      GridColor       =   4210752
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      MousePointer    =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmReceivings.frx":4C1DC
   End
   Begin MSComCtl2.UpDown ScrollBar 
      Height          =   855
      Left            =   10200
      TabIndex        =   36
      Top             =   3840
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   1508
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.Label lblProd 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
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
      Left            =   240
      TabIndex        =   34
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label lblPricePU 
      BackStyle       =   0  'Transparent
      Caption         =   "Price per Unit"
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
      Left            =   6720
      TabIndex        =   33
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Top             =   3000
      Width           =   1815
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
      Left            =   360
      TabIndex        =   31
      Top             =   720
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   10440
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   9120
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
      TabIndex        =   30
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
      TabIndex        =   29
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label lblQty 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Recvd."
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
      Left            =   3240
      TabIndex        =   28
      Top             =   3000
      Width           =   1695
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
Attribute VB_Name = "frmReceivings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock, AddQuantity, ChangeQuantity, MinusQuantity, AddingData As Boolean
Dim sql, Prod_ID, PONo As String
Dim ExistingQuantity, NewQuantity, CurrentQuantity, ppu As Double
Dim nLastKeyAscii As Integer


Private Sub cmdSelProd_Click()
    ParentForm = "frmReceivings1"
    GridSQLString = "SELECT Product_ID,Product,Product_Type,Stock_In_Hand FROM Stock ORDER BY Product_Type"
    SelectedField = 0
    frmDataSelect.Show vbModal
    txtQty.Text = "1"
    txtPrice.Text = "1"
End Sub

Private Sub Form_Load()

    Connect
    SQLString = "SELECT * FROM Receivings ORDER BY TID"
    ShowReceivingsData (SQLString)
    ShowReceivingsGrid (SQLString)
    
    ClearFields
    
    AddQuantity = False
    MinusQuantity = False
    ChangeQuantity = False
    
    GetDate
    GridSet
    
    PrdGrid.Rows = 1
    PrdGrid.Rows = 2
    
    iR = PrdGrid.Rows - 1
    
    Normalize
    txtSearch.Text = ""
    
 'For Int TextBoxes
    Dim tmp1, tmp2, tmp3 As Long
    tmp1 = SetWindowLong(txtQty.hwnd, GWL_STYLE, GetWindowLong(txtQty.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp2 = SetWindowLong(txtPrice.hwnd, GWL_STYLE, GetWindowLong(txtPrice.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp3 = SetWindowLong(txtPricePU.hwnd, GWL_STYLE, GetWindowLong(txtPricePU.hwnd, GWL_STYLE) Or ES_NUMBER)
    
    AddingData = False
End Sub

Private Sub PrdGrid_Click()
    On Error Resume Next
    iR = PrdGrid.Row
    iC = PrdGrid.Col
    
    If iC = 0 Then
        txtTID.Text = PrdGrid.TextMatrix(iR, iC)
        'txtTID.SetFocus
        'SendKeys "{Home}+{End}"
    End If
    If iC = 1 Then
        txtPID.Text = PrdGrid.TextMatrix(iR, iC)
        txtPID.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 2 Then
        txtQty.Text = PrdGrid.TextMatrix(iR, iC)
        txtQty.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 3 Then
        txtPrice.Text = PrdGrid.TextMatrix(iR, iC)
        txtPrice.SetFocus
        SendKeys "{Home}+{End}"
    End If
    If iC = 4 Then
        txtPricePU.Text = PrdGrid.TextMatrix(iR, iC)
        txtPricePU.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

Private Sub ScrollBar_DownClick()
    PrdGrid.Rows = PrdGrid.Rows + 1
End Sub

Private Sub ScrollBar_UpClick()
    If PrdGrid.Rows > 2 Then PrdGrid.Rows = PrdGrid.Rows - 1
    If PrdGrid.Rows = 2 Then PrdGrid.Rows = 1: PrdGrid.Rows = 2
End Sub

Private Sub GridSet()
    With PrdGrid
    .Cols = 5
    .Rows = 2
    .ColWidth(0) = 2500
    .ColWidth(1) = 1100
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    
    .TextMatrix(0, 0) = " Transaction"
    .TextMatrix(0, 1) = " Product"
    .TextMatrix(0, 2) = " Quantity"
    .TextMatrix(0, 3) = " Price"
    .TextMatrix(0, 4) = " Price/Unit"
    End With
End Sub

Private Sub cmdSelectPO_Click()
    ParentForm = "frmReceivings"
    GridSQLString = "SELECT Purchase_Order.PO_No, Purchase_Order.Date, Supplier.Company, Purchase_Order.Delivery_Date FROM Purchase_Order, Supplier WHERE Purchase_Order.Supplier_ID = Supplier.Supplier_ID ORDER BY Purchase_Order.Date,Company"
    SelectedField = 0
    frmDataSelect.Show vbModal
End Sub

Private Sub cmdNew_Click()
    PrdGrid.Rows = 1
    PrdGrid.Rows = 2
    
    iR = PrdGrid.Rows - 1
    
    ClearFields
    EnterNewReceivings
    cmdAd.Enabled = True
    cmdSelectPO.SetFocus
            
    PrdGrid.Visible = True
    ScrollBar.Visible = True
    DataGrid1.Visible = False

End Sub

Private Sub cmdAd_Click()
    If (txtPID.Text = "" Or txtPID.Text = " ") Then
        MsgBox "Please Select a Product !!!", vbOKOnly, "Information Required"
        cmdSelectPO.SetFocus
        Exit Sub
    End If
    If (txtQty.Text = "" Or txtQty.Text = " ") Then
        MsgBox "Please provide Quantity for selected Product !!!", vbOKOnly, "Information Required"
        txtQty.SetFocus
        Exit Sub
    End If
    If (txtPrice.Text = "" Or txtPrice.Text = " ") Then
        MsgBox "Please provide Price for selected Product !!!", vbOKOnly, "Information Required"
        txtPrice.SetFocus
        Exit Sub
    End If
    
'        iR = iR + 1
'
'        If PrdGrid.Rows = iR Then
'            txtPID.Text = ""
'            txtQty.Text = ""
'            txtPrice.Text = ""
'            txtPricePU.Text = ""
'        End If
        AddingData = True
        GenerateID
        cmdSelProd.SetFocus
End Sub

Private Sub cmdAdd_Click()

    NewQuantity = 0
       
    'Checking Fields for Records
    If (txtTID.Text = "" Or txtTID.Text = " ") Then
        'MsgBox "Please provide a Transaction ID !!!", vbOKOnly, "Information Required"
        txtTID.SetFocus
        Exit Sub
    End If
    If (txtPO.Text = "" Or txtPO.Text = " ") Then
        MsgBox "Please Select a Purchase Order !!!", vbOKOnly, "Information Required"
        'txtPO.SetFocus
        Exit Sub
    End If
     If (txtPID.Text = "" Or txtPID.Text = " ") Then
        MsgBox "Please Select a Product !!!", vbOKOnly, "Information Required"
        txtPID.SetFocus
        Exit Sub
    End If
    If (txtQty.Text = "" Or txtQty.Text = " ") Then
        MsgBox "Please provide Quantity for selected Product !!!", vbOKOnly, "Information Required"
        txtQty.SetFocus
        Exit Sub
    End If
    If (txtPrice.Text = "" Or txtPrice.Text = " ") Then
        MsgBox "Please provide Price for selected Product !!!", vbOKOnly, "Information Required"
        txtPrice.SetFocus
        Exit Sub
    End If
    If (txtR.Text = "") Then txtR.Text = "-"
    
    iR = PrdGrid.Rows - 1
    
    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
    PrdGrid.TextMatrix(iR, 1) = txtPID.Text
    PrdGrid.TextMatrix(iR, 2) = txtQty.Text
    PrdGrid.TextMatrix(iR, 3) = txtPrice.Text
    PrdGrid.TextMatrix(iR, 4) = txtPricePU.Text
    
        'Updating Database
        If Len(PrdGrid.TextMatrix(1, 1)) > 0 Then
            
            rn = 1
        
            For rn = 1 To PrdGrid.Rows - 1
            If PrdGrid.TextMatrix(rn, 0) <> "" Then
        
                sql = "INSERT INTO Receivings VALUES("
                sql = sql & "'" & (PrdGrid.TextMatrix(rn, 0)) & "',"
                sql = sql & "'" & txtDate & "','" & txtPO & "',"
                sql = sql & "'" & (PrdGrid.TextMatrix(rn, 1)) & "',"
                sql = sql & "" & (Val(PrdGrid.TextMatrix(rn, 2))) & ","
                sql = sql & "" & (Val(PrdGrid.TextMatrix(rn, 3))) & ","
                sql = sql & "" & (Val(PrdGrid.TextMatrix(rn, 4))) & ","
                sql = sql & "'" & txtR & "');"
                
                Prod_ID = (PrdGrid.TextMatrix(rn, 1))
                NewQuantity = (Val(PrdGrid.TextMatrix(rn, 2)))
                
                'MsgBox SQL
                conn.Execute sql
                
                AddQuantity = True
                UpdateQuantities
                
            End If
            Next
            
            MsgBox "Data Saved Successfully", vbInformation, "POS"
            Normalize
            cmdRDB_Click
            cmdNew.SetFocus

        Else
            MsgBox "Data Not Available", vbInformation, "POS"
        End If
        Exit Sub
        
End Sub
Private Sub UpdateQuantities()
    
    NewQuantity = 0
    Set rsTmp = New ADODB.Recordset
    Query = "SELECT Stock_In_Hand FROM Stock WHERE Product_ID='" & Prod_ID & "'"

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
    
    ExistingQuantity = Val(rsTmp!Stock_In_Hand)
    
    If AddQuantity = True Then
        NewQuantity = ExistingQuantity + (Val(txtQty.Text))
    
    ElseIf MinusQuantity = True Then
        NewQuantity = ExistingQuantity - (Val(txtQty.Text))
    
'    ElseIf (AddQuantity = False And ChangeQuantity = False) Then
'        NewQuantity = ExistingQuantity - Val(txtQty.Text)
    
    ElseIf ChangeQuantity = True Then
        NewQuantity = (ExistingQuantity - CurrentQuantity) + Val(txtQty.Text)
    End If
    
    Query = "UPDATE Stock SET Stock_In_Hand=" & NewQuantity & " WHERE Product_ID='" & Prod_ID & "'"
    '** NOTE, Price of Stock is not Updated because the user might not want to change the price everytime!
    'It can be done if requried!
    'MsgBox Query
    conn.Execute Query
    
    AddQuantity = False
    ChangeQuantity = False
    MinusQuantity = False
    
    rsTmp.Close
    Set rsTmp = Nothing
    
End Sub

Private Sub cmdEdit_Click()
    CurrentQuantity = Val(txtQty.Text)
    ChangeQuantity = True
    
    cmdSelectPO.Enabled = True
    cmdAd.Enabled = True
    cmdSelProd.Enabled = True
    txtQty.Enabled = True
    txtPrice.Enabled = True
    txtR.Enabled = True
    cmdSelectPO.SetFocus
        
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Receivings SET PO_No='" & txtPO.Text & "', Quantity=" & txtQty.Text & ",Price=" & txtPrice.Text & ",Price_Per_Unit=" & txtPricePU.Text & ",Remarks='" & txtR.Text & "' Where TID='" & txtTID.Text & "'"
    'MsgBox sql
    conn.Execute sql
    
    Prod_ID = txtPID.Text
    UpdateQuantities
    
    ShowReceivingsData ("SELECT * FROM Receivings ORDER BY TID")
    Set DataGrid1.DataSource = RsReceivingsGrid
    ShowReceivingsGrid ("SELECT * FROM Receivings ORDER BY TID")
    DataGrid1.Row = Rx

    Normalize
    
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
    
    sql = "DELETE FROM Receivings Where TID='" & txtTID.Text & "'"
    Prod_ID = txtPID.Text
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    conn.Execute sql

    
    AddQuantity = False
    MinusQuantity = True
    UpdateQuantities
    
    ClearFields
    Normalize
    cmdRDB_Click
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    ClearFields
    SQLString = "SELECT * FROM Receivings ORDER BY TID"
    Rx = 0
    ShowReceivingsData (SQLString)
    ShowReceivingsGrid (SQLString)
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowReceivingsData ("SELECT * FROM Receivings ORDER BY TID")
    ShowReceivingsGrid ("SELECT * FROM Receivings ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowReceivingsData ("SELECT * FROM Receivings ORDER BY TID")
    ShowReceivingsGrid ("SELECT * FROM Receivings ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowReceivingsData ("SELECT * FROM Receivings ORDER BY TID")
    ShowReceivingsGrid ("SELECT * FROM Receivings ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowReceivingsData ("SELECT * FROM Receivings ORDER BY TID")
    ShowReceivingsGrid ("SELECT * FROM Receivings ORDER BY TID")
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

    SQLString = "SELECT * FROM Receivings WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
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
    txtDate.Text = Format(rs!Dated, "YYYY-MM-DD")
    txtPO.Text = rs!PO_No
    txtPID.Text = rs!Product_ID
    txtQty.Text = rs!Quantity
    txtPrice.Text = rs!Price
    txtPricePU.Text = rs!Price_Per_Unit
    txtR.Text = rs!Remarks

    End If
    rs.Close
    Set rs = Nothing
Err:
    MsgBox "Invalid Search", vbInformation
    txtSearch.SetFocus

End Sub

Public Sub EnterNewReceivings()

    txtQty.Text = "1"
    txtPrice.Text = "1"
    SetButtons (False)
    SetFields (True)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdNew.Visible = False
    cmdCancel.Enabled = True
    cmdSelectPO.Enabled = True
    cmdSelProd.Enabled = True
    cmdAdd.Enabled = True
    ClearFields
    GenerateID
    GetDate
    txtDate.Text = DateToday
    cmdAd.Enabled = True
    
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
    txtQty.Enabled = TextFieldLock
    txtPrice.Enabled = TextFieldLock
    'txtPricePU.Enabled = TextFieldLock
    txtR.Enabled = TextFieldLock
End Sub

Private Sub ClearFields()
    txtTID.Text = ""
    txtPID.Text = ""
    txtDate.Text = ""
    txtPO.Text = ""
    txtQty.Text = ""
    txtPrice.Text = ""
    txtPricePU.Text = ""
    txtR.Text = ""
End Sub

Private Sub Normalize()
    SetFields (False)
    SetButtons (True)
    cmdNew.Visible = True
    cmdEdit.Visible = True
    cmdDelete.Enabled = True

    cmdSelectPO.Enabled = False
    cmdSelProd.Enabled = False
    cmdAdd.Enabled = False
    cmdAd.Enabled = False
    
    txtSearch.Enabled = True
    ST.Enabled = True
    ClearFields
    
    PrdGrid.Visible = False
    ScrollBar.Visible = False
    DataGrid1.Visible = True

    cmdRDB_Click
End Sub

Private Sub GenerateID()
    txtTID.Text = "T" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub



Private Sub txtPrice_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    ppu = Val(txtPrice.Text) / Val(txtQty.Text)
    txtPricePU.Text = Val(ppu)
End Sub

Private Sub txtQty_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    ppu = Val(txtPrice.Text) / Val(txtQty.Text)
    txtPricePU.Text = Val(ppu)
End Sub

Private Sub txtTID_Change()
'    On Error Resume Next
'    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
End Sub

Private Sub txtPID_Change()
    On Error Resume Next
    If AddingData = True Then
        AddingData = False
        iR = iR + 1

        PrdGrid.Rows = PrdGrid.Rows + 1
        iR = PrdGrid.Rows - 1
    End If
    
    PrdGrid.TextMatrix(iR, 1) = txtPID.Text
    PrdGrid.TextMatrix(iR, 0) = txtTID.Text
End Sub

Private Sub txtQty_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 2) = txtQty.Text
End Sub

Private Sub txtPrice_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 3) = txtPrice.Text
End Sub

Private Sub txtPricePU_Change()
    On Error Resume Next
    PrdGrid.TextMatrix(iR, 4) = txtPricePU.Text
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

Private Sub txtSearch_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtDate_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPrice_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPricePU_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtQty_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSearch_Click
    End If
End Sub

