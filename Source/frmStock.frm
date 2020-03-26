VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: STOCK MANAGEMENT :."
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9165
   Icon            =   "frmStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmStock.frx":0ECA
   ScaleHeight     =   8670
   ScaleWidth      =   9165
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Bar Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Print Current Invoice..."
      Top             =   1200
      Width           =   3255
   End
   Begin VB.ComboBox Company 
      Enabled         =   0   'False
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
      ItemData        =   "frmStock.frx":4B4D7
      Left            =   2280
      List            =   "frmStock.frx":4B4D9
      Sorted          =   -1  'True
      TabIndex        =   11
      Text            =   "Company"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox txtSelPrice 
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
      TabIndex        =   14
      Text            =   "txtSelPrice"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5640
      ScaleHeight     =   825
      ScaleWidth      =   3225
      TabIndex        =   41
      ToolTipText     =   "BarCode has been Copied to Clipboard!"
      Top             =   240
      Width           =   3255
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   7440
      Top             =   8040
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid 
      Height          =   1335
      Left            =   6840
      TabIndex        =   39
      Top             =   7080
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.ComboBox PType 
      Enabled         =   0   'False
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
      ItemData        =   "frmStock.frx":4B4DB
      Left            =   2280
      List            =   "frmStock.frx":4B4DD
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   "ProductType"
      Top             =   1680
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
      TabIndex        =   3
      Top             =   5640
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
      TabIndex        =   0
      Text            =   "txtSearch"
      Top             =   5640
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
      TabIndex        =   2
      Top             =   5640
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
      ItemData        =   "frmStock.frx":4B4DF
      Left            =   3720
      List            =   "frmStock.frx":4B501
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   "Product_ID"
      Top             =   5640
      Width           =   1935
   End
   Begin VB.TextBox txtProduct 
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
      TabIndex        =   8
      Text            =   "txtProduct"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtDescription 
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
      Text            =   "txtDescription"
      Top             =   2160
      Width           =   2295
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
      TabIndex        =   21
      Top             =   4800
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
      TabIndex        =   26
      Top             =   4800
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
      TabIndex        =   25
      Top             =   5160
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
      TabIndex        =   20
      Top             =   5160
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
      TabIndex        =   5
      Top             =   4800
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
      Left            =   2280
      TabIndex        =   7
      Text            =   "txtDate"
      ToolTipText     =   "Date Format yyyy-MM-dd"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox txtPID 
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
      Text            =   "txtPID"
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox txtStock 
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
      Left            =   2280
      TabIndex        =   15
      Text            =   "txtStock"
      Top             =   3240
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
      TabIndex        =   17
      Text            =   "frmStock.frx":4B579
      Top             =   3840
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
      TabIndex        =   24
      Top             =   4800
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
      TabIndex        =   23
      Top             =   4800
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
      TabIndex        =   22
      Top             =   5160
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
      TabIndex        =   19
      Top             =   5160
      Width           =   1455
   End
   Begin VB.TextBox txtPS 
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
      Text            =   "txtPS"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtBuyPrice 
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
      TabIndex        =   13
      Text            =   "txtBuyPrice"
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtROL 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
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
      TabIndex        =   16
      Text            =   "txtROL"
      Top             =   3240
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   27
      Top             =   6120
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
      TabIndex        =   18
      Top             =   4800
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
      TabIndex        =   28
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price"
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
      TabIndex        =   43
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      TabIndex        =   42
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   40
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblPID 
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID"
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
   Begin VB.Label lblST 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock in Hand"
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
      Top             =   3240
      Width           =   1815
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
      Left            =   360
      TabIndex        =   36
      Top             =   1200
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
      Top             =   3840
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
      Left            =   360
      TabIndex        =   34
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblDesc 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label lblPT 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Type"
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
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblPS 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Size"
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
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblPPU 
      BackStyle       =   0  'Transparent
      Caption         =   "Buying Price"
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
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblROL 
      BackStyle       =   0  'Transparent
      Caption         =   "ReOrder Level"
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
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   3720
      Y2              =   3720
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public EticXrow As Integer, NumRow As Integer, TopMarg As Single
Public LeftMarg As Single, HorSpace As Single, VertSpace As Single, NomeArc As String
Public InterX As Single, nomeFile As String
Dim x As Double, y As Double, OddString As String, depX As Double, depY As Double, precX As Double

Dim sql As String
Dim rn As Integer
Dim nLastKeyAscii As Integer
Private iC, iR As Integer
Private TextFieldLock, ButtonLock, isChecked As Boolean

Private Sub cmdPrint_Click()
    Dim numberofcopies As String
    numberofcopies = InputBox("Please provide number of copies you want to PRINT!      [Maximum 22 per page]", "Information Required", 22)
    If numberofcopies = "" Or numberofcopies = " " Then Exit Sub
    
    EticXrow = 2
    NumRow = Val(numberofcopies)
    
    If NumRow > 22 Then
        MsgBox "Only 22 Labels will be printed per page!", vbInformation, "POS"
        NumRow = 22
    End If
    
    TopMarg = 10
    LeftMarg = 15
    HorSpace = 40
    VertSpace = 25 '30
    InterX = 0.325
    If InterX < 0.325 Then InterX = 0.325

    Dim i As Integer, cod As String
    Dim RStampate, xx As Integer, ColStampate As Integer
    x = LeftMarg
    y = TopMarg
    RStampate = 1
    ColStampate = 0
    
    For xx = 1 To NumRow
        cod = UCase(Trim(txtPID.Text))
        If cod = "" Then GoTo Prossimo
        If ColStampate = EticXrow Then
            If RStampate = NumRow Then
                x = LeftMarg
                Printer.NewPage
                y = TopMarg
                RStampate = 1
                ColStampate = 1
            Else
                x = LeftMarg
                y = y + VertSpace
                RStampate = RStampate + 1
                ColStampate = 1
            End If
        Else
            ColStampate = ColStampate + 1
        End If
        
        precX = x
        drawBar2 ("100101101101")
        For i = 1 To Len(cod)
            drawBar2 (retcode(Mid(cod, i, 1), i))
        Next i
        drawBar2 ("100101101101")
        For i = Len(cod) + 1 To 14
            drawBar2 ("000000000000")
        Next i
        depX = Printer.CurrentX
        depY = Printer.CurrentY
        Printer.CurrentX = precX
        Printer.CurrentY = depY + 1
        Printer.FontName = "verdana"
        Printer.FontSize = "8"
        Printer.Print cod
        Printer.CurrentX = precX
        Printer.CurrentY = depY - 15
        Printer.Print txtProduct.Text & " Rs." & txtSelPrice.Text
        Printer.CurrentX = depY
        Printer.CurrentY = depX
        x = x + HorSpace
Prossimo:
    Next
    Printer.EndDoc
End Sub
Private Function retcode(ByVal numero As String, ByVal counter As Integer) As String
    Dim pari As Boolean, RetSRt As String, ODD As Boolean
    RetSRt = ""
    Select Case numero
        Case "0"
            RetSRt = "101001101101"
        Case "1"
            RetSRt = "110100101011"
        Case "2"
            RetSRt = "101100101011"
        Case "3"
            RetSRt = "110110010101"
        Case "4"
            RetSRt = "101001101011"
        Case "5"
            RetSRt = "110100110101"
        Case "6"
            RetSRt = "101100110101"
        Case "7"
            RetSRt = "101001011011"
        Case "8"
            RetSRt = "110100101101"
        Case "9"
            RetSRt = "101100101101"
        Case "A"
            RetSRt = "101100101101"
        Case "B"
            RetSRt = "101101001011"
        Case "C"
            RetSRt = "110110100101"
        Case "D"
            RetSRt = "101011001011"
        Case "E"
            RetSRt = "110101100101"
        Case "F"
            RetSRt = "101101100101"
        Case "G"
            RetSRt = "101010011011"
        Case "H"
            RetSRt = "110101001101"
        Case "I"
            RetSRt = "101101001101"
        Case "J"
            RetSRt = "101011001101"
        Case "K"
            RetSRt = "110101010011"
        Case "L"
            RetSRt = "101101010011"
        Case "M"
            RetSRt = "110110101001"
        Case "N"
            RetSRt = "101011010011"
        Case "O"
            RetSRt = "110101101001"
        Case "P"
            RetSRt = "101101101001"
        Case "Q"
            RetSRt = "101010110011"
        Case "R"
            RetSRt = "110101011001"
        Case "S"
            RetSRt = "101101011001"
        Case "T"
            RetSRt = "101011011001"
        Case "U"
            RetSRt = "110010101011"
        Case "V"
            RetSRt = "100110101011"
        Case "W"
            RetSRt = "110011010101"
        Case "X"
            RetSRt = "100101101011"
        Case "Y"
            RetSRt = "110010110101"
        Case "Z"
            RetSRt = "100110110101"
        Case "-"
            RetSRt = "100101011011"
        Case "."
            RetSRt = "110010101101"
        Case "$"
            RetSRt = "100100100101"
        Case "/"
            RetSRt = "100100101001"
        Case "+"
            RetSRt = "100101001001"
        Case "%"
            RetSRt = "101001001001"
    End Select
    retcode = RetSRt
End Function
Private Sub drawBar2(ByVal stringa As String)
    Dim i As Integer, car As String, stepx As Double, colore As Long, dimy As Double
    stepx = InterX
    dimy = 10
    Printer.ScaleMode = 6
    stringa = stringa & "0"
    For i = 1 To Len(stringa)
        car = Mid(stringa, i, 1)
        If car = "0" Then colore = vbWhite Else colore = vbBlack
        Printer.Line (x, y)-(x + stepx, y + dimy), colore, BF
        x = x + stepx
    Next i
End Sub

Private Sub Form_Load()

    Connect
    ClearFields
    Normalize
    GetDate
    
    CheckROL
    CheckMinusStock

    CheckB4Connect
    ShowStockData (SQLString)
    ShowStockGrid (SQLString)
    
     'For Int TextBoxes
    Dim tmp1, tmp2, tmp3, tmp4, tmp5 As Long
    tmp1 = SetWindowLong(txtBuyPrice.hwnd, GWL_STYLE, GetWindowLong(txtBuyPrice.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp2 = SetWindowLong(txtSelPrice.hwnd, GWL_STYLE, GetWindowLong(txtSelPrice.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp3 = SetWindowLong(txtStock.hwnd, GWL_STYLE, GetWindowLong(txtStock.hwnd, GWL_STYLE) Or ES_NUMBER)
    tmp4 = SetWindowLong(txtROL.hwnd, GWL_STYLE, GetWindowLong(txtROL.hwnd, GWL_STYLE) Or ES_NUMBER)
End Sub

Private Sub CheckB4Connect()
    isChecked = False
    
    SQLString = "SELECT * FROM Stock ORDER BY Product_Type"
    If isReOrder = True Then SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<ReOrder_Level ORDER BY Product_Type;"
    If isStockMinus = True Then SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<0 ORDER BY Product_Type;"
End Sub

Private Sub cmdNew_Click()
    EnterNewProduct
    txtStock.Text = "0"
    txtROL.Text = "10"
End Sub

Private Sub cmdAdd_Click()

    'Checking Fields for Records
    If (txtPID.Text = "" Or txtPID.Text = " ") Then
        MsgBox "Please provide a Product ID !!!", vbOKOnly, "Information Required"
        txtPID.SetFocus
        Exit Sub
    End If
    If (txtProduct.Text = "" Or txtProduct.Text = " ") Then
        MsgBox "Please provide a Product Name !!!", vbOKOnly, "Information Required"
        txtProduct.SetFocus
        Exit Sub
    End If
    If (PType.Text = "" Or PType.Text = " ") Then
        MsgBox "Please provide Product Type for Product " + txtProduct.Text + " !!!", vbOKOnly, "Information Required"
        PType.SetFocus
        Exit Sub
    End If
    If (Company.Text = "" Or Company.Text = " ") Then Company.Text = "-"
    If (txtPS.Text = "" Or txtPS.Text = " ") Then txtPS.Text = "-"
    If (txtBuyPrice.Text = "" Or txtBuyPrice.Text = " ") Then txtBuyPrice.Text = "0"
    If (txtSelPrice.Text = "" Or txtSelPrice.Text = " ") Then txtSelPrice.Text = "0"
    If (txtDescription.Text = "" Or txtDescription.Text = " ") Then txtDescription.Text = "-"
    If (txtStock.Text = "" Or txtStock.Text = " ") Then txtStock.Text = "0"
    If (txtROL.Text = "" Or txtROL.Text = " ") Then txtROL.Text = "10"
    If (txtR.Text = "") Then txtR.Text = "-"
    
    'Updating Database
    If DupCheck("SELECT * FROM Stock WHERE Product_ID='" & txtPID.Text & "' AND Product='" & txtProduct.Text & "' AND Product_Type='" & PType.Text & "' AND Product_Size='" & txtPS.Text & "' AND Company='" & UCase(Company.Text) & "'") = True Then
        MsgBox "Product Already Exists in Stock!!! ", , "General Error"
    Else
        sql = "INSERT INTO Stock VALUES('" & txtPID & "','" & txtDate & "','" & UCase(txtProduct) & "','" & UCase(PType) & "','" & txtPS & "','" & UCase(Company) & "'," & txtStock & ",'" & txtDescription & "'," & txtBuyPrice & "," & txtSelPrice & "," & txtROL & ",'" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
    End If
        
    Normalize
    cmdNew.SetFocus
    Unload Me
    Load Me
    Me.Left = 20
    Me.Top = 20
    Exit Sub
    
End Sub

Private Sub cmdEdit_Click()
    SetFields (True)
    txtProduct.SetFocus
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Stock SET Date='" & txtDate.Text & "',Product='" & txtProduct.Text & "',Product_Type='" & PType.Text & "',Product_Size='" & txtPS.Text & "',Company='" & UCase(Company.Text) & "',Stock_In_Hand='" & txtStock.Text & "',Description='" & txtDescription.Text & "',Buying_Price=" & txtBuyPrice.Text & ",Selling_Price=" & txtSelPrice.Text & ",ReOrder_Level=" & txtROL.Text & ",Remarks='" & txtR.Text & "' Where Product_ID='" & txtPID.Text & "'"
    conn.Execute sql
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_Type")
    Set DataGrid1.DataSource = rsStockGrid
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_Type")
    DataGrid1.Row = Rx

    Normalize
    
End Sub

Private Sub cmdCancel_Click()
    ClearFields
    Normalize
End Sub

Private Sub cmdDelete_Click()
    Dim sqlR, sqlIn, sqlS As String
        
    If MsgBox("This will DELETE Complete Data of the current Product from Database[Receivings, Invoices & Stock]. ARE YOU SURE?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        'Set rsTemp = Nothing
        Exit Sub
    End If
    
    'Deleting Receivings info
    sqlR = "DELETE FROM Receivings WHERE Product_ID='" & txtPID.Text & "'"
'    MsgBox "SQL IS " & sqlR
    conn.Execute sqlR
    
    'Deleting Invoice info
    sqlIn = "DELETE FROM Invoice WHERE Product_ID='" & txtPID.Text & "'"
'    MsgBox "SQL IS " & sqlSA
    conn.Execute sqlIn
    
    'Deleting Purchase_Order info
    sqlS = "DELETE FROM Stock WHERE Product_ID='" & txtPID.Text & "'"
'    MsgBox "SQL IS " & sqlPO
    conn.Execute sqlS
       
    Rx = Rx - 1
    
    Set DataGrid1.DataSource = rsStockGrid
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_ID")
    If (Rx <> 0) Then DataGrid1.Row = Rx
    ClearFields
    Normalize
    cmdRDB_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdRDB_Click()
    ClearFields
    SQLString = "SELECT * FROM Stock ORDER BY Product_ID"
    Rx = 0
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_Type")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_Type")
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_Type")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_Type")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_Type")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_Type")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_Type")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_Type")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowStockData ("SELECT * FROM Stock ORDER BY Product_Type")
    ShowStockGrid ("SELECT * FROM Stock ORDER BY Product_Type")
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
    
    If (ST.Text = "ReOrder_Level" Or ST.Text = "Stock_In_Hand") Then
        SQLString = "SELECT * FROM Stock WHERE " + ST.Text + "=" & Val(txtSearch)
    Else
        SQLString = "SELECT * FROM Stock WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    End If
    
    rs.Open SQLString, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    Set rsStockGrid = New ADODB.Recordset
    rsStockGrid.CursorLocation = adUseClient
    rsStockGrid.CursorType = adOpenStatic
    rsStockGrid.LockType = adLockReadOnly
    rsStockGrid.Open SQLString, conn
    Set DataGrid1.DataSource = rsStockGrid
      
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        
        MsgBox "Record Not Found !!!", vbInformation, ""
        txtSearch.SetFocus
        SendKeys "{Home}+{End}"
        cmdRDB_Click
        Exit Sub
    End If
    If IsNull(rs!Product_ID) Then
        ClearFields
    Else
       
    txtPID.Text = rs!Product_ID
    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
    txtProduct.Text = rs!Product
    PType.Text = rs!Product_Type
    txtPS.Text = rs!Product_Size
    Company.Text = rs!Company
    txtBuyPrice.Text = rs!Buying_Price
    txtSelPrice.Text = rs!selling_Price
    txtDescription.Text = rs!Description
    txtStock.Text = rs!Stock_In_Hand
    txtROL.Text = rs!ReOrder_Level
    txtR.Text = rs!Remarks
    
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub ClearFields()
    txtPID.Text = ""
    txtDate.Text = ""
    txtProduct.Text = ""
    PType.Text = ""
    txtPS.Text = ""
    Company.Text = ""
    txtBuyPrice.Text = ""
    txtSelPrice.Text = ""
    txtDescription.Text = ""
    txtStock.Text = ""
    txtROL.Text = ""
    txtR.Text = ""
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    txtDate.Enabled = TextFieldLock
    txtProduct.Enabled = TextFieldLock
    PType.Enabled = TextFieldLock
    Company.Enabled = TextFieldLock
    txtPS.Enabled = TextFieldLock
    Company.Enabled = TextFieldLock
    txtBuyPrice.Enabled = TextFieldLock
    txtSelPrice.Enabled = TextFieldLock
    txtDescription.Enabled = TextFieldLock
    txtStock.Enabled = TextFieldLock
    txtROL.Enabled = TextFieldLock
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

    txtSearch.Enabled = True
    txtSearch.Text = ""
    ST.Enabled = True
    cmdRDB_Click
    GetComboData
    RemoveComboDuplicates
End Sub

Public Sub EnterNewProduct()
    
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
    txtProduct.SetFocus
    
End Sub
Private Sub GenerateID()
    txtPID.Text = "P" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Function DupCheck(chkID As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic

    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtPID.Text = rs!Product_ID Then
        DupCheck = True
    Else
        DupCheck = False
    End If
    rs.Close
    Set rs = Nothing
End Function

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

Private Sub GetComboData()
    Adodc.ConnectionString = conn
    Adodc.CursorLocation = adUseClient
    Adodc.CursorType = adOpenDynamic
    Adodc.RecordSource = "SELECT Product_Type,Company FROM Stock ORDER BY Product_Type"
    Set DataGrid.DataSource = Adodc
    
    If Adodc.Recordset.BOF Then
        Exit Sub
    Else

    'For Item1 and Item Combo
        Dim x As Integer
        For x = 0 To (Adodc.Recordset.RecordCount - 1)
            PType.AddItem Adodc.Recordset.Fields(0)
            Company.AddItem Adodc.Recordset.Fields(1)
            Adodc.Recordset.MoveNext
        Next x
    End If
    
End Sub

Public Function RemoveComboDuplicates()
    Dim y As Integer
    Dim x As Integer
    y = PType.ListCount + 1
    For x = 1 To PType.ListCount
        y = y - 1
        If PType.List(y) = PType.List(y - 1) Then
            PType.RemoveItem (y)
        End If
    Next
    
    y = Company.ListCount + 1
    For x = 1 To Company.ListCount
        y = y - 1
        If Company.List(y) = Company.List(y - 1) Then
            Company.RemoveItem (y)
        End If
    Next
End Function

Private Sub PType_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(PType)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(PType)
   End Select
End Sub
Private Sub Company_Change()
   Select Case nLastKeyAscii
      Case vbKeyBack
         Call Combo_Lookup(Company)
      Case vbKeyDelete
      Case Else
         Call Combo_Lookup(Company)
   End Select
End Sub
Private Sub PType_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(PType.SelText) <> 0 And PType.SelStart > 0 Then
         PType.SelStart = PType.SelStart - 1
         PType.SelLength = CB_MAXLENGTH
   End If
End Sub
Private Sub Company_KeyDown(KeyCode As Integer, Shift As Integer)
   nLastKeyAscii = KeyCode
   
   If KeyCode = vbKeyBack And Len(Company.SelText) <> 0 And Company.SelStart > 0 Then
         Company.SelStart = Company.SelStart - 1
         Company.SelLength = CB_MAXLENGTH
   End If
End Sub


Private Sub txtDescription_GotFocus()
    SendKeys "{Home}+{End}"
End Sub


Private Sub txtPID_Change()
Call DrawBarcode(txtPID, Picture1)
End Sub

Private Sub txtPricePU_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtProduct_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtPS_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtROL_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSearch_Click
    End If
End Sub

Private Sub txtStock_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub CheckROL()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseServer
    rs.CursorType = adOpenStatic
    rs.LockType = adLockOptimistic
    
    sql = "SELECT COUNT(*) as No FROM Stock WHERE Stock_In_Hand<ReOrder_Level;"
    rs.Open sql, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If
    If Val(rs!No) > 0 Then
        If MsgBox("Some products needs to be ReOrdered!, would you like to have a look?", vbYesNo + vbDefaultButton2, "Stock") = vbYes Then
        'SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<ReOrder_Level;"
        'MsgBox rs!No & " Product(s) needs to be ReOrdered!", vbInformation, "Stock"
        isReOrder = True
        End If
    Else
        isReOrder = False
        Exit Sub
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub CheckMinusStock()
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
    If Val(rs!No) > 0 Then
        If MsgBox("Some product's quantities are in minus in stock, which needs to be urgently ReOrdered!, would you like to have a look?", vbYesNo + vbDefaultButton2, "Stock") = vbYes Then
        'SQLString = "SELECT * FROM Stock WHERE Stock_In_Hand<0;"
        'MsgBox rs!No & " Product(s) in Stock needs to be ReOrdered URGENTLY!", vbCritical, "Stock"
        isStockMinus = True
        End If
    Else
        isStockMinus = False
        Exit Sub
    End If
    rs.Close
    Set rs = Nothing
End Sub
