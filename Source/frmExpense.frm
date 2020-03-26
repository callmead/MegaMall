VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmExpense 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: EXPENSE :."
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9195
   Icon            =   "frmExpense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmExpense.frx":0ECA
   ScaleHeight     =   7335
   ScaleWidth      =   9195
   Begin VB.TextBox txtParticulars 
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
      TabIndex        =   6
      Text            =   "txtParticulars"
      Top             =   1200
      Width           =   2295
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
      ItemData        =   "frmExpense.frx":4B4D7
      Left            =   2280
      List            =   "frmExpense.frx":4B4E4
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "PM"
      Top             =   1200
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   240
      TabIndex        =   22
      Top             =   4800
      Width           =   8655
      _ExtentX        =   15266
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
   Begin VB.ComboBox ET 
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
      ItemData        =   "frmExpense.frx":4B503
      Left            =   2280
      List            =   "frmExpense.frx":4B50D
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "ET"
      Top             =   720
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
      TabIndex        =   20
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
      TabIndex        =   17
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
      TabIndex        =   19
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
      ItemData        =   "frmExpense.frx":4B52C
      Left            =   3720
      List            =   "frmExpense.frx":4B542
      Sorted          =   -1  'True
      TabIndex        =   18
      Text            =   "Supplier"
      Top             =   4320
      Width           =   1935
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
      TabIndex        =   12
      Top             =   3240
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
      TabIndex        =   21
      Top             =   3240
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
      TabIndex        =   16
      Top             =   3600
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
      TabIndex        =   11
      Top             =   3600
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
      Top             =   3240
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
      TabIndex        =   8
      Text            =   "frmExpense.frx":4B587
      Top             =   2280
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
      TabIndex        =   15
      Top             =   3240
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
      TabIndex        =   14
      Top             =   3240
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
      TabIndex        =   13
      Top             =   3600
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
      TabIndex        =   10
      Top             =   3600
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
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
      Text            =   "txtAmount"
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtSupplier 
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
      Text            =   "txtSupplier"
      Top             =   720
      Width           =   2295
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
      TabIndex        =   9
      Top             =   3240
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
      TabIndex        =   23
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
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
      Top             =   1200
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
      Left            =   360
      TabIndex        =   30
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label lblSID 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Top             =   720
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
      TabIndex        =   28
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblET 
      BackStyle       =   0  'Transparent
      Caption         =   "Expense Type"
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
      TabIndex        =   26
      Top             =   2280
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
      TabIndex        =   25
      Top             =   240
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   240
      Y2              =   7920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8880
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label lblAmount 
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
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   8880
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "frmExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TID, sql As String
Private iC, iR, rn As Integer
Private TextFieldLock, ButtonLock, fieldlock As Boolean

Private Sub ET_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub Form_Load()

    Connect
    GetDate
    
    SQLString = "SELECT * from Expenditure ORDER BY TID"
    ShowExpenseData (SQLString)
    ShowExpenseGrid (SQLString)
    
    ClearFields
    
    Normalize
    txtSearch.Text = ""
    
    'For Int TextBoxes
    Dim tmp1 As Long
    tmp1 = SetWindowLong(txtAmount.hwnd, GWL_STYLE, GetWindowLong(txtAmount.hwnd, GWL_STYLE) Or ES_NUMBER)
    fieldlock = False

End Sub

Private Sub cmdNew_Click()
    EnterNewExpenditure
    txtAmount.Text = "0"
    ET.Text = "General"
    PM.Text = "Cash"
    fieldlock = True
End Sub

Private Sub cmdAdd_Click()
    
    CheckFieldsData
    
    'Updating Database
    If DupCheck("SELECT * from Expenditure WHERE TID='" & txtTID.Text & "'") = True Then
        MsgBox "TID Already Exists !!! ", , "General Error"
    Else
        sql = "INSERT INTO Expenditure VALUES('" & txtTID & "','" & txtDate & "','" & ET & "','" & txtSupplier & "','" & PM & "','" & txtParticulars & "'," & txtAmount & ",'" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
    End If
        
    Normalize
    cmdNew.SetFocus
    
End Sub

Private Sub cmdEdit_Click()
    
    SetFields (True)
    ET.SetFocus
    SetButtons (False)
    txtSearch.Enabled = False
    ST.Enabled = False
    cmdEdit.Visible = False
    cmdDelete.Enabled = False
    cmdCancel.Enabled = True
    cmdSave.Enabled = True
    fieldlock = True

End Sub

Private Sub cmdSave_Click()
    
    sql = "UPDATE Expenditure SET Date='" & txtDate.Text & "',Expense_Type='" & ET.Text & "',Supplier='" & txtSupplier.Text & "',Payment_Mode='" & PM.Text & "',Particulars='" & txtParticulars.Text & "',Amount=" & txtAmount.Text & ",Remarks='" & txtR.Text & "' Where TID='" & txtTID.Text & "'"
    conn.Execute sql
    ShowExpenseData (SQLString)
    Set DataGrid1.DataSource = RsSuppGrid
    ShowExpenseGrid ("SELECT * from Expenditure ORDER BY TID")
    DataGrid1.Row = Rx

    ClearFields
    Normalize
    
End Sub

Private Sub cmdCancel_Click()
    Normalize
End Sub

Private Sub cmdDelete_Click()
    sql = "DELETE FROM Expenditure Where TID='" & txtTID.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    conn.Execute sql
    ClearFields
    Normalize
    cmdRDB_Click
End Sub

Private Sub cmdClose_Click()
    Unload Me
    UpdateProfit
End Sub

Private Sub cmdRDB_Click()
    ClearFields
    SQLString = "SELECT * from Expenditure ORDER BY TID"
    Rx = 0
    ShowExpenseData (SQLString)
    ShowExpenseGrid (SQLString)
End Sub

Private Sub cmdMF_Click()
    On Error Resume Next
    Rx = 0
    ShowExpenseData ("SELECT * from Expenditure ORDER BY TID")
    ShowExpenseGrid ("SELECT * from Expenditure ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdML_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowExpenseData ("SELECT * from Expenditure ORDER BY TID")
    ShowExpenseGrid ("SELECT * from Expenditure ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdN_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowExpenseData ("SELECT * from Expenditure ORDER BY TID")
    ShowExpenseGrid ("SELECT * from Expenditure ORDER BY TID")
    DataGrid1.Row = Rx
End Sub

Private Sub cmdP_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowExpenseData ("SELECT * from Expenditure ORDER BY TID")
    ShowExpenseGrid ("SELECT * from Expenditure ORDER BY TID")
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
    
    If (ST.Text = "Amount") Then
        SQLString = "SELECT * from Expenditure WHERE " + ST.Text + "=" & txtSearch
    Else
        SQLString = "SELECT * from Expenditure WHERE " + ST.Text + " LIKE '" & txtSearch & "%'"
    End If
    
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
    If IsNull(rs!TID) Then
        ClearFields
    Else
       
    txtTID.Text = rs!TID
    txtDate.Text = Format(rs!Date, "YYYY-MM-DD")
    ET.Text = rs!Expense_Type
    txtSupplier.Text = rs!Customer
    PM.Text = rs!Payment_Mode
    txtParticulars.Text = rs!Particulars
    txtAmount.Text = rs!Amount
    txtR.Text = rs!Remarks
    
    End If
    rs.Close
    Set rs = Nothing
Err:
    MsgBox "Invalid Search", vbInformation
    txtSearch.SetFocus

End Sub

Private Sub ClearFields()
    txtTID.Text = ""
    txtDate.Text = ""
    ET.Text = ""
    txtSupplier.Text = ""
    PM.Text = ""
    txtParticulars.Text = ""
    txtAmount.Text = ""
    txtR.Text = ""
End Sub

Private Sub SetFields(TextFieldLock As Boolean)
    ET.Enabled = TextFieldLock
    PM.Enabled = TextFieldLock
    txtParticulars.Enabled = TextFieldLock
    txtAmount.Enabled = TextFieldLock
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
    txtSupplier.Enabled = False

    txtSearch.Enabled = True
    ST.Enabled = True
    fieldlock = False
    cmdRDB_Click
End Sub

Public Sub EnterNewExpenditure()
    
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
    ET.SetFocus
    
End Sub
Private Sub GenerateID()
    txtTID.Text = "S" & Trim(Str(Year(Date))) & Trim(Str(Month(Date))) & Trim(Str(Day(Date))) & Trim(Str(Hour(Time))) & Trim(Str(Minute(Time))) & Trim(Str(Second(Time)))
End Sub

Private Sub Form_Unload(Cancel As Integer)
UpdateProfit
End Sub

Private Sub PM_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub PM_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
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

    'sql = "SELECT * from Expenditure WHERE TID='" & chkID & "'"
    rs.Open chkID, conn
    If rs.EOF = True Then
        rs.Close
        Set rs = Nothing
        Exit Function
    End If
    If txtTID.Text = rs!TID Then
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

Private Sub txtAmount_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtParticulars_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtR_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtSearch_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub ET_Change()
    If fieldlock = True Then
        If ET.Text = "Supplier Payment" Then
            txtSupplier.Enabled = True
        Else
            txtSupplier.Enabled = False
        End If
    End If
End Sub

Private Sub ET_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtSupplier_GotFocus()
    SendKeys "{Home}+{End}"
End Sub
Private Sub ET_KeyUp(KeyCode As Integer, Shift As Integer)
    If ET.Text = "Supplier Payment" Then
        txtSupplier.Enabled = True
    Else
        txtSupplier.Enabled = False
    End If
End Sub

Private Sub CheckFieldsData()
    'Checking Fields for Records
    If (ET.Text = "" Or ET.Text = " ") Then
        MsgBox "Please provide Expense Type !!!", vbOKOnly, "Information Required"
        ET.SetFocus
        Exit Sub
    End If
    If (txtSupplier.Text = "" Or txtSupplier.Text = " ") Then txtSupplier.Text = "-"
    If (PM.Text = "" Or PM.Text = " ") Then
        MsgBox "Please provide Payment Mode !!!", vbOKOnly, "Information Required"
        PM.SetFocus
        Exit Sub
    End If
    If (txtParticulars.Text = "" Or txtParticulars.Text = " ") Then txtParticulars.Text = "-"
    If (txtAmount.Text = "" Or txtAmount.Text = "0") Then
        MsgBox "Please provide Amount !!!", vbOKOnly, "Information Required"
        txtAmount.SetFocus
        Exit Sub
    End If
    If (txtR.Text = "") Then txtR.Text = "-"
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSearch_Click
    End If
End Sub

