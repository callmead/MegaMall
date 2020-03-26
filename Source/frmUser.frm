VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSecurity 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Manager"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8985
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmUser.frx":27A2
   ScaleHeight     =   4095
   ScaleWidth      =   8985
   Begin VB.TextBox txtR 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   6
      Text            =   "txtR"
      Top             =   1800
      Width           =   7155
   End
   Begin VB.TextBox txtDesg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5850
      TabIndex        =   4
      Text            =   "txtDesg"
      Top             =   780
      Width           =   2925
   End
   Begin VB.TextBox txtFN 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5850
      TabIndex        =   2
      Text            =   "txtFN"
      Top             =   240
      Width           =   2925
   End
   Begin VB.ComboBox UType 
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
      ItemData        =   "frmUser.frx":15C9E
      Left            =   1620
      List            =   "frmUser.frx":15CAE
      Sorted          =   -1  'True
      TabIndex        =   5
      Text            =   "User Type"
      Top             =   1320
      Width           =   2385
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1620
      TabIndex        =   3
      Text            =   "txtPass"
      Top             =   780
      Width           =   2385
   End
   Begin VB.TextBox txtUser 
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
      Height          =   360
      Left            =   1620
      TabIndex        =   1
      Text            =   "txtUser"
      Top             =   240
      Width           =   2385
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   "Last Rec."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7260
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3420
      Width           =   1485
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3420
      Width           =   1485
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2970
      Width           =   1485
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3420
      Width           =   1485
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2670
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3420
      Width           =   1485
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "First Rec."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3420
      Width           =   1485
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4890
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2970
      Width           =   1485
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2970
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Add User"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2970
      Width           =   1485
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Rec"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2970
      Width           =   1485
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3735
      Left            =   0
      TabIndex        =   16
      Top             =   4320
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6588
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      BackColor       =   16777215
      DefColWidth     =   73
      Enabled         =   -1  'True
      ForeColor       =   0
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
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   240
      X2              =   8760
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label lblR 
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Top             =   1830
      Width           =   1245
   End
   Begin VB.Label lblDesg 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   4470
      TabIndex        =   21
      Top             =   810
      Width           =   1275
   End
   Begin VB.Label lblFN 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   4470
      TabIndex        =   20
      Top             =   270
      Width           =   1245
   End
   Begin VB.Label lblType 
      BackStyle       =   0  'Transparent
      Caption         =   "User Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   1380
      Width           =   1245
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   240
      TabIndex        =   18
      Top             =   840
      Width           =   1245
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   270
      Width           =   1245
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sql As String

Private Sub DataGrid1_Click()
On Error Resume Next
With DataGrid1
    .Col = 0
    txtUser.Text = .Text
    .Col = 1
    txtPass.Text = .Text
    .Col = 2
    UType.Text = .Text
    .Col = 3
    txtFN.Text = .Text
    .Col = 4
    txtDesg.Text = .Text
    .Col = 5
    txtR.Text = .Text

End With

End Sub

Private Sub Form_Load()
    Connect
    SQLString = "SELECT * FROM Login ORDER BY User"
    ShowUserData (SQLString)
    ShowUserGrid
    
    cmdSave.Enabled = False
    
End Sub

Private Sub txtUser_LostFocus()
    txtUser.Text = LCase(txtUser.Text)
End Sub

Private Sub UType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdNew_Click()
'Clearing Text Boxes
    cmdFirst.Enabled = False
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    cmdDel.Enabled = False
    cmdEdit.Enabled = False
    cmdSave.Enabled = True
    UType.Text = "Select Type"
    ClearText
    txtUser.Enabled = True
    txtUser.SetFocus

End Sub

Private Sub cmdSave_Click()
'Check
    If txtUser.Text = "" Then
    MsgBox "Enter User Name"
    txtUser.SetFocus
    Exit Sub

    ElseIf txtPass.Text = "" Then
    MsgBox "Enter User Password"
    txtPass.SetFocus
    Exit Sub
    
    ElseIf txtFN.Text = "" Then
    MsgBox "Enter User Full Name"
    txtFN.SetFocus
    Exit Sub
    
    ElseIf txtDesg.Text = "" Then
    MsgBox "Enter User Designation"
    txtDesg.SetFocus
    Exit Sub
        
    ElseIf UType.Text = "Select Type" Then
    MsgBox "Select User Type"
    UType.SetFocus
    Exit Sub
          
    ElseIf txtR.Text = "" Then
    MsgBox "Please enter some remarks"
    txtR.SetFocus
    Exit Sub
    
    Else
        
        sql = "INSERT INTO Login (User,Password,Type,Name,Designation,Remarks) values('" & txtUser & "','" & txtPass & "','" & UType & "','" & txtFN & "','" & txtDesg & "','" & txtR & "')"
        'MsgBox sql
        conn.Execute sql
    
        cmdFirst.Enabled = True
        cmdLast.Enabled = True
        cmdNext.Enabled = True
        cmdPrevious.Enabled = True
        cmdDel.Enabled = True
        cmdEdit.Enabled = True
        cmdSave.Enabled = False
        txtUser.Enabled = False
        cmdNew.SetFocus
    End If
End Sub

Private Sub cmdFirst_Click()
    On Error Resume Next
    Rx = 0
    ShowUserData (SQLString)
    ShowUserGrid
    DataGrid1.Row = Rx
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    Rx = Rx - 1
    ShowUserData (SQLString)
    ShowUserGrid
    DataGrid1.Row = Rx
End Sub

Private Sub cmdRefresh_Click()
    SQLString = "SELECT * FROM Login ORDER BY User"
    Rx = 0
    ShowUserData (SQLString)
    ShowUserGrid
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    Rx = Rx + 1
    ShowUserData (SQLString)
    ShowUserGrid
    DataGrid1.Row = Rx
End Sub

Private Sub cmdLast_Click()
    On Error Resume Next
    Rx = xCount - 1
    ShowUserData (SQLString)
    ShowUserGrid
    DataGrid1.Row = Rx
End Sub

Private Sub cmdEdit_Click()
    sql = "UPDATE Login SET Password='" & txtPass.Text & "',Type='" & UType.Text & "',Name='" & txtFN.Text & "',Designation='" & txtDesg.Text & "',Remarks='" & txtR.Text & "' Where User='" & txtUser.Text & "'"
    conn.Execute sql
    ShowUserData (SQLString)
    Set DataGrid1.DataSource = RsUserGrid
    ShowUserGrid
    DataGrid1.Row = Rx
End Sub

Private Sub cmdDel_Click()
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorType = adOpenDynamic
    rsTemp.LockType = adLockOptimistic
    rsTemp.CursorLocation = adUseServer
    sql = "DELETE FROM Login Where User='" & txtUser.Text & "'"
 
    If MsgBox("Are you sure that you want to Delete this record?", vbYesNo + vbDefaultButton2 + vbCritical, "Confirm Delete") = vbNo Then
        Set rsTemp = Nothing
        Exit Sub
    End If
    rsTemp.Open sql, conn, adOpenKeyset
    Set rsTemp = Nothing
    
    Rx = Rx - 1
    Set DataGrid1.DataSource = RsUserGrid
    ShowUserGrid
    If (Rx <> 0) Then DataGrid1.Row = Rx

End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub ClearText()
'Clear TextBoxes
    txtUser.Text = ""
    txtPass.Text = ""
    txtFN.Text = ""
    txtDesg.Text = ""
    txtR.Text = ""
End Sub
