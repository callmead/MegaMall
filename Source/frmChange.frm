VERSION 5.00
Begin VB.Form frmChange 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   ":: CHANGE PASSWORD :."
   ClientHeight    =   2220
   ClientLeft      =   5025
   ClientTop       =   4965
   ClientWidth     =   6000
   Icon            =   "frmChange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmChange.frx":0ECA
   ScaleHeight     =   2220
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3150
      TabIndex        =   4
      Top             =   1680
      Width           =   2280
   End
   Begin VB.TextBox txtCP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.TextBox txtNP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   660
      Width           =   2655
   End
   Begin VB.TextBox txtNP2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton mnChange 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   450
      TabIndex        =   3
      Top             =   1710
      Width           =   2280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Password"
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
      Left            =   270
      TabIndex        =   7
      Top             =   300
      Width           =   2085
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
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
      Left            =   270
      TabIndex        =   6
      Top             =   720
      Width           =   2085
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Conform New Password"
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
      TabIndex        =   5
      Top             =   1140
      Width           =   2595
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancel_LostFocus()
    txtCP.SetFocus
End Sub

Private Sub Form_Load()
    txtNP.Text = ""
    txtNP2.Text = ""
    txtCP.Text = ""
End Sub

Private Sub mnChange_Click()
    GetPword
    ChangePword
End Sub
Private Sub GetPword()
    Dim sql As String
    sql = "SELECT * FROM Login WHERE User='" + UserName + "'"
    RsLogin.Open sql, conn, adOpenStatic, adLockReadOnly, adCmdText
    
    If RsLogin.EOF = True Then
        RsLogin.Close
        Set RsLogin = Nothing
        MsgBox "USER NOT FOUND IN DATABASE !!!", vbCritical, "POS"
        Exit Sub
    Else
        Pass = RsLogin!Password
    End If
    RsLogin.Close
    Set RsLogin = Nothing
End Sub
Private Sub ChangePword()
    If (txtCP.Text <> Pass) Then
        MsgBox "NOT A VALID CURRENT PASSWORD !!!", vbCritical, "POS"
        txtCP.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    Else
        If (txtNP.Text = txtNP2.Text) Then
            sql = "UPDATE Login SET Password='" & txtNP.Text & "' Where User='" & UserName & "'"
            conn.Execute sql
        
            MsgBox "Your Password has been changed, Remember to login with new password next time...", vbInformation, "Conformation"
            Unload Me
        Else
            MsgBox "Conform New Password does not match the New Password given!!!", vbInformation, "Change Password"
            txtNP.SetFocus
            SendKeys "{Home}+{End}"
            txtNP2.Text = ""
            Exit Sub
        End If
    End If
End Sub
Private Sub txtCP_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtNP_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtNP2_GotFocus()
    SendKeys "{Home}+{End}"
End Sub

Private Sub txtNP2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ChangePword
    End If
End Sub
