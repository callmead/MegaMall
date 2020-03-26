VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDataSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: DATA SELECTOR :."
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "frmDataSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDataSelect.frx":0ECA
   ScaleHeight     =   5070
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      HotImageList    =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Select"
            Object.ToolTipText     =   "Select"
            ImageKey        =   "Spell Check"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Top"
            Object.ToolTipText     =   "First Record"
            ImageKey        =   "Top"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Prior"
            Object.ToolTipText     =   "Previous Record"
            ImageKey        =   "Prior"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Next"
            Object.ToolTipText     =   "Next Record"
            ImageKey        =   "Next"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bottom"
            Object.ToolTipText     =   "Last Record"
            ImageKey        =   "Bottom"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   4095
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16744576
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   2475
      Top             =   2055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4B4D7
            Key             =   "Bell"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4B671
            Key             =   "Sort Ascending"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4B783
            Key             =   "Misc08"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4BA9D
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4BBAF
            Key             =   "Top"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4C0F1
            Key             =   "Prior"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4C1FB
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4C30D
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4C417
            Key             =   "Bottom"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDataSelect.frx":4C959
            Key             =   "Spell Check"
         EndProperty
      EndProperty
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   240
      X2              =   8400
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   4920
   End
End
Attribute VB_Name = "frmDataSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SetGrid
    ReturnValue = ""
End Sub

'Private Sub Form_Resize()
''On Error Resume Next
'    If Me.WindowState <> 1 Then
'        grdDataGrid.Height = Me.Height - (950)
'        grdDataGrid.Width = Me.Width - (250)
'    End If
'End Sub

Public Sub SetGrid()
    Set RsGrid = New ADODB.Recordset
    RsGrid.CursorLocation = adUseClient
    RsGrid.CursorType = adOpenStatic
    RsGrid.LockType = adLockReadOnly
    RsGrid.Open GridSQLString, conn
    Set grdDataGrid.DataSource = RsGrid
End Sub

Private Sub grdDataGrid_Click()
    On Error Resume Next
    With grdDataGrid
        .Col = SelectedField
        ReturnValue = .Text
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
        Case "Select"
            DataSelect
        Case "Bottom"
            LastRecord
        Case "Next"
            NextRecord
        Case "Prior"
            PreviousRecord
        Case "Top"
            FirstRecord
   End Select
End Sub

Public Sub DataSelect()
    
    If ParentForm = "frmSupplierAccountSID" Then
        frmSupplierAccount.txtSID.Text = ReturnValue
        
    End If
    If ParentForm = "frmSupplierAccountPO" Then
        frmSupplierAccount.txtPO.Text = ReturnValue
        
    End If
    If ParentForm = "frmCustomerAccountCID" Then
        frmCustomerAccount.txtCID.Text = ReturnValue
        
    End If
    If ParentForm = "frmCustomerAccountInv" Then
        frmCustomerAccount.txtInv.Text = ReturnValue
        
    End If
    
    If ParentForm = "frmPurchaseOrder" Then
        frmPurchaseOrder.txtSID.Text = ReturnValue
        'frmPurchaseOrder.txtDD.SetFocus
    End If
    If ParentForm = "frmReceivings" Then
        frmReceivings.txtPO.Text = ReturnValue
        'frmReceivings.txtQty.SetFocus
    End If
    If ParentForm = "frmReceivings1" Then
        frmReceivings.txtPID.Text = ReturnValue
        'frmReceivings.txtQty.SetFocus
    End If
    If ParentForm = "frmInvoice" Then
        frmInvoice.txtCID.Text = ReturnValue
        'frmReceivings.txtQty.SetFocus
    End If
    If ParentForm = "RptPO" Then
        'RptPO.T1.Text = ReturnValue
    End If
    If ParentForm = "RptGSale" Then
        'RptGSale.T1.Text = ReturnValue
    End If
    If ParentForm = "RptStock" Then
        'RptStock.t1.text=ReturnValue
    End If
    If ParentForm = "RptInv" Then
        RptInvoice.T1.Text = ReturnValue
    End If
    
    Unload Me
    
End Sub

Public Sub PreviousRecord()
    If RsGrid.RecordCount <> 0 Then
        If RsGrid.BOF Then
            RsGrid.MoveFirst
        Else
            RsGrid.MovePrevious
        End If
    End If
End Sub

Public Sub FirstRecord()
    If RsGrid.RecordCount <> 0 Then
        RsGrid.MoveFirst
    End If
End Sub

Public Sub LastRecord()
    If RsGrid.RecordCount <> 0 Then
        RsGrid.MoveLast
    End If
End Sub

Public Sub NextRecord()
    If RsGrid.RecordCount <> 0 Then
        If RsGrid.EOF Then
            RsGrid.MoveLast
        Else
            RsGrid.MoveNext
        End If
    End If
End Sub
