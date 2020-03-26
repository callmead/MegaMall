VERSION 5.00
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "sysinfo.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8985
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0BC2
   ScaleHeight     =   6375
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   4200
      Top             =   5760
   End
   Begin VB.Timer TmPB 
      Interval        =   140
      Left            =   3720
      Top             =   5760
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   4680
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5880
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblTitle2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mega Shopping Mall"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   2160
      TabIndex        =   6
      Top             =   480
      Width           =   6165
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Point of Sale"
      BeginProperty Font 
         Name            =   "Pristina"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   3600
      TabIndex        =   5
      Top             =   2160
      Width           =   4245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2007 Mega Mall"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   6120
      TabIndex        =   4
      Top             =   5640
      Width           =   2565
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Software Version 1.0.0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6120
      TabIndex        =   3
      Top             =   5880
      Width           =   2565
   End
   Begin VB.Label lblWarning 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":1608D
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   3720
      TabIndex        =   1
      Top             =   3480
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   5640
      Width           =   3015
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************************

Dim IP, Host, OS, Ver, Build, DiskSerialNo As String
'Getting IP
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const ERROR_SUCCESS       As Long = 0
Private Const WS_VERSION_REQD     As Long = &H101
Private Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD    As Long = 1
Private Const SOCKET_ERROR        As Long = -1

Private Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Private Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Private Declare Function WSAGetLastError Lib "wsock32" () As Long

Private Declare Function WSAStartup Lib "wsock32" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "wsock32" () As Long

Private Declare Function gethostname Lib "wsock32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
   
Private Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)


Private Sub Form_Load()
    'For Mouse
    Me.MousePointer = vbHourglass
    Host = GetIPHostName()
    IP = "IP Address: " + GetIPAddress()
    
    DiskSerial

    'Getting System Info
    Select Case SysInfo.OSPlatform
        Case 0
                OS = "OS Platform: Unknown 32-Bit Windows"
        Case 1
                OS = "OS Platform: Windows 95"
        Case 2
                OS = "OS Platform: Windows NT"
    End Select
    Ver = "OS Version: " & SysInfo.OSVersion
    Build = "OS Build: " & SysInfo.OSBuild
        
    'Checking.........
    'MsgBox "H: " + Host + " DS: " + DiskSerialNo
    
    'If (Host <> "hellboy") Then
    If (Host <> "PHOENIX") Then
    'If (Host <> "MM_1") Then
        MsgBox "Software not Registered for this Computer!!! Please contact your DEVELOPER.", vbCritical, ":: | :: POS :: | :."
        End
        Exit Sub
    End If
    'If (DiskSerialNo <> "403260548") Then
    If (DiskSerialNo <> "342658779") Then
    'If (DiskSerialNo <> "994183920") Then
        MsgBox "Software not Registered for this Computer!!! Please contact your DEVELOPER.", vbCritical, ":: | :: POS :: | :."
        End
        Exit Sub
    End If
    '****************************************************

    'Cursor
    Dim hNewCursor As Long
    hNewCursor = LoadCursorFromFile(App.Path & "\cursor.ani")
    hOldCursor = SetClassLong(hwnd, GCL_HCURSOR, hNewCursor)

End Sub


Private Sub Timer2_Timer()
    Unload Me
    frmLogin.Show
    
End Sub

Private Function GetIPAddress() As String

   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim Host      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
    
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
    
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   CopyMemory Host, lpHost, Len(Host)
   CopyMemory dwIPAddr, Host.hAddrList, 4
   
  'create an array to hold the result
   ReDim tmpIPAddr(1 To Host.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, Host.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To Host.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
  
  'the routine adds a period to the end of the
  'string, so remove it here
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function

Private Function GetIPHostName() As String

    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function

Private Function HiByte(ByVal wParam As Integer) As Byte
  
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function

Private Function LoByte(ByVal wParam As Integer) As Byte

  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function

Private Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub

Private Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
 
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function
'********************************************************************

Private Sub DiskSerial()
'Reference-Microsoft Scripting Runtime
    Dim fldr As Folder
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Set drv = fso.GetDrive(fso.GetDriveName("C:"))
    
    DiskSerialNo = "" & drv.SerialNumber
    'MsgBox "Available space: " & FormatNumber(drv.AvailableSpace / 1024, 0) & " KB" & " Drive Letter: " & drv.DriveLetter & " Drive Type: " & drv.DriveType & " File System: " & drv.FileSystem & " Drive free space: " & FormatNumber(drv.FreeSpace / 1024, 0) & " KB" & " Drive is ready: " & drv.IsReady & " Drive path: " & drv.Path & " Root folder: " & drv.RootFolder & " Serial number: " & drv.SerialNumber & " Share name: " & drv.ShareName & " Total size: " & FormatNumber(drv.TotalSize / 1024, 0) & " KB" & " Volume  name : " & drv.VolumeName
    'Place a List Control on form to see the required Details
    'With List1
    '    .AddItem "Available space: " & FormatNumber(drv.AvailableSpace / 1024, 0) & " KB"
    '    .AddItem "Drive letter: " & drv.DriveLetter
    '    .AddItem "Drive type: " & drv.DriveType
    '    .AddItem "Drive file system: " & drv.FileSystem
    '    .AddItem "Drive free space: " & FormatNumber(drv.FreeSpace / 1024, 0) & " KB"
    '    .AddItem "Drive is ready: " & drv.IsReady
    '    .AddItem "Drive path: " & drv.Path
    '    .AddItem "Root folder: " & drv.RootFolder
    '    .AddItem "Serial number: " & drv.SerialNumber
    '    .AddItem "Share name: " & drv.ShareName
    '    .AddItem "Total size: " & FormatNumber(drv.TotalSize / 1024, 0) & " KB"
    '    .AddItem "Volume  name : " & drv.VolumeName
    'End With
    
End Sub

Private Sub TmPB_Timer()
    PB1.Value = PB1.Value + 5
    If (PB1.Value = PB1.Max) Then
        TmPB.Enabled = False
    End If
End Sub
