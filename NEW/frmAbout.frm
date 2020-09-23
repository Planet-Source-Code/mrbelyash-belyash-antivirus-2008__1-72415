VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Î ïðîãðàììå..."
   ClientHeight    =   5370
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3706.469
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ListView ListView1 
      Height          =   1935
      Left            =   135
      TabIndex        =   8
      Top             =   2235
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   3413
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ôàéë"
         Object.Width           =   3081
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Èíôî/Êîë-âî çàïèñåé â áàçå"
         Object.Width           =   2705
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "CRC"
         Object.Width           =   3456
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4155
      TabIndex        =   0
      Top             =   4440
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   4170
      TabIndex        =   1
      Top             =   4890
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "mrbelyash@yandex.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2490
      TabIndex        =   9
      Top             =   1890
      Width           =   3045
   End
   Begin VB.Image Image1 
      Height          =   1620
      Left            =   90
      Picture         =   "frmAbout.frx":C84A
      Stretch         =   -1  'True
      Top             =   210
      Width           =   1950
   End
   Begin VB.Label Label3 
      Caption         =   "mrbelyash@rambler.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2490
      TabIndex        =   7
      Top             =   1560
      Width           =   3045
   End
   Begin VB.Label Label2 
      Caption         =   "www.mrbelyash.ucoz.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2490
      TabIndex        =   6
      Top             =   1260
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "www.mrbelyash.narod.ru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2490
      TabIndex        =   5
      Top             =   930
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   0
      X2              =   5224.884
      Y1              =   2950.68
      Y2              =   2950.68
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Left            =   2160
      TabIndex        =   3
      Top             =   60
      Width           =   3525
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   14.086
      X2              =   5224.884
      Y1              =   2950.68
      Y2              =   2950.68
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   3090
      TabIndex        =   4
      Top             =   570
      Width           =   1875
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":1390C
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   165
      TabIndex        =   2
      Top             =   4440
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public countZap As Long

Dim a(10) As Long
Public countZapOLD As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
    lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
'-----------------------------------------
' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Const READONLY = &H1
Const HIDDEN = &H2
Const SYSTEM = &H4
Const DIRECTORY = &H10
Const ARCHIVE = &H20
Const NORMAL = &H80
Const COMPRESSED = &H800
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long


Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long


Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long


Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)


Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F For version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
    End Type

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    Call infoModul
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Label1_Click()
On Error GoTo 100
Call ShellExecute(0, "Open", "http://www.mrbelyash.narod.ru", "", "", 1)
Exit Sub
100:
End Sub

Private Sub Label2_Click()
On Error GoTo 100
Call ShellExecute(0, "Open", "http://www.mrbelyash.ucoz.ru", "", "", 1)
Exit Sub
100:
End Sub

Private Sub Label3_Click()
On Error GoTo 100
Call ShellExecute(0, "Open", "mailto:" + "mrbelyash@rambler.ru" + "?Subject=" + "Ïðîáëåìû ïðè ðàáîòå ñ Belyash Shield", "", "", 1)
Exit Sub
100:
End Sub
Public Function CheckFileVersion(FilenameAndPath As Variant) As Variant
    On Error GoTo HandelCheckFileVersionError
    Dim lDummy As Long, lsize As Long, rc As Long
    Dim lVerbufferLen As Long, lVerPointer As Long
    Dim sBuffer() As Byte
    Dim udtVerBuffer As VS_FIXEDFILEINFO
    Dim ProdVer As String
    lsize = GetFileVersionInfoSize(FilenameAndPath, lDummy)
    If lsize < 1 Then Exit Function
    ReDim sBuffer(lsize)
    rc = GetFileVersionInfo(FilenameAndPath, 0&, lsize, sBuffer(0))
    rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
    MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    '**** Determine Product Version number *
    '     ***
    ProdVer = Format$(udtVerBuffer.dwProductVersionMSh) _
    & "." & Format$(udtVerBuffer.dwProductVersionMSl) _
        & "." & Format$(udtVerBuffer.dwProductVersionLSh) _
    & "." & Format$(udtVerBuffer.dwProductVersionLSl)

    
    CheckFileVersion = ProdVer
    Exit Function
HandelCheckFileVersionError:
    CheckFileVersion = "N/A"
    Exit Function
End Function
Sub infoModul()
'If Dir$(App.Path + "\BGAntiVirus.exe", vbNormal) <> "" Then
    'Text1.Text = Text1.Text + "BGAntiVirus.exe" + "-" + CheckFileVersion(App.Path + "\BGAntiVirus.exe") + vbCrLf

countZap = 0
countZapOLD = 0
Dim sNext As String
sNext = Dir$(App.path + "\*.exe")
While sNext <> ""

    getinexe (sNext)
    sNext = Dir$
Wend

Dim sNext1 As String
sNext1 = Dir$(App.path + "\*.bVb")
While sNext1 <> ""

    kolwoVirusowAll (sNext1)
    sNext1 = Dir$
Wend
Dim baZE As String
kolwoVirusow
Dim nm As Long
Dim z As Byte

For z = 0 To 10
nm = nm + a(z)
Next z


Dim X As ListItem
Set X = Me.ListView1.ListItems.Add(, , "Âñåãî çàïèñåé")
X.SubItems(1) = nm + countZap


End Sub

Sub kolwoVirusowAll(baZE As String)
On Error GoTo 100
Static i As Integer
i = i + 1
Dim miNumBase As Integer
miNumBase = FreeFile
Dim sMD15 As String
Open App.path + "\" + baZE For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD15
        a(i) = a(i) + 1
     Wend
 Close #miNumBase
Dim CRC As New clsCRC
CRC.BuildTable
Dim X As ListItem
Set X = Me.ListView1.ListItems.Add(, , baZE)
X.SubItems(1) = a(i)
X.SubItems(2) = CRC.GetCRC(App.path + "\" + baZE)
Set X = Nothing
 Exit Sub
100:
 LogPrint "" + Error$
End Sub




Sub getinexe(s As String)
Dim CRC As New clsCRC
CRC.BuildTable
Dim X As ListItem
Set X = Me.ListView1.ListItems.Add(, , s)
X.SubItems(1) = CheckFileVersion(s)
X.SubItems(2) = CRC.GetCRC(CStr(s))
Set X = Nothing
End Sub

Sub kolwoVirusow()
On Error GoTo 100
If Dir$(App.path + "\old.bmd", vbNormal) = "" Then
    Exit Sub
End If
Dim miNumBase As Integer
miNumBase = FreeFile
Dim sMD15 As String
Open App.path + "\old.bmd" For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD15
        countZap = countZap + 1
     Wend
 Close #miNumBase
Dim CRC As New clsCRC
CRC.BuildTable
Dim X As ListItem
Set X = Me.ListView1.ListItems.Add(, , "old.bmd")
X.SubItems(1) = countZap
X.SubItems(2) = CRC.GetCRC(App.path + "\old.bmd")
Set X = Nothing
 Exit Sub
100:
 LogPrint "" + Error$
End Sub

Private Sub Label4_Click()
On Error GoTo 100
Call ShellExecute(0, "Open", "mailto:" + "mrbelyash@yandex.ru" + "?Subject=" + "Ïðîáëåìû ïðè ðàáîòå ñ Belyash Shield", "", "", 1)
Exit Sub
100:
End Sub
