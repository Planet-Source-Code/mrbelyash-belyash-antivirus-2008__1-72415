Attribute VB_Name = "modMain"
Option Explicit

'CRC32 variable
Public mast_md5 As Boolean
Public nostartUP As Boolean
Public noAU As Boolean

Public dbg1 As Boolean
Public checkmemory As Boolean
Public CRC As New clsCRC
'file size to be scanned virus
Public FileSize As Long
'declare Virus Def & info
Public VSig() As VirusSig
Public base() As VirusSig1
Public VSInfo As VS_Info
'declare variable for scan reg extensions
Public intSettingRegOption As Integer
Public strScanRegExt As String
'for faster DoEvents
Declare Function GetInputState Lib "user32.dll" () As Long
'declaration for Classes of Registry
Public Const HKEY_ALL = &H0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

'new ACCESS KEY for delete startup registry
Private Const KEY_ALL_ACCESS = &H3F '((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

'new DataType for Virus Signature
Public Type VirusSig

    Name As String
    Type As String
    Value As String
    Action As String
    ActtionVal As String
    
End Type
Public Type VirusSig1

    Name As String
    Type As String
    Value As String
    Action As String
    ActtionVal As String
    
End Type
'new DataType for Virus Signature Info
Public Type VS_Info
    
    VirusCount As Long
    LastUpdate As Date
    
End Type

' Transparency Constants
'Const LWA_COLORKEY = &H3
Const LWA_ALPHA = &H3
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
'Private Const HWND_TOPMOST = -1
'Private Const SWP_SHOWWINDOW = &H40
'Private Const SWP_NOOWNERZORDER = &H200
Dim ret As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Type SYSTEM_INFO
dwOemID As Long
dwPageSize As Long
lpMinimumApplicationAddress As Long
lpMaximumApplicationAddress As Long
dwActiveProcessorMask As Long
dwNumberOrfProcessors As Long
dwProcessorType As Long
dwAllocationGranularity As Long
dwReserved As Long
End Type
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Private Type MEMORYSTATUS
dwLength As Long
dwMemoryLoad As Long
dwTotalPhys As Long
dwAvailPhys As Long
dwTotalPageFile As Long
dwAvailPageFile As Long
dwTotalVirtual As Long
dwAvailVirtual As Long
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)


Sub SystemInformation()
Dim msg As String ' Status information.
Dim NewLine As String ' New-line.
Dim ret As Integer ' OS Information
Dim ver_major As Integer ' OS Version
Dim ver_minor As Integer ' Minor Os Version
Dim Build As Long ' OS Build
NewLine = Chr(13) + Chr(10) ' New-line.
' Get operating system and version.
Dim verinfo As OSVERSIONINFO
verinfo.dwOSVersionInfoSize = Len(verinfo)
ret = GetVersionEx(verinfo)
If ret = 0 Then
MsgBox "Îøèáêà ïîëó÷åíèÿ âåðñèè Âèíäîâñ"
End
End If
'MsgBox verinfo.dwPlatformId
Select Case verinfo.dwPlatformId
Case 2
GoTo 2
Case Else
MsgBox "Ïðîãðàììà êîððåêòíî ðàáîòàåò òîëüêî ïîä Windows XP", vbExclamation
End
End Select
2:
End Sub
'startup sub
Sub Main()
Call SystemInformation 'ïðîâåðèì ÕÐ ýòî èëè íåò

'åñëè íåò òî âûõîäèì
    'If Year(Now()) = 2007 Then 'can be used only in 2007 if enable
        Dim FirstStart As String
        FirstStart = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPath")
        'check if it is ever started
        If FirstStart = "" Then
            'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPath", App.Path
            CreateStringValue HKEY_CURRENT_USER, "Software\BGAntivirus", 1, "AppPath", App.Path & "\" & App.exeName
            'default app setting    'default setting cmd
            CreateStringValue HKEY_CURRENT_USER, "Software\BGAntivirus", 1, "RegExt", "OCX, DLL, EXE, VBS, SYS, VXD"
            
            CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate", 10
            CreateStringValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", 1, "RegExt", "OCX, DLL, EXE, VBS, SYS, VXD"
            
     CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Quarantine", val(1)
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log", val(1)
    SaveSetting App.exeName, "Options", "logSize", "100024"
    SaveSetting App.exeName, "Options", "Priority", "1"
     CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", val(1)

        Else
            'check if previous open is the same as the current
            If FirstStart <> App.Path & "\" & App.exeName Then
                'change key if different
                CreateStringValue HKEY_CURRENT_USER, "Software\BGAntivirus", 1, "AppPath", App.Path & "\" & App.exeName
            End If
        End If
        
Dim a2 As String
a2 = CStr(GetSetting(App.exeName, "Options", "Priority", "2"))
Select Case a2
Case "1"
modWinFunctions.SetPriorityClass GetCurrentProcess(), IDLE_PRIORITY_CLASS
Case "2"
modWinFunctions.SetPriorityClass GetCurrentProcess(), NORMAL_PRIORITY_CLASS
Case "3"
modWinFunctions.SetPriorityClass GetCurrentProcess(), HIGH_PRIORITY_CLASS
Case "4"
modWinFunctions.SetPriorityClass GetCurrentProcess(), REALTIME_PRIORITY_CLASS
End Select
Dim iParam As Long
Dim sData() As String
checkmemory = True
nostartUP = True
noAU = True
'trigger the command line parser
CL_Get sData
dbg1 = False
mast_md5 = False
'enumerate the resulting parameters
For iParam = LBound(sData) To UBound(sData)
    Debug.Print Format(iParam + 1, "00") & ": " & sData(iParam)
    'MsgBox "" + Format(iParam + 1, "00") & ": " & sData(iParam)
    If sData(iParam) = "/dbg" Then
        dbg1 = True
        
        
    End If
    If sData(iParam) = "/md5" Then
        'âåñòè ëîã md5 õåøåé ïðîâåðÿåìûõ ôàéëîâ
        mast_md5 = True
    End If
    If sData(iParam) = "/go" Then
        'íå ïðîâåðÿòü ïðè âûáîðî÷íîì ñêàíèðîâàíèè ïðîöåññû
        checkmemory = False
    End If
    
        If sData(iParam) = "/ns" Then
        'íå ïðîâåðÿòü ñòàðòàïû
        nostartUP = False
    End If
If sData(iParam) = "/au" Then
        'íå ïðîâåðÿòü àâòîðàíû
        noAU = False
    End If
    
   
Next iParam
        
        
        frmSplash.Show vbModal
        
    'Else
    '    MsgBox "Beta version is expired. Please find an official release of this software.", vbInformation, "Belyash AntiTrojan 2008 beta Expire"
    'End If
    
End Sub

Function CalculateTime(ByVal interval As Single) As String

    Dim sec As Long, mn As Long, hr As Long
    hr = Int(CSng(interval * 24))
    mn = Int(CSng(interval * 24 * 60))
    sec = Int(CSng(interval * 24 * 60 * 60))
    CalculateTime = hr & " hr " & (mn - (hr * 60)) & " mn " & (sec - (mn * 60)) & " sec."
    
End Function

'Reverse a string
Public Function ReverseString(TheString As String) As String
    Dim i As Integer
    For i = 1 To Len(TheString)
        ReverseString = ReverseString & Mid(Right$(TheString, i), 1, 1)
    Next
End Function

'Returns the long value of the string entered as ROOT_KEYS
Public Function GetClassKey(cls As String) As Variant
    Select Case cls
    Case "HKEY_ALL"
        GetClassKey = HKEY_ALL
    Case "HKEY_CLASSES_ROOT"
        GetClassKey = HKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetClassKey = HKEY_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetClassKey = HKEY_LOCAL_MACHINE
    Case "HKEY_USERS"
        GetClassKey = HKEY_USERS
    Case "HKEY_PERFORMANCE_DATA"
        GetClassKey = HKEY_PERFORMANCE_DATA
    Case "HKEY_CURRENT_CONFIG"
        GetClassKey = HKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        GetClassKey = HKEY_DYN_DATA
    End Select
End Function

'Another Delete Registry Key function
Public Sub DeleteRegKey(ROOTKEYS As ROOT_KEYS, Path As String, sKey As String)
    
    Dim ValKey As String
    Dim SecKey As String, SlashPos As Single
    SlashPos = InStrRev(Path, "\", Compare:=vbTextCompare)
    SecKey = Left(Path, SlashPos - 1)    'This will retreive the section key that I need
    ValKey = Right(Path, Len(Path) - SlashPos)    'This will retreive the ValueKey that I need to delete
    DeleteRegKey2 ROOTKEYS, SecKey, ValKey

End Sub
    
'Another Delete Registry Key function
Public Sub DeleteRegKey2(hKey As ROOT_KEYS, strPath As String, strValue As String)
    Dim ret
    RegCreateKey hKey, strPath, ret
    RegDeleteValue ret, strValue
    RegCloseKey ret
End Sub

Public Function DeleteStartup(lPredefinedKey As Long, sKeyName As String, sValueName As String)

       Dim lRetVal As Long      'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
       
End Function
'Transparent Making area
Public Sub MakeTransparent(ByRef Frm As Form, ByVal alpha As Long)
    ret = GetWindowLong(Frm.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Frm.hwnd, GWL_EXSTYLE, ret
    'change ipAlpha for transparency
    SetLayeredWindowAttributes Frm.hwnd, 0, alpha, LWA_ALPHA
End Sub

