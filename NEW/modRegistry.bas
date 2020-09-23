Attribute VB_Name = "modRegistry"
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long



Enum SEARCH_FLAGS
    KEY_NAME = 0
    VALUE_NAME = 1
    VALUE_VALUE = 2
    WHOLE_STRING = 4
End Enum

Enum FOUND_WHERE
    FOUND_IN_KEY_NAME
    FOUND_IN_VALUE_NAME
    FOUND_IN_VALUE_VALUE
End Enum

Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Const KEY_READ = &H20019    ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
' SYNCHRONIZE))

Private Const ERROR_SUCCESS = 0&
Private Const ERR_MORE_DATA = 234&
Private Const ERROR_NO_MORE_ITEMS = 259&

Enum REG
    HKEY_CURRENT_USER = &H80000001
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Enum TypeStringValue
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_MULTI_SZ = 7
End Enum


Enum TypeBase
    TypeHexadecimal
    TypeDecimal
End Enum

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.


Public Function DeleteKey(hKey As REG, SubKey As String) As Long

    On Error Resume Next
    DeleteKey = RegDeleteKey(hKey, SubKey)
    RegCloseKey ret
    
End Function

Public Function DeleteValue(hKey As REG, SubKey As String, lpValName As String) As Long
    
    Dim ret As Long
    On Error Resume Next
    RegOpenKey hKey, SubKey, ret
    DeleteValue = RegDeleteValue(ret, lpValName)
    RegCloseKey ret
    
End Function

 Public Function CreateDwordValue(hKey As REG, SubKey As String, strValueName As String, dwordData As Long) As Long
    Dim ret As Long
    On Error Resume Next
    RegCreateKey hKey, SubKey, ret
    CreateDwordValue = RegSetValueEx(ret, strValueName, 0, REG_DWORD, dwordData, 4)
    RegCloseKey ret
    
End Function

Public Function CreateStringValue(hKey As REG, SubKey As String, RTypeStringValue As TypeStringValue, strValueName As String, strdata As String) As Long
    Dim ret As Long
    On Error Resume Next
    RegCreateKey hKey, SubKey, ret
    CreateStringValue = RegSetValueEx(ret, strValueName, 0, RTypeStringValue, ByVal strdata, Len(strdata))
    RegCloseKey ret
    
End Function

Public Function ReadValue(hKey As REG, SubKey As String, strValueName As String) As String

    Dim RootKey As Long
    Dim isi As String
    Dim lDataBufSize As Long
    Dim lValueType As Long
    On Error Resume Next
    X = RegOpenKey(hKey, SubKey, RootKey)
    ret = RegQueryValueEx(RootKey, strValueName, 0, lValueType, 0, lDataBufSize)
    isi = String(lDataBufSize, Chr$(0))
    ret = RegQueryValueEx(RootKey, strValueName, 0, 0, ByVal isi, lDataBufSize)
    
    ReadValue = Left$(isi, InStr(1, isi, Chr$(0)) - 1)
    RegCloseKey RootKey

End Function


'used for LoadRegistry

Function getstring(hKey As Long, strPath As String, strValue As String)
'----------------------------------------------------------------------------
'Argument       :   Handlekey, path from the root , Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
'----------------------------------------------------------------------------

    Dim ret
    'Open  key
    RegOpenKey hKey, strPath, ret
    'Get content
    getstring = RegQueryStringValue(ret, strValue)
    'Close the key
    RegCloseKey ret
End Function

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
'----------------------------------------------------------------------------
'Argument       :   Handlekey, Name of the Value in side the key
'Return Value   :   String
'Function       :   To fetch the value from a key in the Registry
'Comments       :   on Success , returns the Value else empty String
    '----------------------------------------------------------------------------
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
    
            strBuf = String(lDataBufSize, Chr$(0))
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
    
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strdata As Integer
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strdata, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strdata
            End If
         ElseIf lValueType = REG_DWORD Then
           
            'retrieve the key's value
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strdata, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strdata
            End If
            
        End If
    End If
End Function
'=========================

Public Sub savekey(hKey As Long, strPath As String)
Dim keyhand&
r = RegCreateKey(hKey, strPath, keyhand&)
r = RegCloseKey(keyhand&)
End Sub
Function getdword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long

r = RegOpenKey(hKey, strPath, keyhand)

 ' Get length/data type
lDataBufSize = 4
    
lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        getdword = lBuf
    End If
'Else
'    Call errlog("GetDWORD-" & strPath, False)
End If

r = RegCloseKey(keyhand)
    
End Function




Public Sub savestring(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(hKey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub



Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
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


Public Function DeleteStartup(lPredefinedKey As Long, sKeyName As String, sValueName As String)

       Dim lRetVal As Long      'result of the SetValueEx function
       Dim hKey As Long         'handle of open key

       lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
       lRetVal = RegDeleteValue(hKey, sValueName)
       RegCloseKey (hKey)
       
End Function
Public Sub CleanReg()

    On Error Resume Next
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
        CreateDwordValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegedit", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", 0
      CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictRun", 0
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
    
    'DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
    'HKEY_CURRENT_USER\
    'DeleteValue HKEY_USERS, ".DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
    'HKEY_USERS\.DEFAULT\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2
    
    CreateStringValue HKEY_CLASSES_ROOT, "exefile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "piffile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "batfile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "comfile\shell\open\command", REG_SZ, "", Chr(&H22) & "%1" & Chr(&H22) & " %*"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", REG_SZ, "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "Userinit", GetSystemPath & "userinit.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", REG_SZ, "Debugger", Chr(&H22) & Left(GetWindowsPath, 3) & "Program Files\Microsoft Visual Studio\Common\MSDev98\Bin\msdev.exe" & Chr(&H22) & " -p %ld -e %ld"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", REG_SZ, "Auto", "0"
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp", "Disabled", 0
    CreateStringValue HKEY_CLASSES_ROOT, "exefile", REG_SZ, "", "Application"
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
    
End Sub
