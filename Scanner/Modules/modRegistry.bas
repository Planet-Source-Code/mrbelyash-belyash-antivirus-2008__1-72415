Attribute VB_Name = "modRegistry"
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
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey&, ByVal lpClass$, lpcbClass&, ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, lpcbMaxClassLen&, lpcValues&, lpcbMaxValueNameLen&, lpcbMaxValueLen&, lpcbSecurityDescriptor&, lpftLastWriteTime As Any) As Long
Public Sub CleanReg()
   On Error Resume Next

   Dim z77 As String
    z77 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanReg"))
   If z77 = 0 Then
   Exit Sub
 
   End If
   
   
   
   Dim z1 As String
    z1 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    If z1 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableRegistryTools(ïîäîçðåíèå íà äåéñòâèå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
        CreateDwordValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
        'LogPrint "Èñïðàâëåíà âåòêà-HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"
    End If
    
    ' CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
        Dim z2 As String
    z2 = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    If z2 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableRegistryTools(ïîäîçðåíèå íà äåéñòâèå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
         CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
         LogPrint "Èñïðàâëåíà âåòêà-HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegistryTools"
    End If
    Dim z3 As String
    z3 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegedit"))
    If z3 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableRegedit(ïîäîçðåíèå íà äåéñòâèå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
                    CreateDwordValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegedit", 0
            LogPrint "Èñïðàâëåíà âåòêà-HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\DisableRegedit"
    End If
    
    Dim z4 As String
    z4 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"))
    If z4 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableTaskMgr(ïîäîçðåíèå íà äåéñòâèå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
        CreateDwordValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
        LogPrint "Èñïðàâëåíà âåòêà-HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr"
    End If
    
        Dim z5 As String
    z5 = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr"))
    If z5 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableTaskMgr(ïîäîçðåíèå íà äåéñòâèå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
        CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", 0
        LogPrint "Èñïðàâëåíà âåòêà-HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\DisableTaskMgr"
    End If
    Dim z6 As String
    z6 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"))
    If z6 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableCMD(ïîäîçðåíèå íà äåéñòâèå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
        CreateDwordValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", 0
        LogPrint "Èñïðàâëåíà âåòêà-HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\DisableCMD"
    End If
 
  
    Dim z7 As String
    z7 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictRun"))
    If z7 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "RestrictRun(ïîäîçðåíèå íà äåéñòâèå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
        CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "RestrictRun", 0
        LogPrint "Èñïðàâëåíà âåòêà-HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictRun"
    End If
   
    
 
   
      Dim z9 As String
    z9 = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"))
    If z9 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableConfig(ïîäîçðåíèå íà íàëè÷èå òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
        CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
        LogPrint "Èñïðàâëåíà âåòêà-HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore\DisableConfig"
    End If
  
    
          Dim z10 As String
    z10 = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"))
    If z10 <> "0" Then
        Set X = frmMain.lvVirusFound.ListItems.Add(, , "DisableMSI(ïîäîçðåíèå íà äåéñòâèÿ òðîÿíà)", 1, 5)
        X.SubItems(1) = "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Installer"
        X.SubItems(2) = "Èñïðàâëåíî"
        Set X = Nothing
         CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
         LogPrint "Èñïðàâëåíà âåòêà-HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Installer\DisableMSI"
    End If
  

    
    
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

    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", 0
   
   Dim lCount As Long, bCount As Long
    Dim i As Long, hKey As Long
    lCount = 0
    'clear previous list in registry
    'from allow list
    hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount
        'use delete registry (function in startup)
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2", CStr(i)
        Debug.Print "111111111111"
    Next
    
    Dim lCount1 As Long, bCount1 As Long
    Dim i1 As Long, hKey1 As Long
    lCount1 = 0
    'clear previous list in registry
    'from allow list
    hKey1 = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2")
    lCount1 = GetCount(hKey1, Values)
    For i = 0 To lCount1
        'use delete registry (function in startup)
        DeleteStartup GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2", CStr(i1)
    Next
End Sub

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
R = RegCreateKey(hKey, strPath, keyhand&)
R = RegCloseKey(keyhand&)
End Sub
Function getdword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim R As Long
Dim keyhand As Long

R = RegOpenKey(hKey, strPath, keyhand)

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

R = RegCloseKey(keyhand)
    
End Function




Public Sub savestring(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim R As Long
R = RegCreateKey(hKey, strPath, keyhand)
R = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
R = RegCloseKey(keyhand)
End Sub



Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim R As Long
    R = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    R = RegCloseKey(keyhand)
End Function


