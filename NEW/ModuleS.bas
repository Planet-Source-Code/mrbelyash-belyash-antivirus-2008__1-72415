Attribute VB_Name = "ModuleS"
Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)


Private Type STARTUPINFO
cb As Long
lpReserved As String
lpDesktop As String
lpTitle As String
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwflags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Long
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type
Private Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessId As Long
dwThreadId As Long
End Type
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Public Sub ExecCmd(cmdline$)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
' Èíèöèàëèçèðóåì ñòðóêòóðó STARTUPINFO:
start.cb = Len(start)
' Çàïóñêàåì ïðèëîæåíèå:
ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
' Æäåì çàâåðøåíèÿ çàïóùåííîãî ïðèëîæåíèÿ:
ret& = WaitForSingleObject(proc.hProcess, INFINITE)
ret& = CloseHandle(proc.hProcess)
End Sub


Sub registry_mon()
Dim monR1 As Boolean
monR1 = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonReg"))
If monR1 = False Then
    Exit Sub
End If


On Error GoTo 100
Dim q1 As Long
Dim q2 As Long
Dim q3 As Long
q1 = Shell(App.Path + "\HKLMRO_Belyash.exe", vbHide)
'Sleep 1000

q2 = Shell(App.Path + "\HKLM_Belyash.exe", vbHide)
'Sleep 1000

q3 = Shell(App.Path + "\HKCURO_Belyash.exe", vbHide)
'Sleep 1000

q4 = Shell(App.Path + "\HKCU_Belyash.exe", vbHide)

Exit Sub
100:
MsgBox "ìîíèòîð ðååñòðà-" + Error, vbCritical

End Sub
