Attribute VB_Name = "ModuleS"
'çàïóñê îáíîâëåíèÿ
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
dwFlags As Long
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
dwProcessID As Long
dwThreadID As Long
End Type
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const NORMAL_PRIORITY_CLASS = &H20&
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

Sub GoObnowl()
 On Error GoTo ErrorMessage
If Dir$(App.Path + "\updater.exe", vbNormal) = "" Then
    MsgBox "Â êàòàëîãå ñ ïðîãðàììîé îòñóòñòâóåò ìîäóëü îáíîâëåíèÿ." + vbCrLf + " Ïåðåóñòàíîâèòå ïðîãðàììó", vbCritical, pr
    Exit Sub
Else

    Dim vf As Long
    vf = Shell(App.Path + "\updater.exe", vbNormalFocus)
        
End If
End
Exit Sub
ErrorMessage:
MsgBox "" + Error$
LogPrint "Îøèáêà çàïóñêà îáíîâëÿëêè-" + Error$


End Sub
