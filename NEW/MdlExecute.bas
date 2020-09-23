Attribute VB_Name = "MdlExecute"
Option Explicit
'Used to hide the dos window (so it doesnt interupt
'anything)
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
'Finds a window
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Sets window possition
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
'Set Long
Public Declare Function SetWindowLong Lib "USER32.DLL" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long  'MODULE 1134

Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_FRAMECHANGED = &H20
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
Public Const WS_EX_TRANSPARENT = &H20&
Public Const GWL_EXSTYLE = (-20)
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
'*************************************************************************************
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40

Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40

Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lpBuffer As String, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
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

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, _
   lpProcessAttributes As Any, lpThreadAttributes As Any, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As Any, lpProcessInformation As Any) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long

Const SW_SHOWMINNOACTIVE = 7
Const STARTF_USESHOWWINDOW = &H1
Const INFINITE = -1&
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&

' to execute the browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWDEFAULT = 10
'delay function
'Declare everything

Dim tcpt As MIB_TCPTABLE
'Used to hide the dos window (so it doesnt interupt
'anything)
'Last declares
Public lngLastSent As Long
Public lngLastRecieved As Long

'Declares for Connection ID
Private Type Connection_
FileName As String
ProcessId As Long
TCPEntryNum As Long
LocalPort As String
RemotePort As String
LocalHost As String
RemoteHost As String
State As String
TCP As Boolean
End Type
Public Connection(2000) As Connection_
Public ConProcessID(2000) As Connection_
Public StatsLen As Long
Public ProgList As String

' ----------------------------
' Support Routines
' ----------------------------

Public Function Execute(ByVal CmdLine As String) As String
CmdLine = "Netstat -o"
    'Executes the command, and when it finish returns value to VB

    Dim proc As PROCESS_INFORMATION, ret As Long, bSuccess As Long
    Dim start As STARTUPINFO
    Dim sa As SECURITY_ATTRIBUTES
    Dim hReadPipe As Long, hWritePipe As Long
    Dim bytesread As Long, mybuff As String
    Dim i As Integer
    Dim retVal As Long
    Dim sReturnStr As String
    
    ' the lenght of the string must be 10 * 1024
    
    mybuff = String(10 * 1024, Chr$(65))
    sa.nLength = Len(sa)
    sa.bInheritHandle = 1&
    sa.lpSecurityDescriptor = 0&
    ret = CreatePipe(hReadPipe, hWritePipe, sa, 0)
    If ret = 0 Then
        '===Error
        sReturnStr = "Error: CreatePipe failed. " & Err.LastDllError
        Exit Function
    End If
    start.cb = Len(start)
    start.hStdOutput = hWritePipe
    start.dwflags = STARTF_USESTDHANDLES + STARTF_USESHOWWINDOW
    start.wShowWindow = SWP_HIDEWINDOW '+ SW_SHOWMINNOACTIVE
    ' Start the shelled application:
    ret& = CreateProcessA(0&, CmdLine$, sa, sa, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    If ret <> 1 Then
        '===Error
        sReturnStr = "Error: CreateProcess failed. " & Err.LastDllError
    End If
    
 
    
    ' Wait for the shelled application to finish:
    ret = WaitForSingleObject(proc.hProcess, INFINITE)
    
    DoEvents
    bSuccess = ReadFile(hReadPipe, mybuff, Len(mybuff), bytesread, 0&)
    If bSuccess = 1 Then
        sReturnStr = Left(mybuff, bytesread)
    Else
        '===Error
        sReturnStr = "Error: ReadFile failed. " & Err.LastDllError
    End If
    ret = CloseHandle(proc.hProcess)
    ret = CloseHandle(proc.hThread)
    ret = CloseHandle(hReadPipe)
    ret = CloseHandle(hWritePipe)
    

    'returns to VB
    Execute = sReturnStr

End Function

Public Sub Parse(Data As String)
Dim SplitData() As String 'Split by vbCrLf (Line Returns)
Dim LineSplit() As String
Dim i As Long
Dim LocP As String
Dim RemP As String
Dim LocA As String
Dim RemA As String
Dim State As String
Dim y As Long
Dim PID As Long

On Error Resume Next

'While there are more than 1 space chrs in a row
'remove them
Do While InStr(1, Data, "  ")
Data = Replace(Data, "  ", " ")
DoEvents
Loop

'Split by vbCrLf (Line Returns)
SplitData = Split(Data, vbCrLf)

'Split by Spaces
For y = 0 To UBound(SplitData)
LineSplit = Split(SplitData(y), " ")
DoEvents
    'Now find everything
    If LineSplit(0) <> "PROTO" Then
        If LineSplit(0) = "TCP " Then
        Connection(i).TCP = True
        Else
        Connection(i).TCP = False
        End If
    
    LocA = Mid(LineSplit(2), 1, InStr(1, LineSplit(2), ":"))
    LocP = Mid(LineSplit(2), InStr(1, LineSplit(2), ":") + 1, Len(LineSplit(2)) - InStr(1, LineSplit(2), ":"))
    RemP = Mid(LineSplit(3), InStr(1, LineSplit(3), ":") + 1, Len(LineSplit(3)) - InStr(1, LineSplit(3), ":"))
    If RemP = "http" Then RemP = "80 (Http)"
    If RemP = "https" Then RemP = "80 (Http)"
    RemA = Mid(LineSplit(3), 1, InStr(1, LineSplit(3), ":"))
    State = LineSplit(4)
    PID = 0
    PID = LineSplit(5)

    If PID <> 0 Then
    Connection(i).LocalHost = Replace(LocA, ":", "")
    Connection(i).LocalPort = LocP
    Connection(i).RemoteHost = Replace(RemA, ":", "")
    Connection(i).RemotePort = RemP
    Connection(i).State = State
    Connection(i).ProcessId = PID
    
    i = i + 1
    End If
    
    
    End If

Next y
StatsLen = i
End Sub
