Attribute VB_Name = "modWindow"
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public wndNames() As String
Public winID As Long
Public winName As String
Public Const SW_SHOW As Long = 5
Public Const SW_HIDE As Long = 0
Public runHWND As Long

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
'ïîä÷èíåííûå îêíà
Dim tempPID As Long
Dim Ret As Long
Dim winText As String
    Ret = GetWindowTextLength(hwnd)
    winText = Space(Ret)
    GetWindowText hwnd, winText, Ret + 1
    If winText <> "" Then
        GetWindowThreadProcessId hwnd, tempPID
        For i = 0 To UBound(procinfo)
            If procinfo(i).th32ProcessID = tempPID Then
                ''If glbPID = procinfo(i).th32ProcessID Then
                        'If IsWindowVisible(hwnd) Then
                ''    ShowWindow hwnd, 0
                            'For z = 0 To 50
                            '    If procHwnds(z) = 0 Then
                            '        procHwnds(z) = hwnd
                            '        Exit For
                            '    End If
                            'Next z
                        'End If
                ''End If
                procinfo(i).childWnd = procinfo(i).childWnd + 1
            End If
        Next i
        'frmTest.lvProcess.AddItem tempPID & "  " & winText
    End If
    EnumWindowsProc = True
    glbHwnd = 0
End Function

Public Function EnumWindowsProc2(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
Dim tempPID As Long
Dim Ret As Long
Dim winText As String
Dim targetID As Long
    targetID = winID
    Ret = GetWindowTextLength(hwnd)
    winText = Space(Ret)
    GetWindowText hwnd, winText, Ret + 1
    If winText <> "" Then
        GetWindowThreadProcessId hwnd, tempPID
        If targetID = tempPID Then
            For i = 0 To UBound(wndNames)
                If wndNames(i) = "" Then
                    wndNames(i) = winText
                    Exit For
                End If
            Next i
        End If
    End If
    EnumWindowsProc2 = True
    glbHwnd = 0
End Function
