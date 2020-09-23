Attribute VB_Name = "mndSS"

'Project:               Show stuff in the taskbar
'Programmer:            Hamman W. Samuel http://samuelsonline.f2g.net [<--This may soon change!]
'Acknowledgements:      Randy Birch http://www.vbnet.com and API Guide http://www.allapi.net
'Description:           The main code of the project is in the module. This submission is a demonstration on
'                       how things like icons and balloons can be added to the task bar tray.
'Features:              Add an icon to task bar, and show a popup balloon.
'                       Capture click events on the tray icon and the balloon too.
'Requirements:          Shell version 6+, and VB6 (successful tests on these)
'License:               You are not only free to modify this code, but also urged to do so!
'                       Easy references to events allow easy customization of the event procedures

'The other project:     I wasn't successful in incorporating animation as well into the same project, so made the
'                       animation project separately. To run either, just right-click on the one of your choice
'                       in the project explorer, and select <SET AS STARTUP>. Now you can run the project with the focus

'----------------------------------------------------------------------------------------------------
'DEUTERONOMY 4:6        FOR THIS IS YOUR WISDOM AND YOUR UNDERSTANDING IN THE SIGHT OF THE NATIONS, WHICH SHALL HEAR ALL THESE STATUTES,
'                       AND SAY, SURELY THIS GREAT NATION IS A WISE AND UNDERSTANDING PEOPLE (KJV)
'----------------------------------------------------------------------------------------------------

Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const NOTIFYICON_VERSION = &H3

Private Const NIM_MESSAGE = &H1
Private Const NIM_ICON = &H2
Private Const NIM_TIP = &H4
Private Const NIM_INFO = &H10

Private Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

Private Const NIM_SETVERSION = &H4
Private Const NIM_SHAREDICON = &H2

Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Const WM_RBUTTONDOWN As Long = &H204

Public Const WM_WNDPROC As Long = (-4)

Private Const WM_USER As Long = &H400
Private Const WM_APP As Long = &H8000&
Public Const WM_MYHOOK As Long = WM_APP + &H15

Public Const WM_SYSTRAY_ID = 999
Public Const WM_BALLOONUSERCLICK = (WM_USER + 5)

Private Const NOTIFYICONDATA_SIZE As Long = 504

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type

Public Enum enBalloonIcons
    none = 0
    Information = 1
    Warning = 2
    Critical = 3
End Enum
      
Public nIconData As NOTIFYICONDATA
Public defWindowProc As Long

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Error GoTo HANDLE_CRASH

Select Case uMsg
    Case WM_MYHOOK
        Select Case lParam
            Case WM_BALLOONUSERCLICK
                Call BalloonClicked
            Case WM_LBUTTONDBLCLK
                Call TrayDblClicked
            Case WM_RBUTTONDOWN
                Call TrayRightClicked
        End Select
    Case Else
        WindowProc = CallWindowProc(defWindowProc, hwnd, uMsg, wParam, lParam)
        Exit Function
End Select

Exit Function
HANDLE_CRASH:
MsgBox Err.Description
End Function

Public Sub ShellTrayAdd(lSourceHwnd As Long, vIconToShow, Optional stTipOnTray As String)
On Error GoTo HANDLE_CRASH

With nIconData
    .cbSize = NOTIFYICONDATA_SIZE
    .hwnd = lSourceHwnd
    .uID = WM_SYSTRAY_ID
    .uFlags = NIM_MESSAGE Or NIM_ICON Or NIM_TIP
    .dwState = NIM_SHAREDICON
    .hIcon = vIconToShow
    .szTip = stTipOnTray & vbNullChar
    .uTimeoutAndVersion = NOTIFYICON_VERSION
    .uCallbackMessage = WM_MYHOOK
End With

If Shell_NotifyIcon(NIM_ADD, nIconData) = 1 Then
    Call Shell_NotifyIcon(NIM_SETVERSION, nIconData)
    defWindowProc = SetWindowLong(lSourceHwnd, WM_WNDPROC, AddressOf WindowProc)
End If
Exit Sub

HANDLE_CRASH:
MsgBox Err.Description
End Sub

Public Sub ShellTrayRemove()

With nIconData
    .cbSize = NOTIFYICONDATA_SIZE
    .hwnd = FrmNet.hwnd
    .uID = WM_SYSTRAY_ID
End With

Call Shell_NotifyIcon(NIM_DELETE, nIconData)

End Sub

Public Sub ShellBalloonShow(lSourceHwnd As Long, ebiBalloonIcon As enBalloonIcons, stBalloonTitle As String, stBalloonText As String)

On Error GoTo HANDLE_CRASH

With nIconData
    .cbSize = NOTIFYICONDATA_SIZE
    .hwnd = lSourceHwnd
    .uID = WM_SYSTRAY_ID
    .uFlags = NIM_INFO
    .dwInfoFlags = ebiBalloonIcon
    .szInfoTitle = stBalloonTitle & vbNullChar
    .szInfo = stBalloonText & vbNullChar
End With

Call Shell_NotifyIcon(NIM_MODIFY, nIconData)
Exit Sub

HANDLE_CRASH:
MsgBox Err.Description
End Sub

Public Sub UnSubClass(lSourceHwnd As Long)
   If defWindowProc <> 0 Then
      SetWindowLong lSourceHwnd, WM_WNDPROC, defWindowProc
      defWindowProc = 0
   End If
End Sub

Sub BalloonClicked()
'You can code this procedure to respond to click events of the balloon
FrmNet.Show
End Sub

Sub TrayDblClicked()
'You can code this procedure to respond to double-click event of the tray icon
FrmNet.Show
End Sub

Sub TrayRightClicked()
'You can code this procedure to respond to right-click event of the tray icon
FrmNet.PopupMenu FrmNet.Menu
End Sub

