VERSION 5.00
Begin VB.UserControl cSysTray 
   CanGetFocus     =   0   'False
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   510
   ClipControls    =   0   'False
   EditAtDesignTime=   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   34
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   34
End
Attribute VB_Name = "cSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'-------------------------------------------------------
' Control Property Globals...
'-------------------------------------------------------
Private NOTIFYICONDATA_SIZE As Long
Private gInTray As Boolean
Private gTrayId As Long
Private gTrayTip As String
Private gTrayHwnd As Long
Private gTrayIcon As StdPicture
Private gAddedToTray As Boolean
Const MAX_SIZE = 510
Public Enum TrayIcon_InfoIcon
        NIIF_NONE = &H0
        NIIF_INFO = &H1
        NIIF_WARNING = &H2
        NIIF_ERROR = &H3
        NIIF_GUID = &H5         ' Íå èñïîëüçóåòñÿ
        NIIF_ICON_MASK = &HF    ' Íå èñïîëüçóåòñÿ
        NIIF_NOSOUND = &H10
   End Enum
Private Const defInTray = False
Private Const defTrayTip = "Gadyash Antivirus 2008 Beta" & vbNullChar

Private Const sInTray = "InTray"
Private Const sTrayIcon = "TrayIcon"
Private Const sTrayTip = "TrayTip"

'-------------------------------------------------------
' Control Events...
'-------------------------------------------------------
Public Event MouseMove(Id As Long)
Public Event MouseDown(Button As Integer, Id As Long)
Public Event MouseUp(Button As Integer, Id As Long)
Public Event MouseDblClick(Button As Integer, Id As Long)

'-------------------------------------------------------
Private Sub UserControl_Initialize()
'-------------------------------------------------------
    gInTray = defInTray                             ' Set global InTray defalt
    gAddedToTray = False                            ' Set default state
    gTrayId = 0                                     ' Set global TrayId default
    gTrayHwnd = Hwnd                                ' Set and keep HWND of user control
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

'-------------------------------------------------------
Private Sub UserControl_InitProperties()
'-------------------------------------------------------
    InTray = defInTray                              ' Init InTray Property
    TrayTip = defTrayTip                            ' Init TrayTip Property
    Set TrayIcon = Picture                          ' Init TrayIcon property
'-------------------------------------------------------
End Sub
'-------------------------------------------------------


'-------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'-------------------------------------------------------
    ' Read in the properties that have been saved into the PropertyBag...
    With PropBag
        InTray = .ReadProperty(sInTray, defInTray)       ' Get InTray
        Set TrayIcon = .ReadProperty(sTrayIcon, Picture) ' Get TrayIcon
        TrayTip = .ReadProperty(sTrayTip, defTrayTip)    ' Get TrayTip
    End With
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

'-------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'-------------------------------------------------------
    With PropBag
        .WriteProperty sInTray, gInTray                 ' Save InTray to propertybag
        .WriteProperty sTrayIcon, gTrayIcon             ' Save TrayIcon to propertybag
        .WriteProperty sTrayTip, gTrayTip               ' Save TrayTip to propertybag
    End With
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

'-------------------------------------------------------
Private Sub UserControl_Resize()
'-------------------------------------------------------
    Height = MAX_SIZE                   ' Prevent Control from being resized...
    Width = MAX_SIZE
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

'-------------------------------------------------------
Private Sub UserControl_Terminate()
'-------------------------------------------------------
    If InTray Then                      ' If TrayIcon is visible
        InTray = False                  ' Cleanup and unplug it.
    End If
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

'-------------------------------------------------------
Public Property Set TrayIcon(Icon As StdPicture)
'-------------------------------------------------------
    Dim Tray As NOTIFYICONDATA                          ' Notify Icon Data structure
    Dim rc As Long                                      ' API return code
'-------------------------------------------------------
    If Not (Icon Is Nothing) Then                       ' If icon is valid...
        If (Icon.Type = vbPicTypeIcon) Then             ' Use ONLY if it is an icon
            If gAddedToTray Then                        ' Modify tray only if it is in use.
                Tray.uID = gTrayId                      ' Unique ID for each HWND and callback message.
                Tray.Hwnd = gTrayHwnd                   ' HWND receiving messages.
                Tray.hIcon = Icon.Handle                ' Tray icon.
                Tray.uFlags = NIF_ICON                  ' Set flags for valid data items
                Tray.cbSize = Len(Tray)                 ' Size of struct.
                
                rc = Shell_NotifyIcon(NIM_MODIFY, Tray) ' Send data to Sys Tray.
            End If
    
            Set gTrayIcon = Icon                        ' Save Icon to global
            Set Picture = Icon                          ' Show user change in control as well(gratuitous)
            PropertyChanged sTrayIcon                   ' Notify control that property has changed.
        End If
    End If
'-------------------------------------------------------
End Property
'-------------------------------------------------------

'-------------------------------------------------------
Public Property Get TrayIcon() As StdPicture
'-------------------------------------------------------
    Set TrayIcon = gTrayIcon                        ' Return Icon value
'-------------------------------------------------------
End Property
'-------------------------------------------------------

'-------------------------------------------------------
Public Property Let TrayTip(Tip As String)
Attribute TrayTip.VB_ProcData.VB_Invoke_PropertyPut = ";Misc"
Attribute TrayTip.VB_UserMemId = -517
'-------------------------------------------------------
    Dim Tray As NOTIFYICONDATA                      ' Notify Icon Data structure
    Dim rc As Long                                  ' API Return code
'-------------------------------------------------------
    If gAddedToTray Then                            ' if TrayIcon is in taskbar
        Tray.uID = gTrayId                          ' Unique ID for each HWND and callback message.
        Tray.Hwnd = gTrayHwnd                       ' HWND receiving messages.
        Tray.szTip = Tip & vbNullChar               ' Tray tool tip
        Tray.uFlags = NIF_TIP                       ' Set flags for valid data items
        Tray.cbSize = Len(Tray)                     ' Size of struct.
        
        rc = Shell_NotifyIcon(NIM_MODIFY, Tray)     ' Send data to Sys Tray.
    End If
    
    gTrayTip = Tip                                  ' Save Tip
    PropertyChanged sTrayTip                        ' Notify control that property has changed
'-------------------------------------------------------
End Property
'-------------------------------------------------------

'-------------------------------------------------------
Public Property Get TrayTip() As String
'-------------------------------------------------------
    TrayTip = gTrayTip                              ' Return Global Tip...
'-------------------------------------------------------
End Property
'-------------------------------------------------------

'-------------------------------------------------------
Public Property Let InTray(Show As Boolean)
Attribute InTray.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
'-------------------------------------------------------
    Dim ClassAddr As Long                           ' Address pointer to Control Instance
'-------------------------------------------------------
    If (Show <> gInTray) Then                       ' Modify ONLY if state is changing!
        If Show Then                                ' If adding Icon to system tray...
            If Ambient.UserMode Then                ' If in RunMode and not in IDE...
                 ' SubClass Controls window proc.
                PrevWndProc = SetWindowLong(gTrayHwnd, GWL_WNDPROC, AddressOf SubWndProc)
                
                ' Get address to user control object
                'CopyMemory ClassAddr, UserControl, 4&
                
                ' Save address to the USERDATA of the control's window struct.
                ' this will be used to get an object refenence to the control
                ' from an HWND in the callback.
                SetWindowLong gTrayHwnd, GWL_USERDATA, ObjPtr(Me) 'ClassAddr
                
                AddIcon gTrayHwnd, gTrayId, TrayTip, TrayIcon ' Add TrayIcon to System Tray...
                gAddedToTray = True                 ' Save state of control used in teardown procedure
            End If
        Else                                        ' If removing Icon from system tray
            If gAddedToTray Then                    ' If Added to system tray then remove...
                DeleteIcon gTrayHwnd, gTrayId       ' Remove icon from system tray
                
                ' Un SubClass controls window proc.
                SetWindowLong gTrayHwnd, GWL_WNDPROC, PrevWndProc
                gAddedToTray = False                ' Maintain the state for teardown purposes
            End If
        End If
        
        gInTray = Show                              ' Update global variable
        PropertyChanged sInTray                     ' Notify control that property has changed
    End If
'-------------------------------------------------------
End Property
'-------------------------------------------------------

'-------------------------------------------------------
Public Property Get InTray() As Boolean
'-------------------------------------------------------
    InTray = gInTray                                ' Return global property
'-------------------------------------------------------
End Property
'-------------------------------------------------------

'-------------------------------------------------------
Private Sub AddIcon(Hwnd As Long, Id As Long, Tip As String, Icon As StdPicture)
'-------------------------------------------------------
    Dim Tray As NOTIFYICONDATA                      ' Notify Icon Data structure
    Dim tFlags As Long                              ' Tray action flag
    Dim rc As Long                                  ' API return code
'-------------------------------------------------------
    Tray.uID = Id                                   ' Unique ID for each HWND and callback message.
    Tray.Hwnd = Hwnd                                ' HWND receiving messages.
    
    If Not (Icon Is Nothing) Then                   ' Validate Icon picture
        Tray.hIcon = Icon.Handle                    ' Tray icon.
        Tray.uFlags = Tray.uFlags Or NIF_ICON       ' Set ICON flag to validate data item
        Set gTrayIcon = Icon                        ' Save icon
    End If
    
    If (Tip <> "") Then                             ' Validate Tip text
        Tray.szTip = Tip & vbNullChar               ' Tray tool tip
        Tray.uFlags = Tray.uFlags Or NIF_TIP        ' Set TIP flag to validate data item
        gTrayTip = Tip                              ' Save tool tip
    End If
    
    Tray.uCallbackMessage = TRAY_CALLBACK           ' Set user defigned message
    Tray.uFlags = Tray.uFlags Or NIF_MESSAGE        ' Set flags for valid data item
    Tray.cbSize = Len(Tray)                         ' Size of struct.
    
    rc = Shell_NotifyIcon(NIM_ADD, Tray)            ' Send data to Sys Tray.
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

'-------------------------------------------------------
Private Sub DeleteIcon(Hwnd As Long, Id As Long)
'-------------------------------------------------------
    Dim Tray As NOTIFYICONDATA                      ' Notify Icon Data structure
    Dim rc As Long                                  ' API return code
'-------------------------------------------------------
    Tray.uID = Id                                   ' Unique ID for each HWND and callback message.
    Tray.Hwnd = Hwnd                                ' HWND receiving messages.
    Tray.uFlags = 0&                                ' Set flags for valid data items
    Tray.cbSize = Len(Tray)                         ' Size of struct.
    
    rc = Shell_NotifyIcon(NIM_DELETE, Tray)         ' Send delete message.
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

'-------------------------------------------------------
Friend Sub SendEvent(MouseEvent As Long, Id As Long)
'-------------------------------------------------------
    Select Case MouseEvent                          ' Dispatch mouse events to control
    Case WM_MOUSEMOVE
        RaiseEvent MouseMove(Id)
    Case WM_LBUTTONDOWN
        RaiseEvent MouseDown(vbLeftButton, Id)
    Case WM_LBUTTONUP
        RaiseEvent MouseUp(vbLeftButton, Id)
    Case WM_LBUTTONDBLCLK
        RaiseEvent MouseDblClick(vbLeftButton, Id)
    Case WM_RBUTTONDOWN
        RaiseEvent MouseDown(vbRightButton, Id)
    Case WM_RBUTTONUP
        RaiseEvent MouseUp(vbRightButton, Id)
    Case WM_RBUTTONDBLCLK
        RaiseEvent MouseDblClick(vbRightButton, Id)
    End Select
'-------------------------------------------------------
End Sub
'-------------------------------------------------------

Public Sub DisplayBalloon(ByVal sTitle As String, ByVal sText As String, Optional ByVal InfoFlags As TrayIcon_InfoIcon)
 Dim Tray As NOTIFYICONDATA                      ' Notify Icon Data structure
    Dim rc As Long                                  ' API return code
If NOTIFYICONDATA_SIZE = 0 Then SetShellVersion
   
   With Tray
      .cbSize = NOTIFYICONDATA_SIZE
      .Hwnd = Hwnd
      .uID = Hwnd
      .uFlags = Tray.uFlags Or NIF_MESSAGE
      '.dwInfoFlags = InfoFlags
     '.szInfoTitle = sTitle & vbNullChar
      .szInfo = "mtt" & vbNullChar
   End With

   ret = Shell_NotifyIcon(NIM_MODIFY, NID)
End Sub
'Óñòàíàâëèâàåì ðàçìåð ïåðåìåííîé NOTIFYICONDATA_SIZE
'â çàâèñèìîñòè îò âåðñèè îáîëî÷êè
Private Sub SetShellVersion()

   Select Case True
      Case IsShellVersion(6)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V3_SIZE '6.0+ structure size
      
      Case IsShellVersion(5)
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V2_SIZE 'pre-6.0 structure size
      
      Case Else
         NOTIFYICONDATA_SIZE = NOTIFYICONDATA_V1_SIZE 'pre-5.0 structure size
   End Select

End Sub

'Îïðåäåëÿåì âåðñèþ îáîëî÷êè Shell32.dll
Private Function IsShellVersion(ByVal version As Long) As Boolean

  'ôóíêöèÿ âîçðàùàåò Èñòèíó åñëè âåðñèÿ îáîëî÷êè
  '(shell32.dll) ðàâíà èëè áîëüøå çàïðàøèâàåìîé
   Dim nBufferSize As Long
   Dim nUnused As Long
   Dim lpBuffer As Long
   Dim nVerMajor As Integer
   Dim bBuffer() As Byte
   
   Const sDLLFile As String = "shell32.dll"
   
   nBufferSize = GetFileVersionInfoSize(sDLLFile, nUnused)
   
   If nBufferSize > 0 Then
    
      ReDim bBuffer(nBufferSize - 1) As Byte
    
      Call GetFileVersionInfo(sDLLFile, 0&, nBufferSize, bBuffer(0))
    
      If VerQueryValue(bBuffer(0), "\", lpBuffer, nUnused) = 1 Then
         
         CopyMemory nVerMajor, ByVal lpBuffer + 10, 2
        
         IsShellVersion = nVerMajor >= version
      
      End If
    
   End If
  
End Function

