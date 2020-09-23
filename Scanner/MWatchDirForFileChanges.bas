Attribute VB_Name = "MWatchDirForFileChanges"
  Option Private Module
  Option Explicit
  ' demo project showing how to use the FindFirstChangeNotification and
  ' FindNextChangeNotification API functions to watch directories for
  ' changes.  by Bryan Stafford of New Vision Software - newvision@mvps.org
  ' this demo is released into the public domain "as is" without warranty or
  ' guaranty of any kind.  In other words, use at your own risk.
  '
  '
  ' NOTE:  there are other routines that can detect changes in a directory.
  ' structure, most notably registering a change notification routine using
  ' the SHChangeNotifyRegister API function.  However, unless the application
  ' making changes to the file system is sending the proper notifications via
  ' the SHChangeNotify API function, only a minimal amount of information is
  ' returned.


  Public Const API_FALSE As Long = &H0&
  Public Const API_TRUE As Long = &H1&
  

  Public Const WS_POPUP As Long = &H80000000

  Public Const WS_EX_TOOLWINDOW As Long = &H80&
  Public Const WS_EX_TRANSPARENT As Long = &H20&

  Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle&, ByVal lpClassName$, ByVal lpWindowName$, ByVal dwStyle&, ByVal x&, ByVal y&, ByVal nWidth&, ByVal nHeight&, ByVal hWndParent&, ByVal hMenu&, ByVal hInstance&, lpParam As Any) As Long

  Public Declare Function SetTimer Lib "user32" (ByVal hWnd&, ByVal nIDEvent&, ByVal uElapse&, ByVal lpTimerFunc&) As Long
  Public Declare Function KillTimer Lib "user32" (ByVal hWnd&, ByVal nIDEvent&) As Long

  Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd&) As Long

  Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes As Long)

' this is the callback function for the API timer that we set in the class object.
' we passed an object pointer to the instance of the object in the idEvent parameter
' so that we can resolve it to the correct instance of the class object here
Public Sub WatchDirTimerProc(ByVal hWnd&, ByVal uMsg&, ByVal idEvent&, ByVal dwTime&)
  WatchDirFromIDEvent(idEvent).CheckWatch
End Sub

Private Function WatchDirFromIDEvent(ByVal idEvent As Long) As cWatchForChanges
  ' resolve the passed object pointer into an object reference.
  ' DO NOT PRESS THE "STOP" BUTTON WHILE IN THIS PROCEDURE!
  
  Dim CwdEx As cWatchForChanges

  CopyMemory CwdEx, idEvent, 4&
  Set WatchDirFromIDEvent = CwdEx
  CopyMemory CwdEx, 0&, 4&
  
End Function

