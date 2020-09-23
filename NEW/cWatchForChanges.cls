VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cWatchForChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
  
  
  
  Public Enum NOTIFY_CHANGE_TYPE
    cwdName = 1&
    cwdSize = 2&
    cwdLastWrite = 4&
    cwdDirName = 8&
  End Enum
  
  Private m_hWnd As Long
  
  Private m_hNotifyChangeName As Long
  Private m_hNotifyChangeSize As Long
  Private m_hNotifyChangeLastWrite As Long
  Private m_hNotifyChangeDirName As Long
  
  ' flag to tell the routine to suspend checking  for changes
  Private m_bSuspendWatch As Boolean
  
  Private Const INVALID_HANDLE_VALUE  As Long = (-1)
  
  Private Const FILE_NOTIFY_CHANGE_FILE_NAME  As Long = &H1&
  Private Const FILE_NOTIFY_CHANGE_DIR_NAME  As Long = &H2&
  Private Const FILE_NOTIFY_CHANGE_SIZE As Long = &H8&
  Private Const FILE_NOTIFY_CHANGE_LAST_WRITE As Long = &H10&


  Private Const WAIT_FAILED As Long = (-1&)
  Private Const WAIT_OBJECT_0 As Long = 0&
  Private Const WAIT_ABANDONED As Long = &H80&
  Private Const WAIT_ABANDONED_0 As Long = &H80&
  Private Const WAIT_TIMEOUT As Long = &H102&
'  Private Const WAIT_IO_COMPLETION As Long = &HC0&
'  Private Const STILL_ACTIVE As Long = &H103&
'  Private Const INFINITE As Long = -1&
  
  Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
  
  Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd&, lpdwProcessId&) As Long
  Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess&, ByVal bInheritHandle&, ByVal dwProcessId&) As Long

  
  Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle&, ByVal dwMilliseconds&) As Long
  
  Private Declare Function FindFirstChangeNotification Lib "kernel32" Alias "FindFirstChangeNotificationA" _
                                                          (ByVal lpPathName$, ByVal bWatchSubTree&, ByVal dwNotifyFilter&) As Long
                                                          
  Private Declare Function FindNextChangeNotification Lib "kernel32" (ByVal hChangeHandle&) As Long
  Private Declare Function FindCloseChangeNotification Lib "kernel32" (ByVal hChangeHandle&) As Long
  
  
  Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20&
  
  Private Const SE_PRIVILEGE_ENABLED As Long = &H2&

  Private Const ANYSIZE_ARRAY As Long = 1&

  Private Type LUID
    LowPart As Long
    HighPart As Long
  End Type

  Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
  End Type

  Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
  End Type

  Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle&, ByVal DesiredAccess&, TokenHandle&) As Long
  Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle&, ByVal DisableAllPrivileges&, NewState As TOKEN_PRIVILEGES, ByVal BufferLength&, ByVal PreviousState&, ReturnLength&) As Long
  Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName$, ByVal lpName$, lpLuid As LUID) As Long

  Event ChangeDetected(ByVal eChangeType As NOTIFY_CHANGE_TYPE)

Public Property Get SuspendWatch() As Boolean
  SuspendWatch = m_bSuspendWatch
End Property
Public Property Let SuspendWatch(ByVal bNewVal As Boolean)
  m_bSuspendWatch = bNewVal
End Property

Friend Sub CheckWatch()
  ' this function is called from the API timer routine in the MWatchDirForFileChanges
  ' BAS module
  
  Dim nChangeType As NOTIFY_CHANGE_TYPE
  
  ' used to adjust the amount of time each check below waits before returning
  Const WAIT_DELAY As Long = 0&
  
  If m_bSuspendWatch = False Then
    ' Check to see if a change occurred. You can change the
    ' delay to whatever is convenient. If it is zero,
    ' it will check for changes and return immediately (i.e.. no delay).
    nChangeType = CheckForChange(m_hNotifyChangeName, cwdName, WAIT_DELAY)
    nChangeType = nChangeType Or CheckForChange(m_hNotifyChangeSize, cwdSize, WAIT_DELAY)
    nChangeType = nChangeType Or CheckForChange(m_hNotifyChangeLastWrite, cwdLastWrite, WAIT_DELAY)
    nChangeType = nChangeType Or CheckForChange(m_hNotifyChangeDirName, cwdDirName, WAIT_DELAY)
  
    If nChangeType Then RaiseEvent ChangeDetected(nChangeType)
  End If ' m_bSuspendWatch = False
  
End Sub

Private Function CheckForChange(ByVal hChangeObject, ByVal eChangeType As NOTIFY_CHANGE_TYPE, ByVal nWait&) As NOTIFY_CHANGE_TYPE
  ' checks for changes in the passed wait object handle

  Select Case WaitForSingleObject(hChangeObject, nWait)
    Case WAIT_OBJECT_0
      ' found a change
      
      ' set the return val for this function
      CheckForChange = eChangeType
      
      ' keep looking for more changes
      Call FindNextChangeNotification(hChangeObject)
      
  End Select

End Function

Public Function StartWatch(ByVal sDir$, Optional ByVal nTimerInterval& = 500&, _
                    Optional ByVal bStartSuspended As Boolean, Optional ByVal bWatchSubTree As Boolean) As Boolean

  ' start watching the passed directory or directory tree.
  
  Dim fWatchSubTree&
  
  StopWatch
  
  m_bSuspendWatch = bStartSuspended
  
  
  ' create an invisible API window to allow us to pass an hWnd to the SetTimer function
  m_hWnd = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_TRANSPARENT, "STATIC", vbNullString, _
                                              WS_POPUP, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, 0&)

  
  ' alter the privileges to allow notification over a network.....
  Dim idProcess&, hProcess&, hToken&, udtPrivileges As TOKEN_PRIVILEGES, nRet&
  
  Call GetWindowThreadProcessId(m_hWnd, idProcess)
  
  If idProcess Then
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, API_FALSE, idProcess)
  
    If hProcess Then
      If OpenProcessToken(hProcess, TOKEN_ADJUST_PRIVILEGES, hToken) Then
        With udtPrivileges
          .PrivilegeCount = 2
          With .Privileges(0)
            .Attributes = SE_PRIVILEGE_ENABLED
            
            nRet = LookupPrivilegeValue(vbNullChar, "SeBackupPrivilege", .pLuid)
          End With
          
          With .Privileges(1)
            .Attributes = SE_PRIVILEGE_ENABLED
            
            nRet = LookupPrivilegeValue(vbNullChar, "SeChangeNotifyPrivilege", .pLuid)
          End With
        End With
        
        If nRet Then
          If AdjustTokenPrivileges(hToken, API_FALSE, udtPrivileges, 0&, 0&, 0&) Then
            ' we don't do anything here cause it won't make any difference if we do....
          End If
        End If
      End If
    End If
  End If
  
  ' specify whether or not to watch the subdirectories also
  If bWatchSubTree Then fWatchSubTree = API_TRUE
  
  ' Establish a notification object for the drive in question
  m_hNotifyChangeName = FindFirstChangeNotification(sDir, fWatchSubTree, FILE_NOTIFY_CHANGE_FILE_NAME)
  m_hNotifyChangeSize = FindFirstChangeNotification(sDir, fWatchSubTree, FILE_NOTIFY_CHANGE_SIZE)
  m_hNotifyChangeLastWrite = FindFirstChangeNotification(sDir, fWatchSubTree, FILE_NOTIFY_CHANGE_LAST_WRITE)
  m_hNotifyChangeDirName = FindFirstChangeNotification(sDir, fWatchSubTree, FILE_NOTIFY_CHANGE_DIR_NAME)
  
  ' If the returned handle was invalid, the notification object
  ' was not created. Return false to indicate that the drive
  ' isn't being monitored.
  If (m_hNotifyChangeName = INVALID_HANDLE_VALUE) Or (m_hNotifyChangeSize = INVALID_HANDLE_VALUE) Or _
      (m_hNotifyChangeLastWrite = INVALID_HANDLE_VALUE) Or (m_hNotifyChangeDirName = INVALID_HANDLE_VALUE) Then
      
    StopWatch
    
    StartWatch = False
    
  Else
  
  
    If m_hWnd Then
      If SetTimer(m_hWnd, ObjPtr(Me), nTimerInterval, AddressOf MWatchDirForFileChanges.WatchDirTimerProc) Then
        StartWatch = True
      Else
        StopWatch
      End If
    End If
  
    
  End If

End Function

Public Sub StopWatch()
  ' clean up all objects....
  
  If m_hWnd Then
    Call KillTimer(m_hWnd, ObjPtr(Me))
  
    Call DestroyWindow(m_hWnd)
  End If
  
  ' Always close the notification objects before exiting
  If m_hNotifyChangeName Then
    Call FindCloseChangeNotification(m_hNotifyChangeName)
  
    m_hNotifyChangeName = 0
  End If
  
  If m_hNotifyChangeSize Then
    Call FindCloseChangeNotification(m_hNotifyChangeSize)
  
    m_hNotifyChangeSize = 0
  End If
  
  If m_hNotifyChangeLastWrite Then
    Call FindCloseChangeNotification(m_hNotifyChangeLastWrite)
  
    m_hNotifyChangeLastWrite = 0
  End If
  
  If m_hNotifyChangeDirName Then
    Call FindCloseChangeNotification(m_hNotifyChangeDirName)
  
    m_hNotifyChangeDirName = 0
  End If
  
  
  m_bSuspendWatch = False
  
End Sub

Private Sub Class_Terminate()
  ' make sure we clean up after ourselves
  StopWatch
End Sub
