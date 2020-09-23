Attribute VB_Name = "mMain"
Option Explicit

Public Declare Function WaitForMultipleObjects Lib "kernel32" _
 _
  (ByVal nCount As Long, lpHandles As Any, ByVal bWaitAll As Long, _
  ByVal dwMilliseconds As Long) As Long

Public Declare Function FindFirstChangeNotification Lib "kernel32" _
  Alias "FindFirstChangeNotificationA" _
  (ByVal lpPathName As String, ByVal bWatchSubtree As Long, _
   ByVal dwNotifyFilter As EFILE_NOTIFY) As Long

Public Declare Function FindNextChangeNotification Lib "kernel32" _
 _
  (ByVal hChangeHandle As Long) As Long

Public Declare Function FindCloseChangeNotification Lib "kernel32" _
 _
  (ByVal hChangeHandle As Long) As Long

Public Enum EFILE_NOTIFY
    FILE_NOTIFY_CHANGE_FILE_NAME = &H1
    FILE_NOTIFY_CHANGE_DIR_NAME = &H2
    FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
    FILE_NOTIFY_CHANGE_SIZE = &H8
    FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
    FILE_NOTIFY_CHANGE_SECURITY = &H100
End Enum

Public Const WAIT_TIMEOUT = 258
Public Const WAIT_FAILED = -1
Public Const INVALID_HANDLE_VALUE = -1

