Attribute VB_Name = "mRegNotifyChange"
'---------------------------------------------------------------------------------------
' Module    : mRegNotifyChange
' DateTime  :
' Author    :
' Mail      :
' Purpose   : Registry Api wraper
' Requirements: None
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' CONSTANTS
'---------------------------------------------------------------------------------------
Private Const WAIT_OBJECT_0 As Long = &H0

Private Const ERROR_SUCCESS As Long = 0&

Private Const REG_NOTIFY_CHANGE_ATTRIBUTES As Long = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET As Long = &H4
Private Const REG_NOTIFY_CHANGE_NAME As Long = &H1
Private Const REG_NOTIFY_CHANGE_SECURITY As Long = &H8
Private Const REG_NOTIFY_CHANGE_ALL As Long = &H8 Or &H1 Or &H2 Or &H4

'---------------------------------------------------------------------------------------
' TYPE
'---------------------------------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

'---------------------------------------------------------------------------------------
' API
'---------------------------------------------------------------------------------------
Private Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal lKey As Long, ByVal lWatchSubtree As Long, ByVal lNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long
Private Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function CreateEvent Lib "kernel32.dll" Alias "CreateEventA" (ByRef lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function ResetEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long
'
'---------------------------------------------------------------------------------------
' Procedure : RegFindFirstChange
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function RegFindFirstChange(lKey As Long, lWatchSubtree As Long, lNotifyFilter As Long) As Long
    Dim lResult As Long
    Dim lChange As Long
    Dim sd As SECURITY_ATTRIBUTES
    
    lChange = CreateEvent(sd, True, False, 0&)

    lResult = RegNotifyChangeKeyValue(lKey, lWatchSubtree, lNotifyFilter, lChange, True)

    If Not (lResult = ERROR_SUCCESS) Then
        SetLastError lResult
        RegFindFirstChange = 0&
        Exit Function
    End If

    If WaitForSingleObject(lChange, 0) = WAIT_OBJECT_0 Then

        lResult = RegNotifyChangeKeyValue(lKey, lWatchSubtree, lNotifyFilter, lChange, True)

        If Not (lResult = ERROR_SUCCESS) Then
            SetLastError lResult
            RegFindFirstChange = 0&
            Exit Function
        End If
    End If

    RegFindFirstChange = lChange

End Function
'
'---------------------------------------------------------------------------------------
' Procedure : RegFindNextChange
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function RegFindNextChange(lChange As Long, lKey As Long, lWatchSubtree As Long, lNotifyFilter As Long) As Boolean
    Dim lResult As Long
    
    If ResetEvent(lChange) = 0 Then
        RegFindNextChange = False
        Exit Function
    End If
    
    lResult = RegNotifyChangeKeyValue(lKey, lWatchSubtree, lNotifyFilter, lChange, True)

    If Not (lResult = ERROR_SUCCESS) Then
        SetLastError lResult
        RegFindNextChange = False
        Exit Function
    Else
        RegFindNextChange = True
    End If
        
End Function
'
'---------------------------------------------------------------------------------------
' Procedure : RegFindCloseChange
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function RegFindCloseChange(lChange As Long) As Boolean
    If lChange Then
        CloseHandle (lChange)
        lChange = 0
    End If
    RegFindCloseChange = True
End Function
