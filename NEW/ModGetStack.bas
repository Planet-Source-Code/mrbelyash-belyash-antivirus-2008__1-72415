Attribute VB_Name = "ModGetStack"
Private Type MIB_TCPROW
    dwState As Long
    dwLocalAddr As Long
    dwLocalPort As Long
    dwRemoteAddr As Long
    dwRemotePort As Long
    End Type


Private Declare Function GetProcessHeap Lib "kernel32" () As Long


Private Declare Function htons Lib "ws2_32.dll" (ByVal dwLong As Long) As Long


Public Declare Function AllocateAndGetTcpExTableFromStack Lib "iphlpapi.dll" (pTcpTableEx As Any, ByVal bOrder As Long, ByVal heap As Long, ByVal zero As Long, ByVal Flags As Long) As Long


Private Declare Function SetTcpEntry Lib "iphlpapi.dll" (pTcpTableEx As MIB_TCPROW) As Long


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private pTablePtr As Long
    Private pDataRef As Long
    Public nRows As Long
    Public oRows As Long
    Private nCurrentRow As Long
    Private udtRow As MIB_TCPROW
    Private nState As Long
    Private nLocalAddr As Long
    Private nLocalPort As Long
    Private nRemoteAddr As Long
    Private nRemotePort As Long
    Private nProcId As Long
    Public nRet As Long

Private Type Connection_
    FileName As String
    ProcessId As Long
    LocalPort As String
    OrigLocalPort As String
    RemotePort As String
    OrigRemotePort As String
    LocalHost As String
    OrigLocalHost As String
    RemoteHost As String
    OrigRemoteHost As String
    State As String
End Type
Public Connection(2000) As Connection_

Public Function GetIPAddress(dwAddr As Long) As String
    Dim arrIpParts(3) As Byte
    CopyMemory arrIpParts(0), dwAddr, 4
    GetIPAddress = CStr(arrIpParts(0)) & "." & _
    CStr(arrIpParts(1)) & "." & _
    CStr(arrIpParts(2)) & "." & _
    CStr(arrIpParts(3))
End Function


Public Function GetPort(ByVal dwPort As Long) As Long
    GetPort = htons(dwPort)
End Function

Public Sub RefreshStack()
Dim i As Long
    pDataRef = 0

For i = 0 To nRows '// read 24 bytes at a time
    CopyMemory nState, ByVal pTablePtr + (pDataRef + 4), 4
    CopyMemory nLocalAddr, ByVal pTablePtr + (pDataRef + 8), 4
    CopyMemory nLocalPort, ByVal pTablePtr + (pDataRef + 12), 4
    CopyMemory nRemoteAddr, ByVal pTablePtr + (pDataRef + 16), 4
    CopyMemory nRemotePort, ByVal pTablePtr + (pDataRef + 20), 4
    CopyMemory nProcId, ByVal pTablePtr + (pDataRef + 24), 4

    DoEvents

        If nRemoteAddr <> 0 Or nRemotePort <> 0 Or nLocalPort <> 0 Then
            Connection(i).State = nState
            Connection(i).LocalHost = GetIPAddress(nLocalAddr)
            Connection(i).OrigLocalHost = nLocalAddr
            Connection(i).LocalPort = GetPort(nLocalPort)
            Connection(i).OrigLocalPort = nLocalPort
            Connection(i).RemoteHost = GetIPAddress(nRemoteAddr)
            Connection(i).OrigRemoteHost = nRemoteAddr
            Connection(i).RemotePort = GetPort(nRemotePort)
            Connection(i).OrigRemotePort = nRemotePort
            Connection(i).ProcessId = nProcId
        End If
        pDataRef = pDataRef + 24

        DoEvents
        Next i
        
End Sub


Public Function GetEntryCount() As Long
    GetEntryCount = nRows - 2 '// The last entry is always an EOF of sorts
End Function

Public Function TerminateThisConnection(rNum As Long) As Boolean
    
    On Error GoTo ErrorTrap
    
    udtRow.dwLocalAddr = Connection(rNum).OrigLocalHost
    udtRow.dwLocalPort = Connection(rNum).OrigLocalPort
    udtRow.dwRemoteAddr = Connection(rNum).OrigRemoteHost
    udtRow.dwRemotePort = Connection(rNum).OrigRemotePort
    udtRow.dwState = 12
    SetTcpEntry udtRow
    
    TerminateThisConnection = True
    Exit Function
    
ErrorTrap:
    TerminateThisConnection = False
    
End Function

Public Function GetRefresh() As Boolean
    nRet = AllocateAndGetTcpExTableFromStack(pTablePtr, 0, GetProcessHeap, 0, 2)


    If nRet = 0 Then
        CopyMemory nRows, ByVal pTablePtr, 4
    Else
        GetRefresh = False
        Exit Function
    End If
    
    If nRows = 0 Or pTablePtr = 0 Then
    GetRefresh = False
    Exit Function
    End If

   If oRows = nRows Then
   GetRefresh = False
   Else
   GetRefresh = True
   End If
   
oRows = nRows

End Function

Function c_state(s) As String
  Select Case s
  Case "0": c_state = "UNKNOWN"
  Case "1": c_state = "CLOSED"
  Case "2": c_state = "LISTENING"
  Case "3": c_state = "SYN_SENT"
  Case "4": c_state = "SYN_RCVD"
  Case "5": c_state = "ESTABLISHED"
  Case "6": c_state = "FIN_WAIT1"
  Case "7": c_state = "FIN_WAIT2"
  Case "8": c_state = "CLOSE_WAIT"
  Case "9": c_state = "CLOSING"
  Case "10": c_state = "LAST_ACK"
  Case "11": c_state = "TIME_WAIT"
  Case "12": c_state = "DELETE_TCB"
  End Select
End Function
