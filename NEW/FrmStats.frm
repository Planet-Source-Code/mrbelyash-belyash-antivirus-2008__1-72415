VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ñòàòèñòèêà"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7725
   Icon            =   "FrmStats.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   7200
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "TCP Statistics"
      Height          =   2145
      Left            =   3750
      TabIndex        =   8
      Top             =   120
      Width           =   3885
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   1920
         Top             =   1170
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1815
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parameter"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "IP Statistics"
      Height          =   2145
      Left            =   30
      TabIndex        =   6
      Top             =   2310
      Width           =   7605
      Begin VB.Timer Timer2 
         Interval        =   250
         Left            =   0
         Top             =   240
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   1815
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   3201
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parameter"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "UDP Statistics"
      Height          =   2145
      Left            =   30
      TabIndex        =   4
      Top             =   120
      Width           =   3645
      Begin VB.Timer Timer3 
         Interval        =   250
         Left            =   840
         Top             =   1920
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   1755
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   3096
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parameter"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ICMP Statistics (IN)"
      Height          =   2505
      Left            =   3870
      TabIndex        =   2
      Top             =   4620
      Width           =   3795
      Begin VB.Timer Timer4 
         Interval        =   250
         Left            =   0
         Top             =   240
      End
      Begin MSComctlLib.ListView ListView4 
         Height          =   2205
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   3889
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parameter"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   1411
         EndProperty
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "ICMP Statistics (OUT)"
      Height          =   2505
      Left            =   30
      TabIndex        =   0
      Top             =   4620
      Width           =   3735
      Begin VB.Timer Timer5 
         Interval        =   250
         Left            =   0
         Top             =   240
      End
      Begin MSComctlLib.ListView ListView5 
         Height          =   2175
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Parameter"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   1411
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim IP As MIB_IPSTATS
Dim tcp As MIB_TCPSTATS
Dim udp As MIB_UDPSTATS
Dim icmp As MIBICMPINFO
Dim tStats As MIB_TCPSTATS

Public OldnUmProc As Long

Private Sub LoadProcessInfo()

LoadAllNTProcesses
DoEvents

If OldnUmProc = ProgEntries Then Exit Sub

'ListView6.ListItems.Clear

'For i = 0 To ProgEntries

'If CurProcesses(i).ProcessID <> "0" Then

 '   Set Item = ListView6.ListItems.Add(, , CurProcesses(i).FileName)
  '  Item.SubItems(1) = CurProcesses(i).ProcessID
   ' Item.Tag = i

'End If

'Next i

OldnUmProc = ProgEntries
End Sub

Private Sub Form_Load()
    '
    'Configure the ListView control
    '
'StatusBar.Panels(1).Width = Me.Width
'ListView6.ColumnHeaders(1).Width = ListView6.Width / 2 + 1000
'ListView6.ColumnHeaders(2).Width = ListView6.Width / 2 - 1800
  

    With ListView1.ListItems
        '
        .Add , , "Timeout algorithm"
        .Add , , "Minimum timeout"
        .Add , , "Maximum timeout"
        .Add , , "Maximum connections"
        .Add , , "Active opens"
        .Add , , "Passive opens"
        .Add , , "Failed attempts"
        .Add , , "Establised connections reset"
        .Add , , "Established connections"
        .Add , , "Segments received"
        .Add , , "Segment sent"
        .Add , , "Segments retransmitted"
        .Add , , "Incoming errors"
        .Add , , "Outgoing resets"
        .Add , , "Cumulative connections"
        '
    End With

    With ListView2.ListItems
    .Add , , "IP forwarding enabled or disabled"
    .Add , , "Default time-to-live"
    .Add , , "Datagrams received"
    .Add , , "Received header errors"
    .Add , , "Received address errors"
    .Add , , "datagrams forwarded"
    .Add , , "datagrams with unknown protocol"
    .Add , , "received datagrams discarded"
    .Add , , "received datagrams delivered"
    .Add , , "outgoing datagrams requested"
    .Add , , "outgoing datagrams discarded"
    .Add , , "sent datagrams discarded"
    .Add , , "datagrams for which no route"
    .Add , , "datagrams for which all frags didn't arrive"
    .Add , , "datagrams requiring reassembly"
    .Add , , "successful reassemblies"
    .Add , , "failed reassemblies"
    .Add , , "successful fragmentations"
    .Add , , "failed fragmentations"
    .Add , , "datagrams fragmented"
    .Add , , "number of interfaces on computer"
    .Add , , "number of IP address on computer"
    .Add , , "number of routes in routing table"
    End With

    With ListView3.ListItems
    .Add , , "received datagrams"
    .Add , , "datagrams for which no port"
    .Add , , "errors on received datagrams"
    .Add , , "sent datagrams"
    .Add , , "number of entries in UDP listener table"
    End With
    
    With ListView4.ListItems
    .Add , , "number of messages"
    .Add , , "number of errors"
    .Add , , "destination unreachable messages"
    .Add , , "time-to-live exceeded messages"
    .Add , , "parameter problem messages"
    .Add , , "source quench messages"
    .Add , , "redirection messages"
    .Add , , "echo requests"
    .Add , , "echo replies"
    .Add , , "timestamp requests"
    .Add , , "timestamp replies"
    .Add , , "address mask requests"
    .Add , , "address mask replies"
    End With
    
    With ListView5.ListItems
    .Add , , "number of messages"
    .Add , , "number of errors"
    .Add , , "destination unreachable messages"
    .Add , , "time-to-live exceeded messages"
    .Add , , "parameter problem messages"
    .Add , , "source quench messages"
    .Add , , "redirection messages"
    .Add , , "echo requests"
    .Add , , "echo replies"
    .Add , , "timestamp requests"
    .Add , , "timestamp replies"
    .Add , , "address mask requests"
    .Add , , "address mask replies"
    End With

Call GetTcpStatistics(tStats)

With tStats
ListView1.ListItems(1).SubItems(1) = .dwRtoAlgorithm
ListView1.ListItems(2).SubItems(1) = .dwRtoMin
ListView1.ListItems(3).SubItems(1) = .dwRtoMax
ListView1.ListItems(4).SubItems(1) = .dwMaxConn
ListView1.ListItems(5).SubItems(1) = .dwActiveOpens
ListView1.ListItems(6).SubItems(1) = .dwPassiveOpens
ListView1.ListItems(7).SubItems(1) = .dwAttemptFails
ListView1.ListItems(8).SubItems(1) = .dwEstabResets
ListView1.ListItems(9).SubItems(1) = .dwCurrEstab
ListView1.ListItems(10).SubItems(1) = .dwInSegs
ListView1.ListItems(11).SubItems(1) = .dwOutSegs
ListView1.ListItems(12).SubItems(1) = .dwRetransSegs
ListView1.ListItems(13).SubItems(1) = .dwInErrs
ListView1.ListItems(14).SubItems(1) = .dwOutRsts
ListView1.ListItems(15).SubItems(1) = .dwNumConns
End With
DoEvents

Call GetIpStatistics(IP)

With IP
ListView2.ListItems(1).SubItems(1) = .dwForwarding
ListView2.ListItems(2).SubItems(1) = .dwDefaultTTL
ListView2.ListItems(3).SubItems(1) = .dwInReceives
ListView2.ListItems(4).SubItems(1) = .dwInHdrErrors
ListView2.ListItems(5).SubItems(1) = .dwInAddrErrors
ListView2.ListItems(6).SubItems(1) = .dwForwDatagrams
ListView2.ListItems(7).SubItems(1) = .dwInUnknownProtos
ListView2.ListItems(8).SubItems(1) = .dwInDiscards
ListView2.ListItems(9).SubItems(1) = .dwInDelivers
ListView2.ListItems(10).SubItems(1) = .dwOutRequests
ListView2.ListItems(11).SubItems(1) = .dwRoutingDiscards
ListView2.ListItems(12).SubItems(1) = .dwOutDiscards
ListView2.ListItems(13).SubItems(1) = .dwOutNoRoutes
ListView2.ListItems(14).SubItems(1) = .dwReasmTimeout
ListView2.ListItems(15).SubItems(1) = .dwReasmReqds
ListView2.ListItems(16).SubItems(1) = .dwReasmOks
ListView2.ListItems(17).SubItems(1) = .dwReasmFails
ListView2.ListItems(18).SubItems(1) = .dwFragOks
ListView2.ListItems(19).SubItems(1) = .dwFragFails
ListView2.ListItems(20).SubItems(1) = .dwFragCreates
ListView2.ListItems(21).SubItems(1) = .dwNumIf
ListView2.ListItems(22).SubItems(1) = .dwNumAddr
ListView2.ListItems(23).SubItems(1) = .dwNumRoutes
End With
DoEvents

Call GetUdpStatistics(udp)

With udp
ListView3.ListItems(1).SubItems(1) = .dwInDatagrams
ListView3.ListItems(2).SubItems(1) = .dwNoPorts
ListView3.ListItems(3).SubItems(1) = .dwInErrors
ListView3.ListItems(4).SubItems(1) = .dwOutDatagrams
ListView3.ListItems(5).SubItems(1) = .dwNumAddrs
End With
DoEvents

Call GetIcmpStatistics(icmp)

With icmp
ListView4.ListItems(1).SubItems(1) = .icmpInStats.dwMsgs
ListView4.ListItems(2).SubItems(1) = .icmpInStats.dwErrors
ListView4.ListItems(3).SubItems(1) = .icmpInStats.dwDestUnreachs
ListView4.ListItems(4).SubItems(1) = .icmpInStats.dwTimeExcds
ListView4.ListItems(5).SubItems(1) = .icmpInStats.dwParmProbs
ListView4.ListItems(6).SubItems(1) = .icmpInStats.dwSrcQuenchs
ListView4.ListItems(7).SubItems(1) = .icmpInStats.dwRedirects
ListView4.ListItems(8).SubItems(1) = .icmpInStats.dwEchos
ListView4.ListItems(9).SubItems(1) = .icmpInStats.dwEchoReps
ListView4.ListItems(10).SubItems(1) = .icmpInStats.dwTimestamps
ListView4.ListItems(11).SubItems(1) = .icmpInStats.dwTimestampReps
ListView4.ListItems(12).SubItems(1) = .icmpInStats.dwAddrMasks
ListView4.ListItems(13).SubItems(1) = .icmpInStats.dwAddrMaskReps
DoEvents
ListView5.ListItems(1).SubItems(1) = .icmpOutStats.dwMsgs
ListView5.ListItems(2).SubItems(1) = .icmpOutStats.dwErrors
ListView5.ListItems(3).SubItems(1) = .icmpOutStats.dwDestUnreachs
ListView5.ListItems(4).SubItems(1) = .icmpOutStats.dwTimeExcds
ListView5.ListItems(5).SubItems(1) = .icmpOutStats.dwParmProbs
ListView5.ListItems(6).SubItems(1) = .icmpOutStats.dwSrcQuenchs
ListView5.ListItems(7).SubItems(1) = .icmpOutStats.dwRedirects
ListView5.ListItems(8).SubItems(1) = .icmpOutStats.dwEchos
ListView5.ListItems(9).SubItems(1) = .icmpOutStats.dwEchoReps
ListView5.ListItems(10).SubItems(1) = .icmpOutStats.dwTimestamps
ListView5.ListItems(11).SubItems(1) = .icmpOutStats.dwTimestampReps
ListView5.ListItems(12).SubItems(1) = .icmpOutStats.dwAddrMasks
ListView5.ListItems(13).SubItems(1) = .icmpOutStats.dwAddrMaskReps
End With
DoEvents

LoadProcessInfo
End Sub

Private Sub ListView6_DblClick()
Dim MsgAnswer

MsgAnswer = MsgBox("Terminate Program : " & Right(CurProcesses(ListView6.SelectedItem.Tag).FileName, Len(CurProcesses(ListView6.SelectedItem.Tag).FileName) - InStrRev(CurProcesses(ListView6.SelectedItem.Tag).FileName, "\")), vbYesNo + vbExclamation, "Terminate Program")

If MsgAnswer = vbNo Then Exit Sub

If KillProcessById(CurProcesses(ListView6.SelectedItem.Tag).ProcessID) = True Then
StatusBar.Panels(1).Text = "Program Terminated: " & Right(CurProcesses(ListView6.SelectedItem.Tag).FileName, Len(CurProcesses(ListView6.SelectedItem.Tag).FileName) - InStrRev(CurProcesses(ListView6.SelectedItem.Tag).FileName, "\"))
Else
StatusBar.Panels(1).Text = "Program Terminate Failed: " & Right(CurProcesses(ListView6.SelectedItem.Tag).FileName, Len(CurProcesses(ListView6.SelectedItem.Tag).FileName) - InStrRev(CurProcesses(ListView6.SelectedItem.Tag).FileName, "\"))
End If

DoEvents
LoadProcessInfo

End Sub


Private Sub Timer1_Timer()
    UpdateStats1
End Sub

Private Sub UpdateStats1()
    '
    Dim tStats          As MIB_TCPSTATS
    Static tStaticStats As MIB_TCPSTATS
    '
    Dim lRetValue       As Long
    '
    Dim blnIsSent       As Boolean
    Dim blnIsRecv       As Boolean
    '
    lRetValue = GetTcpStatistics(tStats)
    '
    With tStats
        '
        If Not tStaticStats.dwRtoAlgorithm = .dwRtoAlgorithm Then _
            ListView1.ListItems(1).SubItems(1) = .dwRtoAlgorithm
        If Not tStaticStats.dwRtoMin = .dwRtoMin Then _
            ListView1.ListItems(2).SubItems(1) = .dwRtoMin
        If Not tStaticStats.dwRtoMax = .dwRtoMax Then _
            ListView1.ListItems(3).SubItems(1) = .dwRtoMax
        If Not tStaticStats.dwMaxConn = .dwMaxConn Then _
            ListView1.ListItems(4).SubItems(1) = .dwMaxConn
        If Not tStaticStats.dwActiveOpens = .dwActiveOpens Then _
            ListView1.ListItems(5).SubItems(1) = .dwActiveOpens
        If Not tStaticStats.dwPassiveOpens = .dwPassiveOpens Then _
            ListView1.ListItems(6).SubItems(1) = .dwPassiveOpens
        If Not tStaticStats.dwAttemptFails = .dwAttemptFails Then _
            ListView1.ListItems(7).SubItems(1) = .dwAttemptFails
        If Not tStaticStats.dwEstabResets = .dwEstabResets Then _
            ListView1.ListItems(8).SubItems(1) = .dwEstabResets
        If Not tStaticStats.dwCurrEstab = .dwCurrEstab Then _
            ListView1.ListItems(9).SubItems(1) = .dwCurrEstab
        If Not tStaticStats.dwInSegs = .dwInSegs Then _
            ListView1.ListItems(10).SubItems(1) = .dwInSegs
        If Not tStaticStats.dwOutSegs = .dwOutSegs Then _
            ListView1.ListItems(11).SubItems(1) = .dwOutSegs
        If Not tStaticStats.dwRetransSegs = .dwRetransSegs Then _
            ListView1.ListItems(12).SubItems(1) = .dwRetransSegs
        If Not tStaticStats.dwInErrs = .dwInErrs Then _
            ListView1.ListItems(13).SubItems(1) = .dwInErrs
        If Not tStaticStats.dwOutRsts = .dwOutRsts Then _
            ListView1.ListItems(14).SubItems(1) = .dwOutRsts
        If Not tStaticStats.dwNumConns = .dwNumConns Then _
            ListView1.ListItems(15).SubItems(1) = .dwNumConns
        '
    End With

    tStaticStats = tStats
    '
End Sub
Private Sub UpdateStats2()
On Error Resume Next
Static ip2 As MIB_IPSTATS
Dim lRetValue       As Long

lRetValue = GetIpStatistics(IP)

With IP
If Not ip2.dwForwarding = .dwForwarding Then _
ListView2.ListItems(1).SubItems(1) = .dwForwarding
If Not ip2.dwDefaultTTL = .dwDefaultTTL Then _
ListView2.ListItems(2).SubItems(1) = .dwDefaultTTL
If Not ip2.dwInReceives = .dwInReceives Then _
ListView2.ListItems(3).SubItems(1) = .dwInReceives
If Not ip2.dwInHdrErrors = .dwInHdrErrors Then _
ListView2.ListItems(4).SubItems(1) = .dwInHdrErrors
If Not ip2.dwInAddrErrors = .dwInAddrErrors Then _
ListView2.ListItems(5).SubItems(1) = .dwInAddrErrors
If Not ip2.dwForwDatagrams = .dwForwDatagrams Then _
ListView2.ListItems(6).SubItems(1) = .dwForwDatagrams
If Not ip2.dwInUnknownProtos = .dwInUnknownProtos Then _
ListView2.ListItems(7).SubItems(1) = .dwInUnknownProtos
If Not ip2.dwInDiscards = .dwInDiscards Then _
ListView2.ListItems(8).SubItems(1) = .dwInDiscards
If Not ip2.dwInDelivers = .dwInDelivers Then _
ListView2.ListItems(9).SubItems(1) = .dwInDelivers
If Not ip2.dwOutRequests = .dwOutRequests Then _
ListView2.ListItems(10).SubItems(1) = .dwOutRequests
If Not ip2.dwRoutingDiscards = .dwRoutingDiscards Then _
ListView2.ListItems(11).SubItems(1) = .dwRoutingDiscards
If Not ip2.dwOutDiscards = .dwOutDiscards Then _
ListView2.ListItems(12).SubItems(1) = .dwOutDiscards
If Not ip2.dwOutNoRoutes = .dwOutNoRoutes Then _
ListView2.ListItems(13).SubItems(1) = .dwOutNoRoutes
If Not ip2.dwReasmTimeout = .dwReasmTimeout Then _
ListView2.ListItems(14).SubItems(1) = .dwReasmTimeout
If Not ip2.dwReasmReqds = .dwReasmReqds Then _
ListView2.ListItems(15).SubItems(1) = .dwReasmReqds
If Not ip2.dwReasmOks = .dwReasmOks Then _
ListView2.ListItems(16).SubItems(1) = .dwReasmOks
If Not ip2.dwReasmFails = .dwReasmFails Then _
ListView2.ListItems(17).SubItems(1) = .dwReasmFails
If Not ip2.dwFragOks = .dwFragOks Then _
ListView2.ListItems(18).SubItems(1) = .dwFragOks
If Not ip2.dwFragFails = .dwFragFails Then _
ListView2.ListItems(19).SubItems(1) = .dwFragFails
If Not ip2.dwFragCreates = .dwFragCreates Then _
ListView2.ListItems(20).SubItems(1) = .dwFragCreates
If Not ip2.dwNumIf = .dwNumIf Then _
ListView2.ListItems(21).SubItems(1) = .dwNumIf
If Not ip2.dwNumAddr = .dwNumAddr Then _
ListView2.ListItems(22).SubItems(1) = .dwNumAddr
If Not ip2.dwNumRoutes = .dwNumRoutes Then _
ListView2.ListItems(23).SubItems(1) = .dwNumRoutes
End With

ip2 = IP
End Sub

Private Sub UpdateStats3()
On Error Resume Next
Dim lRetValue       As Long
Static udp2 As MIB_UDPSTATS

lRetValue = GetUdpStatistics(udp)

With udp
If Not udp2.dwInDatagrams = .dwInDatagrams Then _
ListView3.ListItems(1).SubItems(1) = .dwInDatagrams

If Not udp2.dwNoPorts = .dwNoPorts Then _
ListView3.ListItems(2).SubItems(1) = .dwNoPorts

If Not udp2.dwInErrors = .dwInErrors Then _
ListView3.ListItems(3).SubItems(1) = .dwInErrors

If Not udp2.dwOutDatagrams = .dwOutDatagrams Then _
ListView3.ListItems(4).SubItems(1) = .dwOutDatagrams

If Not udp2.dwNumAddrs = .dwNumAddrs Then _
ListView3.ListItems(5).SubItems(1) = .dwNumAddrs

End With

udp2 = udp
End Sub
Private Sub UpdateStats4()
'On Error Resume Next
Dim lRetValue       As Long
Static icmp2 As MIBICMPINFO

lRetValue = GetIcmpStatistics(icmp)

With icmp
If Not icmp2.icmpOutStats.dwMsgs = .icmpOutStats.dwMsgs Then _
ListView4.ListItems(1).SubItems(1) = .icmpOutStats.dwMsgs
If Not icmp2.icmpOutStats.dwErrors = .icmpOutStats.dwErrors Then _
ListView4.ListItems(2).SubItems(1) = .icmpOutStats.dwErrors
If Not icmp2.icmpOutStats.dwDestUnreachs = .icmpOutStats.dwDestUnreachs Then _
ListView4.ListItems(3).SubItems(1) = .icmpOutStats.dwDestUnreachs
If Not icmp2.icmpOutStats.dwTimeExcds = .icmpOutStats.dwTimeExcds Then _
ListView4.ListItems(4).SubItems(1) = .icmpOutStats.dwTimeExcds
If Not icmp2.icmpOutStats.dwParmProbs = .icmpOutStats.dwParmProbs Then _
ListView4.ListItems(5).SubItems(1) = .icmpOutStats.dwParmProbs
If Not icmp2.icmpOutStats.dwSrcQuenchs = .icmpOutStats.dwSrcQuenchs Then _
ListView4.ListItems(6).SubItems(1) = .icmpOutStats.dwSrcQuenchs
If Not icmp2.icmpOutStats.dwRedirects = .icmpOutStats.dwRedirects Then _
ListView4.ListItems(7).SubItems(1) = .icmpOutStats.dwRedirects
If Not icmp2.icmpOutStats.dwEchos = .icmpOutStats.dwEchos Then _
ListView4.ListItems(8).SubItems(1) = .icmpOutStats.dwEchos
If Not icmp2.icmpOutStats.dwEchoReps = .icmpOutStats.dwEchoReps Then _
ListView4.ListItems(9).SubItems(1) = .icmpOutStats.dwEchoReps
If Not icmp2.icmpOutStats.dwTimestamps = .icmpOutStats.dwTimestamps Then _
ListView4.ListItems(10).SubItems(1) = .icmpOutStats.dwTimestamps
If Not icmp2.icmpOutStats.dwTimestampReps = .icmpOutStats.dwTimestampReps Then _
ListView4.ListItems(11).SubItems(1) = .icmpOutStats.dwTimestampReps
If Not icmp2.icmpOutStats.dwAddrMasks = .icmpOutStats.dwAddrMasks Then _
ListView4.ListItems(12).SubItems(1) = .icmpOutStats.dwAddrMasks
If Not icmp2.icmpOutStats.dwAddrMaskReps = .icmpOutStats.dwAddrMaskReps Then _
ListView4.ListItems(13).SubItems(1) = .icmpOutStats.dwAddrMaskReps
End With

icmp2 = icmp
End Sub
Private Sub UpdateStats5()
On Error Resume Next
Dim lRetValue       As Long
Static icmp2 As MIBICMPINFO

lRetValue = GetIcmpStatistics(icmp)

With icmp
If Not icmp2.icmpInStats.dwMsgs = .icmpInStats.dwMsgs Then _
ListView4.ListItems(1).SubItems(1) = .icmpInStats.dwMsgs
If Not icmp2.icmpInStats.dwErrors = .icmpInStats.dwErrors Then _
ListView4.ListItems(2).SubItems(1) = .icmpInStats.dwErrors
If Not icmp2.icmpInStats.dwDestUnreachs = .icmpInStats.dwDestUnreachs Then _
ListView4.ListItems(3).SubItems(1) = .icmpInStats.dwDestUnreachs
If Not icmp2.icmpInStats.dwTimeExcds = .icmpInStats.dwTimeExcds Then _
ListView4.ListItems(4).SubItems(1) = .icmpInStats.dwTimeExcds
If Not icmp2.icmpInStats.dwParmProbs = .icmpInStats.dwParmProbs Then _
ListView4.ListItems(5).SubItems(1) = .icmpInStats.dwParmProbs
If Not icmp2.icmpInStats.dwSrcQuenchs = .icmpInStats.dwSrcQuenchs Then _
ListView4.ListItems(6).SubItems(1) = .icmpInStats.dwSrcQuenchs
If Not icmp2.icmpInStats.dwRedirects = .icmpInStats.dwRedirects Then _
ListView4.ListItems(7).SubItems(1) = .icmpInStats.dwRedirects
If Not icmp2.icmpInStats.dwEchos = .icmpInStats.dwEchos Then _
ListView4.ListItems(8).SubItems(1) = .icmpInStats.dwEchos
If Not icmp2.icmpInStats.dwEchoReps = .icmpInStats.dwEchoReps Then _
ListView4.ListItems(9).SubItems(1) = .icmpInStats.dwEchoReps
If Not icmp2.icmpInStats.dwTimestamps = .icmpInStats.dwTimestamps Then _
ListView4.ListItems(10).SubItems(1) = .icmpInStats.dwTimestamps
If Not icmp2.icmpInStats.dwTimestampReps = .icmpInStats.dwTimestampReps Then _
ListView4.ListItems(11).SubItems(1) = .icmpInStats.dwTimestampReps
If Not icmp2.icmpInStats.dwAddrMasks = .icmpInStats.dwAddrMasks Then _
ListView4.ListItems(12).SubItems(1) = .icmpInStats.dwAddrMasks
If Not icmp2.icmpInStats.dwAddrMaskReps = .icmpInStats.dwAddrMaskReps Then _
ListView4.ListItems(13).SubItems(1) = .icmpInStats.dwAddrMaskReps
End With

icmp2 = icmp

End Sub

Private Sub Timer2_Timer()
UpdateStats2
End Sub

Private Sub Timer3_Timer()
UpdateStats3
End Sub

Private Sub Timer4_Timer()
UpdateStats4
End Sub

Private Sub Timer5_Timer()
UpdateStats5
End Sub

Private Sub Timer6_Timer()
LoadProcessInfo
End Sub
