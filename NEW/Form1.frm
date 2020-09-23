VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Ñîåäèíåíèÿ"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicQuestion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   255
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5955
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pic16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2040
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2760
      Top             =   5400
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8281
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      Icons           =   "Iml32"
      SmallIcons      =   "Iml16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remote Host"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   2
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "State"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ImageList Iml16 
      Left            =   1440
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu Menu 
      Caption         =   "Ìåíþ"
      Begin VB.Menu RefreshIT 
         Caption         =   "Îáíîâèòü"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu ListeningPortsBtn 
         Caption         =   "Èñïîëüçóåìûå ïîðòû"
      End
      Begin VB.Menu ShowStats 
         Caption         =   "Ñòàòèñòèêà"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitProg 
         Caption         =   "Âûõîä"
      End
   End
   Begin VB.Menu MenuPop 
      Caption         =   "MenuPop"
      Visible         =   0   'False
      Begin VB.Menu CloseCon 
         Caption         =   "Çàêðûòü ñîåäèíåíèå"
      End
      Begin VB.Menu TermProg 
         Caption         =   "Óáèòü ïðîöåññ"
      End
      Begin VB.Menu OptionsBtn 
         Caption         =   "Íàñòðîéêè"
         Begin VB.Menu ICMPBtn 
            Caption         =   "ICMP (ping)"
         End
         Begin VB.Menu ResolveHost 
            Caption         =   "Resolve Hostname"
         End
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Processing As Boolean
Public ShowListening As Boolean

Public Sub RefreshList()
  Dim i
  Dim Item As ListItem
  
  Processing = True
  
RefreshStack
DoEvents

LoadNTProcess
DoEvents

ListView1.ListItems.Clear

For i = 0 To GetEntryCount
     
     If ShowListening = False Then
     If Connection(i).State = "2" Then GoTo IsListening
     End If
     
    If Connection(i).FileName = "" Then
    Set Item = ListView1.ListItems.Add(, , "Íåèçâåñòíîå")
    Item.SubItems(5) = ""
    Else
    Set Item = ListView1.ListItems.Add(, , Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\")))
    Item.SubItems(5) = Connection(i).FileName
    End If
    
    Item.SubItems(1) = Connection(i).LocalPort
    Item.SubItems(2) = Connection(i).RemoteHost
    Item.SubItems(3) = Connection(i).RemotePort
    Item.SubItems(4) = c_state(Connection(i).State)
        Item.Tag = i

IsListening:
Next i

GetAllIcons
DoEvents

ShowIcons
DoEvents

Processing = False
End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
On Error Resume Next
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long

If Connection(ListView1.ListItems(Index).Tag).FileName = "" Then
Set imgObj = Iml16.ListImages.Add(Index, , PicQuestion.Image)
Exit Function
End If

'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

  'Small Icon
  With Pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, Pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = Iml16.ListImages.Add(Index, , Pic16.Image)
End Function

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView1
  '.ListItems.Clear
  .SmallIcons = Iml16   'Small
  For Each Item In .ListItems
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

    ListView1.SmallIcons = Nothing
    Iml16.ListImages.Clear
    
'On Local Error Resume Next
For Each Item In ListView1.ListItems
  FileName = Connection(Item.Tag).FileName

  GetIcon FileName, Item.Index
   
Next

End Sub

Private Sub CloseCon_Click()
If TerminateThisConnection(ListView1.SelectedItem.Tag) = True Then
StatusBar.Panels(2).Text = "Connection Closed: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
Else
StatusBar.Panels(2).Text = "Close Connection Failed: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
End If
End Sub

Private Sub ExitProg_Click()
Do Until Processing = False
DoEvents
Loop
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.Width - 150
ListView1.Left = 0
ListView1.Height = Me.Height - 1050

ListView1.ColumnHeaders(1).Width = 1300
ListView1.ColumnHeaders(2).Width = 1100
ListView1.ColumnHeaders(3).Width = 1100
ListView1.ColumnHeaders(4).Width = 1100
ListView1.ColumnHeaders(5).Width = 1100
ListView1.ColumnHeaders(6).Width = ListView1.Width \ 2 + 1000

StatusBar.Panels(1).Width = Me.Width / 4
StatusBar.Panels(2).Width = Me.Width / 2
StatusBar.Panels(3).Width = Me.Width / 4

Progbar.Width = StatusBar1.Panels(1).Width
Progbar.Height = StatusBar1.Height - 80
End Sub

Private Sub ICMPBtn_Click()
Dim RemoteHostNet
RemoteHostNet = Connection(ListView1.SelectedItem.Tag).RemoteHost
StatusBar.Panels(2).Text = "(ICMP) " & RemoteHostNet & " : " & Ping(RemoteHostNet, 2000)
DoEvents
End Sub

Private Sub ListeningPortsBtn_Click()
If ShowListening = True Then
ListeningPortsBtn.Checked = False
ShowListening = False
RefreshList
Else
ListeningPortsBtn.Checked = True
ShowListening = True
RefreshList
End If
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

Select Case Button

    Case vbLeftButton
    'MsgBox ListView1.SelectedItem.Key

    Case vbRightButton
    CloseCon.Caption = "Ïðåðâàòü ñîåäèíåíèå: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
    TermProg.Caption = "Óáèòü: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
    PopupMenu MenuPop

End Select
End Sub

Private Sub RefreshIT_Click()
RefreshList
End Sub

Private Sub ResolveHost_Click()
Dim HostAddr
Dim hostname

hostname = ResolveHostname(Connection(ListView1.SelectedItem.Tag).RemoteHost)
DoEvents

If hostname = "" Then
StatusBar.Panels(2).Text = "Íå ìîãó äîñòó÷àòüñÿ äî õîñòà"
Exit Sub
End If

HostAddr = Connection(ListView1.SelectedItem.Tag).RemoteHost

StatusBar.Panels(2).Text = hostname & " (" & HostAddr & ")"
End Sub

Private Sub ShowStats_Click()
FrmStats.Show
End Sub

Private Sub TermProg_Click()
If KillProcessById(Connection(ListView1.SelectedItem.Tag).ProcessID) = True Then
StatusBar.Panels(2).Text = "Óáèò ïðîöåññ: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
Else
StatusBar.Panels(2).Text = "Îøèáêà òåðìèíèðîâàíèÿ: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
End If
End Sub

Private Sub Timer1_Timer()
If Processing = True Then Exit Sub

If IsNetConnectOnline = False Then
StatusBar.Panels(3).Text = "Îòêëþ÷åí îò ñåòè"
Exit Sub
Else
StatusBar.Panels(3).Text = "Â ñåòè"
End If

If GetRefresh = True Then RefreshList

End Sub

