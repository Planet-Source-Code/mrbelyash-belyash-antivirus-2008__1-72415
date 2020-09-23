VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   4560
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2370
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3150
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   840
      Left            =   180
      TabIndex        =   6
      Top             =   3420
      Width           =   6855
      Begin Scaner.XPProgressBar pb1 
         Height          =   315
         Left            =   90
         TabIndex        =   12
         Top             =   450
         Visible         =   0   'False
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   556
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE WAIT...LOADING SYSTEM DATABASE..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFC0C0&
         Height          =   525
         Left            =   210
         TabIndex        =   10
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label LD 
         BackStyle       =   0  'Transparent
         Caption         =   "PLEASE WAIT...LOADING SYSTEM DATABASE..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   180
         TabIndex        =   7
         Top             =   225
         Width           =   6495
      End
   End
   Begin VB.ListBox list1 
      Height          =   255
      Left            =   4980
      TabIndex        =   9
      Top             =   390
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.ListBox lstRunning 
      Height          =   450
      Left            =   2070
      TabIndex        =   8
      Top             =   330
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2655
      Top             =   270
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2008"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6630
      TabIndex        =   13
      Top             =   1290
      Width           =   405
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "OS:Win XP or higher"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4590
      TabIndex        =   11
      Top             =   3150
      Width           =   1995
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   2085
      Left            =   315
      Top             =   315
      Width           =   2040
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   315
      Picture         =   "Dialog1.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   2010
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Mr.Belyash && Company Ltd."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4605
      TabIndex        =   5
      Top             =   2880
      Width           =   2355
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright: 2008"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4605
      TabIndex        =   4
      Top             =   2610
      Width           =   2355
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4605
      TabIndex        =   3
      Top             =   2250
      Width           =   2355
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Intelligent Systems"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   3150
      TabIndex        =   2
      Top             =   720
      Width           =   3210
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Licensed To:   FREEWARE"
      Height          =   285
      Left            =   4500
      TabIndex        =   1
      Top             =   225
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Belyash AntiTrojan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   600
      Left            =   2730
      TabIndex        =   0
      Top             =   1305
      Width           =   4155
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   4275
      Left            =   75
      Top             =   120
      Width           =   7080
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private programs As New ProcessList
Public pass As String
Public abortprov As Boolean
Public showOkno As Boolean


Private Sub Form_Activate()
Static v As Integer
v = v + 1
If v > 1 Then
    Exit Sub
Else
    Timer1.Enabled = True
End If
End Sub
Private Sub Form_Load()
If App.PrevInstance = True Then
'ïîçæå çàìåíèòü íà API ôóíêöèþ,ò.ê. õóéíÿ ýòî à íå ïðîâåðêà
    MsgBox "Ïðèëîæåíèå óæå çàïóùåíî", vbCritical
    End
End If
modScanVirus.selfTesting
abortprov = False
'Call regPrilog 'çàðåãåñòðèðîâàíî ëè
'ProcCount = 0
'procCountSpl = 0
'procFoundSpl = 0
'procCleanSpl = 0
Label5.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision & "b"
LD.Caption = "PLEASE WAIT...LOADING SYSTEM DATABASE..."
modScanVirus.LogPrint "Çàïóñê ïðîãðàììû "
'Label2.Caption = "Licensed To "
'Me.lvProcess.ListItems.Clear
'Call GetProcessList(Me.lvProcess)
'pb1.Min = 0
'pb1.Max = ProcCount



'SetTopMostWindow Me.hwnd, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Me.Timer1.Enabled = True Then Me.Timer1.Enabled = False
Set frmSplash = Nothing

End Sub



Private Sub Timer1_Timer()
'Me.checkmemory
'Call ReadSig
modScanVirus.sbor
Check_activitiproc
Timer1.Enabled = False
Load frmMain
Unload Me
End Sub
 Sub Check_activitiproc()
          procCountSpl = 0
     procFoundSpl = 0
    procCleanSpl = 0
'ïðè âêëþ÷åíèè ñêàíèì âñå ïðîöåññû
On Error GoTo 100
'frmMain.lblCount.Caption = 0
'>> get new list of whats running
Frame1.Refresh
programs.CheckProcesses
pb1.Visible = True
'pb1.Min = 1
'pb1.Max = programs.processCount
pb1.Value = 0
Dim dProgress As Integer
dProgress = 100 / programs.processCount
LD.Caption = "PLEASE WAIT...CHECK MEMORY..."
Label4.Caption = "PLEASE WAIT...CHECK MEMORY..."
LD.Refresh
Label4.Refresh
'>> clear our listbox
frmSplash.lstRunning.Clear
List1.Clear
'>> fill our list box
Dim i As Integer
For i = 1 To programs.processCount - 1
pb1.Value = i * dProgress
If abortprov = True Then

    Exit Sub
End If
'procCountSpl = procCountSpl + 1
'pb1.Value = i
procCountSpl = procCountSpl + 1
'frmMain.lblCount.Caption = Int(frmMain.lblCount.Caption) + 1

DoEvents
 lstRunning.AddItem programs.ProcessName(i)
  List1.AddItem programs.ProcessHandle(i)
  If Left(programs.ProcessName(i), 1) <> "\" Then
                '####
                If FileorFolderExists(programs.ProcessName(i)) = False Then
                    'ResumeThreads (hNumrer1)
                    GoTo 2
                End If
      If modScanVirus.ScanFilepR("c:", programs.ProcessName(i), programs.ProcessHandle(i)) = False Then
      Debug.Print "ïðîöåññû" + programs.ProcessName(i) + "==" + CStr(programs.ProcessHandle(i))
             'ResumeThreads (hNumrer1)
            'Exit Sub
       End If
     
      
                
   End If
2:
Next
lstRunning.Clear
List1.Clear
Exit Sub
100:
MsgBox "" + Error$
End Sub

