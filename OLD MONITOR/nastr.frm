VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form nastr 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Íàñòðîéêè ïðîãðàììû"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   Icon            =   "nastr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin monitor.xpcmdbutton Command2 
      Height          =   375
      Left            =   5730
      TabIndex        =   13
      Top             =   180
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Ñïðàâêà"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   6906
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   -2147483628
      TabCaption(0)   =   "Îáùèå"
      TabPicture(0)   =   "nastr.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Îò÷åò"
      TabPicture(1)   =   "nastr.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Text1"
      Tab(1).Control(3)=   "Command3"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Ìîíèòîðèíã"
      TabPicture(2)   =   "nastr.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      Begin monitor.xpcmdbutton Command3 
         Height          =   285
         Left            =   -70110
         TabIndex        =   12
         Top             =   840
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame4 
         Height          =   2805
         Left            =   -74790
         TabIndex        =   4
         Top             =   540
         Width           =   5025
         Begin VB.CheckBox Check6 
            Caption         =   "Backup íàñòðîåê"
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   1050
            Width           =   1755
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Ìîíèòîðèòü ñâîé êàòàëîã"
            Height          =   315
            Left            =   120
            TabIndex        =   6
            Top             =   720
            Width           =   2445
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Ìîíèòîðèòü ðååñòð"
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   330
            Width           =   1935
         End
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -74910
         TabIndex        =   3
         Top             =   840
         Width           =   4695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ðåæèì âåäåíèÿ ëîãà"
         Height          =   1155
         Left            =   -74880
         TabIndex        =   2
         Top             =   1380
         Width           =   5355
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   225
            Left            =   2340
            TabIndex        =   8
            Text            =   "1"
            Top             =   780
            Width           =   510
         End
         Begin monitor.xpradiobutton Option2 
            Height          =   285
            Left            =   150
            TabIndex        =   9
            Top             =   750
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Äîáàâëÿòü(îãðàíè÷èòü )                Ìá"
         End
         Begin monitor.xpradiobutton Option1 
            Height          =   285
            Left            =   150
            TabIndex        =   10
            Top             =   390
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "Ïåðåçàïèñûâàòü"
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3045
         Left            =   210
         TabIndex        =   1
         Top             =   600
         Width           =   5175
         Begin monitor.xpcheckbox ldMonik 
            Height          =   315
            Left            =   90
            TabIndex        =   15
            Top             =   1380
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   1
            Caption         =   "Çàïîìèíàòü ñîñòîÿíèå ìîíèòîðèíãà"
         End
         Begin monitor.xpcheckbox Check3 
            Height          =   315
            Left            =   90
            TabIndex        =   16
            Top             =   1020
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   1
            Caption         =   "Ïåðåìåñòèòü â êàðàíòèí"
         End
         Begin monitor.xpcheckbox lbAuto 
            Height          =   315
            Left            =   90
            TabIndex        =   17
            Top             =   180
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   1
            Caption         =   "Àâòîçàãðóçêà ïðîãðàììû"
         End
         Begin monitor.xpcheckbox lbFon 
            Height          =   315
            Left            =   90
            TabIndex        =   18
            Top             =   600
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   1
            Caption         =   "Çàãðóçêà â ôîíîâîì ðåæèìå"
         End
      End
      Begin monitor.xpcheckbox Check1 
         Height          =   315
         Left            =   -74880
         TabIndex        =   11
         Top             =   420
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Âåñòè ôàéë îò÷åòà"
      End
   End
   Begin monitor.xpcmdbutton Command1 
      Height          =   375
      Left            =   5730
      TabIndex        =   14
      Top             =   630
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Çàêðûòü"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "nastr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Check4_Click()
 CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonReg", Val(Me.Check4.Value)
End Sub

Private Sub Check5_Click()
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonFolder", Val(Me.Check5.Value)

End Sub

Private Sub Check6_Click()
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BackupReg", Val(Me.Check6.Value)

End Sub

Private Sub Command1_Click()
'çàïèñàòü íàñòðîéêè
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Monitoring", Val(Me.ldMonik.Value)
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonQv", Val(Me.Check3.Value)
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartFon", Val(Me.lbFon.Value)
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStartMon", Val(Me.lbAuto.Value)
    If Me.lbAuto.Value = 1 Then
        Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.exeName & ".exe")
        'CreateStringValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.EXEName & ".exe"
    Else
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", "BGAntiVirus"
     End If

CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogMon", Val(Me.Check1.Value)
If Option1.Value = True Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogAppend", Val("1")
    
Else
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogAppend", Val("0")
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogSizeMon", Val(Me.Text2.Text)
End If
205:
Unload Me
End Sub



Private Sub Command2_Click()

ShowTopicID 1, 18
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    ShowTopicID 1, 11
End If
End Sub

Private Sub Form_Load()
Call LoadRegistry
Text1.Text = App.Path
End Sub

Sub LoadRegistry()
Me.lbAuto.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStartMon"))
Me.lbFon.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartFon"))
Me.Check3.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonQv"))
Me.Check1.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogMon"))
Me.ldMonik.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Monitoring"))
Me.Check4.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonReg"))
Me.Check5.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonFolder"))
Me.Check6.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BackupReg"))
Dim maf As Long
maf = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStartMon"))
    If maf = 1 Then
        Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.exeName & ".exe")
        'CreateStringValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.EXEName & ".exe"
    Else
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", "BGAntiVirus"
     End If
Dim appn1 As Integer
appn1 = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogAppend"))
If appn1 = 1 Then
         Option1.Value = True
         Option2.Value = False
         Me.Text2.Enabled = True
Else
         Option2.Value = True
         Option1.Value = False
End If
'    nastr.Option2.Value = False
'    nastr.Option1.Value = True
     Me.Text2.Text = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogSizeMon"))
'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If frmTest.mnuNast.Visible = False Then
frmTest.mnuNast.Visible = True
End If
Command1_Click
Unload Me

End Sub





Private Sub UpDown1_Change()
Text2.Text = UpDown1.Value
End Sub

Private Sub Option1_Click()
Text2.Enabled = False
End Sub

Private Sub Option2_Click()
Text2.Enabled = True
End Sub

Private Sub Text2_Change()
If Text2.Text > 12 Or Text2.Text <= 0 Then
    MsgBox "Ïðåâûøåíèå ìàêñèìàëüíî äîïóñòèìîãî ðàçìåðà ôàéëà ëîãà", vbCritical, pr
    Text2.Text = "12"
    
End If
End Sub
