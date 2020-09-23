VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Belyash AntiTrojan 2008 beta"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   -450
   ClientWidth     =   9420
   Icon            =   "frmMain2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain2.frx":76CA
   ScaleHeight     =   8040
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      DataField       =   "CRC"
      DataSource      =   "Data1"
      Height          =   435
      Left            =   4380
      TabIndex        =   110
      Text            =   "Text2"
      Top             =   180
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Vbase"
      Top             =   750
      Visible         =   0   'False
      Width           =   3540
   End
   Begin VB.ListBox lstRunning 
      Height          =   255
      Left            =   1770
      TabIndex        =   65
      Top             =   120
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.ListBox list1 
      Height          =   255
      Left            =   1800
      TabIndex        =   64
      Top             =   420
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Timer tmrMonitor 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   3780
      Top             =   1080
   End
   Begin VB.Timer tmrAutoRefresh 
      Enabled         =   0   'False
      Left            =   7440
      Top             =   1170
   End
   Begin MSComctlLib.ImageList img16x16 
      Left            =   6570
      Top             =   1050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":119024
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":11909F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":119340
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":1195EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":11988F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":119ACC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":119D70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   60
      Top             =   7755
      Width           =   9420
      _ExtentX        =   16616
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   8644
            MinWidth        =   8644
            Object.ToolTipText     =   "Ñêàíèðóåìûé ôàéë"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1834
            MinWidth        =   1834
            Object.ToolTipText     =   "Êîë-âî ïðîâåðåíûõ"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
            Object.ToolTipText     =   "Êîë-âî íàéäåíûõ"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1059
            MinWidth        =   1059
            Object.ToolTipText     =   "Êîë-âî óäàëåííûõ"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2363
            MinWidth        =   2363
            Object.ToolTipText     =   "Äàòà ïîñëåäíåãî îáíîâëåíèÿ"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1305
            MinWidth        =   1305
            Object.ToolTipText     =   "Êîë-âî çàïèñåé"
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameScan 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   5955
      Left            =   2160
      TabIndex        =   0
      Top             =   1800
      Width           =   7290
      Begin MSComctlLib.ListView lvVirusFound 
         Height          =   2130
         Left            =   60
         TabIndex        =   51
         Top             =   3510
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   3757
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16x16"
         SmallIcons      =   "img16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Èìÿ âèðóñà"
            Object.Width           =   3140
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ôàéë"
            Object.Width           =   6880
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Ñòàòóñ"
            Object.Width           =   2893
         EndProperty
      End
      Begin VB.TextBox txtPath 
         Height          =   375
         Left            =   6075
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   5715
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   4590
         TabIndex        =   1
         Text            =   "512"
         Top             =   3150
         Width           =   1275
      End
      Begin MSComctlLib.ImageList imgMain 
         Left            =   6480
         Top             =   4680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   13
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":11A018
               Key             =   "mycomputer"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":11C7CA
               Key             =   "genericfile"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":11D294
               Key             =   "removabledrive"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":11FA46
               Key             =   "mydocs"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":1221F8
               Key             =   "cdrom"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":1249AA
               Key             =   "closedfolder"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":12715C
               Key             =   "desktop"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":12990E
               Key             =   "openfolder"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":12C0C0
               Key             =   "unknowndrive"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":12E872
               Key             =   "floppydrive"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":131024
               Key             =   "harddrive"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":1337D6
               Key             =   "netdrive"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain2.frx":135F88
               Key             =   "SelFol"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvFolders 
         Height          =   2535
         Left            =   0
         TabIndex        =   58
         Top             =   90
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   4471
         _Version        =   393217
         Style           =   7
         ImageList       =   "imgMain"
         Appearance      =   1
      End
      Begin BGAntiVirus.xpcmdbutton cmdScan 
         Height          =   345
         Left            =   1860
         TabIndex        =   77
         ToolTipText     =   "Çàïóñòèòü ñêàíèðîâàíèå"
         Top             =   2700
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         Caption         =   "&Ñòàðò"
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
      Begin BGAntiVirus.xpcmdbutton cmdStop 
         Height          =   345
         Left            =   3750
         TabIndex        =   78
         ToolTipText     =   "Îñòàíîâèòü ñêàíèðîâàíèå"
         Top             =   2700
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         Caption         =   "Ñ&òîï"
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
      Begin VB.Image Image14 
         Height          =   1335
         Left            =   0
         Picture         =   "frmMain2.frx":1360E2
         Top             =   0
         Width           =   3105
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000040C0&
         Height          =   615
         Left            =   6795
         TabIndex        =   8
         Top             =   6210
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label lblCount 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   3900
         TabIndex        =   7
         Top             =   5820
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblFound 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   255
         Left            =   2445
         TabIndex        =   6
         Top             =   5730
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Îãðàíè÷èòü ðàçìåð ïðîâåðÿåìûõ ôàéëîâ :                         ( KB)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   3180
         Width           =   6255
      End
      Begin VB.Label lblCleaned 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   255
         Left            =   2430
         TabIndex        =   3
         Top             =   6075
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Found:                of"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   1620
         TabIndex        =   9
         Top             =   5820
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Virus Cleaned:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   1620
         TabIndex        =   4
         Top             =   6180
         Visible         =   0   'False
         Width           =   3615
      End
   End
   Begin VB.Frame frameSetting 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   5955
      Left            =   2130
      TabIndex        =   27
      Top             =   1800
      Width           =   7290
      Begin VB.TextBox txtRefreshRate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2940
         TabIndex        =   56
         Text            =   "10"
         Top             =   2880
         Width           =   600
      End
      Begin VB.TextBox txtScanRegExt 
         Height          =   405
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   50
         Top             =   5250
         Width           =   4845
      End
      Begin BGAntiVirus.xpcmdbutton cmdConfigAppBlock 
         Height          =   375
         Left            =   3150
         TabIndex        =   92
         ToolTipText     =   "Ðåäàêòèðîâàòü íàñòðîéêè çàáëîêèðîâàííûõ ïðèëîæåíèé"
         Top             =   3270
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         Caption         =   "App Blocker"
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
      Begin BGAntiVirus.xpcmdbutton cmdRestoreDefault 
         Height          =   375
         Left            =   3120
         TabIndex        =   93
         ToolTipText     =   "Èçìåíèòü íàñòðîéêè ïî óìîë÷àíèþ(ðåêîìåíäóåòñÿ)"
         Top             =   4470
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   661
         Caption         =   "Ïî óìîë÷àíèþ"
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
      Begin BGAntiVirus.xpradiobutton optScanAll 
         Height          =   375
         Left            =   390
         TabIndex        =   98
         Top             =   4470
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ñêàíèðîâàòü âåñü ðååñòð"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpradiobutton optScanExt 
         Height          =   375
         Left            =   390
         TabIndex        =   99
         Top             =   4800
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Ñêàíèðîâàòü óêàçàííûå ðàñøèðåíèÿ (Æåëàòåëüíî)"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkAutoStart 
         Height          =   315
         Left            =   390
         TabIndex        =   100
         Top             =   540
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "Àâòîñòàðò"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkBlockRisk 
         Height          =   315
         Left            =   390
         TabIndex        =   101
         Top             =   2010
         Width           =   6285
         _ExtentX        =   11086
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
         Caption         =   "Àâòîìàòè÷åñêè áëîêèðîâàòü ïîäîçðèòåëüíûå (Óêàæèòå ÷àñòîòó îáíîâëåíèÿ)"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkLog 
         Height          =   315
         Left            =   390
         TabIndex        =   102
         Top             =   1560
         Width           =   4575
         _ExtentX        =   8070
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
         Caption         =   "Îãðàíè÷èòü îò÷åò (ìàêñèìóì 3 Má)"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox mnuQarantine 
         Height          =   315
         Left            =   390
         TabIndex        =   103
         Top             =   1230
         Width           =   4245
         _ExtentX        =   7488
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
         Caption         =   "Ïåðåìåñòèòü èíôèöèðîâàííûå â êàðàíòèí"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkStartMin 
         Height          =   315
         Left            =   390
         TabIndex        =   104
         Top             =   870
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "Ñòàðòîâàòü â ôîíå"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkExeBlock 
         Height          =   315
         Left            =   390
         TabIndex        =   105
         Top             =   3270
         Width           =   4425
         _ExtentX        =   7805
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
         Caption         =   "Âêëþ÷èòü áëîêåð ïðèëîæåíèé"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkAutoRefeshProList 
         Height          =   315
         Left            =   390
         TabIndex        =   106
         Top             =   2430
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "Àâòîìàòè÷åñêè îáíîâëÿòü ïðîöåññû"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkControlAll 
         Height          =   315
         Left            =   960
         TabIndex        =   107
         Top             =   3630
         Width           =   2925
         _ExtentX        =   5159
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
         Caption         =   "Êîíòðîëèðîâàòü ÂÑÅ ïðèëîæåíèÿ"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcheckbox chkAutoScan 
         Height          =   315
         Left            =   4800
         TabIndex        =   108
         Top             =   630
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "&Enable Auto Scan (Under Construction) /  Invisible"
         BackColor       =   14737632
         ForeColor       =   16711680
      End
      Begin BGAntiVirus.xpcmdbutton cmdSave 
         Height          =   375
         Left            =   5670
         TabIndex        =   109
         ToolTipText     =   "Ñîõðàíèòü íàñòðîéêè ïðîãðàììû"
         Top             =   5250
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         Caption         =   "Ñîõðàíèòü"
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
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "×àñòîòà îáíîâëåíèÿ                    ñåêóíä"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   720
         TabIndex        =   57
         Top             =   2880
         Width           =   3795
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Íàñòðîéêè ïðîâåðêè ðååñòðà"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   420
         TabIndex        =   53
         Top             =   4110
         Width           =   6435
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Íàñòðîéêè ïðîãðàììû"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   450
         TabIndex        =   28
         Top             =   60
         Width           =   4755
      End
   End
   Begin VB.Frame frameVirusSig 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   5955
      Left            =   2115
      TabIndex        =   10
      Top             =   1800
      Width           =   7305
      Begin VB.Frame frmameRegUser 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Äëÿ ñïåöèàëèñòîâ(äîáàâèòü â áàçó âèðóñ)"
         Enabled         =   0   'False
         Height          =   855
         Left            =   240
         TabIndex        =   66
         ToolTipText     =   "Äëÿ ñïåöîâ"
         Top             =   300
         Width           =   6795
         Begin VB.TextBox txtCRC32 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   375
            Left            =   210
            TabIndex        =   68
            Top             =   300
            Width           =   1695
         End
         Begin VB.TextBox txtVirusName 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   3390
            TabIndex        =   67
            Top             =   300
            Width           =   1845
         End
         Begin BGAntiVirus.xpcmdbutton cmdCheckCRC 
            Height          =   345
            Left            =   2010
            TabIndex        =   96
            Top             =   300
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
            Caption         =   "&Âûáîð"
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
         Begin BGAntiVirus.xpcmdbutton cmdAddToDef 
            Height          =   345
            Left            =   5430
            TabIndex        =   97
            Top             =   300
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   609
            Caption         =   "&Äîáàâèòü"
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
      Begin MSComctlLib.ListView lvVirusList 
         Height          =   2445
         Left            =   300
         TabIndex        =   24
         Top             =   1470
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   4313
         View            =   2
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img16x16"
         SmallIcons      =   "img16x16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "System"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3120
         Top             =   5130
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin BGAntiVirus.xpcmdbutton cmdOfflineUp 
         Height          =   345
         Left            =   4950
         TabIndex        =   94
         Top             =   4800
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         Enabled         =   0   'False
         Caption         =   "OFF-Line îáíîâèòü"
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
      Begin BGAntiVirus.xpcmdbutton cmdOnlineUp 
         Height          =   345
         Left            =   4950
         TabIndex        =   95
         Top             =   5280
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         Caption         =   "On-Line îáíîâèòü"
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
      Begin VB.Label lblVirusCount 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2190
         TabIndex        =   15
         Top             =   4860
         Width           =   1455
      End
      Begin VB.Label lblLastUpdate 
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2190
         TabIndex        =   14
         Top             =   5235
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Äàòà îáíîâëåíèÿ "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   495
         Left            =   180
         TabIndex        =   13
         Top             =   5220
         Width           =   1995
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Â áàçå :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   150
         TabIndex        =   12
         Top             =   4830
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Áàçà"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   3255
      End
   End
   Begin VB.Frame frameAbout 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   5925
      Left            =   2130
      TabIndex        =   26
      Top             =   1830
      Width           =   7290
      Begin VB.Image Image12 
         Height          =   1365
         Left            =   3990
         Picture         =   "frmMain2.frx":143A14
         Top             =   2820
         Width           =   3105
      End
      Begin VB.Image Image11 
         Height          =   1365
         Left            =   3930
         Picture         =   "frmMain2.frx":151826
         Top             =   870
         Width           =   3105
      End
      Begin VB.Image Image10 
         Height          =   1335
         Left            =   330
         Picture         =   "frmMain2.frx":15F638
         ToolTipText     =   "mrbelyash@rambler.ru"
         Top             =   2730
         Width           =   3105
      End
      Begin VB.Image Image9 
         Height          =   1335
         Left            =   240
         Picture         =   "frmMain2.frx":16CF6A
         ToolTipText     =   "www.mrbelyash.narod.ru"
         Top             =   915
         Width           =   3105
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ÖÅÍÒÐ ÑÏÐÀÂÊÈ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   2190
         TabIndex        =   59
         Top             =   135
         Width           =   3435
      End
   End
   Begin VB.Frame frameTool 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   5955
      Left            =   2160
      TabIndex        =   16
      Top             =   1800
      Width           =   7230
      Begin BGAntiVirus.xpcmdbutton cmdScanInvalidReg 
         Height          =   375
         Left            =   180
         TabIndex        =   81
         Top             =   330
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         Caption         =   "&Èñïðàâèòü ðååñòð"
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
      Begin BGAntiVirus.xpcmdbutton cmdTool_ProcessMan 
         Height          =   375
         Left            =   2040
         TabIndex        =   82
         Top             =   330
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         Caption         =   "&Ïðîöåññû"
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
      Begin BGAntiVirus.xpcmdbutton cmdTool_Startup 
         Height          =   375
         Left            =   3840
         TabIndex        =   83
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   661
         Caption         =   "À&âòîçàãðóçêà"
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
      Begin BGAntiVirus.xpcmdbutton cmdEnableReg 
         Height          =   375
         Left            =   5580
         TabIndex        =   84
         Top             =   330
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Caption         =   "&Ðååñòð"
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
      Begin VB.Frame frameTool_EnableReg 
         BackColor       =   &H00E0E0E0&
         Height          =   5085
         Left            =   150
         TabIndex        =   18
         Top             =   810
         Width           =   6870
         Begin VB.CheckBox chkLMRegTool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îòêëþ÷èòü äîñòóï ê ðååñòðó"
            Height          =   375
            Left            =   420
            TabIndex        =   42
            Top             =   2955
            Width           =   2925
         End
         Begin VB.CheckBox chkLMNoSR 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàïðåòèòü âîññòàíîâëåíèå ñèñòåìû"
            Height          =   375
            Left            =   420
            TabIndex        =   47
            Top             =   3795
            Width           =   5175
         End
         Begin VB.CheckBox chkLmLimitSR 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îãðàíè÷èò ðàçìåð àðõèâà âîññòàíîâëåíèÿ ÎÑ"
            Height          =   375
            Left            =   420
            TabIndex        =   46
            Top             =   4335
            Width           =   5535
         End
         Begin VB.CheckBox chkLMNoMSI 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îòêëþ÷èòü èíñòàëÿøêó"
            Height          =   375
            Left            =   420
            TabIndex        =   45
            Top             =   4650
            Width           =   3615
         End
         Begin VB.CheckBox chkLMNoSRConfig 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îòêëþ÷èòü äîñòóï ê íàñòðîéêàì âîññòàíîâëåíèÿ ÎÑ"
            Height          =   375
            Left            =   420
            TabIndex        =   44
            Top             =   4080
            Width           =   5895
         End
         Begin VB.CheckBox chkLMNoTaskmgr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü Task Manager"
            Height          =   375
            Left            =   420
            TabIndex        =   43
            Top             =   3225
            Width           =   4575
         End
         Begin VB.CheckBox chkLMNoFolderOption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Óáðàòü ""Ñâîéñòâà ïàïêè"""
            Height          =   375
            Left            =   420
            TabIndex        =   41
            Top             =   3510
            Width           =   3375
         End
         Begin VB.CheckBox chkCUNoFolderOption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Óáðàòü ""Ñâîéñòâà ïàïêè"""
            Height          =   375
            Left            =   420
            TabIndex        =   40
            Top             =   930
            Width           =   2925
         End
         Begin VB.CheckBox chkCUNoCmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü êîíñîëü"
            Height          =   375
            Left            =   420
            TabIndex        =   39
            Top             =   2070
            Width           =   3795
         End
         Begin VB.CheckBox chkCURegTool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü ðååñòð"
            Height          =   375
            Left            =   420
            TabIndex        =   38
            Top             =   360
            Width           =   2415
         End
         Begin VB.CheckBox chkCUNoChangePwd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàïðåòèòü èçìåíÿòü ïðîëü ïîëüçîâàòåëÿ"
            Height          =   375
            Left            =   420
            TabIndex        =   37
            Top             =   2355
            Width           =   4575
         End
         Begin VB.CheckBox chkCUNoLock 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Óáðàòü ëîê êîìïà"
            Height          =   375
            Left            =   420
            TabIndex        =   36
            Top             =   1785
            Width           =   3945
         End
         Begin VB.CheckBox chkCUNoCLose 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü ""Âûõîä"""
            Height          =   375
            Left            =   420
            TabIndex        =   35
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkCUNoLogoff 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàïðåòèòü ""Ñìåíó ïîëüçîâàòåëÿ"""
            Height          =   375
            Left            =   420
            TabIndex        =   34
            Top             =   1485
            Width           =   3315
         End
         Begin VB.CheckBox chkCUTaskmgr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü Task Manager"
            Height          =   375
            Left            =   420
            TabIndex        =   33
            Top             =   645
            Width           =   3225
         End
         Begin BGAntiVirus.xpcmdbutton cmdRegis1Now 
            Height          =   375
            Left            =   4800
            TabIndex        =   89
            Top             =   1200
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   661
            Caption         =   "Ð&åäàêòîð"
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
         Begin BGAntiVirus.xpcmdbutton cmdClearAutorun 
            Height          =   375
            Left            =   4800
            TabIndex        =   90
            Top             =   750
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   661
            Caption         =   "&Óäàëèòü Autorun.inf"
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
         Begin BGAntiVirus.xpcmdbutton cmdCleanReg 
            Height          =   375
            Left            =   4800
            TabIndex        =   91
            Top             =   300
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   661
            Caption         =   "&Ñíÿòü âñå"
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
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Òåêóùàÿ ìàøèíà"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   435
            Left            =   420
            TabIndex        =   49
            Top             =   2700
            Width           =   4875
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Àêòèâíûé ïîëüçîâàòåëü"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   375
            Left            =   390
            TabIndex        =   48
            Top             =   150
            Width           =   3705
         End
      End
      Begin VB.Frame frameTool_Process 
         BackColor       =   &H00E0E0E0&
         Height          =   4830
         Left            =   150
         TabIndex        =   30
         Top             =   840
         Width           =   6840
         Begin MSComctlLib.ListView lvProcess 
            Height          =   1935
            Left            =   240
            TabIndex        =   31
            Top             =   840
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3413
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "img16x16"
            SmallIcons      =   "img16x16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Èìÿ ôàéëà"
               Object.Width           =   8820
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ID"
               Object.Width           =   3
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ProID(ToKill)"
               Object.Width           =   1588
            EndProperty
         End
         Begin MSComctlLib.ListView lvProcessDetail 
            Height          =   1515
            Left            =   240
            TabIndex        =   32
            Top             =   2850
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2672
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Window Caption"
               Object.Width           =   5539
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Window Class"
               Object.Width           =   6068
            EndProperty
         End
         Begin BGAntiVirus.xpcmdbutton cmdProcessRefresh 
            Height          =   375
            Left            =   240
            TabIndex        =   79
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "&Îáíîâèòü"
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
         Begin BGAntiVirus.xpcmdbutton cmdProcessEnd 
            Height          =   375
            Left            =   2070
            TabIndex        =   80
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            Caption         =   "&Çàâåðøèòü"
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
         Begin VB.Label lblTotalProcess 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   5865
            TabIndex        =   55
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Ïðîöåññîâ:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   375
            Left            =   4320
            TabIndex        =   54
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frameTool_ScanReg 
         BackColor       =   &H00E0E0E0&
         Height          =   4875
         Left            =   180
         TabIndex        =   17
         Top             =   840
         Width           =   6855
         Begin VB.TextBox txtCurKey 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   52
            Top             =   1080
            Width           =   6015
         End
         Begin MSComctlLib.ListView lvErrorRegKey 
            Height          =   2205
            Left            =   240
            TabIndex        =   23
            Top             =   2280
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   3889
            View            =   3
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "img16x16"
            SmallIcons      =   "img16x16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Íàéäåíî"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "RootKey"
               Object.Width           =   2469
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "SubKey"
               Object.Width           =   3881
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Çíà÷åíèå"
               Object.Width           =   2646
            EndProperty
         End
         Begin BGAntiVirus.xpcmdbutton cmdStartStop 
            Height          =   375
            Left            =   2400
            TabIndex        =   85
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   661
            Caption         =   "Ñòàðò"
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
         Begin BGAntiVirus.xpcmdbutton cmdDeleteInvalidKey 
            Height          =   375
            Left            =   3570
            TabIndex        =   86
            Top             =   180
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
            Caption         =   "&Óäàëèòü"
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
         Begin BGAntiVirus.xpcmdbutton cmdRegistryNow 
            Height          =   375
            Left            =   5310
            TabIndex        =   87
            Top             =   180
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   661
            Caption         =   "Ðåäàêòîð"
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
         Begin VB.Label lblScanRegError 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Height          =   255
            Left            =   1890
            TabIndex        =   22
            Top             =   1920
            Width           =   675
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Êîë-âî îøèáîê :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   1920
            Width           =   1395
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblScanRegStatus 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ñêàíèðóþ :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   360
            TabIndex        =   20
            Top             =   600
            Width           =   2685
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frameTool_Startup 
         BackColor       =   &H00E0E0E0&
         Height          =   4890
         Left            =   180
         TabIndex        =   19
         Top             =   825
         Width           =   6840
         Begin MSComctlLib.ListView lvStartUp 
            Height          =   3165
            Left            =   360
            TabIndex        =   29
            Top             =   1200
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5583
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Èìÿ"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Ôàéë"
               Object.Width           =   6174
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Òèï"
               Object.Width           =   2646
            EndProperty
         End
         Begin BGAntiVirus.xpcmdbutton cmdStartUp_Del 
            Height          =   375
            Left            =   420
            TabIndex        =   88
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   661
            Caption         =   "&Óäàëèòü"
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
   End
   Begin VB.Frame frameLicense 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6015
      Left            =   2160
      TabIndex        =   25
      Top             =   1860
      Width           =   7200
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ëèöåíçèÿ"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5505
         Left            =   270
         TabIndex        =   63
         Top             =   150
         Width           =   6375
      End
   End
   Begin VB.Frame franalize 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ñáîð ñâåäåíèé"
      Height          =   6900
      Left            =   2160
      TabIndex        =   61
      Top             =   1860
      Visible         =   0   'False
      Width           =   7260
      Begin RichTextLib.RichTextBox rtf1 
         Height          =   2610
         Left            =   120
         TabIndex        =   62
         Top             =   3300
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   4604
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmMain2.frx":17A89C
      End
      Begin BGAntiVirus.xpcmdbutton Command1 
         Height          =   345
         Left            =   240
         TabIndex        =   69
         Top             =   2850
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         Caption         =   "Ñîõðàíèòü"
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
      Begin BGAntiVirus.xpcheckbox Check1 
         Height          =   345
         Left            =   300
         TabIndex        =   70
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
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
         Caption         =   "Ïåðåìåííûå"
         BackColor       =   14737632
      End
      Begin BGAntiVirus.xpcheckbox Check2 
         Height          =   315
         Left            =   300
         TabIndex        =   71
         Top             =   570
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Ïðîöåññû"
         BackColor       =   14737632
      End
      Begin BGAntiVirus.xpcheckbox Check3 
         Height          =   315
         Left            =   300
         TabIndex        =   72
         Top             =   870
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Êîëè÷åñòâî ïàìÿòè"
         BackColor       =   14737632
      End
      Begin BGAntiVirus.xpcheckbox Check4 
         Height          =   315
         Left            =   300
         TabIndex        =   73
         Top             =   1200
         Width           =   3765
         _ExtentX        =   6641
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
         Caption         =   "Àêòèâíûå îêíà ïðèëîæåíèé"
         BackColor       =   14737632
      End
      Begin BGAntiVirus.xpcheckbox Check5 
         Height          =   315
         Left            =   300
         TabIndex        =   74
         Top             =   1530
         Width           =   5775
         _ExtentX        =   10186
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
         Caption         =   "Ýêñïîðò ðååñòðà(íå ðåêêîìåíäóåòñÿ-âîçìæíà óòå÷êà äàííûõ)"
         BackColor       =   14737632
      End
      Begin BGAntiVirus.xpcheckbox Check6 
         Height          =   315
         Left            =   300
         TabIndex        =   75
         Top             =   1830
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Àâòîçàïóñê"
         BackColor       =   14737632
      End
      Begin BGAntiVirus.xpcheckbox Check7 
         Height          =   315
         Left            =   300
         TabIndex        =   76
         Top             =   2190
         Width           =   2895
         _ExtentX        =   5106
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
         Caption         =   "Íàñòðîéêè ïðîãðàììû"
         BackColor       =   14737632
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   90
         X2              =   6930
         Y1              =   2700
         Y2              =   2700
      End
   End
   Begin VB.Image imgSetting 
      Height          =   810
      Left            =   420
      Picture         =   "frmMain2.frx":17A920
      ToolTipText     =   "Íàñòðîéêè ïðîãðàììû"
      Top             =   2610
      Width           =   1665
   End
   Begin VB.Image Image3 
      Height          =   780
      Left            =   420
      Picture         =   "frmMain2.frx":17B140
      Top             =   2610
      Width           =   1665
   End
   Begin VB.Image imgVirusDef 
      Height          =   810
      Left            =   420
      MousePointer    =   99  'Custom
      Picture         =   "frmMain2.frx":17B944
      ToolTipText     =   "Äîáàâëåíèå íîâûõ âðåäîíîñîâ"
      Top             =   3390
      Width           =   1680
   End
   Begin VB.Image Image2 
      Height          =   780
      Left            =   420
      Picture         =   "frmMain2.frx":17C122
      Top             =   3390
      Width           =   1695
   End
   Begin VB.Image imgTools 
      Height          =   795
      Left            =   420
      MousePointer    =   99  'Custom
      Picture         =   "frmMain2.frx":17C8ED
      ToolTipText     =   "Ðàçëè÷íûå èíñòðóìåíòû(èñïîëüçîâàòü ïîä àäìèíîì)"
      Top             =   4170
      Width           =   1710
   End
   Begin VB.Image Image4 
      Height          =   795
      Left            =   420
      Picture         =   "frmMain2.frx":17D067
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image imgLicense 
      Height          =   780
      Left            =   420
      MousePointer    =   99  'Custom
      Picture         =   "frmMain2.frx":17D7C1
      ToolTipText     =   "Âñÿêî-ðàçíî....."
      Top             =   4950
      Width           =   1665
   End
   Begin VB.Image Image5 
      Height          =   765
      Left            =   420
      Picture         =   "frmMain2.frx":17DFC2
      Top             =   4980
      Width           =   1695
   End
   Begin VB.Image imgAbout 
      Height          =   780
      Left            =   420
      MousePointer    =   99  'Custom
      Picture         =   "frmMain2.frx":17E7B0
      ToolTipText     =   "Öåíòð ïîìîùè"
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   780
      Left            =   420
      Picture         =   "frmMain2.frx":17EE4A
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Image imgScan 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   405
      Picture         =   "frmMain2.frx":17F4F6
      Stretch         =   -1  'True
      ToolTipText     =   "Çàïóñòèòü ñêàíèðîâàíèå"
      Top             =   1740
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   420
      Picture         =   "frmMain2.frx":17FC41
      Stretch         =   -1  'True
      Top             =   1740
      Width           =   1665
   End
   Begin VB.Image Image7 
      Height          =   795
      Left            =   420
      MousePointer    =   99  'Custom
      Picture         =   "frmMain2.frx":1803A0
      Stretch         =   -1  'True
      ToolTipText     =   "Ñáîð èíôîðìàöèè äëÿ ïîèñêà îøèáîê..."
      Top             =   6510
      Width           =   1665
   End
   Begin VB.Image Image8 
      Height          =   825
      Left            =   390
      Picture         =   "frmMain2.frx":192C76
      Top             =   6540
      Width           =   1725
   End
   Begin VB.Menu mnuop 
      Caption         =   "Op"
      Visible         =   0   'False
      Begin VB.Menu mnuOpSelAll 
         Caption         =   "Âûáðàòü âñ¸"
      End
      Begin VB.Menu mnuopUnSelAll 
         Caption         =   "Ñíÿòü ñî âñåõ"
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "tray"
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "Ïîêàçàòü"
      End
      Begin VB.Menu mnuT55 
         Caption         =   "-"
      End
      Begin VB.Menu s12 
         Caption         =   "Ñàéò ïðîãðàììû"
      End
      Begin VB.Menu mnuAbow 
         Caption         =   "Ñïðàâêà"
      End
      Begin VB.Menu mnuenblsc 
         Caption         =   "Ñêàíèðîâàòü Ìîé êîìïüþòåð"
      End
      Begin VB.Menu mnuT1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Âûãðóçèòü"
      End
   End
   Begin VB.Menu mnuProMenu 
      Caption         =   "pro_menu"
      Visible         =   0   'False
      Begin VB.Menu mnuProMenu_Ban 
         Caption         =   "&Ýòîò ïðîöåññ íå áåçîïàñåí"
      End
      Begin VB.Menu mnuProMenu_Safe 
         Caption         =   "&Áåçîïàñíûé ïðîöåññ"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private programs As New ProcessList
Dim WithEvents cTray As TrayIconAndBalloon
Attribute cTray.VB_VarHelpID = -1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private fso As New FileSystemObject
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
  Private m_lngRetVal As Long
Private Type SYSTEM_INFO
dwOemID As Long
dwPageSize As Long
lpMinimumApplicationAddress As Long
lpMaximumApplicationAddress As Long
dwActiveProcessorMask As Long
dwNumberOfProcessors As Long
dwProcessorType As Long
dwAllocationGranularity As Long
dwReserved As Long
End Type
Private m_typSystemInfo As SYSTEM_INFO
Const TH32CS_SNAPPROCESS As Long = 2&
Const MAX_PATH As Integer = 260
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * MAX_PATH
End Type
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hsnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hsnapshot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
'scan registry variable with EVENTS
Dim WithEvents cReg As cRegSearch
Attribute cReg.VB_VarHelpID = -1
Dim dt As String


Private Type STARTUPINFO
cb As Long
lpReserved As String
lpDesktop As String
lpTitle As String
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Long
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type
Private Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessID As Long
dwThreadID As Long
End Type
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Public Event ErrorDownload(FromPathName As String, ToPathName As String)
Public Event DownloadComplete(FromPathName As String, ToPathName As String)
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
 (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Public StopMycompScan As Boolean


Public Function DownloadFile(FromPathName As String, ToPathName As String)
If URLDownloadToFile(0, FromPathName, ToPathName, 0, 0) = 0 Then
DownloadFile = True
RaiseEvent DownloadComplete(FromPathName, ToPathName)
Else
DownloadFile = False
RaiseEvent ErrorDownload(FromPathName, ToPathName)
End If
End Function
Public Function NumberOfProcessors() As Long
GetSystemInfo m_typSystemInfo
NumberOfProcessors = m_typSystemInfo.dwNumberOfProcessors
End Function



Private Sub cmdClearAutorun_Click()
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    On Error Resume Next    'in case not found, and on cd
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
        Kill drv.DriveLetter & ":\autorun.inf"
    Next
    Call ShowTrayMessage("Óäàëåíèå Autorun.inf", "Autorun.inf íà âñåõ äèñêàõ óäàëåí.")
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
End Sub
Private Sub Command2_Click()
     'ñêàíèðîâàòü ïðîöåññû è èñêàòü çàðàæåííûå
   Me.lvProcess.ListItems.Clear
    Me.lvProcessDetail.ListItems.Clear
    'refresh data
    Call GetProcessList(Me.lvProcess)
        Dim Pro_ID As Long
'áóäåì ìîíèòîðèòü ôàéëû
Dim i As Integer
Dim f As Integer

For i = 1 To frmMain.lvProcess.ListItems.Count - 1
Debug.Print frmMain.lvProcess.ListItems.Item(i).SubItems(1)
    If frmMain.lvProcess.ListItems.Item(i).SubItems(1) <> "System Process" Then
         MsgBox frmMain.lvProcess.ListItems.Item(i).SubItems(1)
        Pro_ID = frmMain.lvProcess.ListItems.Item(i).SubItems(2)
        Call ScanFileProc("c:\", frmMain.lvProcess.ListItems.Item(i).SubItems(1), Pro_ID)
      End If
Next
'MsgBox "çàêîí÷èë ñêàíèòü ïàìÿòü"
Exit Sub
'ìîíèòîðèíã
 
    'Call CheckProcess
End Sub

Private Sub cmdOfflineUp_Click()
'Îáúÿâëÿåì ñòðîêîâóþ ïåðåìåííóþ äëÿ íàçíà÷åíèÿ òèïîâ ôàéëîâ
Dim strFileType As String
'Åñëè âîçíèêíåò îøèáêà, ò.å ïîëüçîâàòåëü íàæåë íà êëàâèøó Cancel,
'îòïðàâèòüñÿ ê îáðàáîò÷èêó îøèáêè - ErrorHandler
On Error GoTo ErrorHandler
'Îáåñïå÷èâàåì ãåíåðàöèþ îùèáêè
CommonDialog1.CancelError = True

'Èíèöèàëèçèðóåì ñòðîêîâóþ ïåðåìåííóþ strFileType
strFileType = strFileType & " Text Files (*.sig)|*.sig|"
'Ïðèñâàèâàåì åå ñâîéñòâó Filter
CommonDialog1.Filter = strFileType
'Óñòàíàâëèâàåì íåîáõîäèìûé èíäåêñ
CommonDialog1.FilterIndex = 2
'Ïðèñâàèâàåì íà÷àëüíóþ äèðåêòîðèþ ñâîñòâó InitDir
CommonDialog1.InitDir = App.Path
'Îáåñïå÷èâàåì çàùèòó îò íåïðàâèëüíîãî ââåäåííîãî ôàéëà èëè äåðèêòîðèè, à òàê æå ñêðûâàåì ôëàæåê Read Only

CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
'Âûçûâàåì äèàëîã Open
CommonDialog1.Action = 1 'Èëè æå CommonDialog1.ShowOpen
'***********
Call obnowlLocal(CommonDialog1.FileName)
MsgBox "Äëÿ ïîäãðóçêè íîâûõ áàç ïðîãðàììà äîëæíà áûòü ïåðåçàïóùåíà", vbCritical, pr
End
'Çäåñü ðàñïîëîãàåòñÿ Âàø êîä.(íå çàáóäòå, ÷òî ïóòü ê âûáðàííîìó ôàéëó Âû ñ÷èòûâàåòå èç ñâîéñòâà FileName)
'**********
Exit Sub
'Îáðàáîòêà ïåðåõâàòûâàåìîé îùèáêè
ErrorHandler:
If Err.Number = 32755 Then
Exit Sub
End If

End Sub
Sub obnowlLocal(meobn As String)

Dim a As String
Dim f As Integer
Dim mePt As String
f = InStrRev(meobn, "\", -1, vbTextCompare)
If f <> 0 Then
mePt = Right(meobn, Len(meobn) - f)
    If Dir$(App.Path + "\" + mePt) = "" Then
    
        FileCopy meobn, App.Path + "\" + mePt
        MsgBox "Îáíîâëåíèå çàâåðøåíî", vbInformation
        Exit Sub
    Else
        MsgBox "Äàííàÿ áàçà óæå åñòü â êàòàëîãå", vbCritical
        Exit Sub
    End If


Else
    Exit Sub
End If
End Sub
Private Sub cmdOnlineUp_Click()
'çàïóñêàåì online on-line îáíîâëÿëêó
   ModuleS.GoObnowl
End Sub

Sub get_IEtemp()
Dim a As String
    a = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Cache")
    Me.txtPath.Text = a 'óêàæó ïóòü
    cmdScan_CLBZ 'ñêàíèðóþ ïóòü
End Sub
Sub Check_critical()
'ïðîâåðèòü êðèòè÷åñêèå îáüåêòû
'DoEvents
StopMycompScan = False
'GoTo 10
Call rescan_proc 'ñêàíèì àêòèâíûå ïðîöåññû
 If StopMycompScan = True Then GoTo 803
    Dim FWinPath As String
    FWinPath = Environ("windir")
    'MsgBox FWinPath
    Me.txtPath.Text = FWinPath 'óêàæó ïóòü
    cmdScan_CLBZ 'ñêàíèðóþ ïóòü
 If StopMycompScan = True Then GoTo 803
10:
    Dim AUsr1 As String
    AUsr1 = Environ("ALLUSERSPROFILE") + "\"
    Me.txtPath.Text = AUsr1 'óêàæó ïóòü
    cmdScan_CLBZ 'ñêàíèðóþ ïóòü
 If StopMycompScan = True Then GoTo 803
    
    'MsgBox "Done check"
    Dim hmp As String
    Dim hmD As String
    hmp = Environ$("HOMEPATH")
    hmD = Environ$("HOMEDRIVE") + hmp
    If hmD <> "" Then
         Me.txtPath.Text = hmD 'óêàæó ïóòü
        cmdScan_CLBZ 'ñêàíèðóþ ïóòü
    End If
 If StopMycompScan = True Then GoTo 803
Call get_IEtemp 'ïðîâåðÿåì êåø ýêñïëîðåðà
803:
  StopMycompScan = True
Call baloon("Ñêàíèðîâàíèå çàâåðøåíî")
Exit Sub



   
'GetAllRun'ïîêà ñòàðòàïû íå áóäó îáðàáàòûâàòü
End Sub
'SysTray Events
Private Sub cSysTray1_MouseDblClick(Button As Integer, Id As Long)
    Me.WindowState = vbNormal
    Show
    cTray.Delete
    'Me.SetFocus
End Sub

Private Sub cSysTray1_MouseDown(Button As Integer, Id As Long)
'    Dim MSG As Long
'    'unknown statement
'    MSG = X / Screen.TwipsPerPixelX
'    Select Case MSG
'        Case WM_RBUTTONDOWN     'right click
'            PopupMenu mnuTray
'        Case WM_LBUTTONDBLCLK   'double click
'            mnuOpen_Click
'    End Select
'   MsgBox Button
    If Button = 2 Then   'right click
        PopupMenu mnuTray
    End If
End Sub



Private Sub cmdRegis1Now_Click()
'çàïóñêàþ âíóòðåíèé ðåäàêòîð ðååñòðà
modFilePath.RegistryWnutr
End Sub

Private Sub cmdRegistryNow_Click()
'çàïóñêàåì âíóòðåííèé ðåäàêòîð ðååñòðà
modFilePath.RegistryWnutr
End Sub

Private Sub cmdSave_Click()
'ñîõðàíèòü íàñòðîéêè
If MsgBox("Ñîõðàíèòü íàñòðîéêè ?", vbYesNo + vbQuestion, pr) = vbNo Then
    Exit Sub
End If
If Me.chkAutoStart.Value = Checked Then
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.exeName & ".exe")
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStart", Val(1)
Else
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStart", Val(0)
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", "BGAntiVirus"
        Call cmdTool_Startup_Click
End If

If Me.chkStartMin.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartMin", Val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartMin", Val(0)
End If

If Me.mnuQarantine.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Quarantine", Val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Quarantine", Val(0)
End If

If Me.chkLog.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log", Val(1)
Else
  CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log", Val(0)
End If

If Me.chkBlockRisk.Value = Checked Then
 CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk", Val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk", Val(0)
End If

If chkAutoRefeshProList.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh", Val(1)
     Me.tmrAutoRefresh.Enabled = True
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh", Val(0)
        Me.tmrAutoRefresh.Enabled = False
End If


    'edit in EXEFile/Shell/Open/Command
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    If Me.chkExeBlock.Value = Checked Then
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock", Val(1)
         sh.regwrite "HKCR\exefile\shell\open\command\original", Chr$(34) + "%1" + Chr$(34) + " %*"
        sh.regwrite "HKCR\exefile\shell\open\command\", App.Path & "\AppBlock.EXE %1 %*"
        'chkControlAll.Enabled = True
    Else
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock", Val(0)
         sh.regwrite "HKCR\exefile\shell\open\command\", Chr$(34) + "%1" + Chr$(34) + " %*"
        'chkControlAll.Enabled = False
    End If
    If Me.chkControlAll.Value = Checked Then
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll", Val(1)
    Else
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll", Val(0)
    End If
End Sub

Private Sub Form_Activate()
If Me.WindowState <> vbMinimized Then
Me.Top = 0
End If
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    ShowTopicID 1, 33
End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'ïåðåäàåì äàííûå â îáúåêò
      cTray.CallEvent x, Y

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim strQuestion As String
Dim intAnswer As Integer
Dim aryMode As Variant

aryMode = Array(pr, "vbFormCode", "vbAppWindows", pr, "vbFormMDIForm")
strQuestion = "Âû äåéñòâèòåëüíî õîòèòå âûéòè èç ïðîãðàììû ?"
intAnswer = MsgBox(strQuestion, vbQuestion + vbYesNo, aryMode(UnloadMode))

If intAnswer = vbNo Then Cancel = -1
End Sub

'========================================'
' FORM EVENTS                            '
'========================================'
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Hide
        'õåíäë îêíà
      cTray.hwnd = hwnd
   'èêîíêà, ÷òî áóäåò îòîáðàæåíà â òðåå
      cTray.Icon = Icon
   'òóëòèïñ (âñïëûâàþùàÿ ïîäñêàçêà)
      cTray.ToolTipText = "Belyash Antitrojan 2008b"
      
   'ñîçäàåì èêîíêó
      cTray.Add
   
    End If
    If Me.WindowState <> vbMinimized Then
        If Me.Width <> 9540 Then
            Me.Width = 9540
            
        End If
        If Me.Height <> 8550 Then
            Me.Height = 8550
              
        End If
     ' frmMain.StartUpPosition = 2
    Dim LeftPos As Integer
Dim TopPos As Integer

LeftPos = Int((Screen.Width - Me.Width) / 2)
TopPos = 0
Me.Top = TopPos
Me.Left = LeftPos
procCleanSpl = 0
'procCountSpl = 0
procFoundSpl = 0
frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + procFoundSpl
frmMain.SB.Panels(3).Text = frmMain.lblFound.Caption
   frmMain.lblCleaned.Caption = Int(frmMain.lblCleaned.Caption) + procCleanSpl
   frmMain.SB.Panels(4).Text = frmMain.lblCleaned.Caption
frmMain.lblCount.Caption = Int(frmMain.lblCount.Caption) + procCountSpl
frmMain.SB.Panels(2).Text = frmMain.lblCount.Caption
 frmMain.SB.Panels(6).Text = frmMain.Data1.Recordset.RecordCount
SB.Refresh

    End If
End Sub

Private Sub Form_Load()

'SetTopMostWindow Me.hwnd, True
'Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    'check if application is already loaded
'  Me.Show
          'ñîçäàåì èíñòàíñ îáúåêòà

     SB.Panels(2).Text = CStr(procCountSpl) 'ïðîâåðåíî ôàéëîëâ
     SB.Panels(3).Text = CStr(procFoundSpl) 'íàéäåíî âèðóñîâ
     SB.Panels(4).Text = CStr(procCleanSpl)  'óäàëåíî
     frmMain.lblCount.Caption = procCountSpl
     frmMain.lblFound.Caption = procFoundSpl
     frmMain.lblCleaned.Caption = procCleanSpl
      ' SB.Panels(6).Text = virCount
       SB.Refresh
    'loading scan page
    '-----------------
    Dim ADrive As Drive
  Dim Icon As String
  Dim Name As String
  Dim AFolder As Folder
  Dim DriveFolders As Folder

  ' Show the drives w/correct icons and names
  For Each ADrive In fso.Drives
    If ADrive.DriveType = CDRom Then
      Icon = "cdrom"
      If ADrive.IsReady Then Name = ADrive.DriveLetter Else Name = "CD-ROM Drive"
    ElseIf ADrive.DriveType = Fixed Then
      Icon = "harddrive"
      If ADrive.IsReady Then Name = ADrive.DriveLetter Else Name = "Hard Drive"
    ElseIf ADrive.DriveType = Remote Then
      Icon = "netdrive"
      If ADrive.IsReady Then Name = ADrive.DriveLetter Else Name = "Network Drive"
    ElseIf ADrive.DriveType = Removable Then
      If ADrive.DriveLetter = "A" Or ADrive.DriveLetter = "B" Then Icon = "floppydrive" Else Icon = "removabledrive"
      If ADrive.IsReady Then
        Name = ADrive.DriveLetter
      Else
        If ADrive.DriveLetter = "A" Or ADrive.DriveLetter = "B" Then Name = "Floppy Drive" Else Name = "Removable Drive"
      End If
    Else
      Icon = "unknowndrive"
      If ADrive.IsReady Then Name = ADrive.DriveLetter Else Name = "Unknown"
    End If
    
    'Add the drives node to the root tree
    'The key is the drives path
   ' tvFolders.Nodes.Add , 0, ADrive.Path, Name & " (" & UCase(ADrive.DriveLetter) & ":)", Icon
    
      tvFolders.Nodes.Add , 0, ADrive.Path, " " & UCase(ADrive.DriveLetter) & ":", Icon
    'If the drive is available grab the drives root directories
    'We do this before the user expands the drive the the plus-minus box shows up right.
    If ADrive.IsReady Then
      Set DriveFolders = fso.GetFolder(ADrive.RootFolder)
      For Each AFolder In DriveFolders.SubFolders
        'Add the folder to the tree, with the drive as it's parent
        'The key is the full path to the folder
        tvFolders.Nodes.Add ADrive.Path, 4, AFolder.Path, AFolder.Name, "closedfolder"
      Next
    End If
  Next
    
    
    
    'set windows to topmost
    'SetTopMostWindow Me.hwnd, True
    
    'frmSplash.lblLoading.Caption = "Loading Scanning Tool ..."
    'frmSplash.Refresh
    'set default to 512 kb
    FileSize = 524288
    Call imgScan_Click

    'loading setting page
    '-----------------------
    Call LoadSetting
    
    'loading virus sig. page
    '-----------------------
    Call RefreshDefList
    
    
    'loading tool page
    '-----------------------
    'scan reg content
    Set cReg = New cRegSearch
    Call cmdScanInvalidReg_Click
    
    'process content
    'Call GetProcessList(Me.lvProcess)
    'Me.lblTotalProcess.Caption = Me.lvProcess.ListItems.Count
    'Call CheckProcess
    
    'registry content
    Call LoadRegistry
    Set cTray = New TrayIconAndBalloon
strDBPath = App.Path + "\Avbase.bel"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    If Me.Tag = "" Then Exit Sub
'    Dim MSG As Long
'    'unknown statement
'    MSG = X / Screen.TwipsPerPixelX
'    Select Case MSG
'        Case WM_RBUTTONDOWN     'right click
'            PopupMenu mnuTray
'        Case WM_LBUTTONDBLCLK   'double click
'            mnuOpen_Click
'    End Select
End Sub







Private Sub Image10_Click()
Call ShellExecute(0, "Open", "mailto:" + "mrbelyash@rambler.ru" + "?Subject=" + "Ïðîáëåìû ïðè ðàáîòå ñ Belyash AntiTrojan", "", "", 1)
End Sub

Private Sub Image11_Click()
ShowTopicID 1, 33
End Sub

Private Sub Image12_Click()
cmdOnlineUp_Click
End Sub
Private Sub Image9_Click()
Call ShellExecute(0, "Open", "http://mrbelyash.narod.ru/antivirus/belyashAV.htm", "", "", 1)
End Sub
Private Sub mnuAbow_Click()
Me.WindowState = vbNormal
    Me.Show
     cTray.Delete
imgAbout_Click
End Sub

Private Sub mnuenblsc_Click()
If Me.mnuenblsc.Caption = "Ñêàíèðîâàòü Ìîé êîìïüþòåð" Then
    Me.mnuenblsc.Caption = "Îñòàíîâèòü ñêàíèðîâàíèå"
    Check_critical
Else
    Me.mnuenblsc.Caption = "Ñêàíèðîâàòü Ìîé êîìïüþòåð"
    cmdStop_Click
End If
End Sub
Private Sub mnuProMenu_Ban_Click()
    Dim f As Long
    f = FreeFile
    Open App.Path & "\AttPro.bin" For Append As f
    Print #f, vbCrLf & Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text
    Close f
    Call cmdProcessRefresh_Click
End Sub

Private Sub mnuProMenu_Safe_Click()
    Dim f As Long
    Dim strTemp As String, temp1 As String, temp2 As String
    f = FreeFile
    Open App.Path & "\AttPro.bin" For Binary As f
    strTemp = Input$(LOF(f), 1)
    Close f
    temp1 = Left$(strTemp, InStr(1, strTemp, Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text) - 2)
    temp2 = Right$(strTemp, Len(strTemp) - InStr(1, strTemp, Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text) - Len(Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).Text) - 1)
    strTemp = temp1 & temp2
    strTemp = Replace(strTemp, vbCrLf & vbCrLf, vbCrLf)
    Open App.Path & "\AttPro.bin" For Output As f
    Print #f, strTemp
    Close f
    Call cmdProcessRefresh_Click
End Sub

'systray popup
Private Sub mnuClose_Click()
    If blnScan = True Then
        If MsgBox("Ïðîãðàììà ñêàíèðóåò êîìïüþòåð.Âû äåéñòâèòåëüíî õîòèòå âûéòè ?", vbYesNoCancel, "Ïðåðâàòü ñêàíèðîâàíèå") = vbYes Then
            Exit Sub
        End If
    End If
    cReg.StopSearch
  cTray.Delete
    End
End Sub
Private Sub mnuOpen_Click()
    Me.WindowState = vbNormal
    Me.Show
     cTray.Delete
   End Sub

'========================================'
' SCAN                                   '
'========================================'

Sub rescan_proc()
     'ñêàíèðîâàòü ïðîöåññû è èñêàòü çàðàæåííûå
   Me.lvProcess.ListItems.Clear
    Me.lvProcessDetail.ListItems.Clear
    'refresh data
    Call GetProcessList(Me.lvProcess)
        Dim Pro_ID As Long
'áóäåì ìîíèòîðèòü ôàéëû
Dim i As Integer
Dim f As Integer
Dim nk1 As Long
'Dim nk2 As Integer
nk1 = Int(frmMain.lblCount.Caption)
For i = 1 To frmMain.lvProcess.ListItems.Count - 1
If blnScan = False Then
    Exit Sub
End If
'DoEvents
frmMain.lblCount.Caption = nk1 + (i - 1)
frmMain.SB.Panels(2).Text = frmMain.lblCount.Caption
frmMain.SB.Refresh
'frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + 1
'frmMain.SB.Panels(2).Text = frmMain.lblFound.Caption
'MsgBox "1"
 LogPrint "found process in memory-" + frmMain.lvProcess.ListItems.Item(i).SubItems(1)
Debug.Print frmMain.lvProcess.ListItems.Item(i).SubItems(1)
'frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + 1
    If frmMain.lvProcess.ListItems.Item(i).SubItems(1) <> "System Process" Then
        ' MsgBox frmMain.lvProcess.ListItems.Item(i).SubItems(1)
        Pro_ID = frmMain.lvProcess.ListItems.Item(i).SubItems(2)
        Call ScanFileProc("c:\", frmMain.lvProcess.ListItems.Item(i).SubItems(1), Pro_ID)
      End If
Next

'frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + ProcCount
'MsgBox "çàêîí÷èë ñêàíèòü ïàìÿòü"
Exit Sub

End Sub
Private Sub cmdScan_Click()
'ñêàíèðîâàíèå êàòàëîãà èëè âñåãî äèñêà-Ãëàâíàÿ
   'check scan status
    On Error GoTo 115
    If Me.mnuQarantine.Value = vbChecked Then
        QarNoDelete = True
    Else
        QarNoDelete = False
    End If
 
       If FileorFolderExists(App.Path + "\quarantine") = False Then
            MkDir App.Path + "\quarantine"
       End If
    If blnScan = False Then 'not scanning
        cmdScan.Enabled = False
        If Me.txtPath.Text <> "" Then   'scan path is set
        SB.Panels(1).Text = "Çàïóùåíî ñêàíèðîâàíèå..."
        Call smdScaning2 '1
        
         LogPrint "Scan-" + Trim$(Me.txtPath.Text)
        Me.txtPath.Text = Trim$(Me.txtPath.Text)
        SB.Panels(1).Text = "Íà÷àëè ñêàíèðîâàíèå"
            lvVirusFound.ListItems.Clear
            Dim ST As Variant, ET As Variant
            'get start time
            ST = Time
            'reset all statistic
            blnScan = True
            Me.Text1.Enabled = False
            'Me.lblCount.Caption = 0
            'Me.lblFound.Caption = 0
            'Me.lblCleaned.Caption = 0
            Me.cmdStop.SetFocus
            ' strScanDetail = "<Font Name='Verdana' Size=3 Color=Blue>Scanning Íà÷àëè â : " & Time$ & "<br>Ñêàíèðóþ : <i>" & Me.txtPath.Text & "</i><br>-----------------------------<br></font>"
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Dim x As ListItem
            Set x = Me.lvVirusFound.ListItems.Add(, , "Íà÷àëè â :", 1, 1)
            x.SubItems(1) = Time$
            Set x = Nothing
            Set x = Me.lvVirusFound.ListItems.Add(, , "Ñêàíèðóþ :", 1, 1)
            x.SubItems(1) = Me.txtPath.Text
            Set x = Nothing
            'start scanning
            Call ScanFile(Me.txtPath.Text)
            
            'finish scanning
            'get end time
            ET = Time
            'calculate time
            Dim x1 As String
            x1 = CalculateTime(ET - ST)
            ' strScanDetail = strScanDetail & "<font size=3 color=BLUE>-----------------------------<br>"
            If blnScan = True Then
                ' strScanDetail = strScanDetail & "Scanning Çàâåðøåíî â :" & Time$ & "</font>"
                Set x = Me.lvVirusFound.ListItems.Add(, , "Çàêîí÷èëè â :", 1, 1)
                x.SubItems(1) = Time$
                'tray message
                Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå çàâåðøåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ: " & x1)
            Else
                ' strScanDetail = strScanDetail & "Scanning was Çàâåðøåíî â :" & Time$ & "</font>"
                Set x = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                x.SubItems(1) = Time$
                'MsgBox "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1
                'tray message
                Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå îñòàíîâëåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1)
            End If
            Set x = Nothing
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Me.Text1.Enabled = True
            blnScan = False
            Me.lblPath.Caption = ""
        End If
    End If
    If cmdScan.Enabled = False Then
        cmdScan.Enabled = True
    End If
    SB.Panels(1).Text = "Ñêàíèðîâàíèå çàâåðøåíî"
    SB.Panels(2).Text = Me.lblCount.Caption
     SB.Panels(3).Text = Me.lblFound.Caption
    SB.Panels(4).Text = Me.lblCleaned.Caption
Exit Sub
115:
MsgBox "" + Error$
End Sub
Private Sub cmdScan_CL()
'ñêàíèðîâàíèå êàòàëîãà èëè âñåãî äèñêà-Ãëàâíàÿ
   'check scan status
    On Error Resume Next
    If Me.mnuQarantine.Value = vbChecked Then
        QarNoDelete = True
    Else
        QarNoDelete = False
    End If
 
       If FileorFolderExists(App.Path + "\quarantine") = False Then
            MkDir App.Path + "\quarantine"
       End If
    If blnScan = False Then 'not scanning
        cmdScan.Enabled = False
        If Me.txtPath.Text <> "" Then   'scan path is set
        'Call rescan_proc '1
        
         LogPrint "Scan-" + Trim$(Me.txtPath.Text)
        Me.txtPath.Text = Trim$(Me.txtPath.Text)
        SB.Panels(1).Text = "Íà÷àëè ñêàíèðîâàíèå"
            lvVirusFound.ListItems.Clear
            Dim ST As Variant, ET As Variant
            'get start time
            ST = Time
            'reset all statistic
            blnScan = True
            Me.Text1.Enabled = False
            'Me.lblCount.Caption = 0
            'Me.lblFound.Caption = 0
            'Me.lblCleaned.Caption = 0
            Me.cmdStop.SetFocus
            ' strScanDetail = "<Font Name='Verdana' Size=3 Color=Blue>Scanning Íà÷àëè â : " & Time$ & "<br>Ñêàíèðóþ : <i>" & Me.txtPath.Text & "</i><br>-----------------------------<br></font>"
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Dim x As ListItem
            Set x = Me.lvVirusFound.ListItems.Add(, , "Íà÷àëè â :", 1, 1)
            x.SubItems(1) = Time$
            Set x = Nothing
            Set x = Me.lvVirusFound.ListItems.Add(, , "Ñêàíèðóþ :", 1, 1)
            x.SubItems(1) = Me.txtPath.Text
            Set x = Nothing
            'start scanning
            Call ScanFile(Me.txtPath.Text)
            
            'finish scanning
            'get end time
            ET = Time
            'calculate time
            Dim x1 As String
            x1 = CalculateTime(ET - ST)
            ' strScanDetail = strScanDetail & "<font size=3 color=BLUE>-----------------------------<br>"
            If blnScan = True Then
                ' strScanDetail = strScanDetail & "Scanning Çàâåðøåíî â :" & Time$ & "</font>"
                Set x = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                x.SubItems(1) = Time$
                'tray message
                Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå çàâåðøåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1)
            Else
                ' strScanDetail = strScanDetail & "Scanning was Çàâåðøåíî â :" & Time$ & "</font>"
                Set x = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                x.SubItems(1) = Time$
                'MsgBox "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1
                'tray message
                Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1)
            End If
            Set x = Nothing
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Me.Text1.Enabled = True
            blnScan = False
            Me.lblPath.Caption = ""
        End If
    End If
    If cmdScan.Enabled = False Then
        cmdScan.Enabled = True
    End If
    SB.Panels(2).Text = Me.lblCount.Caption
     SB.Panels(3).Text = Me.lblFound.Caption
    SB.Panels(4).Text = Me.lblCleaned.Caption
Exit Sub
115:
MsgBox "" + Error$
End Sub
Private Sub cmdScan_CLBZ()
'ñêàíèðîâàíèå êàòàëîãà èëè âñåãî äèñêà-Ãëàâíàÿ
   'check scan status
  
    
    On Error Resume Next
    If Me.mnuQarantine.Value = vbChecked Then
        QarNoDelete = True
    Else
        QarNoDelete = False
    End If
 
       If FileorFolderExists(App.Path + "\quarantine") = False Then
            MkDir App.Path + "\quarantine"
       End If
    If blnScan = False Then 'not scanning
        cmdScan.Enabled = False
        If Me.txtPath.Text <> "" Then   'scan path is set
        'Call rescan_proc '1
        
         LogPrint "Scan-" + Trim$(Me.txtPath.Text)
        Me.txtPath.Text = Trim$(Me.txtPath.Text)
        SB.Panels(1).Text = "Íà÷àëè ñêàíèðîâàíèå"
            'lvVirusFound.ListItems.Clear'çà÷èñòêà ïàíåëè
            Dim ST As Variant, ET As Variant
            'get start time
            ST = Time
            'reset all statistic
            blnScan = True
            Me.Text1.Enabled = False
            'Me.lblCount.Caption = 0
            'Me.lblFound.Caption = 0
            'Me.lblCleaned.Caption = 0
            Me.cmdStop.SetFocus
            ' strScanDetail = "<Font Name='Verdana' Size=3 Color=Blue>Scanning Íà÷àëè â : " & Time$ & "<br>Ñêàíèðóþ : <i>" & Me.txtPath.Text & "</i><br>-----------------------------<br></font>"
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Dim x As ListItem
            Set x = Me.lvVirusFound.ListItems.Add(, , "Íà÷àëè â :", 1, 1)
            x.SubItems(1) = Time$
            Set x = Nothing
            Set x = Me.lvVirusFound.ListItems.Add(, , "Ñêàíèðóþ :", 1, 1)
            x.SubItems(1) = Me.txtPath.Text
            Set x = Nothing
            'start scanning
            Call ScanFile(Me.txtPath.Text)
            
            'finish scanning
            'get end time
            ET = Time
            'calculate time
            Dim x1 As String
            x1 = CalculateTime(ET - ST)
            ' strScanDetail = strScanDetail & "<font size=3 color=BLUE>-----------------------------<br>"
            If blnScan = True Then
                ' strScanDetail = strScanDetail & "Scanning Çàâåðøåíî â :" & Time$ & "</font>"
                Set x = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                x.SubItems(1) = Time$
                'tray message
               'Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå çàâåðøåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1)
            Else
                ' strScanDetail = strScanDetail & "Scanning was Çàâåðøåíî â :" & Time$ & "</font>"
                Set x = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                x.SubItems(1) = Time$
                'MsgBox "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1
                'tray message
               'Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1)
            End If
            Set x = Nothing
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Me.Text1.Enabled = True
            blnScan = False
            Me.lblPath.Caption = ""
        End If
    End If
    If cmdScan.Enabled = False Then
        cmdScan.Enabled = True
    End If
    SB.Panels(2).Text = Me.lblCount.Caption
     SB.Panels(3).Text = Me.lblFound.Caption
    SB.Panels(4).Text = Me.lblCleaned.Caption
Exit Sub
115:
MsgBox "" + Error$
End Sub
Private Sub cmdStop_Click()
    If blnScan = True Then
        If MsgBox("Îñòàíîâèòü ñêàíèðîâàíèå ?", vbYesNo + vbDefaultButton2 + vbExclamation, "Âíèìàíèå") = vbYes Then
           Me.mnuenblsc.Caption = "Ñêàíèðîâàòü Ìîé êîìïüþòåð"
                
            StopMycompScan = True
            blnScan = False
                Me.txtPath.Text = ""
            SB.Panels(1).Text = "Ïðåðâàíî ïîëüçîâàòåëåì"
             If cmdScan.Enabled = False Then
                cmdScan.Enabled = True
            End If
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
        'check scanning
    If blnScan = True Then
        If MsgBox("Ïðîèçâîäèòñÿ ñêàíèðîâàíèå. Âû äåéñòâèòåëüíî õîòèòå âûéòè ?", vbYesNoCancel + vbExclamation, "Ïðåðâàòü ñêàíèðîâàíèå?") = vbYes Then
            cReg.StopSearch
            Set cTray = Nothing
            End
        Else
            Cancel = True
        End If
    End If
    
    'check reg scanning

    
    'stop application
'    On Error Resume Next
'    'show any window that was hidden in the tray
'    Dim twnd As Long
'    twnd = CLng(Mid(Me.Tag, 1, (InStr(1, Me.Tag, "#", vbTextCompare) - 1)))
'    ShowWindow twnd, 5
'    ShowAllWindows
End
End Sub

Private Sub lvVirusFound_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub








Private Sub s12_Click()
Call ShellExecute(0, "Open", "http://mrbelyash.narod.ru/antivirus/belyashAV.htm", "", "", 1)
End Sub



 Sub smdScaning2()
procCleanSpl = 0
procCountSpl = 0
procFoundSpl = 0
Check_activitiproc
frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + procFoundSpl
frmMain.SB.Panels(3).Text = frmMain.lblFound.Caption
   frmMain.lblCleaned.Caption = Int(frmMain.lblCleaned.Caption) + procCleanSpl
   frmMain.SB.Panels(4).Text = frmMain.lblCleaned.Caption
frmMain.lblCount.Caption = Int(frmMain.lblCount.Caption) + procCountSpl
frmMain.SB.Panels(2).Text = frmMain.lblCount.Caption
SB.Refresh

End Sub
Sub Check_activitiproc()
blnScan = True
'ïðè âêëþ÷åíèè ñêàíèì âñå ïðîöåññû
On Error GoTo 100
'frmMain.lblCount.Caption = 0
'>> get new list of whats running
'Frame1.Refresh
programs.CheckProcesses
'pb1.Visible = True
'pb1.Min = 1
'pb1.Max = programs.processCount
'LD.Caption = "PLEASE WAIT...CHECK MEMORY..."
'LD.Refresh
'>> clear our listbox
lstRunning.Clear
list1.Clear
'>> fill our list box
Dim i As Integer
For i = 1 To programs.processCount - 1
If blnScan = False Then
    Exit Sub
End If
'DoEvents


DoEvents
 lstRunning.AddItem programs.ProcessName(i)
  list1.AddItem programs.ProcessHandle(i)
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
list1.Clear
blnScan = False

Exit Sub
100:
MsgBox "" + Error$
End Sub





'textbox determine file size
Private Sub Text1_Change()
    
    If Me.Text1.Text <> "" Then
        If Me.Text1.Text = 0 Then
            FileSize = 52428800
        Else
            FileSize = Int(Me.Text1.Text) * 1024
        End If
    Else
        FileSize = 1024
        Me.Text1.Text = 1
    End If
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    Dim num As String
    num = "0123456789"
    If InStr(1, num, Chr(KeyAscii)) < 1 Then
        KeyAscii = 0
    End If
    
End Sub


'========================================'
' MENU                                   '
'========================================'
Private Sub Image7_Click()
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.Image7.Visible = False
    Me.imgAbout.Visible = True
    'change content
    Me.frameScan.Visible = False
    Me.frameAbout.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.franalize.Visible = True
    Me.franalize.Enabled = True
    Me.franalize.Refresh
    'enable scroll
    
    rtf1.Text = ""
End Sub
Private Sub imgAbout_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = False
    Me.Image7.Visible = True
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.franalize.Visible = False
    Me.frameAbout.Visible = True
    
   
    
End Sub

Private Sub imgLicense_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = False
    Me.imgAbout.Visible = True
    Me.Image7.Visible = True
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = True
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
    Label4.Caption = "Belyash AntiTrojan 2008 beta" & _
vbCrLf + "Âîçìîæíîñòè:" & _
vbCrLf + "- Ñêàíèðîâàíèå óêàçàíûõ îáëàñòåé íà íàëè÷èå viruses, worms, trojans..." & _
vbCrLf + "- Ñêàíèðîâàíèå è ëå÷åíèå ïàìÿòè îò viruses, worms, trojans..." & _
vbCrLf + "- Ðåäàêòèðîâàíèå ïîëèòèê" & _
vbCrLf + "- Àâòîçàïóñê ïðè çàãðóçêå Windows" & _
vbCrLf + "- Ñêàíèðîâàíèå è èñïðàâëåíèå ðååñòðà (ñ óêàçàíèåì ðàñøèðåíèé) " & _
vbCrLf + "- Ìåíåäæåð ïðîöåññîâ" & _
vbCrLf + "- Ìåíåäæåð àâòîçàãðóçêè" & _
vbCrLf + "- Âêëþ÷åíèå íåêîòîðûõ îïöèé â ðååñòðå, çàáëîêèðîâàíûõ âèðóñîì" & _
vbCrLf + "- Îáíîâëåíèå èç ëîêàëüíîé ïàïêè èëè èç Èíòåðåíåòà" & _
vbCrLf + "- Ñêàíèðîâàíèå ïîòåíöèàëüíî îïàñíûõ ìåñò ñèñòåìû" & _
vbCrLf + "- Áëîêåð ïðèëîæåíèé(EXE)" & _
vbCrLf + "- Áëîêåð îïàñíûõ ïðèëîæåíèé " & _
vbCrLf + "              Ñèñòåìíûå òðåáîâàíèÿ:" & _
vbCrLf + "Ðàçðàáîòêà è òåñòèðîâàíèå òîëüêî ïîä Windows XP Sp2-Sp3" & _
vbCrLf + "Âíèìàíèå: Ðàçðàáîò÷èêè íå íåñóò íèêàêîé îòâåòñòâåííîñòè çà âîçìîæíûå áàãè â ïðîãðàììå(õîòÿ ìû ïîñòàðàëèñü ñäåëàòü âñ¸ çàâèñÿùåå îò íàñ, ÷òîáû èõ óñòðàíèòü.);" & _
vbCrLf + "Ïîýòîìó Âû èñïîëüçóåòå ïðîãðàììó íà ñâîé ñòðàõ è ðèñê."

End Sub

Private Sub imgScan_Click()
    
    'change menu
    Me.imgScan.Visible = False
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
    Me.Image7.Visible = True
    'change content
    Me.frameScan.Visible = True
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
End Sub

Private Sub imgSetting_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = False
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
    Me.Image7.Visible = True
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = True
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
End Sub

Private Sub imgTools_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = True
    Me.imgTools.Visible = False
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
     Me.Image7.Visible = True
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = False
    Me.frameTool.Visible = True
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
        'Call RefreshDefList
End Sub

Private Sub imgVirusDef_Click()
    
    'change menu
    Me.imgScan.Visible = True
    Me.imgSetting.Visible = True
    Me.imgVirusDef.Visible = False
    Me.imgTools.Visible = True
    Me.imgLicense.Visible = True
    Me.imgAbout.Visible = True
    Me.Image7.Visible = True
    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
    Me.frameVirusSig.Visible = True
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
   'RefreshDefList
End Sub

'========================================'
' VIRUS SIGNATURE                        '
'========================================'

Private Sub cmdAddToDef_Click()

    If Me.txtVirusName.Text <> "" Then

        Dim v As VirusSig
        v.Name = Me.txtVirusName.Text
        v.Type = "CRC"
        v.Value = Me.txtCRC32.Text
        Call WriteSig(v)
        'clear textboxes
        Me.txtVirusName.Text = ""
        Me.txtCRC32.Text = ""
        'refresh virus list
        'Call RefreshDefList
        
    End If
    
End Sub

Private Sub cmdCheckCRC_Click()
    Dim strFileType As String
strFileType = "All Files (*.*)|*.*|"
strFileType = strFileType & " Executeble files (*.exe)|*.exe|"
strFileType = strFileType & " Executeble files (*.dll)|*.dll|"
strFileType = strFileType & " Executeble files (*.com)|*.com|"
strFileType = strFileType & " Visual Basic Script (*.vbs)|*.vbs|"

Me.CommonDialog1.Filter = strFileType
    Me.CommonDialog1.DialogTitle = "Open to check CRC signature"
    'Me.CommonDialog1.Flags
    Me.CommonDialog1.ShowOpen
    If Me.CommonDialog1.FileName <> "" Then
        Me.txtCRC32.Text = CRC.GetCRC(Me.CommonDialog1.FileName)
        txtVirusName.Text = Me.CommonDialog1.FileName
        
    End If
'    Kill Me.CommonDialog1.FileName
End Sub

Private Sub lvVirusList_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Sub RefreshDefList()
 
    
    
    
    
    If Dir(App.Path + "\update.txt") <> "" Then
            Dim fso As New FileSystemObject, txtfile, fi
            Set fi = fso.GetFile(App.Path + "\update.txt")
        Me.lblLastUpdate.Caption = Format(fi.DateLastModified, "dd mmmm yyyy")
        Me.SB.Panels(5).Text = Format(fi.DateLastModified, "dd mmmm yyyy")
        Set fi = Nothing
        Set fso = Nothing
        
    Else
        If Dir$(App.Path + "\Avbase.bel", vbNormal) <> "" Then
                            Dim fso1 As New FileSystemObject, fi8
                      Set fi8 = fso1.GetFile(App.Path + "\Avbase.bel")
                    Me.lblLastUpdate.Caption = Format(fi8.DateLastModified, "dd mmmm yyyy")
                    Me.SB.Panels(5).Text = Format(fi8.DateLastModified, "dd mmmm yyyy")
                 Set fi8 = Nothing
                 Set fso1 = Nothing
          Else
                Me.lblLastUpdate.Caption = Format(Now, "dd mmmm yyyy")
                Me.SB.Panels(5).Text = Format(Now, "dd mmmm yyyy")
        End If
    End If
    
   
    Me.lvVirusList.Refresh
End Sub





Private Sub tmrMonitor_Timer()
'MsgBox "1"
Debug.Print Format(Now, "hh:mm:ss")
'tmrMonitor.Enabled = False
End Sub

Private Sub tvFolders_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
 dt = tvFolders.SelectedItem.Image
   
    If Right(tvFolders.SelectedItem.FullPath, 1) = ":" Then
         'MsgBox "" + tvFolders.SelectedItem.FullPath + "\"
         txtPath = tvFolders.SelectedItem.FullPath + "\"
    Else
        'MsgBox "" + tvFolders.SelectedItem.FullPath
        txtPath = tvFolders.SelectedItem.FullPath
    End If
    '
  
    'tvFolders.SelectedItem.Image = "SelFol"
    'tvFolders.Refresh
End Sub

Private Sub txtVirusName_KeyPress(KeyAscii As Integer)
    
    If Chr(KeyAscii) = "," Then
        KeyAscii = 0
        MsgBox "Disallowed character for name.", vbCritical
    End If
    
End Sub

'========================================'
' SETTING                                '
'========================================'

Private Sub txtScanRegExt_Change()
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", 1, "RegExt", Me.txtScanRegExt.Text)
End Sub

Private Sub chkAutoScan_Click()
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", 1
End Sub





Private Sub cmdRestoreDefault_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate", 10
    Me.optScanExt.Value = True
    Me.optScanAll.Value = False
    Me.txtScanRegExt.Text = "OCX, DLL, EXE, VBS, SYS, VXD"
    CreateStringValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", 1, "RegExt", "OCX, DLL, EXE, VBS, SYS, VXD"
    frameSetting.Refresh
    'Call LoadSetting
End Sub

Private Sub optScanAll_Click()
    intSettingRegOption = 1
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", 1
End Sub

Private Sub optScanExt_Click()
    intSettingRegOption = 0
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption", 0
End Sub

Public Sub LoadSetting()
    'check general setting
    Me.chkAutoStart.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoStart"))
    Me.chkStartMin.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "StartMin"))
    Me.mnuQarantine.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Quarantine"))
    'log
    Me.chkLog.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log"))
    
 '    strString = getstring(nm, "Software\BGAntivirus", "LogSize1")
'    txtSizelog.Text = strString
    'EXE File Block
    Me.chkExeBlock.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock"))
    Me.chkControlAll.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll"))
    'check refresh rate
    txtRefreshRate.Text = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate")
    Me.tmrAutoRefresh.interval = txtRefreshRate.Text * 1000
    'check auto refresh status
    Me.chkAutoRefeshProList.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh"))
'    If chkAutoRefeshProList.Value = 1 Then
'        Me.tmrAutoRefresh.Enabled = True
'    Else
'        Me.tmrAutoRefresh.Enabled = False
'    End If
    
    'process block
    Me.chkBlockRisk.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk"))
    'start minimized
    If Me.chkStartMin.Value = 1 Then
        Call HideToTray
'        Hide
        'õåíäë îêíà
'      cTray.hwnd = hwnd
   'èêîíêà, ÷òî áóäåò îòîáðàæåíà â òðåå
'      cTray.Icon = Icon
   'òóëòèïñ (âñïëûâàþùàÿ ïîäñêàçêà)
'      cTray.ToolTipText = "Belyash Antitrojan 2008"
      
   'ñîçäàåì èêîíêó
'      cTray.Add
   
    End If
    'Me.chkAutoScan.Value = ""
    'check options
    intSettingRegOption = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption"))
    If intSettingRegOption = 1 Then 'scan all
        Me.optScanAll.Value = True
    Else    'scan specific file extension
        Me.optScanExt.Value = True
    End If
    Me.txtScanRegExt.Text = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RegExt")
    'get scan reg ext to variable for scanning
    strScanRegExt = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RegExt")
End Sub







Private Sub chkControlAll_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll", Val(Me.chkControlAll.Value)
End Sub

Private Sub cmdConfigAppBlock_Click()
GetAllKeys71
    'frmEditAppBlock.Show vbModal
End Sub
Private Sub txtRefreshRate_KeyUp(KeyCode As Integer, Shift As Integer)
    'Dim strTemp As String
    'strTemp = "1234567890" & vbBack
    'MsgBox KeyCode
    'If InStr(1, strTemp, Chr$(KeyCode)) > 0 Then
    If Len(Me.txtRefreshRate.Text) >= 4 Then Me.txtRefreshRate.Text = Left$(Me.txtRefreshRate.Text, 3)
    If KeyCode = 8 Or (KeyCode >= 96 And KeyCode <= 105) Then
        If Len(Me.txtRefreshRate.Text) = 0 Then Me.txtRefreshRate.Text = 0
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate", Me.txtRefreshRate.Text
        Me.tmrAutoRefresh.interval = Me.txtRefreshRate.Text * 1000
    End If
End Sub
'========================================'
' TOOLS                                  '
'========================================'

'Sub Menu in Tools
'-----------------
Private Sub cmdEnableReg_Click()
    
 backupRegnow
    Me.frameTool_ScanReg.Visible = False
    Me.frameTool_Process.Visible = False
    Me.frameTool_Startup.Visible = False
    Me.frameTool_EnableReg.Visible = True
    
    Call LoadRegistry
End Sub
Sub backupRegnow()
 If MsgBox("Ñäåëàòü backup ðååñòðà?", vbYesNo + vbQuestion) = vbYes Then
    SB.Panels(1).Text = "Äåëàþ êîïèþ ðååñòðà...Ïîäîæäèòå"
        create_Backup "1"
         create_Backup "2"
          create_Backup "3"
           create_Backup "4"
            create_Backup "5"
    SB.Panels(1).Text = "Êîïèÿ ðååñòðà ñîçäàíà"
    End If
End Sub
 Sub ExecCmd(cmdline$)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
Dim ret As Long
' Èíèöèàëèçèðóåì ñòðóêòóðó STARTUPINFO:
start.cb = Len(start)
' Çàïóñêàåì ïðèëîæåíèå:
ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
' Æäåì çàâåðøåíèÿ çàïóùåííîãî ïðèëîæåíèÿ:
ret& = WaitForSingleObject(proc.hProcess, INFINITE)
ret& = CloseHandle(proc.hProcess)
End Sub
Sub create_Backup(GoBackup1 As String)
On Error GoTo 100
Dim fHCR As String
Dim fHCU As String
Dim fHLM As String
Dim fHU As String
Dim fHCC As String
'If Check1.Value = vbChecked Or Check1.Value = vbGrayed Then
'basRegistry.regPiS "", "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegedit", 0
'basRegistry.regPiS "HKEY_CURRENT_USER", "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegedit", 0
CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
'CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0

fHCR = "HCR" + CStr(Format(Now, "dd-mm-yy" + "_" + "hh.mm.ss")) + ".reg"
fHCU = "HCU" + CStr(Format(Now, "dd-mm-yy" + "_" + "hh.mm.ss")) + ".reg"
fHLM = "HLM" + CStr(Format(Now, "dd-mm-yy" + "_" + "hh.mm.ss")) + ".reg"
fHU = "HU" + CStr(Format(Now, "dd-mm-yy" + "_" + "hh.mm.ss")) + ".reg"
fHCC = "HCC" + CStr(Format(Now, "dd-mm-yy" + "_" + "hh.mm.ss")) + ".reg"
Select Case GoBackup1
 Case "1"

    ExecCmd ("regedit" + " " + "/e" + " " + fHCR + " HKEY_CLASSES_ROOT\")
  Case "2"

    ExecCmd ("regedit" + " " + "/e" + " " + fHCU + " HKEY_CURRENT_USER\")
  Case "3"

    ExecCmd ("regedit" + " " + "/e" + " " + fHLM + " HKEY_LOCAL_MACHINE\")
  Case "4"

    ExecCmd ("regedit" + " " + "/e" + " " + fHU + " HKEY_USERS\")
  Case "5"

    ExecCmd ("regedit" + " " + "/e" + " " + fHCC + " HKEY_CURRENT_CONFIG\")
    Case Else
    GoTo 100
End Select
'LogPrint "BACKUP is succesful"
'MsgBox "BACKUP is succesful", vbInformation
'End If
Exit Sub
100:
MsgBox "BACKUP íå ñîçäàí" + vbCrLf + Error$, vbCritical + vbApplicationModal
'LogPrint "BACKUP not create-" + Error$
End Sub


Private Sub cmdScanInvalidReg_Click()
'backupRegnow
    Me.frameTool_ScanReg.Visible = True
    Me.frameTool_Process.Visible = False
    Me.frameTool_Startup.Visible = False
    Me.frameTool_EnableReg.Visible = False
End Sub

Private Sub cmdTool_Startup_Click()
    Me.frameTool_ScanReg.Visible = False
    Me.frameTool_Process.Visible = False
    Me.frameTool_Startup.Visible = True
    Me.frameTool_EnableReg.Visible = False
    'retrieve all startup reg
    Call GetAllRun
End Sub

Private Sub cmdTool_ProcessMan_Click()
'ïîëó÷àåì ïðîöåññû
    Me.frameTool_ScanReg.Visible = False
    Me.frameTool_Process.Visible = True
    Me.frameTool_Startup.Visible = False
    Me.frameTool_EnableReg.Visible = False
    Me.lvProcess.ListItems.Clear
    Call programs.CheckProcesses
    Dim i As Integer
    Dim x As ListItem
    For i = 1 To programs.processCount - 1
       Set x = Me.lvProcess.ListItems.Add(, , programs.ProcessName(i), 7)
        x.SubItems(1) = programs.ProcessName(i)
        x.SubItems(2) = programs.ProcessHandle(i)
     'Me.lvProcess.ListItems.Add programs.ProcessName(i)
  '  lvProcess.ListItems.Add programs.ProcessHandle(i)
    Next i
    Set x = Nothing
    CheckProcess
    Me.lblTotalProcess.Caption = programs.processCount - 1
    lvProcessDetail.ListItems.Clear
    ModuleA.DEnumWindows
End Sub

'Scan Registry
'-------------

Private Sub lvErrorRegKey_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.Checked = False
    Else
        Item.Checked = True
    End If
End Sub

Private Sub lvErrorRegKey_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lvErrorRegKey_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'right click
    If Button = 2 Then
        'select item first
        'Call lvErrorRegKey_ItemClick
        PopupMenu mnuop, , Me.lvErrorRegKey.Left + Me.frameTool.Left + Me.frameTool_ScanReg.Left + x, Me.lvErrorRegKey.Top + Me.frameTool.Top + Me.frameTool_ScanReg.Top + Y
    End If
End Sub

Private Sub mnuOpSelAll_Click()
    Dim i As Long
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'checked
        .ListItems.Item(i).Checked = True
    Next
    End With
End Sub

Private Sub mnuopUnSelAll_Click()
    Dim i As Long
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'unchecked
        .ListItems.Item(i).Checked = False
    Next
    End With
End Sub


Private Sub cmdStartStop_Click()
    
    'button CAPTION
    If cmdStartStop.Caption = "Ñòàðò" Then 'start => stop
        cmdStartStop.Caption = "Ñòîï"
    Else
        cmdStartStop.Caption = "Ñòàðò"     'stop => start
        cReg.StopSearch
        txtCurKey.Text = ""
        Exit Sub
    End If
    
    'clear items
    Me.lvErrorRegKey.ListItems.Clear
    txtCurKey.Text = ""
    lblScanRegStatus.Caption = "Ïðîâåðÿþ :"
    lblScanRegError.Caption = 0
    
    'SEARCH START
    '============
    '0=HKEY_ALL
    cReg.RootKey = 0
    'Don't search in any specific subkey (Search in all subkeys)
    cReg.SubKey = ""
    'Only find errors in value names and value values
    cReg.SearchFlags = KEY_NAME * 0 + VALUE_NAME * 1 + VALUE_VALUE * 1 + WHOLE_STRING * 0
    'Search for registry values with the suffix "C:\"
    cReg.SearchString = "C:\"
    'Start searching for invalid registry values
    cReg.DoSearch
    '=============
    'SEARCH FINISH
    
    txtCurKey.Text = ""
End Sub

Private Sub cmdDeleteInvalidKey_Click()
'Call backupRegnow 'create backup
    Dim removed As Long, i As Integer
    'I don't think this is necessary, but if the registry backup takes a while, this program tells the user to wait.
    txtCurKey.FontSize = 12
    txtCurKey.FontBold = True
    'txtCurKey.Text = "Creating Registry Backup..."
    BackupReg
    'change status
    'txtCurKey.Text = "Registry Backup completed. Cleaning Errors..."
        
    'Loop through every item in lvwRegErrors
    With Me.lvErrorRegKey
    For i = 1 To .ListItems.Count
        'checked to be deleted
        If .ListItems.Item(i).Checked = True Then
            'Delete the registry error and mark the item as removed
            DeleteRegKey GetClassKey(.ListItems.Item(i).SubItems(1)), .ListItems.Item(i).SubItems(2), .ListItems.Item(i).SubItems(3)
            '.ListItems.Item(i).Text = "Cleaned"
            .ListItems.Item(i).Icon = 2
            .ListItems.Item(i).SmallIcon = 2
            removed = removed + 1
        End If
    Next
    End With
    'no deletion
    If removed = 0 Then GoTo endSub
    'change last status
    txtCurKey.Text = "×èñòêà ðååñòðà çàâåðøåíà"
    Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Cleaning Registry Errors completed and backup. Cleaned " & removed & " of " & Me.lvErrorRegKey.ListItems.Count & " .")
endSub:
    txtCurKey.FontSize = 8
    txtCurKey.FontBold = False
    txtCurKey.Text = ""
    
End Sub

'Create a backup of the registry, using the "regedit.exe /e" command takes too long.
Public Sub BackupReg()

    Dim i As Integer
    Dim TheKey As String
    Dim TheValue As String
    Dim DefaultValue As Boolean
    Dim BackupFilename As String
    Dim f As Long
    
    'check folder backup
    If FileorFolderExists(App.Path & "\RegBak") = False Then MkDir App.Path & "\RegBak"
    
    BackupFilename = App.Path & "\RegBak\Backup_" & Format(Date, "dd-mm-yyyy") & "_" & Format(Time, "hh-nn-ss") & ".reg"
    'MsgBox BackupFilename
    
    'open file to write
    f = FreeFile
    Open BackupFilename For Output As #f
    Print #f, "REGEDIT4" & vbCrLf
    'Loops through all the checked items and saves the values reg file
    With lvErrorRegKey
    For i = 1 To .ListItems.Count
        If .ListItems.Item(i).Checked = True Then
        
            TheKey = ReverseString(.ListItems.Item(i).SubItems(1) & "\" & .ListItems.Item(i).SubItems(2))
            'the value might ends with a "\", then it's the default value for that key
            If Right$(TheKey, 1) = "\" Then DefaultValue = True: TheKey = Mid(TheKey, 2)
            TheValue = Chr(34) & Replace(ReverseString(Mid(TheKey, 1, InStr(1, TheKey, "\") - 1)), "\", "\\") & Chr(34)
            TheKey = ReverseString(Mid(TheKey, InStr(1, TheKey, "\") + 1))
            If DefaultValue = True Then TheValue = "@"
            'add key to .reg file
            Print #f, "[" & TheKey & "]" '& vbCrLf
            Print #f, TheValue & "=" & Chr(34) & .ListItems.Item(i).SubItems(3) & Chr(34) '& vbCrLf
            
        End If
    Next
    Close #f
    End With
    
End Sub


'class cRegSearch event
Private Sub cReg_SearchFound(ByVal sRootKey As String, ByVal sKey As String, ByVal sValue As Variant, ByVal lFound As FOUND_WHERE)
    
    Dim KN As String    'KeyName
    Dim FileorPath As String  'File Path
    Dim x As ListItem
    
    'WHERE
    Select Case lFound
    Case FOUND_IN_KEY_NAME
        KN = "KEY_NAME"
    Case FOUND_IN_VALUE_NAME
        KN = "VALUE NAME"
    Case FOUND_IN_VALUE_VALUE
        KN = "VALUE VALUE"
    End Select

    FileorPath = sValue
    
    'Condition !
    'If Right$(FileorPath, 4) = ".EXE" Or Right$(FileorPath, 4) = ".exe" Or Right$(FileorPath, 4) = ".DLL" Or Right$(FileorPath, 4) = ".dll" Or Right$(FileorPath, 4) = ".OCX" Or Right$(FileorPath, 4) = ".ocx" Or Right$(FileorPath, 4) = ".SYS" Or Right$(FileorPath, 4) = ".sys" Or Right$(FileorPath, 4) = ".VXD" Or Right$(FileorPath, 4) = ".vxd" Or Right$(FileorPath, 3) = ".AX" Or Right$(FileorPath, 3) = ".ax" Then
    
    'check if actual file exist as in registry
    If FileorFolderExists(FormatValue(FileorPath)) = False Then 'not exist => invalid key
        
        If intSettingRegOption = 1 Then 'scan all
            'add to list for any key
            With Me.lvErrorRegKey
                Set x = .ListItems.Add(, , KN, 5, 5)
                x.SubItems(1) = sRootKey
                x.SubItems(2) = sKey
                x.SubItems(3) = sValue
            End With
            Set x = Nothing
            'add to counter
            Me.lblScanRegError.Caption = Int(Me.lblScanRegError.Caption) + 1
        Else    'scan specific extension
            'MsgBox FileorPath
            'MsgBox Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)
            'MsgBox InStr(1, LCase(Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)), LCase(strScanRegExt))
            'If InStr(1, Right$(FileorPath, 3), strScanRegExt, vbTextCompare) > 0 Then    'found in extension
            If InStr(1, strScanRegExt, Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3), vbTextCompare) > 0 Then    'found in extension
                With Me.lvErrorRegKey
                    Set x = .ListItems.Add(, , KN, 5, 5)
                    x.SubItems(1) = sRootKey
                    x.SubItems(2) = sKey
                    x.SubItems(3) = sValue
                End With
                Set x = Nothing
                'add to counter
                Me.lblScanRegError.Caption = Int(Me.lblScanRegError.Caption) + 1
            End If
        End If
    End If
    
End Sub

'class cRegSearch event
Private Sub cReg_SearchFinished(ByVal lReason As Long)
    
    If lReason = 0 Then
        Me.lblScanRegStatus.Caption = "Ïðîâåðêà çàâåðøåíà"
    ElseIf lReason = 1 Then
        Me.lblScanRegStatus.Caption = "Ïðîâåðêà ïðåðâàíà"
    Else
        Me.lblScanRegStatus.Caption = "Îøèáêà ïðîâåðêè"
    End If
    cmdStartStop.Caption = "Ñòàðò"
End Sub

'class cRegSearch event, when change key to search
Private Sub cReg_SearchKeyChanged(ByVal sFullKeyName As String)
    txtCurKey.Text = sFullKeyName
End Sub

'========================================'
' Process                                '
'========================================'

'check all process
Sub CheckProcess()
        
    Dim f As Long, strAttPro As String
    f = FreeFile
    Open App.Path & "\AttPro.bin" For Binary As f
    strAttPro = Input$(LOF(f), 1)
    Close f
    'MsgBox strAttPro
    
    Dim intBlockRisk As Integer, haveRisk As Boolean
    intBlockRisk = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk"))
    
    Dim i As Long
    With Me.lvProcess
        For i = 1 To .ListItems.Count
            If Right(.ListItems(i).SubItems(1), 11) = "svchost.exe" Then
                'MsgBox .ListItems(i).Text & " SYSTEM"
                .ListItems(i).SmallIcon = 7
            Else
                If InStr(1, strAttPro, .ListItems(i).Text, vbTextCompare) > 0 Then 'check in virus list
                    .ListItems(i).SmallIcon = 5 'mark risk
                    haveRisk = True 'mark variable
                Else
                    .ListItems(i).SmallIcon = 6
                End If
            End If
        Next
        'run only when Enable Auto Block, and Have Risk too
        If haveRisk = True And intBlockRisk = 1 Then
            For i = 1 To .ListItems.Count
                If .ListItems(i).SmallIcon = 5 Then  'risk item
                    Process_Kill .ListItems(i).SubItems(2)  'kill process
                End If
            Next
            'after kill all risk, refresh list
           ' Call GetProcessList(Me.lvProcess)
           
        End If
    End With
    
End Sub

'process clicked => get detail
Private Sub lvProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)
    'MsgBox ""
    Me.lvProcessDetail.ListItems.Clear
    pname = ""
 
    GetProcessesPids Trim(LCase(Item.Text)), procpids
    'all application instances
    'e.g multiple internet explorer windows
    Dim i As Integer
    i = 1
    While procpids(i) <> -1
        PID = procpids(i)
        GetWindowList Me.lvProcessDetail
        i = i + 1
    Wend
End Sub

'right click for dropdown menu

Private Sub lvProcess_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
        If Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).SmallIcon = 7 Then 'check on icon instead (can check on subitem(1)) 'system process
            mnuProMenu_Ban.Enabled = False
            mnuProMenu_Safe.Enabled = False
            PopupMenu mnuProMenu
        Else
            If Me.lvProcess.ListItems(Me.lvProcess.SelectedItem.Index).SmallIcon = 5 Then   'attention process
                mnuProMenu_Ban.Enabled = False
                mnuProMenu_Safe.Enabled = True
                PopupMenu mnuProMenu
            Else    'safe process
                mnuProMenu_Ban.Enabled = True
                mnuProMenu_Safe.Enabled = False
                PopupMenu mnuProMenu
            End If
        End If
    End If

End Sub

Private Sub cmdProcessRefresh_Click()
    'clear items from all listviews
    Me.lvProcess.ListItems.Clear
    Me.lvProcessDetail.ListItems.Clear
    'refresh data
    Call programs.CheckProcesses
    Dim i As Integer
    Dim x As ListItem
    For i = 1 To programs.processCount - 1
    
    Set x = Me.lvProcess.ListItems.Add(, , programs.ProcessName(i), 7)
        x.SubItems(1) = programs.ProcessName(i)
        x.SubItems(2) = programs.ProcessHandle(i)
     'Me.lvProcess.ListItems.Add programs.ProcessName(i)
  '  lvProcess.ListItems.Add programs.ProcessHandle(i)
    Next i
    Set x = Nothing
    CheckProcess
    Me.lblTotalProcess.Caption = programs.processCount - 1
End Sub

Private Sub tmrAutoRefresh_Timer()
    Call cmdProcessRefresh_Click
    If Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk")) = 1 Then
        Call CheckProcess
    End If
End Sub

Private Sub cmdProcessEnd_Click()
    If MsgBox("Do you really want to end this process?", vbYesNo + vbQuestion, "End Process : " & Me.lvProcess.SelectedItem.Text) = vbYes Then
        Dim Pro_ID As Long
        Pro_ID = Me.lvProcess.SelectedItem.SubItems(2)
        Process_Kill Pro_ID
        Call cmdProcessRefresh_Click
    End If
End Sub

'========================================'
' STARTUP                                '
'========================================'

'Enumerate from all RUN
Private Sub GetAllRun()
    On Error Resume Next
    Dim x As ListItem, hKey As Long, lCount As Long, i As Long
    lvStartUp.ListItems.Clear
    'Enumerate from HKEY_LOCAL_MACHINE , Run
    hKey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set x = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        x.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        x.SubItems(2) = "HKEY_LOCAL_MACHINE"
        Set x = Nothing
    Next i
    
    'Enumerate from HKEY_LOCAL_MACHINE , RunServices
    hKey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\RunServices")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set x = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        x.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        x.SubItems(2) = "HKEY_LOCAL_MACHINE (Service)"
        Set x = Nothing
    Next i
    
    'Enumerate from HKEY_CURRENT_USER , Run
    hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set x = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        x.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        x.SubItems(2) = "HKEY_CURRENT_USER"
        Set x = Nothing
    Next i
    
    'get startup from tasks
    Dim wdP As String
    wdP = Environ$("windir") + "\Tasks"
    'MsgBox wdP
    Dim fso As New FileSystemObject
    Dim sFolder As Folder
    Dim sFiles As Files
    Dim sFile As File
    Set sFolder = fso.GetFolder(wdP)
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            If sFile.Name <> "desktop.ini" Then
                Set x = Me.lvStartUp.ListItems.Add(, , sFile.Name)
                x.SubItems(1) = sFile.Path
                x.SubItems(2) = "Tasks"
                Set x = Nothing
            End If
        Next
    End If
    'get startup from current user startup
    Dim strUserProfile As String
    strUserProfile = Environ$("UserProfile") & "\Start Menu\Programs\Startup"
    'Set sFolder = fso.GetFolder("%userprofile%\Start Menu\Programs\Startup")
    Set sFolder = fso.GetFolder(strUserProfile)
    'MsgBox sFolder.Path
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            If sFile.Name <> "desktop.ini" Then
                'Set X = Me.lvStartUp.ListItems.Add(, , "User Startup")
                Set x = Me.lvStartUp.ListItems.Add(, , sFile.Name)
                x.SubItems(1) = sFile.Path
                x.SubItems(2) = "User Startup"
                Set x = Nothing
            End If
        Next
    End If
    'get startup from current all user startup
    Dim AUsr1 As String
    AUsr1 = Environ("ALLUSERSPROFILE") + "\Start Menu\Programs\Startup"
    
    Set sFolder = fso.GetFolder(AUsr1)
    Set sFiles = sFolder.Files
    If sFiles.Count > 0 Then
        For Each sFile In sFiles
            If sFile.Name <> "desktop.ini" Then
                 Set x = Me.lvStartUp.ListItems.Add(, , sFile.Name)
                 x.SubItems(1) = sFile.Path
                x.SubItems(2) = "All User Startup"
                Set x = Nothing
            End If
        Next
    End If
End Sub
Sub startup_log()
analize "-------------Start Autoruns----------------"
GetAllRun
Dim i As Integer
For i = 1 To Me.lvStartUp.ListItems.Count

analize Me.lvStartUp.ListItems.Item(i).ListSubItems(1)
Next i
analize "-------------End Autoruns----------------"
End Sub
'show Startup Folder /TASK?
Sub StartUpFolder()
'    txtCmdLine.Text = ""
'    txtName.Text = ""
'    If optStartMenu.Value = True Then
'        optStartMenu.FontBold = True
'        optRunServices.FontBold = False
'        optRun.FontBold = False
'        optRun2.FontBold = False
'        optWinINI.FontBold = False
'        Option1.FontBold = False
'    End If
'    ShellExecute 0, "open", CheckFolderID(StartUp), "", CheckFolderID(StartUp), 1
End Sub

'show Win.ini file
Sub ShowWinINIFile()
'    txtCmdLine.Text = ""
'    txtName.Text = ""
'    ShellExecute 0, "open", "notepad.exe", WinDir & "\win.ini", "", 1
'    If optWinINI.Value = True Then
'        optWinINI.FontBold = True
'        optRunServices.FontBold = False
'        optRun.FontBold = False
'        optStartMenu.FontBold = False
'        optRun2.FontBold = False
'        Option1.FontBold = False
'    End If
End Sub

'show System.ini file
Sub ShowSystemINIFile()
'    txtCmdLine.Text = ""
'    txtName.Text = ""
'    If Option1.Value = True Then
'        Option1.FontBold = True
'        optRunServices.FontBold = False
'        optRun.FontBold = False
'        optWinINI.FontBold = False
'        optStartMenu.FontBold = False
'        optRun2.FontBold = False
'    End If
'
'    MsgBox "Please note that the line n3 must be something like this" & vbCrLf & " [shell=Explorer.exe]  if you have something else" & vbCrLf & "Its possible that the system is loading in abnormal way", vbInformation, "Warning do not edit this if you dont know"
'    ShellExecute 0, "open", "notepad.exe", WinDir & "\system.ini", "", 1
End Sub

Private Sub cmdStartUp_Del_Click()
    With Me.lvStartUp
    Dim i As Long
    Dim tmp As Long
    Dim fso As New FileSystemObject
    For i = 1 To .ListItems.Count
        'checked to be deleted
        If .ListItems.Item(i).Checked = True Then
            'Delete startup
            'If .ListItems.Item(i).SubItems(2) = "ScheduledTask" Then    'Schedule Task   'not added yet
            'DeleteStartup GetClassKey("HKEY_LOCAL_MACHINE"), "Software/Microsoft/Windows/CurrentVersion/Policies/Explorer/Run", TASK
            If .ListItems.Item(i).SubItems(2) = "HKEY_LOCAL_MACHINE (Service)" Then 'run service startup
                DeleteStartup GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\RunServices", .ListItems.Item(i).Text
            ElseIf .ListItems.Item(i).SubItems(2) = "HKEY_LOCAL_MACHINE" Or .ListItems.Item(i).SubItems(2) = "HKEY_CURRENT_USER" Then 'normal startup
                If .ListItems.Item(i).Text = "BGAntivirus" Then
                    Me.chkAutoStart.Value = 0    'delete autostart
                    Exit For
                End If
                DeleteStartup GetClassKey(.ListItems.Item(i).SubItems(2)), "Software\Microsoft\Windows\CurrentVersion\Run", .ListItems.Item(i).Text
            Else    'If .ListItems.Item(i).SubItems(2) = "Startup" Then  'startup folder, else all are file
                'Kill .ListItems.Item(i).SubItems(1)  'kill filepath
                fso.DeleteFile .ListItems.Item(i).SubItems(1), True 'file system need to be deleted by force
            End If
        End If
    Next
    Set fso = Nothing
    End With
    'refresh startup run
    Call GetAllRun
End Sub

'========================================'
' REGISTRY                               '
'========================================'

Sub LoadRegistry()
    'Check in registry
    'Current User
    'Me.chkCUTaskmgr.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
'    Me.chkCUNoLogoff.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff"))
'    Me.chkCUNoCLose.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose"))
'    Me.chkCUNoLock.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation"))
'    Me.chkCUNoChangePwd.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword"))
'    Me.chkCURegTool.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
'    Me.chkCUNoCmd.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"))
'    Me.chkCUNoFolderOption.Value = Val(ReadValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
    Me.chkCUTaskmgr.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
    Me.chkCUNoLogoff.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff"))
    Me.chkCUNoCLose.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose"))
    Me.chkCUNoLock.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation"))
    Me.chkCUNoChangePwd.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword"))
    Me.chkCURegTool.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    Me.chkCUNoCmd.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"))
    Me.chkCUNoFolderOption.Value = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))

    'Local Machine
'    Me.chkLMNoFolderOption.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
'    Me.chkLMRegTool.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
'    Me.chkLMNoTaskmgr.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
'    Me.chkLMNoSRConfig.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"))
'    Me.chkLMNoSR.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR"))
'    Me.chkLmLimitSR.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"))
'    Me.chkLMNoMSI.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"))
    Me.chkLMNoFolderOption.Value = Val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
    Me.chkLMRegTool.Value = Val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    Me.chkLMNoTaskmgr.Value = Val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
    Me.chkLMNoSRConfig.Value = Val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"))
    Me.chkLMNoSR.Value = Val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR"))
    Me.chkLmLimitSR.Value = Val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"))
    Me.chkLMNoMSI.Value = Val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"))

End Sub

Private Sub chkCUNoChangePwd_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword", Val(Me.chkCUNoChangePwd.Value)
End Sub

Private Sub chkCUNoCLose_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose", Val(Me.chkCUNoCLose.Value)
End Sub

Private Sub chkCUNoCmd_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", Val(Me.chkCUNoCmd.Value)
End Sub

Private Sub chkCUNoFolderOption_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", Val(Me.chkCUNoFolderOption.Value)
End Sub

Private Sub chkCUNoLock_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation", Val(Me.chkCUNoLock.Value)
End Sub

Private Sub chkCUNoLogoff_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff", Val(Me.chkCUNoLogoff.Value)
End Sub

Private Sub chkCURegTool_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", Val(Me.chkCURegTool.Value)
    
End Sub

Private Sub chkCUTaskmgr_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", Val(Me.chkCUTaskmgr.Value)
End Sub

Private Sub chkLmLimitSR_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", Val(Me.chkLmLimitSR.Value)
End Sub

Private Sub chkLMNoFolderOption_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", Val(Me.chkLMNoFolderOption.Value)
End Sub

Private Sub chkLMNoMSI_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", Val(Me.chkLMNoMSI.Value)
End Sub

Private Sub chkLMNoSR_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", Val(Me.chkLMNoSR.Value)
End Sub

Private Sub chkLMNoSRConfig_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", Val(Me.chkLMNoSRConfig.Value)
End Sub

Private Sub chkLMNoTaskmgr_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", Val(Me.chkLMNoTaskmgr.Value)
End Sub

Private Sub chkLMRegTool_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", Val(Me.chkLMRegTool.Value)
End Sub

Private Sub cmdCleanReg_Click()
    Call CleanReg
    Me.chkCUNoCLose.Value = 0
    Me.chkCUNoLogoff.Value = 0
    Me.chkCUNoLock.Value = 0
    Me.chkCUNoChangePwd.Value = 0
    Call LoadRegistry
End Sub

'About
'=============================


Private Sub tvFolders_Collapse(ByVal Node As MSComctlLib.Node)
  'When a folder is collapsed, change the icon to a closed folder. :-)
  If Node.Image = "openfolder" Then Node.Image = "closedfolder"
End Sub

Private Sub tvFolders_Expand(ByVal Node As MSComctlLib.Node)
  On Error Resume Next
  Dim SubSubFolder As Folder
  Dim SubFolder As Folder
  Dim AFolder As Folder
  'When the folder is expanded, change the icon to an open folder
  If Node.Image = "closedfolder" Then Node.Image = "openfolder"
  
  'And build the tree further
  'We're actually adding sub-nodes to the children (making the expanded nodes "Grand Children")
  Set AFolder = fso.GetFolder(Node.Key & "\") 'Add the backslash :-)
  txtPath = AFolder
  For Each SubFolder In AFolder.SubFolders
    For Each SubSubFolder In SubFolder.SubFolders
      'Add children to the expanded nodes children
      tvFolders.Nodes.Add SubFolder.Path, 4, SubSubFolder.Path, SubSubFolder.Name, "closedfolder"
    Next
  Next
End Sub

Private Sub tvFolders_NodeClick(ByVal Node As MSComctlLib.Node)
  On Error GoTo ErrorHandler
  Dim AFolder As Folder
  Dim AFile As File
  
  'When a folder is clicked, show the files in that folder
  'lvFiles.ListItems.Clear
  Set AFolder = fso.GetFolder(Node.Key & "\") 'Add the backslash :-)
  
  For Each AFile In AFolder.Files
       'The added items key is the full path and filename
       'lvFiles.ListItems.Add , AFolder.Path & "\" & AFile.Name, AFile.Name, "genericfile", "genericfile"
  Next
ErrorHandler:
  ' If there's an error, call it out (such as permission denied, disk not ready, etc)
  If Err.Number <> 0 Then MsgBox Err.Description, vbCritical, "Error Number: " & Err.Number
End Sub

'Runs the files associated program
Private Function MyShell(PathAndFile As String, Optional Parameters As String = "", Optional ShowCmd As Long = vbNormalNoFocus) As Long
  Dim Path As String
  Dim File As String
  
  'Get everything up to and including the last backslash
  Path = Left(PathAndFile, InStrRev(PathAndFile, "\"))
  
  'Make absolutely sure there's no baskslashs on the end :-)
  While (Right$(Path, 1) = "\")
    Path = Left(Path, Len(Path) - 1)
  Wend
  
  'Grab everything from the last backslash, on to the end
  File = Mid$(PathAndFile, InStrRev(PathAndFile, "\") + 1)
  
  'Grab the results of the API Call
  MyShell = ShellExecute(0, vbNullString, File, Parameters, Path, ShowCmd)
  
  'If there's an error, let's try VB's Shell
  'If this doesn't work, it will yell out an error
  If MyShell < 32 Then Shell PathAndFile, ShowCmd
End Function
'========================================================
'ìîè ïðîöû ïîãíàëè
Private Sub Command1_Click()
On Error Resume Next
If Dir$(App.Path + "\analize.log", vbNormal) <> "" Then
    Kill App.Path + "\analize.log"
End If
Dim n As Integer
n = 0
rtf1.Text = ""
If Check1.Value = vbChecked Then
    SB.Panels(1).Text = "Îáðàáàòûâàþ ïåðåìåííûå îêðóæåíèÿ...Æäèòå..."
    peremen
    n = n + 1
End If
If Check4.Value = vbChecked Then
    SB.Panels(1).Text = "Îáðàáàòûâàþ àêòèâíûå îêíà...Æäèòå..."
    ModuleA.Ac_Wind
    n = n + 1
End If
If Check2.Value = vbChecked Then
    SB.Panels(1).Text = "Îáðàáàòûâàþ àêòèâíûå ïðîöåññû...Æäèòå..."
    poci12
    n = n + 1
End If
If Check5.Value = vbChecked Then
    'ýêñïîðò ðåãèñòðà
    SB.Panels(1).Text = "Îáðàáàòûâàþ ðååñòð ...Æäèòå..."
        create_Backup "1"
         create_Backup "2"
          create_Backup "3"
           create_Backup "4"
            create_Backup "5"
            n = n + 1
End If
If Check7.Value = vbChecked Then
     SB.Panels(1).Text = "Ñ÷èòûâàþ íàñòðîéêè ...Æäèòå..."
    GetAllKeys1
    n = n + 1
End If

If Check6.Value = vbChecked Then
     SB.Panels(1).Text = "Îáðàáàòûâàþ àâòîçàãðóçêó ...Æäèòå..."
    startup_log
    n = n + 1
End If
If n > 0 Then
    GoTo 9
Else
    SB.Panels(1).Text = "Ëîã íå ñîçäàí.Âûáåðèòå îïöèè"
    GoTo 800
   
End If
9:

Dim dgFile As Integer
dgFile = FreeFile
Open App.Path + "\analize.log" For Append As #dgFile
Print #dgFile, rtf1.Text
Close #1
     SB.Panels(1).Text = "Ëîã ñîçäàí è ñîõðàíåí â ôàéëå analize.log"
     MsgBox "Äàííûå ñîõðàíåíû â ôàéëå" + vbCrLf + App.Path + "\analize.log", vbInformation
     Exit Sub
800:

    MsgBox "Ëîã íå ñîçäàí" + vbCrLf + Error$, vbCritical, pr
    Debug.Print Err

End Sub
Sub analiz_fil(kmessage As String)
On Error GoTo 100
Dim bFile As Integer
bFile = FreeFile
Open App.Path + "\analize.log" For Append As #bFile
Print #bFile, kmessage
Close #bFile
Exit Sub
100:
MsgBox "Ðåïîðò íå ñîçäàí" + vbCrLf + Error$, vbCritical
End Sub
Public Function GetAllKeys1()
Dim x As ListItem, hKey As Long, lCount As Long, i As Long
analize "-----------------Options programm start-------------------"
Dim apppa As String
Dim RegExt As String


'apppa = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPath")
'analize "AppPath=" & apppa
    'Enumerate from HKEY_LOCAL_MACHINE , Run
    hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        'Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        'X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        'Set X = Nothing
        analize Trim$(CStr(EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", EnumValue(hKey, i))))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
    analize "---------------Ban---------------"
hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        'Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        'X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        'Set X = Nothing
        analize Trim$(CStr(EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", EnumValue(hKey, i))))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
        analize "---------------Allow---------------"
hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        'Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        'X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        'Set X = Nothing
        analize Trim$(CStr(EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", EnumValue(hKey, i))))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i

analize "-----------------Options programm end--------------------"

End Function
Public Function GetAllKeys71()
Dim x As ListItem, hKey As Long, lCount As Long, i As Long
' "-------------Options programm start----------------"
Dim apppa As String
Dim RegExt As String


'apppa = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPath")
'analize "AppPath=" & apppa
    'Enumerate from HKEY_LOCAL_MACHINE , Run
    'hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus")
    'lCount = GetCount(hKey, Values)
    'For i = 0 To lCount - 1
       ' Set X = frmEditAppBlock.lvAppList.ListItems.Add(, , EnumValue(hKey, i))
       ' X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
       ' Set X = Nothing

        ' frmEditAppBlock.List1.AddItem CStr(EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", EnumValue(hKey, i)))
'         frmEditAppBlock.lvAppList.ListItems.Add CStr(EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", EnumValue(hKey, i)))
       'frmEditAppBlock.lvAppList.ListItems.Add GetKeyValue(hKey, EnumValue(hKey, i))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    'Next i
'    analize "-----Ban-----"
hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
    
        'frmEditAppBlock.List1.AddItem , , EnumValue(hKey, i)
      
        frmEditAppBlock.list1.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
       
        Set x = Nothing
      '  frmEditAppBlock.lvAppList.ListItems.Add EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", EnumValue(hKey, i))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
       ' analize "-----Allow-----"
hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
  
        'frmEditAppBlock.List1.AddItem , , EnumValue(hKey, i)
         frmEditAppBlock.List2.AddItem GetKeyValue(hKey, EnumValue(hKey, i))
'        X.SubItems(2) = "HKEY_LOCAL_MACHINE"
   
        Set x = Nothing
        'frmEditAppBlock.lvAppList.ListItems.Add EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", EnumValue(hKey, i))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i

' "-------------Options programm end----------------"
frmEditAppBlock.Show
End Function

Sub analize(dmesd As String)
'ëîãèì
'analiz_fil (dmesd) + vbCrLf
If Trim$(dmesd) <> "" Then
    frmMain.rtf1.Text = frmMain.rtf1.Text + Trim$(dmesd) + vbCrLf
End If
End Sub
Sub peremen()
Dim m As Integer
Dim EnvString As String
Dim a As Long
'Dim a_minute As Long
analize "-------------Info start----------------"
'analize app.Minor+app.Minor+app.
analize "Version: " & App.Major & "." & App.Minor & "." & App.Revision & "b"
m = 1
Do
EnvString = Environ(m)
analize Environ(m)
m = m + 1
Loop Until EnvString = ""

Dim a_hour, a_minute, a_second
a = Format(GetTickCount() / 1000, "0") 'âñåãî ñåêóíä
a_hour = Int(a / 3600)
a = a - a_hour * 3600
a_minute = Int(a / 60)
a_second = a - a_minute * 60
analize "Âðåìÿ ðàáîòû " & Str(a_hour) & " ÷àñ(à)îâ " & Str(a_minute) & " ìèíóò" & Str(a_second) & " ñåêóíä"
analize "-------------Info start----------------"
End Sub

Private Sub poci12()
analize "-------------Process start----------------"
Me.lvProcess.ListItems.Clear
Call programs.CheckProcesses
       Dim i As Integer
    For i = 1 To programs.processCount - 1
       analize Trim$(programs.ProcessName(i)) + "-PID(" + CStr(programs.ProcessHandle(i)) + ")"
   Next i
analize "-------------Process end----------------"
End Sub
Sub HideToTray()

    'F10 button minimize any window to tray
        Dim pt As POINTAPI
        Dim i As Long, tmp As Long, fhwnd As Long
        Dim traydata As NOTIFYICONDATA
        Dim wndtext As String * 256
        Dim wndlen As String
        Dim clslen As Long
        Dim clsname As String * 260
        Dim clsinfo As WNDCLASS
        Dim tpid As Long, hProc As Long
        Dim n As Long
        Dim Icon As Long
        
'        fhwnd = GetForegroundWindow
'        If fhwnd = 0 Then
'                'Timer1.Enabled = True
'            Exit Sub
'        End If
        
        'get cursor position
        GetCursorPos pt
        tmp = WindowFromPoint(pt.x, pt.Y)
        'is there a window or not
        If tmp = 0 Then Exit Sub
        fhwnd = GetTopLevelWindow(tmp)
        
        Dim hform As Form
        Set hform = New frmMain
        
        'setup attributes for the tray icon
        'extract the icon of the exe
        GetWindowThreadProcessId fhwnd, tpid
      'õåíäë îêíà
    Me.WindowState = vbMinimized
    Form_Resize
    '  cTray.hwnd = Me.hwnd
   'èêîíêà, ÷òî áóäåò îòîáðàæåíà â òðåå
'      cTray.Icon = Icon
   'òóëòèïñ (âñïëûâàþùàÿ ïîäñêàçêà)
'      cTray.ToolTipText = "Belyash Antitrojan 2008"
      
   'ñîçäàåì èêîíêó
'    cTray.Add
   'Me.Hide
End Sub
'èâåíò ñðàáàòûâàåò ïðè äåéñòâèÿõ íà èêîíêå â òðåå
Private Sub cTray_OnIcon(MouseButton As Integer)
   'îáëàäî÷íàÿ èíôîðìàöèÿ
      Debug.Print MouseButton
   
   'ëåâûé äâîéíîé êëèê
  If MouseButton = TRAYICON_MOUSE_LEFTDBLCLICK Then
  'MsgBox "LeftDoubleClick on TrayIcon"
    Me.WindowState = vbNormal
    Me.Show
    Me.Top = 0
    'Me.Left = 100
  End If
   'îòæàòèå ïðàâîé êíîïêè ìûøè
      If MouseButton = TRAYICON_MOUSE_RIGHTUP Then
        cTray.CallPopupMenu Me, mnuTray, 2, , , mnuOpen
      End If
      End Sub
Public Sub baloon(sCmex As String)
cTray.DisplayBalloon "Belyash AntiTrojan 2008 beta", sCmex, NIIF_INFO ' + NIIF_NOSOUND
End Sub



