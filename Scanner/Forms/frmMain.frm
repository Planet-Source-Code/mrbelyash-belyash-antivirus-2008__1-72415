VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Belyash AntiTrojan 2008 beta"
   ClientHeight    =   8475
   ClientLeft      =   0
   ClientTop       =   -900
   ClientWidth     =   8910
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0356
   ScaleHeight     =   8475
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Scaner.xpcmdpicture imgAbout 
      Height          =   720
      Left            =   120
      TabIndex        =   99
      Top             =   5970
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1270
      State           =   1
      Picture         =   "frmMain.frx":B8C8
      Blend           =   0   'False
   End
   Begin Scaner.xpcmdpicture imgLicense 
      Height          =   720
      Left            =   120
      TabIndex        =   98
      Top             =   5160
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1270
      State           =   1
      Picture         =   "frmMain.frx":B8E4
      Blend           =   0   'False
   End
   Begin Scaner.xpcmdpicture Image7 
      Height          =   720
      Left            =   120
      TabIndex        =   97
      Top             =   4350
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1270
      State           =   1
      Picture         =   "frmMain.frx":B900
      Blend           =   0   'False
   End
   Begin Scaner.xpcmdpicture imgTools 
      Height          =   720
      Left            =   120
      TabIndex        =   96
      Top             =   3540
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1270
      State           =   1
      Picture         =   "frmMain.frx":B91C
      Blend           =   0   'False
   End
   Begin Scaner.xpcmdpicture imgSetting 
      Height          =   720
      Left            =   120
      TabIndex        =   95
      Top             =   2730
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1270
      State           =   1
      Picture         =   "frmMain.frx":B938
      Blend           =   0   'False
   End
   Begin Scaner.xpcmdpicture imgScan 
      Height          =   715
      Left            =   120
      TabIndex        =   94
      Top             =   1920
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1270
      State           =   1
      Picture         =   "frmMain.frx":B954
      Blend           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   60
      ScaleHeight     =   285
      ScaleWidth      =   8745
      TabIndex        =   86
      Top             =   8100
      Width           =   8775
      Begin VB.Label lblLastUpdate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7830
         TabIndex        =   91
         ToolTipText     =   "Êîëè÷åñòâî âèðóñíûõ çàïèñåé"
         Top             =   30
         Width           =   915
      End
      Begin VB.Label lblCleaned 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6930
         TabIndex        =   90
         ToolTipText     =   "Êîë-âî âûëå÷åíûõ"
         Top             =   30
         Width           =   870
      End
      Begin VB.Label lblFound 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6180
         TabIndex        =   89
         ToolTipText     =   "Êîë-âî íàéäåíûõ"
         Top             =   30
         Width           =   690
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5130
         TabIndex        =   88
         ToolTipText     =   "Êîë-âî ïðîâåðåííûõ"
         Top             =   30
         Width           =   1005
      End
      Begin VB.Label SB 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   87
         Top             =   0
         Width           =   4965
      End
   End
   Begin Scaner.xVistaForm xVistaForm1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   688
      Caption         =   "Belyash Anti-Trojan 2008b"
      DisplayIcon     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      Icon            =   "frmMain.frx":B970
      ShowSytemTrayIcon=   -1  'True
   End
   Begin VB.TextBox Text2 
      DataField       =   "CRC"
      DataSource      =   "Data1"
      Height          =   435
      Left            =   4350
      TabIndex        =   81
      Text            =   "Text2"
      Top             =   180
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.ListBox lstRunning 
      Height          =   255
      Left            =   1740
      TabIndex        =   47
      Top             =   120
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.ListBox list1 
      Height          =   255
      Left            =   1770
      TabIndex        =   46
      Top             =   420
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Timer tmrAutoRefresh 
      Enabled         =   0   'False
      Left            =   7410
      Top             =   1170
   End
   Begin MSComctlLib.ImageList img16x16 
      Left            =   6540
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
            Picture         =   "frmMain.frx":1304A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":130C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13366
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13615
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":138B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13D96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame frameScan 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6135
      Left            =   1710
      TabIndex        =   0
      Top             =   1920
      Width           =   7170
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   630
         Top             =   3195
      End
      Begin Scaner.EviProgressBar EviProgressBar2 
         Height          =   285
         Left            =   90
         TabIndex        =   118
         Top             =   5805
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   503
         BrushStyle      =   0
         Color           =   12937777
         Style           =   2
         Color2          =   12937777
      End
      Begin MSComctlLib.ListView lvVirusFound 
         Height          =   2565
         Left            =   90
         TabIndex        =   35
         Top             =   3150
         Width           =   7050
         _ExtentX        =   12435
         _ExtentY        =   4524
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
            Object.Width           =   3492
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   4200
         Visible         =   0   'False
         Width           =   1710
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
               Picture         =   "frmMain.frx":1403E
               Key             =   "mycomputer"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":167F0
               Key             =   "genericfile"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":172BA
               Key             =   "removabledrive"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":19A6C
               Key             =   "mydocs"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C21E
               Key             =   "cdrom"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E9D0
               Key             =   "closedfolder"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":21182
               Key             =   "desktop"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":23934
               Key             =   "openfolder"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":260E6
               Key             =   "unknowndrive"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28898
               Key             =   "floppydrive"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B04A
               Key             =   "harddrive"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D7FC
               Key             =   "netdrive"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2FFAE
               Key             =   "SelFol"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView tvFolders 
         Height          =   2535
         Left            =   60
         TabIndex        =   42
         Top             =   60
         Width           =   5040
         _ExtentX        =   8890
         _ExtentY        =   4471
         _Version        =   393217
         Style           =   7
         ImageList       =   "imgMain"
         Appearance      =   1
      End
      Begin Scaner.xpcmdbutton cmdScan 
         Height          =   345
         Left            =   1860
         TabIndex        =   56
         ToolTipText     =   "Çàïóñòèòü ñêàíèðîâàíèå"
         Top             =   2700
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
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
      Begin Scaner.xpcmdbutton cmdStop 
         Height          =   345
         Left            =   3750
         TabIndex        =   57
         ToolTipText     =   "Îñòàíîâèòü ñêàíèðîâàíèå"
         Top             =   2700
         Width           =   1365
         _ExtentX        =   2408
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
      End
      Begin VB.Label lblNormal09 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Time:0"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5340
         TabIndex        =   109
         Top             =   2100
         Width           =   1470
      End
      Begin VB.Label lblNormal112 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:0"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   5340
         TabIndex        =   108
         Top             =   1830
         Width           =   1275
      End
      Begin VB.Label lblNormal0 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ñòàòèñòèêà"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5610
         TabIndex        =   107
         Top             =   120
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         Height          =   2565
         Left            =   5190
         Top             =   60
         Width           =   1905
      End
      Begin VB.Label lblSize 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Size:0"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5340
         TabIndex        =   106
         Top             =   2340
         Width           =   1635
      End
      Begin VB.Label lblSyst 
         BackColor       =   &H00E0E0E0&
         Caption         =   "System:0"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5340
         TabIndex        =   105
         Top             =   1230
         Width           =   1635
      End
      Begin VB.Label lblRead 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ReadOnly:0"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5340
         TabIndex        =   104
         Top             =   1530
         Width           =   1635
      End
      Begin VB.Label lblArch 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Archive:0"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5340
         TabIndex        =   103
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label LblHiden 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hidden:0"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5340
         TabIndex        =   102
         Top             =   690
         Width           =   1635
      End
      Begin VB.Label lblNormal 
         BackColor       =   &H00E0E0E0&
         Caption         =   "UnScaning:0"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   5340
         TabIndex        =   101
         ToolTipText     =   "Íå ïðîâåðåíûå ôàéëû (ôàéëû íóëåâîé äëèííû)"
         Top             =   420
         Width           =   1635
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000040C0&
         Height          =   615
         Left            =   6795
         TabIndex        =   2
         Top             =   6210
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.Frame frameAbout 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6165
      Left            =   1680
      TabIndex        =   12
      Top             =   1890
      Width           =   7140
      Begin VB.Image Image12 
         Height          =   1365
         Left            =   3990
         Picture         =   "frmMain.frx":30108
         Top             =   2820
         Width           =   3105
      End
      Begin VB.Image Image11 
         Height          =   1365
         Left            =   3930
         Picture         =   "frmMain.frx":3DF1A
         Top             =   870
         Width           =   3105
      End
      Begin VB.Image Image10 
         Height          =   1335
         Left            =   330
         Picture         =   "frmMain.frx":4BD2C
         ToolTipText     =   "mrbelyash@rambler.ru"
         Top             =   2730
         Width           =   3105
      End
      Begin VB.Image Image9 
         Height          =   1335
         Left            =   240
         Picture         =   "frmMain.frx":5965E
         ToolTipText     =   "www.mrbelyash.narod.ru"
         Top             =   900
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
         TabIndex        =   43
         Top             =   135
         Width           =   3435
      End
   End
   Begin VB.Frame frameTool 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6225
      Left            =   1710
      TabIndex        =   3
      Top             =   1860
      Width           =   7080
      Begin Scaner.xpcmdbutton cmdScanInvalidReg 
         Height          =   375
         Left            =   180
         TabIndex        =   60
         Top             =   120
         Width           =   1725
         _ExtentX        =   3043
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
      End
      Begin Scaner.xpcmdbutton cmdTool_ProcessMan 
         Height          =   375
         Left            =   1980
         TabIndex        =   61
         Top             =   120
         Width           =   1635
         _ExtentX        =   2884
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
      End
      Begin Scaner.xpcmdbutton cmdTool_Startup 
         Height          =   375
         Left            =   3690
         TabIndex        =   62
         Top             =   120
         Width           =   1605
         _ExtentX        =   2831
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
      End
      Begin Scaner.xpcmdbutton cmdEnableReg 
         Height          =   375
         Left            =   5400
         TabIndex        =   63
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
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
      End
      Begin VB.Frame frameTool_EnableReg 
         BackColor       =   &H00E0E0E0&
         Height          =   5565
         Left            =   180
         TabIndex        =   5
         Top             =   570
         Width           =   6825
         Begin VB.CheckBox chkLMRegTool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îòêëþ÷èòü äîñòóï ê ðååñòðó"
            Height          =   375
            Left            =   360
            TabIndex        =   113
            Top             =   3150
            Width           =   3225
         End
         Begin VB.CheckBox chkCURegTool 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü ðååñòð"
            Height          =   375
            Left            =   360
            TabIndex        =   112
            Top             =   360
            Width           =   3225
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   4590
            Top             =   1980
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CheckBox chkLMNoSR 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàïðåòèòü âîññòàíîâëåíèå ñèñòåìû"
            Height          =   375
            Left            =   360
            TabIndex        =   31
            Top             =   4005
            Width           =   5175
         End
         Begin VB.CheckBox chkLmLimitSR 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îãðàíè÷èò ðàçìåð àðõèâà âîññòàíîâëåíèÿ ÎÑ"
            Height          =   375
            Left            =   360
            TabIndex        =   30
            Top             =   4545
            Width           =   5535
         End
         Begin VB.CheckBox chkLMNoMSI 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îòêëþ÷èòü èíñòàëÿøêó"
            Height          =   375
            Left            =   360
            TabIndex        =   29
            Top             =   4860
            Width           =   3615
         End
         Begin VB.CheckBox chkLMNoSRConfig 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Îòêëþ÷èòü äîñòóï ê íàñòðîéêàì âîññòàíîâëåíèÿ ÎÑ"
            Height          =   375
            Left            =   360
            TabIndex        =   28
            Top             =   4290
            Width           =   5895
         End
         Begin VB.CheckBox chkLMNoTaskmgr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü Task Manager"
            Height          =   375
            Left            =   360
            TabIndex        =   27
            Top             =   3450
            Width           =   4575
         End
         Begin VB.CheckBox chkLMNoFolderOption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Óáðàòü ""Ñâîéñòâà ïàïêè"""
            Height          =   375
            Left            =   360
            TabIndex        =   26
            Top             =   3720
            Width           =   3375
         End
         Begin VB.CheckBox chkCUNoFolderOption 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Óáðàòü ""Ñâîéñòâà ïàïêè"""
            Height          =   375
            Left            =   360
            TabIndex        =   25
            Top             =   930
            Width           =   2925
         End
         Begin VB.CheckBox chkCUNoCmd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü êîíñîëü"
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   2070
            Width           =   3795
         End
         Begin VB.CheckBox chkCUNoChangePwd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàïðåòèòü èçìåíÿòü ïðîëü ïîëüçîâàòåëÿ"
            Height          =   375
            Left            =   360
            TabIndex        =   23
            Top             =   2355
            Width           =   4575
         End
         Begin VB.CheckBox chkCUNoLock 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Óáðàòü ëîê êîìïà"
            Height          =   375
            Left            =   360
            TabIndex        =   22
            Top             =   1785
            Width           =   3945
         End
         Begin VB.CheckBox chkCUNoCLose 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü ""Âûõîä"""
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   1200
            Width           =   2655
         End
         Begin VB.CheckBox chkCUNoLogoff 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàïðåòèòü ""Ñìåíó ïîëüçîâàòåëÿ"""
            Height          =   375
            Left            =   360
            TabIndex        =   20
            Top             =   1485
            Width           =   3315
         End
         Begin VB.CheckBox chkCUTaskmgr 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Çàáëîêèðîâàòü Task Manager"
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   645
            Width           =   3225
         End
         Begin Scaner.xpcmdbutton cmdClearAutorun 
            Height          =   375
            Left            =   4800
            TabIndex        =   67
            Top             =   750
            Width           =   1785
            _ExtentX        =   3149
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
         End
         Begin Scaner.xpcmdbutton cmdCleanReg 
            Height          =   375
            Left            =   4800
            TabIndex        =   68
            Top             =   300
            Width           =   1785
            _ExtentX        =   3149
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
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   690
            TabIndex        =   33
            Top             =   2910
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
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   810
            TabIndex        =   32
            Top             =   150
            Width           =   3705
         End
      End
      Begin VB.Frame frameTool_Process 
         BackColor       =   &H00E0E0E0&
         Height          =   5550
         Left            =   180
         TabIndex        =   16
         Top             =   600
         Width           =   6840
         Begin MSComctlLib.ListView lvProcess 
            Height          =   3075
            Left            =   240
            TabIndex        =   17
            Top             =   840
            Width           =   6285
            _ExtentX        =   11086
            _ExtentY        =   5424
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
            TabIndex        =   18
            Top             =   3960
            Width           =   6285
            _ExtentX        =   11086
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
         Begin Scaner.xpcmdbutton cmdProcessRefresh 
            Height          =   375
            Left            =   240
            TabIndex        =   58
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
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
         End
         Begin Scaner.xpcmdbutton cmdProcessEnd 
            Height          =   375
            Left            =   2070
            TabIndex        =   59
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame frameTool_ScanReg 
         BackColor       =   &H00E0E0E0&
         Height          =   5535
         Left            =   180
         TabIndex        =   4
         Top             =   600
         Width           =   6855
         Begin VB.TextBox txtCurKey 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FF0000&
            Height          =   735
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   1080
            Width           =   6015
         End
         Begin MSComctlLib.ListView lvErrorRegKey 
            Height          =   3075
            Left            =   240
            TabIndex        =   10
            Top             =   2280
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   5424
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
         Begin Scaner.xpcmdbutton cmdStartStop 
            Height          =   375
            Left            =   3840
            TabIndex        =   64
            Top             =   180
            Width           =   1035
            _ExtentX        =   1826
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
         End
         Begin Scaner.xpcmdbutton cmdDeleteInvalidKey 
            Height          =   375
            Left            =   5040
            TabIndex        =   65
            Top             =   180
            Width           =   1635
            _ExtentX        =   2884
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   600
            Width           =   2685
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame frameTool_Startup 
         BackColor       =   &H00E0E0E0&
         Height          =   5490
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   6840
         Begin MSComctlLib.ListView lvStartUp 
            Height          =   4215
            Left            =   360
            TabIndex        =   15
            Top             =   1200
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   7435
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
         Begin Scaner.xpcmdbutton cmdStartUp_Del 
            Height          =   375
            Left            =   420
            TabIndex        =   66
            Top             =   240
            Width           =   1635
            _ExtentX        =   2884
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
         End
      End
   End
   Begin VB.Frame franalize 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ñáîð ñâåäåíèé"
      Height          =   6225
      Left            =   1740
      TabIndex        =   44
      Top             =   1860
      Visible         =   0   'False
      Width           =   7065
      Begin VB.ListBox List4 
         Height          =   255
         Left            =   3900
         TabIndex        =   100
         Top             =   210
         Visible         =   0   'False
         Width           =   2025
      End
      Begin VB.TextBox rtf1 
         Appearance      =   0  'Flat
         Height          =   2580
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   45
         Top             =   3330
         Width           =   6795
      End
      Begin Scaner.xpcmdbutton Command1 
         Height          =   345
         Left            =   240
         TabIndex        =   48
         Top             =   2850
         Width           =   1545
         _ExtentX        =   2725
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
      End
      Begin Scaner.xpcheckbox Check1 
         Height          =   345
         Left            =   300
         TabIndex        =   49
         Top             =   240
         Width           =   2835
         _ExtentX        =   5001
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox Check2 
         Height          =   315
         Left            =   300
         TabIndex        =   50
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox Check3 
         Height          =   315
         Left            =   300
         TabIndex        =   51
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox Check4 
         Height          =   315
         Left            =   300
         TabIndex        =   52
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox Check5 
         Height          =   315
         Left            =   300
         TabIndex        =   53
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox Check6 
         Height          =   315
         Left            =   300
         TabIndex        =   54
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox Check7 
         Height          =   315
         Left            =   300
         TabIndex        =   55
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
         BackColor       =   14737632
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   210
         X2              =   6930
         Y1              =   2700
         Y2              =   2700
      End
   End
   Begin VB.Frame frameLicense 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6255
      Left            =   1620
      TabIndex        =   11
      Top             =   1830
      Width           =   7110
      Begin VB.PictureBox PictureLOGO2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5490
         Left            =   150
         ScaleHeight     =   366
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   465
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   510
         Width           =   6975
      End
      Begin VB.PictureBox PictureLOGO1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   5460
         Left            =   150
         Picture         =   "frmMain.frx":66F90
         ScaleHeight     =   364
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   466
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   510
         Width           =   6990
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Î ïðîãðàììå..."
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
         Left            =   2370
         TabIndex        =   116
         Top             =   90
         Width           =   2805
      End
   End
   Begin VB.Frame frameSetting 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Last Update :"
      Height          =   6165
      Left            =   1680
      TabIndex        =   13
      Top             =   1860
      Width           =   7110
      Begin VB.TextBox txtSizeChk 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3660
         TabIndex        =   110
         Text            =   "22478208"
         Top             =   2520
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ïðèîðèòåò"
         ForeColor       =   &H00FF0000&
         Height          =   705
         Left            =   4050
         TabIndex        =   83
         Top             =   390
         Width           =   2985
         Begin MSComctlLib.Slider Slider1 
            Height          =   315
            Left            =   150
            TabIndex        =   84
            Top             =   270
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            LargeChange     =   1
            Max             =   3
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   2070
         TabIndex        =   82
         Text            =   "100024"
         Top             =   870
         Width           =   990
      End
      Begin VB.TextBox txtRefreshRate 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   2820
         TabIndex        =   40
         Text            =   "10"
         Top             =   3510
         Width           =   600
      End
      Begin VB.TextBox txtScanRegExt 
         Height          =   555
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         Top             =   5580
         Width           =   4845
      End
      Begin Scaner.xpcmdbutton cmdConfigAppBlock 
         Height          =   375
         Left            =   4500
         TabIndex        =   69
         ToolTipText     =   "Ðåäàêòèðîâàòü íàñòðîéêè çàáëîêèðîâàííûõ ïðèëîæåíèé"
         Top             =   2910
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
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
      End
      Begin Scaner.xpcmdbutton cmdRestoreDefault 
         Height          =   375
         Left            =   2910
         TabIndex        =   70
         ToolTipText     =   "Èçìåíèòü íàñòðîéêè ïî óìîë÷àíèþ(ðåêîìåíäóåòñÿ)"
         Top             =   5070
         Width           =   1755
         _ExtentX        =   3096
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
      End
      Begin Scaner.xpradiobutton optScanAll 
         Height          =   315
         Left            =   240
         TabIndex        =   71
         Top             =   4950
         Width           =   2415
         _ExtentX        =   4260
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
         BackColor       =   14737632
      End
      Begin Scaner.xpradiobutton optScanExt 
         Height          =   315
         Left            =   240
         TabIndex        =   72
         Top             =   5250
         Width           =   4965
         _ExtentX        =   8758
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox chkBlockRisk 
         Height          =   315
         Left            =   390
         TabIndex        =   73
         Top             =   1590
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox chkLog 
         Height          =   315
         Left            =   390
         TabIndex        =   74
         Top             =   900
         Width           =   3045
         _ExtentX        =   5371
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox mnuQarantine 
         Height          =   315
         Left            =   390
         TabIndex        =   75
         Top             =   570
         Width           =   3555
         _ExtentX        =   6271
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox chkExeBlock 
         Height          =   315
         Left            =   300
         TabIndex        =   76
         Top             =   3900
         Width           =   5715
         _ExtentX        =   10081
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox chkAutoRefeshProList 
         Height          =   315
         Left            =   390
         TabIndex        =   77
         Top             =   1920
         Width           =   2865
         _ExtentX        =   5054
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox chkControlAll 
         Height          =   315
         Left            =   600
         TabIndex        =   78
         Top             =   4260
         Width           =   5865
         _ExtentX        =   10345
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox chkAutoScan 
         Height          =   315
         Left            =   390
         TabIndex        =   79
         Top             =   1230
         Width           =   5295
         _ExtentX        =   9340
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
         BackColor       =   14737632
      End
      Begin Scaner.xpcmdbutton cmdSave 
         Height          =   375
         Left            =   5520
         TabIndex        =   80
         ToolTipText     =   "Ñîõðàíèòü íàñòðîéêè ïðîãðàììû"
         Top             =   5580
         Width           =   1425
         _ExtentX        =   2514
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
      End
      Begin Scaner.xpcheckbox xpcheAutoran 
         Height          =   315
         Left            =   390
         TabIndex        =   115
         Top             =   2220
         Width           =   2955
         _ExtentX        =   5212
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
         Caption         =   "Çà÷èñòêà àâòîðàíîâ"
         BackColor       =   14737632
      End
      Begin Scaner.xpcheckbox xpcheckbox71 
         Height          =   315
         Left            =   390
         TabIndex        =   117
         Top             =   2790
         Width           =   2865
         _ExtentX        =   5054
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
         Caption         =   "Ñêàíèðîâàòü âñå ôàéëû"
         BackColor       =   14737632
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Áëîêåð ïðèëîæåíèé"
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
         Left            =   360
         TabIndex        =   114
         Top             =   3090
         Width           =   5505
      End
      Begin VB.Label lblLabel2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Îãðàíè÷èòü ðàçìåð ïðîâåðÿåìîãî ôàéëà                                  áàéò"
         Height          =   195
         Left            =   420
         TabIndex        =   111
         Top             =   2550
         Width           =   5040
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
         Left            =   630
         TabIndex        =   41
         Top             =   3510
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
         Left            =   240
         TabIndex        =   37
         Top             =   4560
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
         TabIndex        =   14
         Top             =   60
         Width           =   4755
      End
   End
   Begin VB.Image Image2 
      Height          =   1230
      Left            =   1650
      Picture         =   "frmMain.frx":E3672
      Stretch         =   -1  'True
      Top             =   510
      Width           =   7005
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      X1              =   1470
      X2              =   1470
      Y1              =   1680
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000004&
      BorderWidth     =   2
      X1              =   60
      X2              =   1470
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   60
      X2              =   8820
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   60
      Picture         =   "frmMain.frx":FE0F4
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1410
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Declare Function GetTickCount Lib "Kernel32" () As Long
Public BB As Boolean
Public bw As Integer
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Dim stopMe As Boolean
Dim neg1 As Long
Private programs As New ProcessList
'Dim WithEvents cTray As TrayIconAndBalloon
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const flags = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
dwThreadId As Long
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

Public Function NumberOfProcessors() As Long
GetSystemInfo m_typSystemInfo
NumberOfProcessors = m_typSystemInfo.dwNumberOfProcessors
End Function





Private Sub cmdClearAutorun_Click()
autorans_killNow
End Sub
Sub autorans_killNow()
If Me.xpcheAutoran.Value = Unchecked Then
    Exit Sub
End If
Dim X As ListItem
Dim wseDisk As String
wseDisk = ""
    Dim fso As New FileSystemObject
    Dim drv As Drive
    Dim drvs As Drives
    On Error Resume Next    'in case not found, and on cd
    Dim k As Byte
    k = 0
    Set drvs = fso.Drives
    For Each drv In drvs
        DoEvents
       autorunsdelorni (drv.DriveLetter & ":\autorun.inf")
       If FileorFolderExists(drv.DriveLetter & ":\autorun.inf") = True Then
            k = 1
        End If
        If Trim$(drv.DriveLetter) <> "" Then
            wseDisk = wseDisk & drv.DriveLetter & ": +"
        End If
    Next
    If k = 0 Then
        Set X = Me.lvVirusFound.ListItems.Add(, , "Ïîèñê autorun.inf", 1, 2)
         X.SubItems(1) = wseDisk
        X.SubItems(2) = "Äèñêè î÷èùåíû"
    Else
        Set X = Me.lvVirusFound.ListItems.Add(, , "Ïîèñê autorun.inf", 1, 5)
        X.SubItems(1) = wseDisk
        X.SubItems(2) = "Îøèáêà ïðè î÷èñòêå"
    End If
    LogPrint "Autorun.inf íà âñåõ äèñêàõ óäàëåí."
    Set fso = Nothing
    Set drv = Nothing
    Set drvs = Nothing
    Set X = Nothing
End Sub
Function autorunsdelorni(ptnau As String)
On Error GoTo 100
Kill ptnau
autorunsdelorni = True
100:
End Function
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
       ' Call ScanFileProc("c:\", frmMain.lvProcess.ListItems.Item(i).SubItems(1), Pro_ID)
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
strFileType = strFileType & " Text Files (*.bel)|*.bel|"
'Ïðèñâàèâàåì åå ñâîéñòâó Filter
CommonDialog1.Filter = strFileType
'Óñòàíàâëèâàåì íåîáõîäèìûé èíäåêñ
CommonDialog1.FilterIndex = 2
'Ïðèñâàèâàåì íà÷àëüíóþ äèðåêòîðèþ ñâîñòâó InitDir
CommonDialog1.InitDir = App.Path
'Îáåñïå÷èâàåì çàùèòó îò íåïðàâèëüíîãî ââåäåííîãî ôàéëà èëè äåðèêòîðèè, à òàê æå ñêðûâàåì ôëàæåê Read Only

CommonDialog1.flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
'Âûçûâàåì äèàëîã Open
CommonDialog1.Action = 1 'Èëè æå CommonDialog1.ShowOpen
'***********
Call obnowlLocal(CommonDialog1.Filename)
If MsgBox("Äëÿ ïîäãðóçêè íîâûõ áàç ïðîãðàììà äîëæíà áûòü ïåðåçàïóùåíà. Çàâåðøèòü ðàáîòó ñ ïðîãðàììîé ?", vbCritical + vbYesNo, pr) = vbYes Then
    End
End If
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
Exit Sub
  
'GetAllRun'ïîêà ñòàðòàïû íå áóäó îáðàáàòûâàòü
End Sub
Private Sub cmdSave_Click()
'ñîõðàíèòü íàñòðîéêè
If MsgBox("Ñîõðàíèòü íàñòðîéêè ?", vbYesNo + vbQuestion, pr) = vbNo Then
    Exit Sub
End If
If Len(txtSizeChk.Text) = 0 Or val(txtSizeChk.Text) = 0 Then
    MsgBox "Íåïðàâèëüíûé ðàçìåð îãðàíè÷åíèÿ ïðîâåðêè ôàéëîâ", vbCritical
Exit Sub
End If
SaveSetting App.exeName, "Options", "chkSizeFile", Trim$(Me.txtSizeChk.Text)
If Me.xpcheAutoran.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Autorans", val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Autorans", val(0)
End If
If Me.xpcheckbox71.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanAllFiles", val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanAllFiles", val(0)
End If



If Me.mnuQarantine.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Quarantine", val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Quarantine", val(0)
End If

If Me.chkLog.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log", val(1)
    SaveSetting App.exeName, "Options", "logSize", Me.Text3.Text
Else
  CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log", val(0)
  SaveSetting App.exeName, "Options", "logSize", "100024"
End If

If Me.chkAutoScan.Value = Checked Then
 CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanReg", val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanReg", val(0)
End If
SaveSetting App.exeName, "Options", "Priority", Slider1.Value

If Me.chkBlockRisk.Value = Checked Then
 CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk", val(1)
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk", val(0)
End If

If chkAutoRefeshProList.Value = Checked Then
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh", val(1)
     Me.tmrAutoRefresh.Enabled = True
Else
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh", val(0)
        Me.tmrAutoRefresh.Enabled = False
End If
    'edit in EXEFile/Shell/Open/Command
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    If Me.chkExeBlock.Value = Checked Then
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock", val(1)
         sh.regwrite "HKCR\exefile\shell\open\command\original", Chr$(34) + "%1" + Chr$(34) + " %*"
        sh.regwrite "HKCR\exefile\shell\open\command\", App.Path & "\AppBlock.EXE %1 %*"
        'chkControlAll.Enabled = True
    Else
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock", val(0)
         sh.regwrite "HKCR\exefile\shell\open\command\", Chr$(34) + "%1" + Chr$(34) + " %*"
        'chkControlAll.Enabled = False
    End If
    If Me.chkControlAll.Value = Checked Then
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll", val(1)
    Else
        CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll", val(0)
    End If
    

Select Case Slider1.Value
Case "0"
modWinFunctions.SetPriorityClass GetCurrentProcess(), IDLE_PRIORITY_CLASS
Case "1"
modWinFunctions.SetPriorityClass GetCurrentProcess(), NORMAL_PRIORITY_CLASS
Case "2"
modWinFunctions.SetPriorityClass GetCurrentProcess(), HIGH_PRIORITY_CLASS
Case "3"
modWinFunctions.SetPriorityClass GetCurrentProcess(), REALTIME_PRIORITY_CLASS
End Select
   Debug.Print Slider1.Value
   
End Sub

Sub mcheck_startap2()
On Error GoTo 100
GetAllRun
Dim X As ListItem
Set X = Me.lvVirusFound.ListItems.Add(, , "Ïðîâåðêà ñòàðòàïîâ", 1, 2)
For i = 1 To lvStartUp.ListItems.Count - 1
Debug.Print lvStartUp.ListItems.Item(i).SubItems(1)
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
   If FileorFolderExists(Trim$(lvStartUp.ListItems.Item(i).SubItems(1))) = True Then
   
      modScanVirus.ScanFile (Trim$(lvStartUp.ListItems.Item(i).SubItems(1)))
   End If
Next i
Set X = Nothing
Exit Sub
100:
LogPrint "Îøèáêà ïðè ïðîâåðêå ñòàðòàïîâ-" + Error$

End Sub





Private Sub Form_Activate()
If Me.WindowState <> vbMinimized Then
Me.Top = 10
End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    ShowTopicID 1, 33
End If
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call ReleaseCapture
Call SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
'ïåðåäàåì äàííûå â îáúåêò
'      cTray.CallEvent X, Y

End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim strQuestion As String
Dim intAnswer As Integer
Dim aryMode As Variant

aryMode = Array(pr, "Âíèìàíèå", "Âíèìàíèå", pr, "Âíèìàíèå")
strQuestion = "Âû äåéñòâèòåëüíî õîòèòå âûéòè èç ïðîãðàììû ?"
intAnswer = MsgBox(strQuestion, vbQuestion + vbYesNo, aryMode(UnloadMode))

If intAnswer = vbNo Then Cancel = -1
End Sub

'========================================'
' FORM EVENTS                            '
'========================================'

Private Sub Form_Load()
On Error Resume Next
Me.Top = 10
lblLastUpdate.Caption = countZapOLD
If CLng(countZapOLD) = 0 Then
  
    MsgBox "Îòñóòñòâóþ áàçû.Ðàáîòà äàëüøå íå èìååò ñìûñëà", vbCritical
    End
   
End If
loadCapt
neg1 = 1
Me.Caption = "Belyash AntiTrojan 2008 beta v." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
'SetTopMostWindow Me.hwnd, True
'Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)

     frmMain.lblCount.Caption = procCountSpl
     frmMain.lblFound.Caption = procFoundSpl
     frmMain.lblCleaned.Caption = procCleanSpl
       SB.Refresh
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
     'SetTopMostWindow Me.hwnd, True
    'FileSize = 524288
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
   Me.SB.Caption = "Îæèäàíèå äåéñòâèé ïîëüçîâàòåëÿ"

If dbg1 = True Then
    Me.Check1.Value = Checked
    Me.Check2.Value = Checked
    Me.Check3.Value = Checked
    Me.Check4.Value = Checked
    Me.Check5.Value = Checked
    Me.Check6.Value = Checked
    Me.Check7.Value = Checked
    Call sbor_perem
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Me.MousePointer = ccArrow
End Sub













Private Sub Form_Resize()
loadCapt
End Sub

Private Sub frameScan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = ccArrow
End Sub



Private Sub frameSetting_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = ccArrow
End Sub






Private Sub frameTool_Process_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = ccArrow
End Sub



Private Sub frameTool_ScanReg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = ccArrow
End Sub



Private Sub frameTool_Startup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = ccArrow
End Sub

Private Sub franalize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = ccArrow
End Sub

Private Sub Image10_Click()
Call ShellExecute(0, "Open", "mailto:" + "mrbelyash@rambler.ru" + "?Subject=" + "Ïðîáëåìû ïðè ðàáîòå ñ Belyash AntiTrojan", "", "", 1)
End Sub

Private Sub Image11_Click()
ShowTopicID 1, 11
End Sub

Private Sub Image12_Click()
cmdOnlineUp_Click
End Sub



Private Sub Image9_Click()
Call ShellExecute(0, "Open", "http://mrbelyash.narod.ru/antivirus/belyashAV.htm", "", "", 1)
End Sub




Private Sub mnuCl2_Click()
Me.lvVirusFound.ListItems.Clear
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

    End
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
       ' Call ScanFileProc("c:\", frmMain.lvProcess.ListItems.Item(i).SubItems(1), Pro_ID)
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
    Call modScanVirus.ochistStatistic
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
        Timer2.Enabled = True
        Me.SB.Caption = "Çàïóùåíî ñêàíèðîâàíèå..."
        Me.Picture1.Refresh
            Dim X As ListItem
            Set X = Me.lvVirusFound.ListItems.Add(, , "Íà÷àëè â :", 1, 1)
            X.SubItems(1) = Time$
            'Set x = Nothing
           
            
If checkmemory = True Then
Set X = Me.lvVirusFound.ListItems.Add(, , "Ïðîâåðêà ïðîöåññîâ", 1, 2)
        Call smdScaning2 'ñêàíèðóþ ïàìÿòü
Else
    Set X = Me.lvVirusFound.ListItems.Add(, , "Ïðîâåðêà ïðîöåññîâ çàïðåùåíà êëþ÷¸ì êîììàíäíîé ñòðîêè :", 1, 1)
End If
 If nostartUP = True Then
        Call mcheck_startap2 'ñêàíèðóþ ñòàðòàï
Else
 Set X = Me.lvVirusFound.ListItems.Add(, , "Ïðîâåðêà ñòàðòàïîâ çàïðåùåíà êëþ÷¸ì êîììàíäíîé ñòðîêè :", 1, 1)
End If
  If noAU = True Then
        Call autorans_killNow 'óäàëÿþ àâòîðàíû
Else
    Set X = Me.lvVirusFound.ListItems.Add(, , "Ïðîâåðêà àâòîðàíîâ çàïðåùåíà êëþ÷¸ì êîììàíäíîé ñòðîêè :", 1, 1)
End If
   Set X = Me.lvVirusFound.ListItems.Add(, , "Ïðîâåðêà ðååñòðà", 1, 2)


Call modRegistry.CleanReg 'çàïóñòèë ÷èñòêó ðååñòðà
            Set X = Me.lvVirusFound.ListItems.Add(, , "Ñêàíèðóþ :", 1, 1)
            X.SubItems(1) = Me.txtPath.Text
            Set X = Nothing
         LogPrint "Scan-" + Trim$(Me.txtPath.Text)
        Me.txtPath.Text = Trim$(Me.txtPath.Text)
            'lvVirusFound.ListItems.Clear
            Dim ST As Variant, ET As Variant
            'get start time
            ST = Time
            'reset all statistic
            blnScan = True
            
            'Me.lblCount.Caption = 0
            'Me.lblFound.Caption = 0
            'Me.lblCleaned.Caption = 0
            Me.cmdStop.SetFocus
            ' strScanDetail = "<Font Name='Verdana' Size=3 Color=Blue>Scanning Íà÷àëè â : " & Time$ & "<br>Ñêàíèðóþ : <i>" & Me.txtPath.Text & "</i><br>-----------------------------<br></font>"
            ' Call UpdateDetail(strScanDetail, WebBrowser1)

            'start scanning
            
            Call ScanFile(Me.txtPath.Text)
            
            'finish scanning
            'get end time
            ET = Time
            'calculate time
            Dim X1 As String
            X1 = CalculateTime(ET - ST)
            ' strScanDetail = strScanDetail & "<font size=3 color=BLUE>-----------------------------<br>"
            If blnScan = True Then
                ' strScanDetail = strScanDetail & "Scanning Çàâåðøåíî â :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Çàêîí÷èëè â :", 1, 1)
                X.SubItems(1) = Time$
                'tray message
               ' Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå çàâåðøåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ: " & X1)
            Else
                ' strScanDetail = strScanDetail & "Scanning was Çàâåðøåíî â :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                X.SubItems(1) = Time$
                'MsgBox "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1
                'tray message
               ' Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå îñòàíîâëåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & X1)
            End If
            Set X = Nothing
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
       
            blnScan = False
            Me.lblPath.Caption = ""
        End If
    End If
    If cmdScan.Enabled = False Then
        cmdScan.Enabled = True
    End If
    SB.Caption = "Ñêàíèðîâàíèå çàâåðøåíî"
    Me.Picture1.Refresh
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
        SB.Caption = "Íà÷àëè ñêàíèðîâàíèå"
            lvVirusFound.ListItems.Clear
            Dim ST As Variant, ET As Variant
            'get start time
            ST = Time
            'reset all statistic
            blnScan = True
          
            'Me.lblCount.Caption = 0
            'Me.lblFound.Caption = 0
            'Me.lblCleaned.Caption = 0
            Me.cmdStop.SetFocus
            ' strScanDetail = "<Font Name='Verdana' Size=3 Color=Blue>Scanning Íà÷àëè â : " & Time$ & "<br>Ñêàíèðóþ : <i>" & Me.txtPath.Text & "</i><br>-----------------------------<br></font>"
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Dim X As ListItem
            Set X = Me.lvVirusFound.ListItems.Add(, , "Íà÷àëè â :", 1, 1)
            X.SubItems(1) = Time$
            Set X = Nothing
            Set X = Me.lvVirusFound.ListItems.Add(, , "Ñêàíèðóþ :", 1, 1)
            X.SubItems(1) = Me.txtPath.Text
            Set X = Nothing
            'start scanning
            Call ScanFile(Me.txtPath.Text)
            
            'finish scanning
            'get end time
            ET = Time
            'calculate time
            Dim X1 As String
            X1 = CalculateTime(ET - ST)
            ' strScanDetail = strScanDetail & "<font size=3 color=BLUE>-----------------------------<br>"
            If blnScan = True Then
                ' strScanDetail = strScanDetail & "Scanning Çàâåðøåíî â :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                X.SubItems(1) = Time$
                'tray message
               ' Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå çàâåðøåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & X1)
            Else
                ' strScanDetail = strScanDetail & "Scanning was Çàâåðøåíî â :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                X.SubItems(1) = Time$
                'MsgBox "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1
                'tray message
                'Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & X1)
            End If
            Set X = Nothing
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
          
            blnScan = False
            Me.lblPath.Caption = ""
        End If
    End If
    If cmdScan.Enabled = False Then
        cmdScan.Enabled = True
    End If
    
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
        SB.Caption = "Íà÷àëè ñêàíèðîâàíèå"
            'lvVirusFound.ListItems.Clear'çà÷èñòêà ïàíåëè
            Dim ST As Variant, ET As Variant
            'get start time
            ST = Time
            'reset all statistic
            blnScan = True
        
            'Me.lblCount.Caption = 0
            'Me.lblFound.Caption = 0
            'Me.lblCleaned.Caption = 0
            Me.cmdStop.SetFocus
            ' strScanDetail = "<Font Name='Verdana' Size=3 Color=Blue>Scanning Íà÷àëè â : " & Time$ & "<br>Ñêàíèðóþ : <i>" & Me.txtPath.Text & "</i><br>-----------------------------<br></font>"
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
            Dim X As ListItem
            Set X = Me.lvVirusFound.ListItems.Add(, , "Íà÷àëè â :", 1, 1)
            X.SubItems(1) = Time$
            Set X = Nothing
            Set X = Me.lvVirusFound.ListItems.Add(, , "Ñêàíèðóþ :", 1, 1)
            X.SubItems(1) = Me.txtPath.Text
            Set X = Nothing
            'start scanning
            Call ScanFile(Me.txtPath.Text)
            
            'finish scanning
            'get end time
            ET = Time
            'calculate time
            Dim X1 As String
            X1 = CalculateTime(ET - ST)
            ' strScanDetail = strScanDetail & "<font size=3 color=BLUE>-----------------------------<br>"
            If blnScan = True Then
                ' strScanDetail = strScanDetail & "Scanning Çàâåðøåíî â :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                X.SubItems(1) = Time$
                'tray message
               'Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå çàâåðøåíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1)
            Else
                ' strScanDetail = strScanDetail & "Scanning was Çàâåðøåíî â :" & Time$ & "</font>"
                Set X = Me.lvVirusFound.ListItems.Add(, , "Çàâåðøåíî â :", 1, 1)
                X.SubItems(1) = Time$
                'MsgBox "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1
                'tray message
               'Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Ñêàíèðîâàíèå ïðåðâàíî" & vbCrLf & "Âðåìÿ ñêàíèðîâàíèÿ " & x1)
            End If
            Set X = Nothing
            ' Call UpdateDetail(strScanDetail, WebBrowser1)
       
            blnScan = False
            Me.lblPath.Caption = ""
        End If
    End If
    If cmdScan.Enabled = False Then
        cmdScan.Enabled = True
    End If
    
Exit Sub
115:
MsgBox "" + Error$
End Sub
Private Sub cmdStop_Click()
    If blnScan = True Then
        If MsgBox("Îñòàíîâèòü ñêàíèðîâàíèå ?", vbYesNo + vbDefaultButton2 + vbExclamation, "Âíèìàíèå") = vbYes Then
         '  Me.mnuenblsc.Caption = "Ñêàíèðîâàòü Ìîé êîìïüþòåð"
                Timer2.Enabled = False
                Me.EviProgressBar2.Value = 0
                bw = 0
            StopMycompScan = True
            blnScan = False
                Me.txtPath.Text = ""
            SB.Caption = "Ïðåðâàíî ïîëüçîâàòåëåì"
             If cmdScan.Enabled = False Then
                cmdScan.Enabled = True
            End If
        End If
    End If
     StopMycompScan = False
     blnScan = False
    SB.Caption = "Ïðîâåðêà ïðåðâàíà ïîëüçîâàòåëåì..."
End Sub
Private Sub Timer2_Timer()
         DoEvents
If BB = False Then
DoEvents
    bw = bw + 1
    If bw > Me.EviProgressBar2.Max Then
        BB = True
    End If
Else
DoEvents
    bw = bw - 1
    If bw = 0 Then
        BB = False
    End If
End If
Me.EviProgressBar2.Value = bw
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'check scanning
    If blnScan = True Then
        If MsgBox("Ïðîèçâîäèòñÿ ñêàíèðîâàíèå. Âû äåéñòâèòåëüíî õîòèòå âûéòè ?", vbYesNoCancel + vbExclamation, "Ïðåðâàòü ñêàíèðîâàíèå?") = vbYes Then
            cReg.StopSearch
            cmdStop_Click
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
      If modScanVirus.ScanFileCMD("c:", programs.ProcessName(i), programs.ProcessHandle(i)) = False Then
      Me.SB.Caption = programs.ProcessName(i) + ":" + CStr(programs.ProcessHandle(i))
      
      Debug.Print "ïðîöåññû-" + programs.ProcessName(i) + "==" + CStr(programs.ProcessHandle(i))
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





Private Sub lvVirusFound_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lvVirusFound.MousePointer = ccArrow
End Sub

Private Sub Slider1_Change()
Select Case Slider1.Value
Case "0"

Slider1.ToolTipText = "Íèçêèé"
Case "1"
Slider1.ToolTipText = "Íîðìàëüíûé"
Case "2"
Slider1.ToolTipText = "Âûñîêèé"
Case "3"
Slider1.ToolTipText = "Ìàêñèìàëüíûé"

End Select
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
    'ñáîð ñâåäåíèé

    Me.rtf1.Text = ""
    Me.frameScan.Visible = False
    Me.frameAbout.Visible = False
    Me.frameSetting.Visible = False
   
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.franalize.Visible = True
    Me.franalize.Enabled = True
    Me.franalize.Refresh
    'enable scroll
    
    rtf1.Text = ""
End Sub
Private Sub imgAbout_Click()
'î ïðîãðàììå

    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False

    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.franalize.Visible = False
    Me.frameAbout.Visible = True
    
   
    
End Sub

Private Sub imgLicense_Click()
'ëèöåíçèÿ


    'change content
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False

    Me.frameTool.Visible = False
    Me.frameLicense.Visible = True
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
    
 Dim pCol As Long
  Dim pW As Long
  Dim pH As Long

  
  PictureLOGO2.cls
  pW = PictureLOGO1.Width
  pH = PictureLOGO1.Height
  For Y = 0 To pH - 1
    For X = 0 To pW - 1
      pCol = GetPixel(PictureLOGO1.hdc, X, Y)
      PictureLOGO2.Line (pW / 2, pH)-(X, Y), pCol * neg1
      SetPixel PictureLOGO2.hdc, X, Y, pCol
    Next X
    Sleep 5
    PictureLOGO2.Refresh

  Next Y
  
End Sub

Private Function Sleep(Delay As Long)
  Dim start As Long
  start = GetTickCount
  While (start + Delay) > GetTickCount
    DoEvents
  Wend
End Function
Private Sub imgScan_Click()
    
    'ñàíèðîâàíèå

    'change content
    Me.frameScan.Visible = True
    Me.frameSetting.Visible = False
  
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
End Sub

Private Sub imgSetting_Click()
    

    'íàñòðîéêè
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = True
   
    Me.frameTool.Visible = False
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
End Sub

Private Sub imgTools_Click()
    

    'èíñòðóìåíòû
    Me.frameScan.Visible = False
    Me.frameSetting.Visible = False
   
    Me.frameTool.Visible = True
    Me.frameLicense.Visible = False
    Me.frameAbout.Visible = False
    Me.franalize.Visible = False
    'disanable scroll
    'Me.tmrAbout.Enabled = False
        'Call RefreshDefList
End Sub




Private Sub lvVirusList_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Sub RefreshDefList()
  
        
    
    If Dir(App.Path + "\update.txt") <> "" Then
            Dim fso As New FileSystemObject, txtfile, fi
            Set fi = fso.GetFile(App.Path + "\update.txt")
        Me.lblLastUpdate.Caption = Format(fi.DateLastModified, "dd/mm/yy")
        'Me.SB.Panels(5).Text = Format(fi.DateLastModified, "dd mmmm yyyy")
        Set fi = Nothing
        Set fso = Nothing
    
    End If
    
   
   ' Me.lvVirusList.Refresh
End Sub







Private Sub tvFolders_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
tvFolders.MousePointer = ccArrow
End Sub

Private Sub tvFolders_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
   ' SaveSetting App.exeName, "Options", "Priority", Slider1.Value
  
Slider1.Value = val(GetSetting(App.exeName, "Options", "Priority", "2"))
    Me.txtSizeChk.Text = GetSetting(App.exeName, "Options", "chkSizeFile", "22478208")
   
    Me.mnuQarantine.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Quarantine"))
    'log
    Me.chkLog.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log"))
     Me.xpcheAutoran.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Autorans"))
    
  Me.xpcheckbox71.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanAllFiles"))
    
    
 '    strString = getstring(nm, "Software\BGAntivirus", "LogSize1")
'    txtSizelog.Text = strString
    'EXE File Block
    Me.chkExeBlock.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppBlock"))
    Me.chkControlAll.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll"))
    'check refresh rate
    txtRefreshRate.Text = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "RefreshRate")
    Me.tmrAutoRefresh.interval = txtRefreshRate.Text * 1000
    'check auto refresh status
    Me.chkAutoRefeshProList.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AutoRefresh"))
    

    Me.chkAutoScan.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanReg"))
    
'    If chkAutoRefeshProList.Value = 1 Then
'        Me.tmrAutoRefresh.Enabled = True
'    Else
'        Me.tmrAutoRefresh.Enabled = False
'    End If
    
    'process block
    Me.chkBlockRisk.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk"))
    'start minimized
    'If Me.chkStartMin.Value = 1 Then
       ' Call HideToTray
'        Hide
        'õåíäë îêíà
'      cTray.hwnd = hwnd
   'èêîíêà, ÷òî áóäåò îòîáðàæåíà â òðåå
'      cTray.Icon = Icon
   'òóëòèïñ (âñïëûâàþùàÿ ïîäñêàçêà)
'      cTray.ToolTipText = "Belyash Antitrojan 2008"
      
   'ñîçäàåì èêîíêó
'      cTray.Add
   
   ' End If
    'Me.chkAutoScan.Value = ""
    'check options
    intSettingRegOption = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ScanRegOption"))
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
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "ControlAll", val(Me.chkControlAll.Value)
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
    SB.Caption = "Äåëàþ êîïèþ ðååñòðà...Ïîäîæäèòå"
        create_Backup "1"
         create_Backup "2"
          create_Backup "3"
           create_Backup "4"
            create_Backup "5"
    SB.Caption = "Êîïèÿ ðååñòðà ñîçäàíà"
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
    Dim X As ListItem
    For i = 1 To programs.processCount - 1
       Set X = Me.lvProcess.ListItems.Add(, , programs.ProcessName(i), 7)
        X.SubItems(1) = programs.ProcessName(i)
        X.SubItems(2) = programs.ProcessHandle(i)
     'Me.lvProcess.ListItems.Add programs.ProcessName(i)
  '  lvProcess.ListItems.Add programs.ProcessHandle(i)
    Next i
    Set X = Nothing
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
   ' Call ShowTrayMessage("Belyash AntiTrojan 2008 beta", "Cleaning Registry Errors completed and backup. Cleaned " & removed & " of " & Me.lvErrorRegKey.ListItems.Count & " .")
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
    Dim X As ListItem
    
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
                Set X = .ListItems.Add(, , KN, 5, 5)
                X.SubItems(1) = sRootKey
                X.SubItems(2) = sKey
                X.SubItems(3) = sValue
            End With
            Set X = Nothing
            'add to counter
            Me.lblScanRegError.Caption = Int(Me.lblScanRegError.Caption) + 1
        Else    'scan specific extension
            'MsgBox FileorPath
            'MsgBox Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)
            'MsgBox InStr(1, LCase(Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3)), LCase(strScanRegExt))
            'If InStr(1, Right$(FileorPath, 3), strScanRegExt, vbTextCompare) > 0 Then    'found in extension
            If InStr(1, strScanRegExt, Mid(FileorPath, InStr(1, FileorPath, ".") + 1, 3), vbTextCompare) > 0 Then    'found in extension
                With Me.lvErrorRegKey
                    Set X = .ListItems.Add(, , KN, 5, 5)
                    X.SubItems(1) = sRootKey
                    X.SubItems(2) = sKey
                    X.SubItems(3) = sValue
                End With
                Set X = Nothing
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
    intBlockRisk = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk"))
    
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
    'Me.lvProcessDetail.ListItems.Clear
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



Private Sub cmdProcessRefresh_Click()
    'clear items from all listviews
    Me.lvProcess.ListItems.Clear
    Me.lvProcessDetail.ListItems.Clear
    'refresh data
    Call programs.CheckProcesses
    Dim i As Integer
    Dim X As ListItem
    For i = 1 To programs.processCount - 1
    
    Set X = Me.lvProcess.ListItems.Add(, , programs.ProcessName(i), 7)
        X.SubItems(1) = programs.ProcessName(i)
        X.SubItems(2) = programs.ProcessHandle(i)
     'Me.lvProcess.ListItems.Add programs.ProcessName(i)
  '  lvProcess.ListItems.Add programs.ProcessHandle(i)
    Next i
    Set X = Nothing
    CheckProcess
    ModuleA.DEnumWindows
    Me.lblTotalProcess.Caption = programs.processCount - 1
End Sub

Private Sub tmrAutoRefresh_Timer()
    Call cmdProcessRefresh_Click
    If val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "BlockRisk")) = 1 Then
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
    Dim X As ListItem, hKey As Long, lCount As Long, i As Long
    lvStartUp.ListItems.Clear
    'Enumerate from HKEY_LOCAL_MACHINE , Run
    hKey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        Set X = Nothing
    Next i
    
    'Enumerate from HKEY_LOCAL_MACHINE , RunServices
    hKey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\RunServices")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        X.SubItems(2) = "HKEY_LOCAL_MACHINE (Service)"
        Set X = Nothing
    Next i
    
    'Enumerate from HKEY_CURRENT_USER , Run
    hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run")
    lCount = GetCount(hKey, Values)
    For i = 0 To lCount - 1
        Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        X.SubItems(2) = "HKEY_CURRENT_USER"
        Set X = Nothing
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
                Set X = Me.lvStartUp.ListItems.Add(, , sFile.Name)
                X.SubItems(1) = sFile.Path
                X.SubItems(2) = "Tasks"
                Set X = Nothing
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
                Set X = Me.lvStartUp.ListItems.Add(, , sFile.Name)
                X.SubItems(1) = sFile.Path
                X.SubItems(2) = "User Startup"
                Set X = Nothing
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
                 Set X = Me.lvStartUp.ListItems.Add(, , sFile.Name)
                 X.SubItems(1) = sFile.Path
                X.SubItems(2) = "All User Startup"
                Set X = Nothing
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
                    'Me.chkAutoStart.Value = 0    'delete autostart
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
    Me.chkCUTaskmgr.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
    Me.chkCUNoLogoff.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff"))
    Me.chkCUNoCLose.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose"))
    Me.chkCUNoLock.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation"))
    Me.chkCUNoChangePwd.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword"))
    'Me.chkCURegTool.Value = CBool(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    'Me.chkCUNoCmd.Value = CBool(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD"))
    'Me.chkCUNoFolderOption.Value = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))

    'Local Machine
'    Me.chkLMNoFolderOption.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
'    Me.chkLMRegTool.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
'    Me.chkLMNoTaskmgr.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
'    Me.chkLMNoSRConfig.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"))
'    Me.chkLMNoSR.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR"))
'    Me.chkLmLimitSR.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"))
'    Me.chkLMNoMSI.Value = Val(ReadValue(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"))
    Me.chkLMNoFolderOption.Value = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"))
    'Me.chkLMRegTool.Value = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools"))
    Me.chkLMNoTaskmgr.Value = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr"))
    Me.chkLMNoSRConfig.Value = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig"))
    Me.chkLMNoSR.Value = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR"))
    Me.chkLmLimitSR.Value = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing"))
    Me.chkLMNoMSI.Value = val(getstring(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI"))

End Sub

Private Sub chkCUNoChangePwd_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableChangePassword", val(Me.chkCUNoChangePwd.Value)
End Sub

Private Sub chkCUNoCLose_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoClose", val(Me.chkCUNoCLose.Value)
End Sub

Private Sub chkCUNoCmd_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableCMD", val(Me.chkCUNoCmd.Value)
End Sub

Private Sub chkCUNoFolderOption_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", val(Me.chkCUNoFolderOption.Value)
End Sub

Private Sub chkCUNoLock_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableLockWorkstation", val(Me.chkCUNoLock.Value)
End Sub

Private Sub chkCUNoLogoff_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\Explorer", "NoLogoff", val(Me.chkCUNoLogoff.Value)
End Sub

Private Sub chkCURegTool_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", val(Me.chkCURegTool.Value)
    
End Sub

Private Sub chkCUTaskmgr_Click()
    CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", val(Me.chkCUTaskmgr.Value)
End Sub

Private Sub chkLmLimitSR_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", val(Me.chkLmLimitSR.Value)
End Sub

Private Sub chkLMNoFolderOption_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions", val(Me.chkLMNoFolderOption.Value)
End Sub

Private Sub chkLMNoMSI_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", val(Me.chkLMNoMSI.Value)
End Sub

Private Sub chkLMNoSR_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", val(Me.chkLMNoSR.Value)
End Sub

Private Sub chkLMNoSRConfig_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", val(Me.chkLMNoSRConfig.Value)
End Sub

Private Sub chkLMNoTaskmgr_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "software\microsoft\windows\currentversion\policies\system", "DisableTaskMgr", val(Me.chkLMNoTaskmgr.Value)
End Sub

Private Sub chkLMRegTool_Click()
    CreateDwordValue GetClassKey("HKEY_LOCAL_MACHINE"), "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", val(Me.chkLMRegTool.Value)
   
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

sbor_perem
End Sub
Public Sub sbor_perem()
On Error Resume Next
If Dir$(App.Path + "\analize.log", vbNormal) <> "" Then
    Kill App.Path + "\analize.log"
End If
Dim n As Integer
n = 0
rtf1.Text = ""
If Check1.Value = vbChecked Then
    SB.Caption = "Îáðàáàòûâàþ ïåðåìåííûå îêðóæåíèÿ...Æäèòå..."
    peremen
    n = n + 1
End If

If Check3.Value = vbChecked Then
    SB.Caption = "Îáðàáàòûâàþ BHO"
   check_BHO
    n = n + 1
End If


If Check4.Value = vbChecked Then
    SB.Caption = "Îáðàáàòûâàþ àêòèâíûå îêíà...Æäèòå..."
    ModuleA.Ac_Wind
    n = n + 1
End If
If Check2.Value = vbChecked Then
    SB.Caption = "Îáðàáàòûâàþ àêòèâíûå ïðîöåññû...Æäèòå..."
    poci12
    n = n + 1
End If
If Check5.Value = vbChecked Then
    'ýêñïîðò ðåãèñòðà
    SB.Caption = "Îáðàáàòûâàþ ðååñòð ...Æäèòå..."
        create_Backup "1"
         create_Backup "2"
          create_Backup "3"
           create_Backup "4"
            create_Backup "5"
            n = n + 1
End If
If Check7.Value = vbChecked Then
     SB.Caption = "Ñ÷èòûâàþ íàñòðîéêè ...Æäèòå..."
    GetAllKeys1
    n = n + 1
End If

If Check6.Value = vbChecked Then
     SB.Caption = "Îáðàáàòûâàþ àâòîçàãðóçêó ...Æäèòå..."
    startup_log
    n = n + 1
End If
If n > 0 Then
    GoTo 9
Else
    SB.Caption = "Ëîã íå ñîçäàí.Âûáåðèòå îïöèè"
    GoTo 800
   
End If
9:

Dim dgFile As Integer
dgFile = FreeFile
Open App.Path + "\analize.log" For Append As #dgFile
Print #dgFile, rtf1.Text
Close #1
     SB.Caption = "Ëîã ñîçäàí è ñîõðàíåí â ôàéëå analize.log"
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
Dim X As ListItem, hKey As Long, lCount As Long, i As Long
analize "-----------------Options programm start-------------------"
Dim apppa As String
Dim RegExt As String

   
   analize "Check file size maximum=" & GetSetting(App.exeName, "Options", "chkSizeFile", "22478208")
    
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
        analize Trim$(CStr(EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", EnumValue(hKey, i)))) + vbCrLf
        
        analize GetKeyValue(hKey, EnumValue(hKey, i)) + vbCrLf
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
Dim X As ListItem, hKey As Long, lCount As Long, i As Long
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
       
        Set X = Nothing
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
   
        Set X = Nothing
        'frmEditAppBlock.lvAppList.ListItems.Add EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", EnumValue(hKey, i))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i

' "-------------Options programm end----------------"
frmEditAppBlock.Show
End Function

Sub analize(dmesd As String)
'ëîãèì
'analiz_fil (dmesd) + vbCrLf
'If Trim$(dmesd) <> "" Then
    frmMain.rtf1.Text = frmMain.rtf1.Text + Trim$(dmesd) + vbCrLf
'End If
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
        tmp = WindowFromPoint(pt.X, pt.Y)
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

    '  cTray.hwnd = Me.hwnd
   'èêîíêà, ÷òî áóäåò îòîáðàæåíà â òðåå
'      cTray.Icon = Icon
   'òóëòèïñ (âñïëûâàþùàÿ ïîäñêàçêà)
'      cTray.ToolTipText = "Belyash Antitrojan 2008"
      
   'ñîçäàåì èêîíêó
'    cTray.Add
   'Me.Hide
End Sub

Sub loadCapt()
imgScan.Caption = "Ñêàíèðîâàíèå"
imgSetting.Caption = "Íàñòðîéêè"
imgTools.Caption = "Èíñòðóìåíòû"
Image7.Caption = "Àíàëèç"
imgLicense.Caption = "Ëèöåíçèÿ"
imgAbout.Caption = "Ïîìîùü"
cmdCleanReg.Caption = "Ïî óìîë÷àíèþ"
cmdClearAutorun.Caption = "Óäàëèòü àâòîðàíû"
cmdScanInvalidReg.Caption = "Èñïðàâèòü ðååñòð"
cmdTool_ProcessMan.Caption = "Ïðîöåññû"
cmdTool_Startup.Caption = "Àâòîçàãðóçêà"
cmdEnableReg.Caption = "Ðååñòð"
cmdProcessRefresh.Caption = "Îáíîâèòü"
cmdProcessEnd.Caption = "Çàâåðøèòü"
cmdStartStop.Caption = "Ñòàðò"
cmdDeleteInvalidKey.Caption = "Óäàëèòü"
cmdStartUp_Del.Caption = "Óäàëèòü"

xpcheckbox71.Caption = "Ñêàíèðîâàòü âñå ôàéëû"
mnuQarantine.Caption = "Ïåðåìåùàòü â êàðàíòèí"
chkLog.Caption = "Îãðàíè÷èòü îò÷åò"
chkAutoScan.Caption = "Èñïðàâëÿòü ðååñòð"
xpcheAutoran.Caption = "Çà÷èñòêà àâòîðàíîâ"
chkBlockRisk.Caption = "Áëîêèðîâàòü ïîäîçðèòåëüíûå"
chkAutoRefeshProList.Caption = "Àâòîîáíîâëåíèå ïðîöåññîâ"
chkExeBlock.Caption = "Âêëþ÷èòü áëîêèðàòîð ïðèëîæåíèé"
chkControlAll.Caption = "Êîíòðîëèðîâàòü çàïóñê âñåõ ïðîãðàìì"
optScanAll.Caption = "Âñå"
optScanExt.Caption = "Óêàçàíûå ðàñøèðåíèÿ"
cmdRestoreDefault.Caption = "Ïî óìîë÷àíèþ"
cmdSave.Caption = "Ñîõðàíèòü"
Command1.Caption = "Ñîõðàíèòü"
Check1.Caption = "Ïåðåìåííûå îêðóæåíèÿ"
Check2.Caption = "Àêòèâíûå ïðîöåññû"
Check3.Caption = "Íàäñòðîéêè IE"
Check4.Caption = "Àêòèâíûå îêíà"
Check5.Caption = "Ýêñïîðòèðîâàòü ðååñòð"
Check6.Caption = "Àâòîçàãðóçêà"
Check7.Caption = "Íàñòðîéêè"
cmdScan.Caption = "Ñòàðò"
cmdStop.Caption = "Ñòîï"
lblLabel2.Caption = "Îãðàíè÷èòü ðàçìåð ïðîâåðÿåìîãî ôàéëà                                  áàéò"
End Sub

Sub check_BHO()
'÷èòàåì BHO äëÿ IE
    On Error Resume Next
    List4.Clear
    analize "-------------BHO----------------"

    GetRegKeys &H80000002, "SOFTWARE\Microsoft\Internet Explorer\Extensions", List4
    Dim z As Integer
   LogPrint "--------Start BHO---------"
    For z = 0 To List4.ListCount - 1
    Dim hKey As Long, lCount As Long, i As Long
    hKey = OpenKey(GetClassKey("HKEY_LOCAL_MACHINE"), "SOFTWARE\Microsoft\Internet Explorer\Extensions\" + List4.list(z))
    lCount = GetCount(hKey, Values)
    LogPrint "-----" + CStr(z + 1) + "------"
    analize "-----" + CStr(z + 1) + "------"
    For i = 0 To lCount - 1
    Dim d1 As String
    Dim d2 As String
    d1 = EnumValue(hKey, i) + "="
    d2 = GetKeyValue(hKey, EnumValue(hKey, i))
         '  Debug.Print EnumValue(hKey, i) + "=",
        'Debug.Print GetKeyValue(hKey, EnumValue(hKey, i))
        Dim lngPos As Integer
        lngPos = InStr(d1, Chr(0))
        If lngPos > 0 Then
            d1 = Left$(d1, lngPos - 1)
        End If
      Dim lngPos1 As Integer
        lngPos1 = InStr(d2, Chr(0))
        If lngPos1 > 0 Then
            d2 = Left$(d2, lngPos1 - 1)
        End If
        LogPrint d1 + "=" + d2
        analize d1 + "=" + d2
        Debug.Print d1 + "=" + d2
        'Set x = Nothing
    Next i
    
Next z
LogPrint "--------End BHO---------"
End Sub



Private Sub txtSizeChk_Validate(Cancel As Boolean)
If Not IsNumeric(txtSizeChk) Then MsgBox "Ââîäèòå òîëüêî ÷èñëà", vbCritical
End Sub


