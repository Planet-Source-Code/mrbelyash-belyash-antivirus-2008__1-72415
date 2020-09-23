VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmNet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Belyash Shield"
   ClientHeight    =   7710
   ClientLeft      =   150
   ClientTop       =   240
   ClientWidth     =   10170
   Icon            =   "FormNet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCheck 
      Interval        =   100
      Left            =   180
      Top             =   3090
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7455
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10186
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ñåòåâîé ìîíèòîð"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Ôàéëîâûé ìîíèòîð"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Êîë-âî çàïèñåé"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pic16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   660
      ScaleHeight     =   420
      ScaleWidth      =   540
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   3810
   End
   Begin VB.Frame Frame2 
      Height          =   6045
      Left            =   45
      TabIndex        =   2
      Top             =   1125
      Width           =   10065
      Begin VB.Timer tmrProcRef 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   315
         Top             =   1560
      End
      Begin VB.Frame Frame3 
         Height          =   5025
         Left            =   -45
         TabIndex        =   14
         Top             =   990
         Width           =   10050
         Begin ComctlLib.ListView ListView1 
            Height          =   4425
            Left            =   -30
            TabIndex        =   38
            Top             =   120
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   7805
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
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   8
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
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Process"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   7
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Direction"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Timer TimerDir 
            Enabled         =   0   'False
            Interval        =   200
            Left            =   870
            Top             =   690
         End
         Begin VB.PictureBox PicQuestion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   2700
            Picture         =   "FormNet.frx":038A
            ScaleHeight     =   255
            ScaleWidth      =   255
            TabIndex        =   39
            Top             =   3450
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.ListBox lstBox 
            Height          =   1230
            Left            =   1920
            TabIndex        =   40
            Top             =   2550
            Visible         =   0   'False
            Width           =   4575
         End
         Begin VB.ListBox lvProcess 
            Height          =   1230
            Left            =   390
            TabIndex        =   41
            Top             =   2700
            Visible         =   0   'False
            Width           =   945
         End
      End
      Begin VB.CommandButton Command9 
         Height          =   795
         Left            =   5340
         Picture         =   "FormNet.frx":06CC
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Çàêðûòü âñå ïîðòû"
         Top             =   180
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.CommandButton Command8 
         Height          =   795
         Left            =   9180
         Picture         =   "FormNet.frx":242E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Äåòàëè"
         Top             =   150
         Width           =   765
      End
      Begin VB.CommandButton Command7 
         Height          =   795
         Left            =   4350
         Picture         =   "FormNet.frx":28AE
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Âêëþ÷èòü ìîíèòîð"
         Top             =   150
         Width           =   765
      End
      Begin VB.CommandButton Command6 
         Height          =   795
         Left            =   3510
         Picture         =   "FormNet.frx":2F4B
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Áëîêåð ïðèëîæåíèé"
         Top             =   150
         Width           =   765
      End
      Begin VB.CommandButton Command5 
         Height          =   795
         Left            =   2610
         Picture         =   "FormNet.frx":3485
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Îáíîâèòü"
         Top             =   150
         Width           =   765
      End
      Begin VB.CommandButton Command4 
         Height          =   795
         Left            =   1770
         Picture         =   "FormNet.frx":39C5
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Ñîîáùèòü îá îøèáêå"
         Top             =   150
         Width           =   765
      End
      Begin VB.CommandButton Command3 
         Height          =   795
         Left            =   900
         Picture         =   "FormNet.frx":3D9B
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Äåòàëè"
         Top             =   150
         Width           =   765
      End
      Begin VB.CommandButton Command2 
         Height          =   795
         Left            =   60
         Picture         =   "FormNet.frx":42D9
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Âêëþ÷èòü Net Shield"
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   1950
         TabIndex        =   42
         Top             =   1320
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   4035
         Left            =   3420
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   4245
         Begin VB.CommandButton cmdMon 
            Caption         =   "Start Monitor"
            Height          =   1305
            Left            =   3180
            TabIndex        =   53
            Top             =   2070
            Width           =   795
         End
         Begin VB.TextBox txtURL 
            Height          =   285
            Left            =   1110
            TabIndex        =   10
            Top             =   630
            Width           =   2715
         End
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   1080
            TabIndex        =   9
            Top             =   270
            Width           =   2715
         End
         Begin VB.TextBox txtMonitor 
            Height          =   285
            Left            =   1080
            TabIndex        =   8
            Top             =   990
            Width           =   2775
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Firefox"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   7
            Top             =   1350
            Width           =   855
         End
         Begin VB.CheckBox Chk 
            Caption         =   "IE"
            Height          =   255
            Index           =   1
            Left            =   3270
            TabIndex        =   6
            Top             =   1350
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.ListBox lstURL 
            Height          =   1950
            IntegralHeight  =   0   'False
            Left            =   120
            TabIndex        =   5
            Top             =   1740
            Width           =   2925
         End
         Begin VB.CheckBox MonitorOpera 
            Caption         =   "Opera"
            Height          =   195
            Left            =   1350
            TabIndex        =   4
            Top             =   1350
            Width           =   795
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            Caption         =   "Title"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   13
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            Caption         =   "URL"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   630
            Width           =   855
         End
         Begin VB.Label Lbl 
            Alignment       =   2  'Center
            Caption         =   "Monitor"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   11
            Top             =   990
            Width           =   855
         End
      End
      Begin ComctlLib.ImageList ImageList1 
         Left            =   7230
         Top             =   210
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   327682
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4260
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4710
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.Frame Frame6 
      Height          =   5445
      Left            =   45
      TabIndex        =   15
      Top             =   1215
      Width           =   10050
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Text            =   "1024"
         Top             =   3240
         Width           =   975
      End
      Begin monitor.xpradiobutton Option2 
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   3240
         Width           =   2715
         _ExtentX        =   4789
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
      End
      Begin monitor.xpradiobutton Option1 
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   2970
         Width           =   3015
         _ExtentX        =   5318
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
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   255
         Left            =   210
         TabIndex        =   19
         Top             =   2070
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   450
         _Version        =   327682
         BorderStyle     =   1
         LargeChange     =   1
         Max             =   3
      End
      Begin monitor.xpcheckbox lbAuto 
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   600
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
      End
      Begin monitor.xpcheckbox lbFon 
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   900
         Width           =   2595
         _ExtentX        =   4577
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
      End
      Begin monitor.xpcheckbox Check3 
         Height          =   375
         Left            =   3630
         TabIndex        =   22
         Top             =   540
         Width           =   2475
         _ExtentX        =   4366
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
      Begin monitor.xpcheckbox ldMonik 
         Height          =   255
         Left            =   3630
         TabIndex        =   23
         Top             =   945
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   450
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
      Begin monitor.xpcheckbox chkBlock 
         Height          =   375
         Left            =   3630
         TabIndex        =   24
         Top             =   1230
         Width           =   3015
         _ExtentX        =   5318
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
      Begin monitor.xpcheckbox Check6 
         Height          =   345
         Left            =   120
         TabIndex        =   25
         Top             =   1230
         Width           =   1905
         _ExtentX        =   3360
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
      Begin monitor.xpcheckbox chkDisk1 
         Height          =   345
         Left            =   3630
         TabIndex        =   26
         Top             =   1980
         Width           =   2865
         _ExtentX        =   5054
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
      Begin monitor.xpcheckbox Check5 
         Height          =   315
         Left            =   3630
         TabIndex        =   27
         Top             =   2340
         Width           =   2805
         _ExtentX        =   4948
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
      End
      Begin monitor.xpcheckbox Check4 
         Height          =   285
         Left            =   3630
         TabIndex        =   28
         Top             =   1590
         Width           =   2055
         _ExtentX        =   3625
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
      End
      Begin monitor.xpcheckbox Check1 
         Height          =   285
         Left            =   90
         TabIndex        =   29
         Top             =   2640
         Width           =   2865
         _ExtentX        =   5054
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
      End
      Begin monitor.xpcmdbutton xpcmdbutton1 
         Height          =   315
         Left            =   8400
         TabIndex        =   30
         Top             =   4440
         Width           =   1455
         _ExtentX        =   2566
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
      End
      Begin monitor.xpcheckbox chkmonSyte 
         Height          =   405
         Left            =   7050
         TabIndex        =   35
         Top             =   960
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
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
      Begin monitor.xpcheckbox chkTrojan 
         Height          =   405
         Left            =   7050
         TabIndex        =   36
         Top             =   1350
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   714
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
      Begin monitor.xpcheckbox chkmonNet 
         Height          =   375
         Left            =   7065
         TabIndex        =   37
         Top             =   630
         Width           =   2895
         _ExtentX        =   5106
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
      Begin VB.Line Line5 
         X1              =   180
         X2              =   9990
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   6780
         X2              =   6780
         Y1              =   150
         Y2              =   4170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         X1              =   3480
         X2              =   3480
         Y1              =   150
         Y2              =   4230
      End
      Begin VB.Label Label5 
         Caption         =   "NET ìîíèòîð"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7770
         TabIndex        =   34
         Top             =   180
         Width           =   1755
      End
      Begin VB.Shape Shape1 
         Height          =   825
         Left            =   120
         Top             =   1650
         Width           =   2655
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000011&
         X1              =   6750
         X2              =   6750
         Y1              =   120
         Y2              =   4200
      End
      Begin VB.Label Label2 
         Caption         =   "Ôàéëîâûé ìîíèòîð"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3810
         TabIndex        =   33
         Top             =   150
         Width           =   2265
      End
      Begin VB.Label Label4 
         Caption         =   "Îáùèå"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   32
         Top             =   150
         Width           =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   3510
         X2              =   3510
         Y1              =   150
         Y2              =   4200
      End
      Begin VB.Label Label3 
         Caption         =   "Ïðèîðèòåò"
         Height          =   255
         Left            =   930
         TabIndex        =   31
         Top             =   1740
         Width           =   885
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Íàñòðîéêè"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   5880
      TabIndex        =   51
      Top             =   750
      Width           =   3675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Ãëàâíàÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1890
      TabIndex        =   50
      Top             =   750
      Width           =   3675
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   0
      Picture         =   "FormNet.frx":472E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10125
   End
   Begin ComctlLib.ImageList Iml16 
      Left            =   120
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu Menu 
      Caption         =   "Ìåíþ"
      Visible         =   0   'False
      Begin VB.Menu mnuSh 
         Caption         =   "Ïîêàçàòü"
      End
      Begin VB.Menu l7 
         Caption         =   "-"
      End
      Begin VB.Menu monFile 
         Caption         =   "Âêëþ÷èòü ìîíèòîð"
      End
      Begin VB.Menu ntsh 
         Caption         =   "Âêë Net Shield"
      End
      Begin VB.Menu blckpr 
         Caption         =   "Áëîêåð ïðèëîæåíèé"
      End
      Begin VB.Menu abpc 
         Caption         =   "Î ïðîãðàììå..."
      End
      Begin VB.Menu upd_TR 
         Caption         =   "Îáíîâèòü"
      End
      Begin VB.Menu RefreshIT 
         Caption         =   "Îáíîâèòü"
         Visible         =   0   'False
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
Attribute VB_Name = "FrmNet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(100) As Long
Public countZapOLD As Long
Public countZap As Long


Public baze_File As String
Dim arr_hndls(0) As Long
Dim path(0) As String
'===========================
  ' instance of the watch object
  'Private WithEvents m_oWatchDir As cWatchForChanges
  Dim hChangeHandle As Long
Dim hWatched As Long
Dim terminateFlag As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
    lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
'---------------------------------------------------------------------------------------
' CONSTANTS
'---------------------------------------------------------------------------------------
'Constants for topmost.
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long



Private Const TIME_OUT = &H102
  Private Const HKEY_CLASSES_ROOT As Long = &H80000000
  Private Const HKEY_CURRENT_USER As Long = &H80000001
  Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
  Private Const HKEY_USERS As Long = &H80000003

Private Const INFINITE As Long = &HFFFFFFFF

Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const KEY_READ As Long = KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY

Private Const ERROR_SUCCESS As Long = 0&

Private Const REG_NOTIFY_CHANGE_ATTRIBUTES As Long = &H2
Private Const REG_NOTIFY_CHANGE_LAST_SET As Long = &H4
Private Const REG_NOTIFY_CHANGE_NAME As Long = &H1
Private Const REG_NOTIFY_CHANGE_SECURITY As Long = &H8
Private Const REG_NOTIFY_CHANGE_ALL As Long = &H8 Or &H1 Or &H2 Or &H4

Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const LANG_NEUTRAL As Long = &H0

Private Const MONITOR_TOPKEY  As Long = HKEY_CURRENT_USER
Private Const MONITOR_SUBKEY  As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"

'---------------------------------------------------------------------------------------
' APIS
'---------------------------------------------------------------------------------------
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal samDesired As Long, phkResult As Long) As Long

'---------------------------------------------------------------------------------------
' MEMBER VARIABLES
'---------------------------------------------------------------------------------------
Private f_lChange As Long
Private f_lKey As Long

'----------------
Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)

Public newAplic As String
Public Allow As Boolean
 ' Private m_sWatchedDir As String
  
  'Private m_udtFT As FILETIME
  
  ' array to track files in the watched directory
 ' Private m_audtDirContents() As WIN32_FIND_DATA
  
  ' instance of the watch object
 'Private WithEvents m_oWatchDir As cWatchForChanges
  


Public namePort As String
Public nameBlockSyte As String
Public nameBlockSyteTwo As String
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Public Processing As Boolean
Public ShowListening As Boolean
Private WithEvents URLMon As clsURLMon
Attribute URLMon.VB_VarHelpID = -1
Private Const RAS_MAXENTRYNAME As Integer = 256
Private Const RAS_MAXDEVICETYPE As Integer = 16
Private Const RAS_MAXDEVICENAME As Integer = 128
Private Const RAS_RASCONNSIZE As Integer = 412
'Private Const ERROR_SUCCESS = 0&
Private Type RasEntryName
dwSize As Long
szEntryName(RAS_MAXENTRYNAME) As Byte
End Type
Private Type RasConn
dwSize As Long
hRasConn As Long
szEntryName(RAS_MAXENTRYNAME) As Byte
szDeviceType(RAS_MAXDEVICETYPE) As Byte
szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type
Private Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long
Private Declare Function RasHangUp Lib "rasapi32.dll" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long
Private gstrISPName As String
Public ReturnCode As Long
'Public WithEvents cTray As mdlMain
Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeoutAndVersion As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
  guidItem As GUID
End Type


'Private WithEvents sysTray As SystemTray.Application


Public Sub HangUp()
Dim i As Long
Dim lpRasConn(255) As RasConn
Dim lpcb As Long
Dim lpcConnections As Long
Dim hRasConn As Long
lpRasConn(0).dwSize = RAS_RASCONNSIZE
lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
lpcConnections = 0
ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, lpcConnections)
If ReturnCode = ERROR_SUCCESS Then
For i = 0 To lpcConnections - 1
If Trim(ByteToString(lpRasConn(i).szEntryName)) = Trim(gstrISPName) Then
hRasConn = lpRasConn(i).hRasConn
ReturnCode = RasHangUp(ByVal hRasConn)
End If
Next i
End If
End Sub

Public Function ByteToString(bytString() As Byte) As String
Dim i As Integer
ByteToString = ""
i = 0
While bytString(i) = 0&
ByteToString = ByteToString & Chr(bytString(i))
i = i + 1
Wend
End Function


Private Sub abpc_Click()
frmAbout.Show
End Sub

Private Sub blckpr_Click()
frmBlocked.Show
End Sub

Private Sub Chk_Click(Index As Integer)
    Select Case Index
        Case 0 ' firefox
            URLMon.MonitorFireFox = Chk(0).Value
        Case 1 ' internet explorer
            URLMon.MonitorIE = Chk(1).Value

            
            
    End Select
End Sub



Private Sub cmdMon_Click()

Dim z2 As String
z2 = CStr(GetSetting("Belyash AV", "Options", "Monsyte", "True"))
If z2 = "False" Then
    Exit Sub
End If

'ìîíèòîðèíã url
Dim f As Integer
DoEvents
FillProcessListNT
For f = 0 To Me.lstBox.ListCount - 1

syteGlaw = Spliting(Connection(i).FileName, "\")
Select Case syteGlaw
Case "IEXPLORE.EXE"
   Chk(1).Value = Checked
   Exit For
Case "firefox.exe"
 '  Chk(0).Value = Checked
   'Exit For
Case "opera.exe"
   MonitorOpera.Value = Checked
  
Exit For
End Select
Next f
  '  Debug.Print "==============" + syteGlaw
    If URLMon.Active Then
        lstURL.Clear
        URLMon.StopMon
        'cmdMon.Caption = "Start Monitor"
    Else
        URLMon.StartMon txtMonitor
        'cmdMon.Caption = "Stop Monitor"
        lstURL.Clear
    End If
End Sub







Public Sub message(snumer As Integer, dmesd As String)
    'None = 0
    'Information = 1
    'Warning = 2
    'Critical = 3
Select Case snumer
    Case 0
        ShellBalloonShow Me.hwnd, none, "Belyash Shield v." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision), dmesd
    Case 1
        ShellBalloonShow Me.hwnd, Information, "Belyash Shield v." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision), dmesd
    Case 2
    ShellBalloonShow Me.hwnd, Warning, "Belyash Shield v." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision), dmesd
    Case 3
    ShellBalloonShow Me.hwnd, Critical, "Belyash Shield v." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision), dmesd
End Select
LogPrint dmesd

End Sub







Private Sub Command1_Click()
Frame1.Visible = True
End Sub



Private Sub Command2_Click()
If Command2.ToolTipText = "Âêëþ÷èòü Net Shield" Then
   Me.Timer1.Enabled = True
    Command2.ToolTipText = "Âûêëþ÷èòü Net Shield"
    ntsh.Caption = "Âûêëþ÷èòü Net Shield"
  '  URLMon.Active = True
    cmdMon_Click
    SaveSetting "Belyash AV", "Options", "MonNET", "False"
       Me.chkmonNet.Value = Unchecked
      ntsh.Caption = "Âûêëþ÷èòü Net Shield"
     Me.StatusBar.Panels(2).Text = "OFF"
     Me.StatusBar.Refresh
Else
'âêëþ÷åíî
Command2.ToolTipText = "Âêëþ÷èòü Net Shield"
  'monFile.Caption = "Âêëþ÷èòü Net Shield"
 Me.Timer1.Enabled = False
 ListView1.ListItems.Clear
'URLMon.Active = False
cmdMon_Click
'URLMon.StartMon txtMonitor
   Me.chkmonNet.Value = Checked
   ntsh.Caption = "Âêëþ÷èòü Net Shield"
     SaveSetting "Belyash AV", "Options", "MonNET", "True"
     Me.StatusBar.Panels(2).Text = "ON"
     Me.StatusBar.Refresh
        

End If
End Sub

Private Sub Command3_Click()
FrmStats.Show
End Sub

Private Sub Command4_Click()
On Error GoTo 100
Call ShellExecute(0, "Open", "mailto:" + "mrbelyash@rambler.ru" + "?Subject=" + "Ïðîáëåìû ïðè ðàáîòå ñ Belyash Anti-Trojan", "", "", 1)
Exit Sub
100:
End Sub

Private Sub Command5_Click()
onn1
End Sub

Private Sub Command6_Click()
frmBlocked.Show
End Sub

Private Sub Command7_Click()
'âêëþ÷èòü ìîíèòîðèíã

Call setMonitor
FillProcessListNT
End Sub

Private Sub Command8_Click()
ShowTopicID 1, 11
End Sub

Private Sub Command9_Click()

BlockInput True
End Sub

Private Sub Form_Load()
      
If App.PrevInstance = True Then
 End
 End If
 modScanVirus.selfTesting
 sbor
 StatusBar.Panels(4).Text = countZapOLD
 If countZapOLD = 0 Then
    MsgBox "Îòñóòñòâóþò áàçû", vbCritical
    End
 End If
 '=========
 baze_File = App.path + "\baze.txt"
 Me.Caption = "Belyash Shield v." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " b"
 Call ShellTrayAdd(Me.hwnd, Me.Icon, "Belyash Shield v." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision))
 le_refresh

  path(0) = Environ$("systemroot")
  'path(1) = Environ$("systemroot") + "\system32"
  'path(2) = Environ$("systemroot") + "\system32\drivers"
 'Me.Caption = "Belyash Shield v." + CStr(App.Major) + vbCrLf + CStr(App.Minor) + vbCrLf + CStr(App.Revision)
 Dim lResult As Long
    
    lResult = RegOpenKeyEx(MONITOR_TOPKEY, MONITOR_SUBKEY, 0, KEY_READ, f_lKey)

    If Not (lResult = ERROR_SUCCESS) Then
        DisplayError "RegOpenKeyEx", lResult
    End If
    
    If f_lChange Then
        mRegNotifyChange.RegFindCloseChange f_lChange
    End If
     
    f_lChange = mRegNotifyChange.RegFindFirstChange(f_lKey, True, REG_NOTIFY_CHANGE_ALL)
    
    If f_lChange = 0 Then
        DisplayError "RegFindFirstChange", GetLastError
    End If
   
 '-------
   'Me.Hide
  If tmrON = True Then
    Me.StatusBar.Panels(3).Text = "ON"
   Else
    Me.StatusBar.Panels(3).Text = "OFF"
   End If
   
    noList = False
    tmrON = False
    keyOn = 0
    monitorOn = False
    firstRun = True
    refProc = True
    unloadOK = False
    logOn = True
    logNew = False
    protectOpt = False
    protectAccess = False
    showGo = False
    hotkeyPrompt = False
    taskmgrFrozen = False
    tempAccPass = False
    protectPass = ""
    prevIndex = 1
    prevCapt = "Process Home"
    'Me.BackColor = 14078416
    ReDim procinfo(150) As PROCESSENTRY32
    ReDim jailInfo(1) As jailedProc

    Call enumProc
    firstRun = False
    Set URLMon = New clsURLMon
    'Chk(0).Value = 1
   ' URLMon.MonitorFireFox = True
   
    
    Chk(1).Value = 1
   URLMon.MonitorIE = True
   URLMon.MonitorOpera = True
    MonitorOpera.Value = 1
     'centers form
   ' Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

    'sets cbSize to the Length of TrayIcon
 
   'õåíäë îêíà
    
   'èêîíêà, ÷òî áóäåò îòîáðàæåíà â òðåå
   
   'òóëòèïñ (âñïëûâàþùàÿ ïîäñêàçêà)

      
   'ñîçäàåì èêîíêó

   Call LoadRegistry
    Dim a2 As String
a2 = CStr(GetSetting("Belyash AV", "Options", "Priority", "2"))
Select Case a2
Case "1"
modThread.SetPriorityClass GetCurrentProcess(), IDLE_PRIORITY_CLASS
Case "2"
modThread.SetPriorityClass GetCurrentProcess(), NORMAL_PRIORITY_CLASS
Case "3"
modThread.SetPriorityClass GetCurrentProcess(), HIGH_PRIORITY_CLASS
Case "4"
modThread.SetPriorityClass GetCurrentProcess(), REALTIME_PRIORITY_CLASS
Case Else
 modThread.SetPriorityClass GetCurrentProcess(), HIGH_PRIORITY_CLASS
End Select

Dim fn As String
     fn = CStr(GetSetting("Belyash AV", "Options", "StartFon", "False"))
If fn = "True" Then
    Me.Hide
Else
    Me.Show

End If

Dim n1 As Boolean
n1 = CStr(GetSetting("Belyash AV", "Options", "Monitoring", "True"))
 If n1 = "True" Then
    Call setMonitor
    FillProcessListNT 'âêëþ÷àåì ìîíèòîðèíã
   ' ModuleS.registry_mon 'ìîíèòîðèì ðååñòð
    'Call cmdStartWatch 'âêëþ÷èòü ñëåæåíèå çà êàòàëîãîì âèíäû
 Else
    tmrON = False
   ' Call setMonitor
    'FillProcessListNT
End If

Dim nz As Boolean
nz = CStr(GetSetting("Belyash AV", "Options", "MonNET", "True"))
 If nz = "True" Then
    Command2.ToolTipText = "Âûêëþ÷èòü Net Shield"
    ntsh.Caption = "Âûêëþ÷èòü Net Shield"
    Me.Timer1.Enabled = True
    Processing = False
   ' cmdMon_Click
    StatusBar.Panels(2).Text = "ON"
   
Else
Command2.ToolTipText = "Âêëþ÷èòü Net Shield"
ntsh.Caption = "Âêëþ÷èòü Net Shield"
Me.Timer1.Enabled = False
cmdMon_Click
Processing = True
StatusBar.Panels(2).Text = "OFF"

  End If


Dim nz7 As Boolean
nz7 = CStr(GetSetting("Belyash AV", "Options", "MonReg", "True"))
 If nz = "True" Then
    tmrCheck.Enabled = True
End If
End Sub

Sub LoadRegistry()
Dim a1 As String
a1 = CStr(GetSetting("Belyash AV", "Options", "AutoStartMon", "True"))
If a1 = "True" Then
    Me.lbAuto.Value = Checked
Else
    Me.lbAuto.Value = Unchecked
End If


Dim z1 As String
z1 = CStr(GetSetting("Belyash AV", "Options", "MonNET", "True"))
If z1 = "True" Then
    Me.chkmonNet.Value = Checked
Else
    Me.chkmonNet.Value = Unchecked
End If


Dim z2 As String
z2 = CStr(GetSetting("Belyash AV", "Options", "Monsyte", "True"))
If z2 = "True" Then
    Me.chkmonSyte.Value = Checked
Else
    Me.chkmonSyte.Value = Unchecked
End If


Dim z3 As String
z3 = CStr(GetSetting("Belyash AV", "Options", "MonTrojan", "True"))
If z3 = "True" Then
    Me.chkTrojan.Value = Checked
Else
    Me.chkTrojan.Value = Unchecked
End If

Dim a2 As String
a2 = CStr(GetSetting("Belyash AV", "Options", "StartFon", "True"))
If a2 = "True" Then
    Me.lbFon.Value = Checked
Else
    Me.lbFon.Value = Unchecked
End If

Dim a29 As String
a29 = CStr(GetSetting("Belyash AV", "Options", "Block", "True"))
If a29 = "True" Then
    Me.chkBlock.Value = Checked
Else
    Me.chkBlock.Value = Unchecked
End If





Dim a3 As String
a3 = CStr(GetSetting("Belyash AV", "Options", "MonQv", "True"))
If a3 = "True" Then
    Me.Check3.Value = Checked
Else
    Me.Check3.Value = Unchecked
End If

Dim a4 As String
a4 = CStr(GetSetting("Belyash AV", "Options", "LogMon", "True"))
If a4 = "True" Then
    Me.Check1.Value = Checked
Else
    Me.Check1.Value = Unchecked
End If

Dim a5 As String
a5 = CStr(GetSetting("Belyash AV", "Options", "MonReg", "True"))
If a5 = "True" Then
    Me.Check4.Value = Checked
Else
    Me.Check4.Value = Unchecked
End If

Dim a781 As String
a781 = CStr(GetSetting("Belyash AV", "Options", "Notification", "True"))
If a781 = "True" Then
    Me.Check6.Value = Checked
Else
    Me.Check6.Value = Unchecked
End If


Dim a6 As String
a6 = CStr(GetSetting("Belyash AV", "Options", "MonFolder", "True"))
If a6 = "True" Then
    Me.Check5.Value = Checked
Else
    Me.Check5.Value = Unchecked
End If

Dim a61 As String
a61 = CStr(GetSetting("Belyash AV", "Options", "MonDiskActiv", "True"))
If a61 = "True" Then
    Me.chkDisk1.Value = Checked
Else
    Me.chkDisk1.Value = Unchecked
End If

    Dim n1 As Boolean
     n1 = CBool(GetSetting("Belyash AV", "Options", "Monitoring", "True"))
 If n1 = True Then
    Me.ldMonik.Value = Checked
 Else
    Me.ldMonik.Value = Unchecked
 End If
     
 Dim a8 As String
a8 = CStr(GetSetting("Belyash AV", "Options", "AutoStartMon", "True"))
If a8 = "True" Then
   Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BelyashAV", App.path & "\" & App.exeName & ".exe")
Else
     DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", "BelyashAV"
End If

     
Dim a9 As String
a9 = CStr(GetSetting("Belyash AV", "Options", "LogAppend", "True"))
If a9 = "True" Then
         Option1.Value = True
         Option2.Value = False
         Me.Text2.Enabled = True
Else
         Option2.Value = True
         Option1.Value = False
         Me.Text2.Enabled = False
End If
     
Me.Text2.Text = Val(GetSetting("Belyash AV", "Options", "LogSizeMon", "1024496"))

Slider1.Value = Val(GetSetting("Belyash AV", "Options", "Priority", "2"))

Select Case Slider1.Value
Case "0"

modThread.SetPriorityClass GetCurrentProcess(), IDLE_PRIORITY_CLASS
Case "1"
modThread.SetPriorityClass GetCurrentProcess(), NORMAL_PRIORITY_CLASS
Case "2"
modThread.SetPriorityClass GetCurrentProcess(), HIGH_PRIORITY_CLASS
Case "3"
modThread.SetPriorityClass GetCurrentProcess(), REALTIME_PRIORITY_CLASS
End Select
   Debug.Print Slider1.Value
  
End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim strQuestion As String
Dim intAnswer As Integer
Dim aryMode As Variant

aryMode = Array(, , , , "")
strQuestion = "Âû äåéñòâèòåëüíî õîòèòå çàêðûòü ïðèëîæåíèå ?"
intAnswer = MsgBox(strQuestion, vbQuestion + vbYesNo, aryMode(UnloadMode))

If intAnswer = vbNo Then
    Me.Hide
    Cancel = -1
Else
xpcmdbutton1_Click
StopNotify
  ShellTrayRemove
UnSubClass Me.hwnd
FrmNet.tmrCheck.Enabled = False
FrmNet.Timer1.Enabled = False
'cmdMon_Click

End If

  
 
End Sub


Private Sub Label1_Click()
Frame6.Visible = False

Frame2.Visible = True
Me.Image1.Picture = LoadPicture(App.path + "\s1.bmp")
Image1.Refresh
End Sub
Sub le_refresh()
lbAuto.Caption = "Àâòîçàãðóçêà"
lbFon.Caption = "Çàãðóçêà â ôîíîâîì ðåæèìå"
Check6.Caption = "Âûäàâàòü òóëòèï"
Check1.Caption = "Îãðàíè÷èòü ðàçìåð ëîãà"
Option1.Caption = "Ïåðåçàïèñûâàòü"
Option2.Caption = "Îáðåçàòü"
Check3.Caption = "Ïåðåìåùàòü â êàðàíòèí"
ldMonik.Caption = "Çàïîìèíàòü ñîñòîÿíèå"
chkBlock.Caption = "Áëîêåð ïðèëîæåíèé"
Check4.Caption = "Ìîíèòîðèòü ðååñòð"
chkDisk1.Caption = "Ìîíèòîðèòü äèñêîâóþ àêòèâíîñòü"
Check5.Caption = "Ìîíèòîðèòü ñèñòåìíûé êàòàëîã"
chkmonNet.Caption = "Çàïîìèíàòü ñîñòîÿíèå"
chkmonSyte.Caption = "Ìîíèòîðèòü îïàñíûå ñàéòû"
chkTrojan.Caption = "Ìîíèòîðèòü òðîÿíñêèå ïîðòû"
xpcmdbutton1.Caption = "OK"
FrmNet.Refresh

End Sub
Private Sub Label6_Click()
le_refresh


Frame2.Visible = False
Frame6.Visible = True
Me.Image1.Picture = LoadPicture(App.path + "\s2.bmp")
Image1.Refresh
End Sub

Private Sub lstBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
If Button = vbRightButton Then
    For i = 0 To lstBox.ListCount - 1
        If lstBox.Selected(i) Then
                
        End If
    Next i
End If
End Sub

Private Sub mnuSh_Click()
           
            
            'Delete our icon from the tray.
            'Shell_NotifyIconA NIM_DELETE, nidTray
            'ZoomForm ZOOM_FROM_TRAY, FrmNet.hwnd
            DoEvents
            FrmNet.Show
            DoEvents
End Sub

Private Sub monFile_Click()
setMonitor
End Sub

Private Sub ntsh_Click()
Command2_Click
End Sub

Private Sub Option2_Click()
Text2.Enabled = True
End Sub

Private Sub upd_TR_Click()
Command5_Click
End Sub

Private Sub URLMon_TitleChanged(Title As String)
    txtTitle = Title
End Sub

Private Sub URLMon_URLChanged(URL As String)

If yesBlockSyte(URL) = True Then
    FrmNet.message 1, "Íàéäåí çàáëîêèðîâàííûé ñàéò" + vbCrLf + nameBlockSyte
End If

    If lstURL.ListCount >= 20 Then
       ' MsgBox "1"
        lstURL.Clear
    End If
    
    Dim i As Integer
     txtURL = URL
   
        lstURL.AddItem URL
     
    
    lstURL.ListIndex = lstURL.ListCount - 1
End Sub
Public Sub RefreshList()
 On Error GoTo 500
  Dim i
  Dim Item As ListItem
  Dim syteGlaw As String
  Dim arrport As String
  Processing = True
  Dim contrlCRC As String
Dim mi_namepr As String

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
    If Connection(i).LocalPort = Connection(i).RemotePort And Connection(i).LocalPort <> "" Then Item.SubItems(7) = "Incomming" Else Item.SubItems(7) = "Outgoing"
    Item.SubItems(1) = Connection(i).LocalPort
    Item.SubItems(2) = Connection(i).RemoteHost
   
    Item.SubItems(3) = Connection(i).RemotePort
    Item.SubItems(4) = c_state(Connection(i).State)
        Item.SubItems(6) = Connection(i).ProcessID
        Item.Tag = i
         contrlCRC = modFileManipulation.GetMD5(CStr(Connection(i).FileName))

       
Dim z43 As String
z43 = CStr(GetSetting("Belyash AV", "Options", "MonTrojan", "True"))
If z43 = "False" Then
    GoTo 222
End If

       arrport = CStr((Connection(i).LocalPort))
    Dim myport As String
    myport = Chr(34) + arrport + Chr(34)
        If yesVirZ(myport) = True Then
              If namePort <> "" Then
                message 1, "Îáíàðóæåíî ñîåäèíåíèå íà òðîÿíñêèé ïîðò: " + Connection(i).LocalPort + vbCrLf + "Òðîÿí: " + namePort + vbCrLf + "Ïðèëîæåíèå: " + CStr(ListView1.ListItems(i))
                   
                   ' TerminateThisConnection (ListView1.SelectedItem.Tag)
                    'Exit For
                 ' End If
                End If
        End If
222:
   
   If trustApl(contrlCRC) = False Then
        mi_namepr = Connection(i).FileName
        If Trim$(mi_namepr) = "" Then
            GoTo 600
   End If
     modThread.Thread_Suspend (Connection(i).ProcessID)
     Allow = False
    frmAlert.Label4.Caption = Connection(i).ProcessID
    frmAlert.Label3.Caption = mi_namepr
    frmAlert.Label5.Caption = i
   newAplic = Connection(i).FileName
    'frmAlert.Resize
    frmAlert.Show
    End If



600:
IsListening:
Next i

GetAllIcons
DoEvents

ShowIcons
DoEvents

Processing = False
Exit Sub
500:


End Sub
Function BlockedProg(pr77 As String) As Boolean
'On Error GoTo 10
FrmNet.Data1.DatabaseName = App.path + "\inetbase.mdb"
FrmNet.Data1.RecordSource = "Blocked"
FrmNet.Data1.Refresh


BlockedProg = False
FrmNet.Data1.Recordset.FindFirst "crc = '" _
& Trim(pr77) & "'"
If FrmNet.Data1.Recordset.NoMatch Then
'MsgBox "Èìÿ íå íàéäåíî"
BlockedProg = False
Else
'MsgBox "Íàøåë====" + FrmNet.Data1.Recordset("name") + "====" + FrmNet.Data1.Recordset("CRC")

    BlockedProg = True
Exit Function

BlockedProg = False
End If

Exit Function
10:
MsgBox "" + Error$ + CStr(Err)
Debug.Print Error$
End Function

Function trustApl(CZX As String) As Boolean
'On Error GoTo 10
FrmNet.Data1.DatabaseName = App.path + "\inetbase.mdb"
FrmNet.Data1.RecordSource = "applic"
FrmNet.Data1.Refresh


trustApl = False
FrmNet.Data1.Recordset.FindFirst "crc = '" _
& Trim(CZX) & "'"
If FrmNet.Data1.Recordset.NoMatch Then
'MsgBox "Èìÿ íå íàéäåíî"
trustApl = False
Else
'MsgBox "Íàøåë====" + FrmNet.Data1.Recordset("name") + "====" + FrmNet.Data1.Recordset("CRC")

    trustApl = True
Exit Function

trustApl = False
End If

Exit Function
10:
MsgBox "" + Error$ + CStr(Err)
Debug.Print Error$
End Function
Function yesVirZ(CFc As String) As Boolean
'On Error GoTo 10
FrmNet.Data1.DatabaseName = App.path + "\inetbase.mdb"
FrmNet.Data1.RecordSource = "portTable"
FrmNet.Data1.Refresh


yesVirZ = False
FrmNet.Data1.Recordset.FindFirst "port = '" _
& Trim(CFc) & "'"
If FrmNet.Data1.Recordset.NoMatch Then
'MsgBox "Èìÿ íå íàéäåíî"
yesVirZ = False
Else
'MsgBox "Íàøåë====" + FrmNet.Data1.Recordset("name") + "====" + FrmNet.Data1.Recordset("CRC")
namePort = FrmNet.Data1.Recordset("opisan")
    yesVirZ = True
Exit Function

yesVirZ = False
End If

Exit Function
10:
MsgBox "" + Error$ + CStr(Err)
Debug.Print Error$
End Function
Function yesBlockSyte(CFx As String) As Boolean
'On Error GoTo 10
Dim baze_File1 As String
   baze_File1 = App.path + "\inet.txt"
'On Error GoTo 10
yesBlockSyte = False
    Dim miNumBase As Integer
    Dim sMD5 As String
    Dim zb As String
    Dim bm As Integer
    miNumBase = FreeFile
     Open baze_File1 For Input As #miNumBase
        While Not EOF(miNumBase)
             Line Input #miNumBase, sMD5
             
            bm = InStr(1, sMD5, CFx, vbTextCompare)
    If bm <> 0 Then
        
        yesBlockSyte = True
        nameBlockSyte = CFx
       'MsgBox "" + CStr(CFx)
        Close #miNumBase1
        Exit Function
     End If
     Wend
10:

   yesBlockSyte = False
  Close #miNumBase1

End Function
Function yesBlockSyteTwo(CDFx As String) As Boolean
'On Error GoTo 10
'On Error GoTo 10
Dim baze_File1 As String
   baze_File1 = App.path + "\inet.txt"
'On Error GoTo 10
yesBlockSyteTwo = False
    Dim miNumBase As Integer
    Dim sMD5 As String
    Dim zb As String
    Dim bm As Integer
    miNumBase = FreeFile
     Open baze_File1 For Input As #miNumBase
        While Not EOF(miNumBase)
             Line Input #miNumBase, sMD5
             
            bm = InStr(1, sMD5, CFx, vbTextCompare)
    If bm <> 0 Then
        
        yesBlockSyteTwo = True
        nameBlockSyteTwo = CFx
       'MsgBox "" + CStr(CFx)
        Close #miNumBase1
        Exit Function
     End If
     Wend
10:

   yesBlockSyteTwo = False
  Close #miNumBase1





End Function
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
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, Pic16.hdc, 0, 0, ILD_TRANSPARENT)
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
On Error GoTo 100
If TerminateThisConnection(ListView1.SelectedItem.Tag) = True Then
StatusBar.Panels(1).Text = "Connection Closed: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
Else
StatusBar.Panels(1).Text = "Close Connection Failed: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
End If
'Debug.Print ""; ListView1.SelectedItem.Tag
100:
End Sub

Private Sub ExitProg_Click()
Unload Me
'End
End Sub
Private Function Spliting(sFullPath As String, point As String)
On Error GoTo 10
Dim str1() As String
str1 = Split(sFullPath, point)
Spliting = str1(UBound(str1))
Exit Function
10:
Spliting = "Íåèçâåñòíîå"
End Function
Private Sub Form_Resize()
On Error Resume Next
'ListView1.Width = Me.Width - 150
ListView1.Left = 100
'ListView1.Height = Me.Height - 3050

ListView1.ColumnHeaders(1).Width = 2300
ListView1.ColumnHeaders(2).Width = 800
ListView1.ColumnHeaders(3).Width = 2100
ListView1.ColumnHeaders(4).Width = 800
ListView1.ColumnHeaders(5).Width = 1600
ListView1.ColumnHeaders(6).Width = ListView1.Width \ 2 + 800

'StatusBar.Panels(1).Width = Me.Width / 4
'StatusBar.Panels(2).Width = Me.Width / 3
'StatusBar.Panels(3).Width = Me.Width / 5
'StatusBar.Panels(4).Width = Me.Width / 6

Progbar.Width = StatusBar1.Panels(1).Width
Progbar.Height = StatusBar1.Height - 80
   'If Me.WindowState = vbMinimized Then
    'Command9_Click
   'End If
If Me.WindowState <> vbMinimized Then
    If Me.Width <> 10260 Then
        Me.Width = 10260
    End If


    If Me.Height <> 8190 Then
        Me.Height = 8190
    End If


End If
'Me.Caption = "Belyash Shield v." + CStr(App.Major) + vbCrLf + CStr(App.Minor) + vbCrLf + CStr(App.Revision)
14:
 
End Sub

Private Sub ICMPBtn_Click()
On Error GoTo 100
Dim RemoteHostNet
RemoteHostNet = Connection(ListView1.SelectedItem.Tag).RemoteHost
StatusBar.Panels(1).Text = "(ICMP) " & RemoteHostNet & " : " & Ping(RemoteHostNet, 2000)
DoEvents
100:

End Sub

Private Sub ListeningPortsBtn_Click()
On Error Resume Next
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

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Select Case Button

    Case vbLeftButton
    'MsgBox ListView1.SelectedItem.Key

    Case vbRightButton
    CloseCon.Caption = "Ïðåðâàòü ñîåäèíåíèå: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
    TermProg.Caption = "Óáèòü: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
    Me.PopupMenu MenuPop

End Select
End Sub

Private Sub RefreshIT_Click()
'Command5_Click
RefreshList
End Sub

Private Sub ResolveHost_Click()
On Error GoTo label100
Dim HostAddr
Dim hostname

hostname = ResolveHostname(Connection(ListView1.SelectedItem.Tag).RemoteHost)
DoEvents

If hostname = "" Then
StatusBar.Panels(1).Text = "Íå ìîãó äîñòó÷àòüñÿ äî õîñòà"
Exit Sub
End If

HostAddr = Connection(ListView1.SelectedItem.Tag).RemoteHost

StatusBar.Panels(1).Text = hostname & " (" & HostAddr & ")"
label100:

End Sub

Private Sub ShowStats_Click()
FrmStats.Show
End Sub

Private Sub TermProg_Click()
On Error GoTo 200
If KillProcessById(Connection(ListView1.SelectedItem.Tag).ProcessID) = True Then
StatusBar.Panels(1).Text = "Óáèò ïðîöåññ: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
Else
StatusBar.Panels(1).Text = "Îøèáêà òåðìèíèðîâàíèÿ: " & Right(Connection(ListView1.SelectedItem.Tag).FileName, Len(Connection(ListView1.SelectedItem.Tag).FileName) - InStrRev(Connection(ListView1.SelectedItem.Tag).FileName, "\"))
End If
200:
End Sub

Private Sub Timer1_Timer()
If Processing = True Then Exit Sub

If IsNetConnectOnline = False Then
StatusBar.Panels(2).Text = "OFF"

Exit Sub
Else
StatusBar.Panels(2).Text = "ON"

End If

If GetRefresh = True Then RefreshList

End Sub

Private Sub xpcmdbutton1_Click()
'çàïèñàòü íàñòðîéêè
Label1_Click
If Me.Text2.Text = "" Then
MsgBox "Âû íå óêàçàëè ðàçìåð îò÷åòà", vbCritical, pr
Exit Sub
End If
If Me.chkmonNet.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "MonNET", "True"
Else
SaveSetting "Belyash AV", "Options", "MonNET", "false"
End If

If Me.chkmonSyte.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "Monsyte", "True"
Else
SaveSetting "Belyash AV", "Options", "Monsyte", "false"
End If


If Me.chkTrojan.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "MonTrojan", "True"
Else
SaveSetting "Belyash AV", "Options", "MonTrojan", "False"
End If





If Me.ldMonik.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "Monitoring", "True"
Else
SaveSetting "Belyash AV", "Options", "Monitoring", "false"
End If
If Me.chkBlock.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "Block", "True"
Else
SaveSetting "Belyash AV", "Options", "Block", "false"
End If
 If Me.Check4.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "MonReg", "True"
Else
    SaveSetting "Belyash AV", "Options", "MonReg", "false"
End If


If Me.Check6.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "Notification", "True"
Else
    SaveSetting "Belyash AV", "Options", "Notification", "false"
End If


If Me.Check3.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "MonQv", "True"
Else
    SaveSetting "Belyash AV", "Options", "MonQv", "false"
End If
 
 If Me.lbFon.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "StartFon", "True"
Else
    SaveSetting "Belyash AV", "Options", "StartFon", "false"
End If
 If Me.lbAuto.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "AutoStartMon", "True"
Else
    SaveSetting "Belyash AV", "Options", "AutoStartMon", "false"
End If
  If Me.Check5.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "MonFolder", "True"
Else
    SaveSetting "Belyash AV", "Options", "MonFolder", "False"
End If
  If Me.chkDisk1.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "MonDiskActiv", "True"
Else
    SaveSetting "Belyash AV", "Options", "MonDiskActiv", "False"
End If



    If Me.lbAuto.Value = 1 Then
        Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BelyashAV", App.path & "\" & App.exeName & ".exe")
        'CreateStringValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", 1, "BGAntivirus", App.Path & "\" & App.EXEName & ".exe"
    Else
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\Microsoft\Windows\CurrentVersion\Run", "BelyashAV"
     End If

  If Me.Check1.Value = Checked Then
   SaveSetting "Belyash AV", "Options", "LogMon", "True"
Else
    SaveSetting "Belyash AV", "Options", "LogMon", "false"
End If

If Me.Check4.Value = vbUnchecked Then
    ModuleDop.sn89
End If



If Option1.Value = True Then
    'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogAppend", Val("1")
    SaveSetting "Belyash AV", "Options", "LogAppend", Val("1")
    
Else
SaveSetting "Belyash AV", "Options", "LogAppend", Val("0")
'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogAppend", Val("0")
'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogSizeMon", Val(Me.Text2.Text)
SaveSetting "Belyash AV", "Options", "LogSizeMon", Val(Me.Text2.Text)
End If
SaveSetting "Belyash AV", "Options", "Priority", Slider1.Value

Select Case Slider1.Value
Case "0"
modThread.SetPriorityClass GetCurrentProcess(), IDLE_PRIORITY_CLASS
Case "1"
modThread.SetPriorityClass GetCurrentProcess(), NORMAL_PRIORITY_CLASS
Case "2"
modThread.SetPriorityClass GetCurrentProcess(), HIGH_PRIORITY_CLASS
Case "3"
modThread.SetPriorityClass GetCurrentProcess(), REALTIME_PRIORITY_CLASS
End Select
   Debug.Print Slider1.Value




205:

End Sub


'======================================
'###########################################################3
'ìîíèòîð
'++++++++++++++++++++++++++
Private Sub tmrProcRef_Timer()
ReDim procinfo(150) As PROCESSENTRY32
    If refProc = True Then
        lvProcess.Clear
        Call enumProc
        refProc = False
    Else
        Call enumProc
    End If
End Sub


Private Sub setMonitor()

    If tmrON = True Then
        tmrProcRef.Enabled = False
        Me.StatusBar.Panels(3).Text = "OFF"
'        lblMonitor.ForeColor = &HC0&
        Command7.ToolTipText = "Âêëþ÷èòü ìîíèòîð"
        
        monFile.Caption = "Âêëþ÷èòü ìîíèòîð"
        ' Frame3.Visible = True
         'Frame4.Visible = False
        tmrON = False
        If tmrCheck.Enabled = True Then
        tmrCheck.Enabled = False
        End If
         StopNotify
        

        
    Else
        monitorOn = True
        tmrProcRef.Enabled = True
        Call uc7Interval_Clicked
         Me.StatusBar.Panels(3).Text = "ON"
        start_regMon 'âêëþ÷àåì ñëæåíèå çà ðååñòðîì
        Command7.ToolTipText = "Âûêëþ÷èòü ìîíèòîð"
        monFile.Caption = "Âûêëþ÷èòü ìîíèòîð"
    '     Frame4.Visible = True
     '   Frame3.Visible = False
       
        monFolder
        'cmdBegin 'íàáëþäàåì çà êàòàëîãîì âèíäû
        tmrON = True
    End If

End Sub
Sub monFolder()
On Error GoTo 10
Dim a76 As String
a76 = CStr(GetSetting("Belyash AV", "Options", "MonFolder", "False"))
If a76 = "False" Then
    'MsgBox ""
    Exit Sub
End If
  'message ""
  arr_hndls(0) = FindFirstChangeNotification _
    (path(0), True, FILE_NOTIFY_CHANGE_FILE_NAME)
  If arr_hndls(0) = INVALID_HANDLE_VALUE Then message 3, ("Error #1!")
  'arr_hndls(1) = FindFirstChangeNotification _
  '  (path(1), True, FILE_NOTIFY_CHANGE_FILE_NAME Or FILE_NOTIFY_CHANGE_SIZE)
  'If arr_hndls(1) = INVALID_HANDLE_VALUE Then message ("Error #1!")
  
  'arr_hndls(2) = FindFirstChangeNotification _
  '  (path(2), True, FILE_NOTIFY_CHANGE_FILE_NAME Or FILE_NOTIFY_CHANGE_SIZE)
  'If arr_hndls(2) = INVALID_HANDLE_VALUE Then message ("Error #1!")
  TimerDir.Enabled = True
  Exit Sub
10:
End Sub
Private Sub TimerDir_Timer()

Dim Status As Long, retVal As Long

  Status = WaitForMultipleObjects(1, arr_hndls(0), False, 100)
  Select Case Status
    Case WAIT_TIMEOUT
    Case 0 'To 2
      retVal = FindNextChangeNotification(arr_hndls(Status))
      If retVal = False Then message 3, "Error#3!"
      
      message 2, "Îáíàðóæåíû èçìåíåíèÿ â êàòàëîãå " + vbCrLf & path(Status) & _
        vbCrLf & Time & vbCrLf
    Case WAIT_FAILED
      MsgBox "Error: wait failed!"
    Case Else
      MsgBox "Error: unknow error!"
  End Select
End Sub
Private Sub StopNotify()
Dim i

  TimerDir.Enabled = False
 ' For i = 0 To UBound(arr_hndls)
    FindCloseChangeNotification arr_hndls(0)
 ' Next
End Sub

Private Sub uc7Interval_Clicked()
    If txtProcRef = "" Then
        tmrON = True
'        Call setMonitor
    ElseIf Trim(txtProcRef) = 0 Then
        'tmrProcRef.Interval = 250
    Else
        'tmrProcRef.Interval = txtProcRef.Text * 1000
    End If
End Sub









Public Sub Alert(dmess As String)
frmAlert.Label2.Caption = dmess
frmAlert.Resize
frmAlert.Show
End Sub
Sub onn1()

On Error GoTo ErrorMessage
If Dir$(App.path + "\updater.exe", vbNormal) = "" Then
    MsgBox "Â êàòàëîãå ñ ïðîãðàììîé îòñóòñòâóåò ìîäóëü îáíîâëåíèÿ." + vbCrLf + " Ïåðåóñòàíîâèòå ïðîãðàììó", vbCritical, pr
    Exit Sub
Else
    If FrmNet.WindowState <> vbMinimized Then
        
        FrmNet.WindowState = vbMinimized
        'WindowState = vbMinimized
    End If
    ExecCmd (App.path + "\updater.exe")
        
End If
'Call modBinaryFile.ReadSig
'Call frmTest.RefreshDefList
Exit Sub
ErrorMessage:
MsgBox "Îøèáêà ïðè çàïóñêå îáíîâëåíèÿ-" + Error$
LogPrint "Îøèáêà çàïóñêà îáíîâëÿëêè-" + Error$

End Sub


'ëîâèì èâåíòû íà ôîðìå
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   'ïåðåäàåì äàííûå â îáúåêò
      cTray.CallEvent X, Y
      
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
Else
        cTray.CallPopupMenu Me, Menu, 2, , , mnuSh
  End If
   'îòæàòèå ïðàâîé êíîïêè ìûøè
      If MouseButton = 0 Or MouseButton = 2 Then
        cTray.CallPopupMenu Me, Menu, 2, , , mnuSh
      End If
      
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
    ShowTopicID 1, 11
End If
End Sub
Private Sub tmrCheck_Timer()
On Error GoTo 10
        Dim nz7 As Boolean
nz7 = CStr(GetSetting("Belyash AV", "Options", "MonReg", "True"))
 If nz = "False" Then
  Exit Sub
End If
'Debug.Print "start registry mon"
    Dim lRet As Long
    Static lCount As Long
    
    lRet = WaitForSingleObject(f_lChange, 0)
    If Not lRet = TIME_OUT Then
        lCount = lCount + 1
        
        'Label1.Caption = Now & "-Registry change notification received "
        Me.message 2, "Îáíàðóæåíî èçìåíåíèå â ðååñòðå.Çàáëîêèðîâàíî"
        modRegistry.CleanReg
        mRegNotifyChange.RegFindNextChange f_lChange, f_lKey, True, REG_NOTIFY_CHANGE_ALL
    End If
Exit Sub
10:
LogPrint "Îøèáêà ñëåæåíèÿ çà ðååñòðîì-" + Error$
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : DisplayError
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub DisplayError(szAPI As String, dwError As Long)
    Dim Buffer As String
    Buffer = Space(200)
    
    FormatMessage FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM, 0, dwError, LANG_NEUTRAL, Buffer, 0, 0&
    
    Debug.Print "Failed:", szAPI
    Debug.Print "Error code:", dwError
    Debug.Print "Message:", Buffer
    'End
End Sub

Sub start_regMon()
     
Dim a59 As String
a59 = CStr(GetSetting("Belyash AV", "Options", "MonReg", "True"))
If a59 = "True" Then
    tmrCheck.Enabled = True
End If
End Sub


Sub kolwoVirusow()
On Error GoTo 100
If Dir$(App.path + "\old.bmd", vbNormal) = "" Then
    Exit Sub
End If
Dim miNumBase As Integer
miNumBase = FreeFile
Dim sMD15 As String
Open App.path + "\old.bmd" For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD15
        countZap = countZap + 1
     Wend
 Close #miNumBase
 Exit Sub
100:
 LogPrint "" + Error$
End Sub

Sub kolwoVirusowAll(baZE As String)
On Error GoTo 100
Static i As Integer
i = i + 1
Dim miNumBase As Integer
miNumBase = FreeFile
Dim sMD15 As String
Open App.path + "\" + baZE For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD15
        a(i) = a(i) + 1
     Wend
 Close #miNumBase

 Exit Sub
100:
 LogPrint "" + Error$
End Sub
Sub sbor()
Dim sNext1 As String
sNext1 = Dir$(App.path + "\*.bVb")
While sNext1 <> ""

    kolwoVirusowAll (sNext1)
    sNext1 = Dir$
Wend
Dim baZE As String
kolwoVirusow
Dim nm As Long
Dim z As Byte

For z = 0 To 100
nm = nm + a(z)
Next z
countZapOLD = nm + countZap
End Sub

