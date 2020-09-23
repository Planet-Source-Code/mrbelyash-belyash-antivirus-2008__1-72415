VERSION 5.00
Begin VB.Form cmdAllow 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Belyash AT Bloker"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AppBlock.xpcmdbutton cmdPAllow 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Äàííîå ïðèëîåíèå áóäåò âñåãäà áåç ïðîâåðêè çàïóùåíî"
      Top             =   2220
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      Caption         =   "Âñåãäà ðàçðåøàòü"
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
   Begin VB.TextBox txtFilePath 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   3240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   5295
   End
   Begin AppBlock.xpcmdbutton cmdAllow 
      Height          =   375
      Left            =   2220
      TabIndex        =   5
      ToolTipText     =   "×òî-òî íå ïîíÿòíî?"
      Top             =   2220
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      Caption         =   "Ïðîïóñòèòü åäèíîæäû"
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
   Begin AppBlock.xpcmdbutton cmdDeny 
      Height          =   375
      Left            =   4350
      TabIndex        =   6
      ToolTipText     =   "Ñåé÷àñ çàþëîêèðîâàòü,à â ñëåäóþùèé ðàç ñíîâà ñïðîñèòü"
      Top             =   2220
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   661
      Caption         =   "Åäèíîæäû çàáëîêèðîâàòü"
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
   Begin AppBlock.xpcmdbutton cmdPDeny 
      Height          =   375
      Left            =   6750
      TabIndex        =   7
      ToolTipText     =   "ïîíÿòíî?Ñì.ñïðàâêó"
      Top             =   2220
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      Caption         =   "Íàâñåãäà çàáë."
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
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Îæèäàíèå äåéñòâèÿ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3210
      TabIndex        =   3
      Top             =   1470
      Width           =   5295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ñîñòîÿíèå : "
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
      Left            =   1830
      TabIndex        =   1
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   240
      Picture         =   "frmAlert.frx":0CCA
      Top             =   240
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ïðèëîæåíèå :"
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
      Height          =   495
      Left            =   1860
      TabIndex        =   0
      Top             =   360
      Width           =   1365
   End
End
Attribute VB_Name = "cmdAllow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TextChange As Boolean
'Dim fso As New FileSystemObject
'Dim fn As File

Private Sub Form_Load()
    'Dim fso As New FileSystemObject
    'MsgBox fso.GetBaseName(Command$)
    'MsgBox fso.GetFileName(Command$)
    
    'avoid open it own file
    If Command$ = "" Then End
    
    ' check setting
    If ReadRegShell("HKCU\Software\BGAntivirus\AppBlock") = 0 Then
        'not enable => allow all file
        Call cmdAllow_Click
    End If
    
    'checked => manage file
    '======================
    'manage all files
    '======================
    Dim bCount As Long
    Dim i As Long
    
    If Val(ReadRegShell("HKCU\Software\BGAntivirus\ControlAll")) = 1 Then
        Me.txtFilePath.Text = Command$
        'Text1.Text = "BelyashAV1:" + Me.txtFilePath.Text
        'BelyashAV1 "BelyashAV1:" + Me.txtFilePath.Text
        ' check setting
        ' allow
        Dim aCount As Long
        aCount = ReadRegShell("HKCU\Software\BGAntivirus\Allow\aCount")
        
        For i = 1 To aCount
            'in allow list
            'MsgBox GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", CStr(i))
            If Command$ = ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Allow\" & CStr(i)) Then
                Call cmdAllow_Click
                'automatically END
            End If
        Next
        
        ' deny
        
        'Dim i As Long
        bCount = ReadRegShell("HKCU\Software\BGAntivirus\Ban\bCount")
        
        For i = 1 To bCount
            'in ban list
            'MsgBox GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", CStr(i))
            'Set fn = fso.GetFile(ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Ban\" & CStr(i)))
            'MsgBox fn.ShortPath
            If Command$ = ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Ban\" & CStr(i)) Then
                'Call cmdDeny_Click
                Me.lblStatus.Caption = "Ïðèëîæåíèå çàáëîêèðîâàíî"
                Me.cmdPDeny.Enabled = False
                Me.cmdDeny.Enabled = False
                Me.cmdPAllow.Enabled = False
            End If
        Next
    Else
        'manage only BLOCKED Files
        bCount = ReadRegShell("HKCU\Software\BGAntivirus\Ban\bCount")
        
        For i = 1 To bCount
            'in ban list
            If Command$ = ReadRegShell("HKEY_CURRENT_USER\Software\BGAntivirus\Ban\" & CStr(i)) Then
                'Call cmdDeny_Click
                Me.lblStatus.Caption = "Ïðèëîæåíèå çàáëîêèðîâàíî"
                Me.cmdPDeny.Enabled = False
                Me.cmdDeny.Enabled = False
                Me.cmdPAllow.Enabled = False
                Exit For
            End If
        Next
        'check if blocked or not
        If Me.lblStatus.Caption <> "Ïðèëîæåíèå çàáëîêèðîâàíî" Then
            'not found, not block => allow
            Call cmdAllow_Click
        End If
    End If
    'New app run normal
    
End Sub

Private Sub cmdAllow_Click()
On Error GoTo 10
    Shell Command$, vbNormalFocus
    'MsgBox "alloe. end"
    End
10:
If Err = 53 Then
    'MsgBox "" + CStr(Err)
    End
End If
End Sub

Private Sub cmdDeny_Click()
    End
End Sub

Private Sub cmdPAllow_Click()
    Dim aCount As Long
    aCount = Val(ReadRegShell("HKCU\Software\BGAntivirus\Allow\aCount"))
        
    'update count
    aCount = aCount + 1
    
    'update in registry
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", 1, "aCount", CStr(aCount))
    'Call CreateDwordValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", "aCount", aCount)
    'WriteRegShell "HKEY_CURRENT_USER\Software\BGAntivirus\Allow\aCount", CStr(aCount)
    
    'add new app to allow list
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", 1, CStr(aCount), Command$)
    Call cmdAllow_Click
    'End
End Sub

Private Sub cmdPDeny_Click()
    Dim bCount As Long
    bCount = Val(ReadRegShell("HKCU\Software\BGAntivirus\Ban\bCount"))
    
    'update count
    bCount = bCount + 1
    
    'update in registry
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", 1, "bCount", CStr(bCount))
    'Call CreateDwordValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "bCount", bCount)
    'WriteRegShell "HKEY_CURRENT_USER\Software\BGAntivirus\Ban\bCount", CStr(bCount)
    
    'add new app to allow list
    Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", 1, CStr(bCount), Command$)
    Call cmdDeny_Click
    'End
End Sub
