VERSION 5.00
Begin VB.Form frmTrayMsg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Scaner.xpcmdbutton cmdClose 
      Height          =   375
      Left            =   1740
      TabIndex        =   2
      Top             =   2190
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      Caption         =   "OK"
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
   Begin VB.Timer TimerUnload1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   540
      Top             =   2220
   End
   Begin VB.Timer TimerUnloadTime 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3360
      Top             =   2040
   End
   Begin VB.Timer TimerLoad 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3960
      Top             =   2040
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   90
      Picture         =   "frmTrayMsg.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   870
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
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
      Height          =   975
      Left            =   60
      TabIndex        =   1
      Top             =   900
      Width           =   4425
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   3465
   End
End
Attribute VB_Name = "frmTrayMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim j As Long

Private Sub cmdClose_Click()
    'Me.TimerUnload1.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    SetTopMostWindow Me.hwnd, True
  ' Call SetWindowPos(Me.Hwnd, HWND_TOPMOST, 0, 0, 0, 0, -1&)
    Call MakeTransparent(Me, 230)
    Me.Left = Screen.Width - 4605
  '  Me.Top = Screen.Height - Me.Height
    i = 0
    j = 0
    Me.TimerLoad.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Me.TimerLoad.Enabled = False
'    Me.TimerUnload1.Enabled = False
'    Me.TimerUnloadTime.Enabled = False
End Sub



Private Sub Image2_Click()
    Me.TimerUnload1.Enabled = True
End Sub



Private Sub Picture1_Click()
 Me.TimerUnload1.Enabled = True
End Sub

Private Sub TimerLoad_Timer()
    If i < 2715 Then
        Me.Top = Screen.Height - i - 500   '500 for taskbar
        i = i + 20
    Else
        Me.TimerLoad.Enabled = False
        Me.TimerUnloadTime.Enabled = True
    End If
End Sub

Private Sub TimerUnload1_Timer()
    If j < 230 Then
        'Me.Top = Screen.Height - j '- 500    '500 for taskbar
        'fade out
        Call MakeTransparent(Me, 230 - j)
        j = j + 5
    Else
        Me.TimerUnload1.Enabled = False
        Unload Me
    End If
End Sub

Private Sub TimerUnloadTime_Timer()
    Me.TimerUnload1.Enabled = True
    Me.TimerUnloadTime.Enabled = False
End Sub
