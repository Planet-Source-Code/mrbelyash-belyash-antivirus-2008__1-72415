VERSION 5.00
Begin VB.Form frmSplash1 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Dialog Caption"
   ClientHeight    =   3975
   ClientLeft      =   2715
   ClientTop       =   3315
   ClientWidth     =   7305
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2655
      Top             =   270
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
      Picture         =   "Dialog.frx":0000
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
      Left            =   4455
      TabIndex        =   6
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
      Left            =   4455
      TabIndex        =   5
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
      Left            =   4455
      TabIndex        =   4
      Top             =   2250
      Width           =   2355
   End
   Begin VB.Label LD 
      BackStyle       =   0  'Transparent
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
      Left            =   270
      TabIndex        =   3
      Top             =   3465
      Width           =   4740
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
      Caption         =   "Licensed To"
      Height          =   285
      Left            =   4500
      TabIndex        =   1
      Top             =   225
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gadyash Antivirus"
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
      Left            =   2835
      TabIndex        =   0
      Top             =   1305
      Width           =   4155
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   3735
      Left            =   135
      Top             =   90
      Width           =   7020
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pass As String
Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Unload Me
End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    MsgBox "Ïðèëîæåíèå óæå çàïóùåíî", vbCritical
    End
End If
Label5.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
LD.Caption = "Load system database..."
frmmain.LogPrint "Çàïóñê ïðîãðàììû "
Label2.Caption = "Licensed To "
Timer1.Enabled = True
 'SetTopMostWindow Me.hwnd, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
If Timer1.Enabled = True Then Timer1.Enabled = False
Set frmSplash = Nothing
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
