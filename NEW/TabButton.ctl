VERSION 5.00
Begin VB.UserControl TabButton 
   Alignable       =   -1  'True
   ClientHeight    =   5205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11565
   ScaleHeight     =   5205
   ScaleWidth      =   11565
   Begin VB.Image ImgTab2 
      Height          =   495
      Left            =   6510
      Top             =   930
      Width           =   4245
   End
   Begin VB.Image ImgTab1 
      Height          =   495
      Left            =   2250
      Top             =   960
      Width           =   4065
   End
   Begin VB.Label LblTab2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Íàñòðîéêè"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   7530
      TabIndex        =   1
      Top             =   1110
      Width           =   2295
   End
   Begin VB.Label LblTab1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ãëàâíàÿ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2910
      TabIndex        =   0
      Top             =   1140
      Width           =   2295
   End
   Begin VB.Image ImgTabs2 
      Height          =   1440
      Left            =   0
      Picture         =   "TabButton.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11490
   End
   Begin VB.Image ImgTabs1 
      Height          =   1440
      Left            =   0
      Picture         =   "TabButton.ctx":BFC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11520
   End
End
Attribute VB_Name = "TabButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event TabClick1()
Event TabClick2()

Private Sub ImgTab1_Click()
TabClick1
End Sub

Private Sub ImgTab2_Click()
TabClick2
End Sub









Private Sub UserControl_Resize()
UserControl.Width = ImgTabs1.Width
UserControl.Height = ImgTabs1.Height
End Sub

Public Sub SetText(TabIndex As Integer, Text As String)
If TabIndex = 1 Then
LblTab1.Caption = Text
End If
If TabIndex = 2 Then
LblTab2.Caption = Text
End If
End Sub

Private Sub TabClick1()
ImgTabs1.Visible = False
ImgTabs2.Visible = True
FrmNet.Frame1.Visible = False
FrmNet.Frame2.Visible = True
RaiseEvent TabClick1
End Sub

Private Sub TabClick2()
ImgTabs1.Visible = True
ImgTabs2.Visible = False
FrmNet.Frame2.Visible = False
FrmNet.Frame3.Visible = True
RaiseEvent TabClick2
End Sub
