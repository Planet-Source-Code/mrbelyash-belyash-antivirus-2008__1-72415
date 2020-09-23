VERSION 5.00
Begin VB.UserControl XPProgressBar 
   BackStyle       =   0  'Transparent
   ClientHeight    =   5760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   417
   ToolboxBitmap   =   "Progress.ctx":0000
   Begin VB.PictureBox picpgb2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   15
      ScaleHeight     =   19
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   415
      TabIndex        =   0
      Top             =   15
      Width           =   6220
   End
   Begin VB.Image imgpgb 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   13
      Left            =   0
      Picture         =   "Progress.ctx":0312
      Top             =   5040
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   12
      Left            =   0
      Picture         =   "Progress.ctx":0B58
      Top             =   4680
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   11
      Left            =   0
      Picture         =   "Progress.ctx":139E
      Top             =   4320
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   10
      Left            =   0
      Picture         =   "Progress.ctx":1BE4
      Top             =   3960
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   9
      Left            =   0
      Picture         =   "Progress.ctx":242A
      Top             =   3600
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   1
      Left            =   0
      Picture         =   "Progress.ctx":2C70
      Top             =   720
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   8
      Left            =   0
      Picture         =   "Progress.ctx":34B6
      Top             =   3240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   7
      Left            =   0
      Picture         =   "Progress.ctx":3CFC
      Top             =   2880
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   5
      Left            =   0
      Picture         =   "Progress.ctx":4542
      Top             =   2160
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   4
      Left            =   0
      Picture         =   "Progress.ctx":4D88
      Top             =   1800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   3
      Left            =   0
      Picture         =   "Progress.ctx":55CE
      Top             =   1440
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   2
      Left            =   0
      Picture         =   "Progress.ctx":5E14
      Top             =   1080
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Height          =   285
      Index           =   6
      Left            =   0
      Picture         =   "Progress.ctx":6182
      Top             =   2520
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image imgpgb 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   0
      Picture         =   "Progress.ctx":651C
      Top             =   360
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "XPProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Enum constColor
Green = 0
Rad = 1
Greenmetal = 2
Blue = 3
Magenta = 4
Yellow = 5
Orange = 6
White = 7
Black = 8
Bluemetal = 9
Magentametal = 10
Gold = 11
Cyan = 12
Metal = 13
End Enum

Const m_def_Color = 0
Const m_def_Value = 0

Dim distance As Integer
Dim m_Value As Variant
Dim Value2
Dim m_Color As constColor
Dim m_Scroling As Boolean

Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub picpgb2_Click()
RaiseEvent Click
End Sub

Private Sub picpgb2_DblClick()
RaiseEvent DblClick
End Sub

Private Sub picpgb2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picpgb2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picpgb2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_Resize()
If UserControl.Height <> 315 Then UserControl.Height = 315
picpgb2.Width = UserControl.ScaleWidth - 2
Call Add
End Sub

Private Sub UserControl_Show()
    distance = 4
    picpgb2.PaintPicture imgpgb(0), 0, 0, 4, 19, 0, 0, 4, 19
    picpgb2.PaintPicture imgpgb(0), 4, 0, picpgb2.Width - 9, 19, 4, 0, 10, 19
    picpgb2.PaintPicture imgpgb(0), picpgb2.Width - 5, 0, 5, 19, 14, 0, 5, 19
End Sub

Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    Call Add
    PropertyChanged "Value"
End Property

Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
    m_Color = m_def_Color
End Sub

Public Property Get Color() As constColor
    Color = m_Color
End Property

Public Property Let Color(ByVal New_Color As constColor)
    m_Color = New_Color
    Call Add
    PropertyChanged "Color"
End Property

Public Property Get Scroling() As Boolean
    Scroling = m_Scroling
End Property

Public Property Let Scroling(ByVal New_Scroling As Boolean)
    m_Scroling = New_Scroling
    Call Add
    PropertyChanged "Scroling"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    m_Color = PropBag.ReadProperty("Color", m_def_Color)
    m_Scroling = PropBag.ReadProperty("Scroling", False)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Color", m_Color, m_def_Color)
    Call PropBag.WriteProperty("Scroling", m_Scroling, False)
End Sub

Private Sub Add()
    Dim colvo
    picpgb2.Cls
    UserControl_Show
    If Value > 100 Then Value = 100
    If Value < 0 Then Value = 0
    d = (Fix((picpgb2.Width) / 10) / 100) * Value
    If Value = 100 Then d = Fix((picpgb2.Width) / 10)
    If m_Scroling = True Then
    d = d * 2
    colvo = 4.6
    Else
    d = d
    colvo = 10
    End If
    For i = 1 To d
    picpgb2.PaintPicture imgpgb(m_Color).Picture, distance, 4, 8, 12, 23, 5, 8, 12
    distance = distance + colvo
    Next i
End Sub
