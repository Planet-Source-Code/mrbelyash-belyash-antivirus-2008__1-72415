VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl xpcmdpicture 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1695
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   59
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   113
   ToolboxBitmap   =   "xpcmdpicture.ctx":0000
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1080
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   600
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox lbl 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   120
   End
   Begin PicClip.PictureClip pc 
      Left            =   120
      Top             =   480
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "xpcmdpicture.ctx":0312
   End
End
Attribute VB_Name = "xpcmdpicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Created by Akopjan Dmitry(akopjan@inbox.ru) 2005

Option Explicit


Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Private Declare Function GdiAlphaBlend Lib "gdi32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINT_API) As Long
Private Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Enum State_b2
    Normal_ = 0
    Default_ = 1
End Enum

Dim State_Value As Integer
Dim m_Font As Font
Dim m_TransparentColor As OLE_COLOR
Dim m_Caption As String
Dim dFont As Font
Dim m_ForeColor As OLE_COLOR
Dim m_Blend As Boolean
Dim tw As Integer

Const m_Def_State = State_b2.Normal_

Private Type POINT_API
    X As Long
    Y As Long
End Type

Dim S As Integer
Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub lbl_Click()
    UserControl_Click
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    GetCursorPos pnt
    ScreenToClient UserControl.hwnd, pnt

    If pnt.X < UserControl.ScaleLeft Or _
       pnt.Y < UserControl.ScaleTop Or _
       pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
       
        Timer1.Enabled = False
        RaiseEvent MouseOut
        statevalue_pic
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
Set dFont = lbl.Font
    UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
    State_Value = 0
    Enabled = True
    MaskColor = vbWhite
    Caption = UserControl.Ambient.DisplayName
    Set Picture = LoadPicture()
    Set Font = dFont
    Blend = True
    ForeColor = vbBlack
    UserControl_Resize
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    make_xpbutton 1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X >= 0 And Y >= 0 And _
       X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
        If Button = vbLeftButton Then
            make_xpbutton 1
        Else: make_xpbutton 3
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    statevalue_pic
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    State_Value = PropBag.ReadProperty("State", 0)
    Enabled = PropBag.ReadProperty("Enabled", True)
    MaskColor = PropBag.ReadProperty("MaskColor", vbWhite)
    Caption = PropBag.ReadProperty("Caption", UserControl.Ambient.DisplayName)
    ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    Blend = PropBag.ReadProperty("Blend", True)
    Set Picture = PropBag.ReadProperty("Picture", LoadPicture())
    Set Font = PropBag.ReadProperty("Font", dFont)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    statevalue_pic
    'If Enabled = True Then lbl.ForeColor = vbBlack Else lbl.ForeColor = RGB(161, 161, 146)
End Property

Public Property Let MaskColor(ByVal C As OLE_COLOR)
m_TransparentColor = C
Capt
PropertyChanged "MaskColor"
End Property

Public Property Get MaskColor() As OLE_COLOR
MaskColor = m_TransparentColor
End Property

Private Sub UserControl_Resize()
lbl.Top = (UserControl.ScaleHeight - lbl.Height) / 2
    lbl.Left = (UserControl.ScaleWidth - lbl.Width) / 2
    statevalue_pic

End Sub

Private Sub UserControl_Show()
    statevalue_pic
End Sub

Private Sub UserControl_Terminate()
    statevalue_pic
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("State", State_Value, 0)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", lbl.Picture, "")
    Call PropBag.WriteProperty("Font", m_Font, dFont)
    Call PropBag.WriteProperty("MaskColor", m_TransparentColor, vbWhite)
    Call PropBag.WriteProperty("Caption", m_Caption, UserControl.Ambient.DisplayName)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, vbBlack)
    Call PropBag.WriteProperty("Blend", m_Blend, True)
    UserControl_Resize
End Sub

Public Property Get State() As State_b2
Attribute State.VB_Description = "Returns/sets the state of the command button when mouse_out."
Attribute State.VB_ProcData.VB_Invoke_Property = ";Misc"
    State = State_Value
End Property

Public Property Let State(ByVal vNewValue As State_b2)
    State_Value = vNewValue
    PropertyChanged "State"
    statevalue_pic
End Property

Private Sub statevalue_pic()
    If State_Value = 1 Then
        S = 4
    ElseIf State_Value = 0 Then
        S = 0
    End If
    
    If UserControl.Enabled = True Then
        make_xpbutton S
    Else: make_xpbutton 2
    End If
End Sub

Private Sub make_xpbutton(z As Integer)
   Dim BF As BLENDFUNCTION, lBF As Long
   Dim Alp As Integer
    UserControl.ScaleMode = 3 'Draw in pixels
    Dim brx, bry, bw, bh As Integer
    'Short cuts
    brx = UserControl.ScaleWidth - 3 'right x
    bry = UserControl.ScaleHeight - 3 'right y
    bw = UserControl.ScaleWidth - 6 'border width - corners width
    bh = UserControl.ScaleHeight - 6 'border height - corners height
    'Draws button
    'Goes clockwise first for corners(first four)
    'followed by borders(next four) and center(last step).
    UserControl.PaintPicture pc.GraphicCell(z), 0, 0, 3, 3, 0, 0, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), brx, 0, 3, 3, 15, 0, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), brx, bry, 3, 3, 15, 18, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), 0, bry, 3, 3, 0, 18, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), 3, 0, bw, 3, 3, 0, 12, 3
    UserControl.PaintPicture pc.GraphicCell(z), brx, 3, 3, bh, 15, 3, 3, 15
    UserControl.PaintPicture pc.GraphicCell(z), 0, 3, 3, bh, 0, 3, 3, 15
    UserControl.PaintPicture pc.GraphicCell(z), 3, bry, bw, 3, 3, 18, 12, 3
    UserControl.PaintPicture pc.GraphicCell(z), 3, 3, bw, bh, 3, 3, 12, 15
    
    
    If z = 3 Or Not Blend Then Alp = 255 Else Alp = 128
    Picture1.Move lbl.Left, lbl.Top, lbl.Width, lbl.Height
    Picture2.Move lbl.Left, lbl.Top, lbl.Width, lbl.Height
    
    With BF
        .BlendOp = 0
        .BlendFlags = 0
        .SourceConstantAlpha = Alp
        .AlphaFormat = 0
    End With
    RtlMoveMemory lBF, BF, 4
   
    GdiTransparentBlt Picture1.hdc, 0, 0, lbl.ScaleWidth, lbl.ScaleHeight, lbl.hdc, 0, 0, lbl.ScaleWidth, lbl.ScaleHeight, m_TransparentColor
    GdiAlphaBlend Picture2.hdc, 0, 0, lbl.ScaleWidth, lbl.ScaleHeight, Picture1.hdc, 0, 0, lbl.ScaleWidth, lbl.ScaleHeight, lBF
    GdiTransparentBlt UserControl.hdc, lbl.Left, lbl.Top, lbl.ScaleWidth, lbl.ScaleHeight, Picture2.hdc, 0, 0, lbl.ScaleWidth, lbl.ScaleHeight, vbWhite
       
    Picture1.Cls
    Picture2.Cls
End Sub

Private Sub Capt()
lbl.Width = 0
lbl.Height = 0
Set lbl.Picture = Picture
If lbl.Height < lbl.TextHeight(m_Caption) Then lbl.Height = lbl.TextHeight(m_Caption)
tw = lbl.Width
    lbl.Width = lbl.Width + lbl.TextWidth(m_Caption) + 5
    lbl.CurrentX = tw + 5
    lbl.CurrentY = (lbl.Height - lbl.TextHeight(m_Caption)) / 2
    lbl.ForeColor = m_ForeColor
    'lbl.BackColor = m_TransparentColor
    lbl.Print m_Caption
UserControl_Resize
End Sub

Public Property Get Picture() As StdPicture
    Set Picture = lbl.Picture
End Property

Public Property Set Picture(ByVal vNewCaption As StdPicture)
    Set lbl.Picture = vNewCaption
    Capt
    PropertyChanged "Picture"
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Set lbl.Font = m_Font
    Capt
    PropertyChanged "Font"
    
End Property

Public Property Let Caption(ByVal vNewCaption As String)
m_Caption = vNewCaption
Capt
PropertyChanged "Caption"
End Property

Public Property Get Caption() As String
Caption = m_Caption
End Property

Public Property Get Blend() As Boolean
Blend = m_Blend
End Property

Public Property Let Blend(vNewBlend As Boolean)
m_Blend = vNewBlend
UserControl_Resize
PropertyChanged "Blend"
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(vNewForeColor As OLE_COLOR)
m_ForeColor = vNewForeColor
Capt
PropertyChanged "ForeCOlor"
End Property
