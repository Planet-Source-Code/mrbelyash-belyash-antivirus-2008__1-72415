VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "VB ����"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   11.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "����������� ����..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3210
      TabIndex        =   4
      Top             =   360
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "������ ���� ����"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3210
      TabIndex        =   3
      Top             =   660
      Width           =   1755
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   0
      Top             =   5160
      Width           =   5025
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   13
         TabIndex        =   2
         Top             =   30
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   390
         TabIndex        =   1
         Top             =   60
         Width           =   525
      End
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   12
      Left            =   4650
      Picture         =   "Form1.frx":0000
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   11
      Left            =   4320
      Picture         =   "Form1.frx":04F2
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   10
      Left            =   3990
      Picture         =   "Form1.frx":09E4
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   9
      Left            =   3660
      Picture         =   "Form1.frx":0ED6
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   8
      Left            =   3330
      Picture         =   "Form1.frx":13C8
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   7
      Left            =   3000
      Picture         =   "Form1.frx":18BA
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   6
      Left            =   2670
      Picture         =   "Form1.frx":1DAC
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   600
      Index           =   5
      Left            =   1920
      Picture         =   "Form1.frx":229E
      Top             =   30
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   4
      Left            =   1590
      Picture         =   "Form1.frx":3960
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   3
      Left            =   1260
      Picture         =   "Form1.frx":3E52
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   930
      Picture         =   "Form1.frx":4344
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image CheckImage 
      Height          =   240
      Left            =   0
      Picture         =   "Form1.frx":4836
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   600
      Picture         =   "Form1.frx":4B78
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   270
      Picture         =   "Form1.frx":506A
      Top             =   30
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Menu File 
      Caption         =   "����"
      Begin VB.Menu FileItem 
         Caption         =   "�������"
         Index           =   0
         Shortcut        =   ^A
      End
      Begin VB.Menu FileItem 
         Caption         =   "�������"
         Index           =   1
         Shortcut        =   ^L
      End
      Begin VB.Menu FileItem 
         Caption         =   "���������"
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu FileItem 
         Caption         =   "��������� ���..."
         Index           =   3
         Shortcut        =   ^E
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu FileItem 
         Caption         =   "������"
         Index           =   5
         Shortcut        =   ^P
      End
      Begin VB.Menu FileItem 
         Caption         =   "�����"
         Checked         =   -1  'True
         Index           =   6
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu FileItem 
         Caption         =   "����������� ����"
         Enabled         =   0   'False
         Index           =   8
      End
      Begin VB.Menu FileItem 
         Caption         =   "���� � �����..."
         Checked         =   -1  'True
         Index           =   9
      End
      Begin VB.Menu FileItem 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu FileItem 
         Caption         =   "�����"
         Index           =   11
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "��������"
      Begin VB.Menu EditItem 
         Caption         =   "��������"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu EditItem 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu EditItem 
         Caption         =   "��������"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu EditItem 
         Caption         =   "����������"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu EditItem 
         Caption         =   "��������"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu EditItem 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu EditItem 
         Caption         =   "�����"
         Index           =   6
         Shortcut        =   {F3}
      End
      Begin VB.Menu EditItem 
         Caption         =   "�������� ���"
         Index           =   7
      End
   End
   Begin VB.Menu Help 
      Caption         =   "�������"
      Begin VB.Menu HelpItem 
         Caption         =   "������..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu HelpItem 
         Caption         =   "� ���������"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ����������������������������������������������������������������������������������������������������������������������������������
' ��                                                                                                                              ��
' ��                            ������ ������������� ������� � ����������� VB ����                                                  ��
' ��                                        �����: �������� �����                                                                 ��
' ��                                                                                                                              ��
' ����������������������������������������������������������������������������������������������������������������������������������

' ����������������������������������������������������������������������������������������������������������������������������������
' ��
' ��                                              ������ � �������:
' ��
' ��                                 ������ ������ ���������� ����������������.
' ��                             ������� �� VB6 � �������������� API (��� ����������)
' ��                              !!! � ������� ��� �� ������ On Error Resume next!!!
' ��                                    ���������� ��� - ���� ��� ���� ���������...
' ��
' ��                                    ��� ������ ������� ��������� ������������
' ��                                            ����������� ���� Windows
' ��                                   ����������� ���� ���������� ����� �������
' ��                                          ������ "����������� ����"
' ��                                              ��� ��� ���������...
' ��
' ��                                      �������� �������������� ���������
' ��
' ��                                                WM_MEASUREITEM
' ��                                �������� �� ���������� ��������� ���� (������, ������)
' ��
' ��                                                  WM_DRAWITEM
' ��                                      �������� �� ���������� ��������� ����
' ��                                        (ODA_DRAWENTIRE ���������� ��� measure,
' ��                                            ODS_SELECTED ������� ���� ������,
' ��                              ODS_SELECTED and ODA_SELECT ����� �� ������ �������� ����)
' ��
' ��                                                 WM_MENUSELECT
' ��                                              ������� ���� �� ����...
' ��
' ��                            ��� ���������, ������� �������� ���������. ��������� �� ��������!
' ��                               �������� ����� � ��. ��������� � ��������� Command2_Click
' ��                                  ������� � ������, ������������ �������� ����.
' ��                                        ���� ��� �� ��� ��������� ���...
' ��
' ��                                                   ����������!
' ��                          � XP ��� ���������� ������� "�������������" � ����� "���������� ����
' ��                                ������������� ����" ���� �� ���� �� ���������� - �������???
' ��                                  ��� ���������� ������� "����������" - ��� � �������.
' ��
' ��
' ����������������������������������������������������������������������������������������������������������������������������������

' ���� ������������ ����������
Public i               As Long
Public allMenu         As Long
Private m_ExMenu       As Boolean

' ������������ ������������� ���� � ������ �� �����...
Private Sub Command1_Click()
    Me.PopupMenu File, 8, Command1.Left + Command1.Width, Command1.Top + Command1.Height
End Sub
' ������ ��������� ����������� ������������ ��� ������ ����
Private Sub Command2_Click()
    '/////////////////////////////////////////////////////////////////////////////////////////////////
    ' ������������� �������� ��� ������ ����
    m_ExMenu = True
    Command2.Caption = "���������!!!"
    Command2.Enabled = False
    ' ��� �������� ����� �����������...
    m_Margin = 3
    m_CheckAreaMenuColor = GetSysColor(COLOR_BTNFACE)
    m_PictureAreaMenuColor = RGB(250, 249, 245)
    m_CaptionAreaMenuColor = GetSysColor(COLOR_MENU)
    m_TextMenuColor = GetSysColor(COLOR_MENUTEXT)
    m_SelectTextMenuColor = RGB(255, 255, 255) '(128, 64, 64)
    m_HotKeyTextMenuColor = RGB(255, 203, 151)
    m_SeparatopColor = RGB(255, 203, 151)
    m_PictureMaskColor = RGB(255, 255, 255)
    m_ShadowColor = RGB(128, 128, 255)
    m_FrameMenuColor = RGB(128, 128, 255)
    m_FrameMenuBackColor = RGB(206, 206, 255)
    m_GrayFrameMenuColor = GetSysColor(COLOR_GRAYTEXT)
    m_GrayFrameMenuBackColor = RGB(250, 249, 245)
    m_LabelRC.Right = 7
    m_LabelBackColor = RGB(255, 203, 151)
    m_LabelForeColor = 0
    
    ' �������� ����� � ������
    Set Module1.MenuForm = Me
    
    Dim topMenu As Long
    Dim subMenu As Long
        
    ' ������� ����� ���������� ����
    allMenu = 0
    For topMenu = 0 To GetMenuItemCount(GetMenu(hwnd)) - 1
        allMenu = allMenu + GetMenuItemCount(GetSubMenu(GetMenu(hwnd), topMenu))
    Next
    ' � ������ ���������� ��������� ������� ����
    allMenu = allMenu + GetMenuItemCount(GetMenu(hwnd))
    ' ������ �� OwnerDrawMenu
    For i = 0 To allMenu
        CreateOwnerDrawMenu GetMenu(hwnd), i, i + 2
    Next
    ' ������� ������ ��������
    ReDim itemPicture(allMenu + 2)
    For i = 0 To allMenu
        Set itemPicture(i) = New StdPicture
    Next
    
    ' ��������� � ������ ������ �������� �� Image...
    ' ������ ����
    Set itemPicture(2) = Image1(0).Picture
    Set itemPicture(3) = Image1(1).Picture
    Set itemPicture(4) = Image1(2).Picture
    Set itemPicture(5) = Image1(3).Picture
    Set itemPicture(7) = Image1(4).Picture
    Set itemPicture(8) = Image1(5).Picture
    Set itemPicture(11) = Image1(11).Picture
    Set itemPicture(13) = Image1(12).Picture
    ' ������ ��������
    Set itemPicture(15) = Image1(6).Picture
    Set itemPicture(17) = Image1(7).Picture
    Set itemPicture(18) = Image1(8).Picture
    Set itemPicture(19) = Image1(9).Picture
    Set itemPicture(21) = Image1(10).Picture
    
    
    ' ��������� ������������ ������ ��������
    m_MaxPictureWidth = 0
    For i = 0 To allMenu
        If m_MaxPictureWidth < ScaleX(itemPicture(i).Width) Then m_MaxPictureWidth = ScaleX(itemPicture(i).Width)
    Next
    ' ��������� ��������
    wlOldProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf MsgProc)
    
End Sub

Public Sub Form_Load()
    
End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' ����������� ����
    If Button = 2 Then Me.PopupMenu Edit, 8, X, Y
End Sub

Public Sub Form_Unload(Cancel As Integer)
    If m_ExMenu Then
        ' ���������� ������ ������� ���������
        If wlOldProc <> 0 Then SetWindowLong hwnd, GWL_WNDPROC, wlOldProc
        ' ����������� ������ �� ��������
        For i = 0 To allMenu
            Set itemPicture(i) = Nothing
        Next
    End If
End Sub

Public Sub Picture2_Change()
    ' ������ ������ ������
    Picture1.Height = m_Margin + Picture2.Height + m_Margin
    ' ������� ���� ��������
    Picture2.Move m_Margin, m_Margin
    ' ������� �����
    Label1.Move Picture2.Left + Picture2.Width + m_Margin, (Picture1.Height - Label1.Height) / 2
End Sub

