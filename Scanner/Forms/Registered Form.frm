VERSION 5.00
Begin VB.Form frmRegistered1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registration Information"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   Icon            =   "Registered Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Registered Form.frx":76CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   2325
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton countinue 
      Caption         =   "&OK"
      Height          =   375
      Left            =   7080
      TabIndex        =   21
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton quit 
      Caption         =   "&Cancel"
      Height          =   285
      Left            =   4110
      TabIndex        =   20
      Top             =   570
      Width           =   1215
   End
   Begin VB.CommandButton checkme 
      Caption         =   "&OK"
      Height          =   285
      Left            =   4110
      TabIndex        =   19
      Top             =   180
      Width           =   1215
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      MaxLength       =   40
      TabIndex        =   14
      Top             =   1290
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1650
      MaxLength       =   40
      TabIndex        =   13
      Top             =   810
      Width           =   2055
   End
   Begin VB.TextBox info 
      Height          =   285
      Left            =   6000
      MaxLength       =   7
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox Checkid 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1650
      MaxLength       =   13
      TabIndex        =   1
      Top             =   1770
      Width           =   2055
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   3930
      X2              =   3690
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   3930
      X2              =   3930
      Y1              =   690
      Y2              =   2130
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00808080&
      X1              =   90
      X2              =   90
      Y1              =   690
      Y2              =   2130
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   90
      X2              =   3930
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   90
      X2              =   3930
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label Label17 
      Caption         =   "Please enter user name, company and serial key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   330
      Width           =   3855
   End
   Begin VB.Label Label16 
      Caption         =   "You must enter required information for registration "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   90
      Width           =   3855
   End
   Begin VB.Label Label11 
      Caption         =   "Enter Name and Company For Registration"
      Height          =   255
      Left            =   6840
      TabIndex        =   18
      Top             =   7320
      Width           =   3015
   End
   Begin VB.Label Label10 
      Caption         =   "Serial Key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   330
      TabIndex        =   17
      Top             =   1770
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   330
      TabIndex        =   16
      Top             =   1290
      Width           =   975
   End
   Begin VB.Label Label8 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   330
      TabIndex        =   15
      Top             =   810
      Width           =   1215
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   7680
      X2              =   7560
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   7680
      X2              =   7680
      Y1              =   5520
      Y2              =   6000
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00404040&
      X1              =   5160
      X2              =   5160
      Y1              =   5520
      Y2              =   6000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   5160
      X2              =   7680
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   5160
      X2              =   7680
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label7 
      Caption         =   "Capture Registration Information"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label timeup 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   7680
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label paidfor 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label5 
      Caption         =   "When Expired"
      ForeColor       =   &H00B5542F&
      Height          =   255
      Left            =   4560
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "When Registed"
      ForeColor       =   &H00B5542F&
      Height          =   255
      Left            =   4560
      TabIndex        =   8
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Data"
      ForeColor       =   &H00B5542F&
      Height          =   255
      Left            =   4560
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Current Id Code:"
      ForeColor       =   &H00B5542F&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label code 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   6960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Location :"
      ForeColor       =   &H00B5542F&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label location 
      BackColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   6600
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lblDay 
      AutoSize        =   -1  'True
      Caption         =   "Date Check is checking for date......"
      Height          =   315
      Left            =   6600
      TabIndex        =   0
      Top             =   7200
      Width           =   2565
   End
End
Attribute VB_Name = "frmRegistered1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WNetGetUserA Lib "mpr.dll" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long
Function GetComputerName() As String
Dim sBuffer As String * 255
If GetComputerNameA(sBuffer, 255&) <> 0 Then
GetComputerName = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
End If
End Function
Function GetUserName() As String
Dim sUserNameBuff As String * 255
sUserNameBuff = Space(255)
Call WNetGetUserA(vbNullString, sUserNameBuff, 255&)
GetUserName = Left$(sUserNameBuff, InStr(sUserNameBuff, vbNullChar) - 1)
End Function

Private Sub Form_Load()
Dim GDriv As String
GDriv = Environ("homedrive")
txtPwd.Text = dmReg.VolumeSerialNumber(GDriv + "\") 'Shows the serial number of your Hard Disk
txtName.Text = GetUserName
End Sub

Private Sub quit_Click()
Unload Me
End Sub

Private Sub txtName_GotFocus()
txtName.BackColor = &HF1EFF0
End Sub

Private Sub txtName_LostFocus()
txtName.BackColor = vbWindowBackground
End Sub
Private Sub txtPwd_GotFocus()
txtPwd.BackColor = &HF1EFF0
End Sub
Private Sub txtPwd_LostFocus()
txtPwd.BackColor = vbWindowBackground
End Sub
Private Sub Checkid_GotFocus()
Checkid.BackColor = &HF1EFF0
End Sub
Private Sub Checkid_LostFocus()
Checkid.BackColor = vbWindowBackground
End Sub

