VERSION 5.00
Begin VB.Form frmAlert 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Al Ameen"
   ClientHeight    =   2985
   ClientLeft      =   4500
   ClientTop       =   6330
   ClientWidth     =   4695
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4695
   Begin VB.ListBox lstBox 
      Height          =   450
      Left            =   390
      TabIndex        =   7
      Top             =   1710
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2730
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   270
      Visible         =   0   'False
      Width           =   1140
   End
   Begin monitor.xpcmdpicture xpcmdpicture1 
      Height          =   405
      Left            =   690
      TabIndex        =   3
      Top             =   2340
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   714
      Picture         =   "frmAlert.frx":0000
      Caption         =   "Ðàçðåøèòü"
      Blend           =   0   'False
   End
   Begin monitor.xpcmdpicture xpcmdpicture2 
      Default         =   -1  'True
      Height          =   405
      Left            =   2460
      TabIndex        =   4
      Top             =   2340
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
      Picture         =   "frmAlert.frx":001C
      Caption         =   "Çàáëîêèðîâàòü"
      Blend           =   0   'False
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF0000&
      Height          =   2775
      Left            =   90
      Top             =   90
      Width           =   4485
   End
   Begin VB.Label Label5 
      Height          =   315
      Left            =   2730
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label4 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   210
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label Label3 
      Height          =   285
      Left            =   2640
      TabIndex        =   2
      Top             =   210
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Âíèìàíèå !!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   270
      Width           =   2085
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   180
      Picture         =   "frmAlert.frx":0038
      Stretch         =   -1  'True
      Top             =   180
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   960
      Left            =   180
      TabIndex        =   0
      Top             =   1290
      Width           =   4305
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare the API Calls for Makin Transparent Window


'******************************************************************************************
'******************************************************************************************
'************************************HELLO PROGRAMMERS*************************************
'******************************************************************************************
'******************************************************************************************
'******************************************************************************************


 '   Hello My Name is Shoaib Mohammed and Iam Doin my I BE on Electrical & Electronics
'Engineering.This Small Program Show how to create a High Class Notification Box on The
'task bar as some programs do (example Norton Antivirus). It was Norton which inspired me
'create this. This is not fully error free and needs a lot of maintainance to produce
'a error free message.

'    You have to constantly monitor the message showed. U cannot show a message until
'a previous one  closes completely. other than that there would be no problem , i think.
'i created this program in a hurry ( i am getting late for the date....just kidding...have
'to hang out with friends and prepare for tommororws exam)and so i was no able to add any
'detailed comments.
    
'    The Best Part of this is that it does no use any api call ( Except to make the
'form transparent). It uses Pure VB and Mathematical calculation (easy to undr stand).


    'I think This woudld be much useful and would provide a Better Interface for ur'
'program.

 '   This is just a base class and u are free to change it any way as you like to
'your own style. U have the full permission of using this code or modifying it or using
'it in any commercial package(but please send ur comments that what keeps programmers
'like me on the GO.)


 '   PLEASE SEND IN UR COMMENTS OR BUGS TO

  '      shoaib_134@ rediffmail.com

'*********************************** BYE *******************************************
 Private Const GWL_EXSTYLE = (-20)
 Private Const WS_EX_TRANSPARENT = &H20&
 Private Const SWP_FRAMECHANGED = &H20
 Private Const SWP_NOMOVE = &H2
 Private Const SWP_NOSIZE = &H1
 Private Const SWP_SHOWME = SWP_FRAMECHANGED Or _
 SWP_NOMOVE Or SWP_NOSIZE

 Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
 

Const n = vbNewLine
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2


Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, _
ByVal cx As Long, ByVal cy As Long, _
ByVal wFlags As Long) As Long


Dim ret As Long










Private Sub Form_Activate()
Me.Label2.Caption = "Ïðèëîæåíèå " + Spliting(Me.Label3.Caption, "\") + " ïûòàåòñÿ âûéòè â ñåòü. Ðàçðåøèòü?"
xpcmdpicture1.Caption = "Ðàçðåøèòü"
xpcmdpicture2.Caption = "Çàáëîêèðîâàòü"
End Sub
Private Sub Form_Load()
'Get the Picture
'Image1.Picture = LoadPicture(App.Path & "\error.gif")
lstBox.Clear
'Main Part
'Set The Position of the form
Me.Top = Screen.Height - Me.Height
Me.Left = Screen.Width - Me.Width
'Me.Label2.Caption = "Ïðèëîæåíèå " + newAplic + " ïûòàåòñÿ âûéòè â ñåòü. Ðàçðåøèòü?"
'Make the Form Transparent

'SetWindowLong Me.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
'Call Trans

'Check if The form had been loaded , if loaded then unload it
If Loaded Then Unload Me: Loaded = False Else Loaded = True
End Sub

'Transparent Making area
Private Sub Trans()
'ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
'ret = ret Or WS_EX_LAYERED
'SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
'SetLayeredWindowAttributes Me.hwnd, 0, 200, LWA_ALPHA

'Dim lngReturnValue As Long
 'lngReturnValue = SetWindowPos(Me.hwnd, -2, 0, 0, 0, 0, 3)
End Sub

'This sub calculates the total height necessary for the form
'This sub resizes the form according to the message. This should me called before
'-the form is shown
Public Sub Resize()
Me.Height = Me.Height + Label2.Height
'Command1.Top = Me.Height - 500 '+ Label2.Height - 500

End Sub



Private Sub Form_Terminate()

Unload Me 'unload the form
End Sub





 



Private Sub xpcmdpicture1_Click()
Dim mi_na As String
Me.Data1.DatabaseName = App.path + "\inetbase.mdb"
Me.Data1.RecordSource = "applic"
Me.Data1.Refresh

If Trim$(Label3.Caption) = "" Then
    Label3.Caption = "Íåèçâåñòíîå"


End If


mi_na = modFileManipulation.GetMD5(CStr(Label3.Caption))
10:
If Trim$(Label3.Caption) = "" Then
    Label3.Caption = "Íåèçâåñòíîå"
End If
Me.Data1.Recordset.AddNew 'Äîáàâèòü íîâóþ çàïèñü
Me.Data1.Recordset("name") = Spliting(Label3.Caption, "\")
Me.Data1.Recordset("crc") = mi_na
Me.Data1.Recordset("trust") = True
Me.Data1.Recordset.Update ' Ñîõðàíèòü èçìåíåíèÿ



FillProcessListNT_All
modThread.Thread_Resume (Me.Label4.Caption)
'Allow = True
'set that the form was  loaded once
Loaded = True
Unload Me
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
Private Sub xpcmdpicture2_Click()
        TerminateThisConnection (Me.Label5.Caption)
        modThread.Thread_Resume (Me.Label4.Caption)
        
        ModlLoadProcess.KillProcessById (Me.Label4.Caption)
'set that the form was  loaded once
Loaded = True
'enable the Timer to hide the form
'Timer2.Enabled = True
Unload Me
End Sub
