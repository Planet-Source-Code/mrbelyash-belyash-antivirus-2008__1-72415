VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBlocked 
   AutoRedraw      =   -1  'True
   Caption         =   "Áëîêåð ïðèëîæåíèé"
   ClientHeight    =   4140
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "frmBlocked.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   4140
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Height          =   330
      Left            =   0
      Top             =   3450
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   2
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   630
      Top             =   1140
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   7350
      TabIndex        =   1
      Top             =   3840
      Width           =   7350
      Begin VB.CommandButton Command1 
         Caption         =   "Íàéòè ïðèëîæåíèå"
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   0
         Width           =   1785
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Çàêðûòü"
         Height          =   300
         Left            =   4230
         TabIndex        =   6
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Îáíîâèòü"
         Height          =   300
         Left            =   3075
         TabIndex        =   5
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Óäàëèòü"
         Height          =   300
         Left            =   1920
         TabIndex        =   4
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Çàïèñàòü"
         Height          =   300
         Left            =   630
         TabIndex        =   3
         Top             =   0
         Width           =   315
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Äîáàâèòü"
         Height          =   300
         Left            =   660
         TabIndex        =   2
         Top             =   0
         Width           =   315
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Align           =   1  'Align Top
      Bindings        =   "frmBlocked.frx":058A
      Height          =   3435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   6059
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1049
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmBlocked"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function Spliting(sFullPath As String, point As String)
Dim str1() As String
str1 = Split(sFullPath, point)
Spliting = str1(UBound(str1))
End Function
Private Sub Command1_Click()
'Îáúÿâëÿåì ñòðîêîâóþ ïåðåìåííóþ äëÿ íàçíà÷åíèÿ òèïîâ ôàéëîâ
Dim strFileType As String
'Åñëè âîçíèêíåò îøèáêà, ò.å ïîëüçîâàòåëü íàæåë íà êëàâèøó Cancel,
'îòïðàâèòüñÿ ê îáðàáîò÷èêó îøèáêè - ErrorHandler
On Error GoTo ErrorHandler
'Îáåñïå÷èâàåì ãåíåðàöèþ îùèáêè
CommonDialog1.CancelError = True

'Èíèöèàëèçèðóåì ñòðîêîâóþ ïåðåìåííóþ strFileType
strFileType = "All Files (*.*)|*.*|"
strFileType = strFileType & " Èñïîëíÿåìûé ôàéë (*.exe)| *.exe|"

'Ïðèñâàèâàåì åå ñâîéñòâó Filter
CommonDialog1.Filter = strFileType
'Óñòàíàâëèâàåì íåîáõîäèìûé èíäåêñ
CommonDialog1.FilterIndex = 2
'Ïðèñâàèâàåì íà÷àëüíóþ äèðåêòîðèþ ñâîñòâó InitDir
CommonDialog1.InitDir = App.path
'Îáåñïå÷èâàåì çàùèòó îò íåïðàâèëüíîãî ââåäåííîãî ôàéëà èëè äåðèêòîðèè, à òàê æå ñêðûâàåì ôëàæåê Read Only

CommonDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
'Âûçûâàåì äèàëîã Open
CommonDialog1.Action = 1 'Èëè æå CommonDialog1.ShowOpen
'***********
'Çäåñü ðàñïîëîãàåòñÿ Âàø êîä.(íå çàáóäòå, ÷òî ïóòü ê âûáðàííîìó ôàéëó Âû ñ÷èòûâàåòå èç ñâîéñòâà FileName)
Dim sCRC As String





datPrimaryRS.Recordset.AddNew 'Äîáàâèòü íîâóþ çàïèñü
datPrimaryRS.Recordset("CRC") = modFileManipulation.GetMD5(CStr(CommonDialog1.FileName))
datPrimaryRS.Recordset("Name") = Spliting(CommonDialog1.FileName, "\")
datPrimaryRS.Recordset("trust") = True
datPrimaryRS.Recordset.Update
' Ñîõðàíèòü èçìåíåíèÿ





'**********
Exit Sub
'Îáðàáîòêà ïåðåõâàòûâàåìîé îùèáêè
ErrorHandler:
If Err.Number = 32755 Then
Exit Sub
End If

End Sub



Private Sub Form_Load()
On Error GoTo 100
'With Adodc1
 '   .ConnectionString = "FILE NAME=C:\VB-DB\Nwind.UDL"
  '  .RecordSource = "SELECT * FROM Products WHERE ProductID"
  'End With
Dim file_name As String
file_name = App.path + "\inetbase.mdb"

datPrimaryRS.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;" & _
       "Data Source=" & file_name & _
       ";Mode=ReadWrite;Persist Security Info=False"
   datPrimaryRS.RecordSource = "Blocked"
datPrimaryRS.Refresh
'Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=F:\MY_Prog\Firewall\NEW\inetbase.mdb;Mode=ReadWrite
Exit Sub
100:
MsgBox "" + Error$, vbCritical

End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  grdDataGrid.Height = Me.ScaleHeight - datPrimaryRS.Height - 30 - picButtons.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'This is where you would put error handling code
  'If you want to ignore errors, comment out the next line
  'If you want to trap them, add code here to handle them
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.MoveLast
  grdDataGrid.SetFocus
  SendKeys "{down}"

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

