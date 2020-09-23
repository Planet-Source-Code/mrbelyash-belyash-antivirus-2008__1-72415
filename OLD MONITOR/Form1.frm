VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ListFolder 
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   765
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hChangeHandle As Long
Dim hWatched As Long
Dim terminateFlag As Long
Private Sub Form_Load()
Me.Show
enblMonFold

End Sub


Private Function WatchCreate(lpPathName As String, flags As Long) As Long

  'FindFirstChangeNotification members:
  '
  '  lpPathName: folder to watch
  '  bWatchSubtree:
  '     True = watch specified folder and its sub folders
  '     False = watch the specified folder only
  '  flags: OR'd combination of the FILE_NOTIFY_ flags to apply
  
   WatchCreate = FindFirstChangeNotification(lpPathName, False, flags)

End Function


Private Sub WatchDelete(hWatched As Long)
    
   Dim r As Long
   
   terminateFlag = True
   DoEvents

   r = FindCloseChangeNotification(hWatched)
 
End Sub


Private Function WatchDirectory(hWatched As Long, interval As Long) As Long

  'Poll the watched folder.
  'The Do..Loop will exit when:
  '   r = 0, indicating a change has occurred
  '   terminateFlag = True, set by the WatchDelete routine
  
   Dim r As Long
   
   Do
   
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   
   Loop While r <> 0 And terminateFlag = False
   
   WatchDirectory = r
   
End Function


Private Function WatchResume(hWatched As Long, interval) As Boolean

   Dim r As Long
   
   r = FindNextChangeNotification(hWatched)
   
   Do
      
      r = WaitForSingleObject(hWatched, interval)
      DoEvents
   
   Loop While r <> 0 And terminateFlag = False
   
   WatchResume = r
   
End Function


Private Sub WatchChangeAction(fPath As String)

   Dim fName As String
   
   ListFolder.Clear

   fName = Dir(fPath & "\" & "*.txt")
   
   If fName > "" Then
   
      ListFolder.AddItem "path: " & vbTab & fPath
      ListFolder.AddItem "file: " & vbTab & fName
      ListFolder.AddItem "size: " & vbTab & FileLen(fPath & "\" & fName)
      ListFolder.AddItem "attr: " & vbTab & GetAttr(fPath & "\" & fName)
   
   End If

End Sub




Private Sub Form_Unload(Cancel As Integer)
   Call WatchDelete(hWatched)
   hWatched = 0
End Sub
Sub enblMonFold()

   Dim r As Long
   Dim watchPath As String
   Dim watchStatus As Long
   
   watchPath = App.Path
   
   terminateFlag = False
   'cmdBegin.Enabled = False
   
   'lbMsg = "Using Explorer and Notepad, create, modify, rename, delete or "
   'lbMsg = lbMsg & "change the attributes of a text file in the watched directory."""

  'get the first file text attributes to the listbox (if any)
   WatchChangeAction watchPath
   
  'show a msgbox to indicate the watch is starting
   'MsgBox "Beginning watching of folder " & watchPath & " .. press OK"
   
  'create a watched directory
   hWatched = WatchCreate(watchPath, FILE_NOTIFY_FLAGS)
   
  'poll the watched folder
   watchStatus = WatchDirectory(hWatched, 100)
  
  'if WatchDirectory exited with watchStatus = 0,
  'then there was a change in the folder.
   If watchStatus = 0 Then
   DoEvents
      'update the listbox for the first file found in the
      'folder and indicate a change took place.
       WatchChangeAction watchPath
                            frmTest.Text1.Text = "Îáíàðóæåíî èçìåíåíèå â êàòàëîãå ïðîãðàììû.Ïðîäîëæàþ íàáëþäåíèå..."
                            Call frmTest.message
       'MsgBox "The watched directory has been changed.  Resuming watch..."
       
      '(perform actions)
      'this is where you'd actually put code to perform a
      'task based on the folder changing.
      'FindNextChangeNotification API, again exiting if
      'watchStatus indicates the terminate flag was set
       Do
       DoEvents
         watchStatus = WatchResume(hWatched, 100)
         
         If watchStatus = -1 Then
              'watchStatus must have exited with the terminate flag
                           frmTest.Text1.Text = " Íàáëþäåíèå çà " & vbCrLf + watchPath + vbCrLf + "ïðåðâàíî"
                          Call frmTest.message
               'MsgBox "Watching has been terminated for " & watchPath
         
         Else: WatchChangeAction watchPath
               
               frmTest.Text1.Text = "Ñíîâà îáíàðóæåíî èçìåíåíèå â êàòàëîãå ñ ïðîãðàììîé."
                          Call frmTest.message
                          'Exit Do
                          DoEvents
                          'MsgBox "The watched directory has been changed again."
              
              '(perform actions)
              'this is where you'd actually put code to perform a
              'task based on the folder changing.
               
         End If
         
       Loop While watchStatus = 0
   
   
   Else
     'watchStatus must have exited with the terminate flag
      frmTest.Text1.Text = "Íàáëþäåíèå çà" & watchPath + " ïðåðâàíî"
                          Call frmTest.message
      'MsgBox "Watching has been terminated for " & watchPath
   
   End If
   
End Sub
