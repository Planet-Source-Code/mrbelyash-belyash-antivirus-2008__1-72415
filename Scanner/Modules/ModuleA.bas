Attribute VB_Name = "ModuleA"
Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function apiGetWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
Private Const mcGWCHILD = 5
Private Const mcGWHWNDNEXT = 2
Private Const mcGWLSTYLE = (-16)
Private Const mcWSVISIBLE = &H10000000
Private Const mconMAXLEN = 255





Function fEnumWindows()
Dim lngx As Long, lngLen As Long
Dim lngStyle As Long, strCaption As String
lngx = apiGetDesktopWindow()
'Return the first child to Desktop
lngx = apiGetWindow(lngx, mcGWCHILD)
Do While Not lngx = 0
strCaption = fGetCaption(lngx)
If Len(strCaption) > 0 Then
lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
'enum visible windows only
If lngStyle And mcWSVISIBLE Then
frmMain.analize "Class = " & fGetClassName(lngx)
frmMain.analize "Caption = " & fGetCaption(lngx)
Debug.Print "Class = " & fGetClassName(lngx),
Debug.Print "Caption = " & fGetCaption(lngx)
End If
End If
lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
Loop
End Function
Function DEnumWindows()
Dim lngx As Long, lngLen As Long
Dim lngStyle As Long, strCaption As String
lngx = apiGetDesktopWindow()
'Return the first child to Desktop
lngx = apiGetWindow(lngx, mcGWCHILD)
Do While Not lngx = 0
strCaption = fGetCaption(lngx)
If Len(strCaption) > 0 Then
lngStyle = apiGetWindowLong(lngx, mcGWLSTYLE)
'enum visible windows only
If lngStyle And mcWSVISIBLE Then
Dim i As Integer
    Dim x As ListItem
    
       Set x = frmMain.lvProcessDetail.ListItems.Add(, , fGetCaption(lngx))
        x.SubItems(1) = fGetClassName(lngx)
        'x.SubItems(2) = fGetCaption(lngx)
     'Me.lvProcess.ListItems.Add programs.ProcessName(i)
  '  lvProcess.ListItems.Add programs.ProcessHandle(i)

    Set x = Nothing
    
'frmMain.lvProcessDetail.ListItems.Add fGetClassName(lngx)
'frmMain.lvProcessDetail.ListItems.Add fGetCaption(lngx)
Debug.Print "Class = " & fGetClassName(lngx),
Debug.Print "Caption = " & fGetCaption(lngx)
End If
End If
lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
Loop
End Function
Private Function fGetClassName(hwnd As Long)
Dim strBuffer As String
Dim intCount As Integer
strBuffer = String$(mconMAXLEN - 1, 0)
intCount = apiGetClassName(hwnd, strBuffer, mconMAXLEN)
If intCount > 0 Then
fGetClassName = Left$(strBuffer, intCount)
End If
End Function
Private Function fGetCaption(hwnd As Long)
Dim strBuffer As String
Dim intCount As Integer
strBuffer = String$(mconMAXLEN - 1, 0)
intCount = apiGetWindowText(hwnd, strBuffer, mconMAXLEN)
If intCount > 0 Then
fGetCaption = Left$(strBuffer, intCount)
End If
End Function
 Sub Ac_Wind()
'ïåðåáîð îêîí
Call fEnumWindows

End Sub
