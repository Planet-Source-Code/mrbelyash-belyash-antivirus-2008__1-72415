Attribute VB_Name = "ModuleA"
Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassname As String, ByVal nMaxCount As Long) As Long
Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function apiGetWindow Lib "user32" Alias "GetWindow" (ByVal Hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function apiGetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal aint As Long) As Long
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
frmMain.analizeMe "Class = " & fGetClassName(lngx)
frmMain.analizeMe "Caption = " & fGetCaption(lngx)
'Debug.Print "Class = " & fGetClassName(lngx)
'Debug.Print "Caption = " & fGetCaption(lngx)
End If
End If
lngx = apiGetWindow(lngx, mcGWHWNDNEXT)
Loop
End Function
Private Function fGetClassName(Hwnd As Long)
Dim strBuffer As String
Dim intCount As Integer
strBuffer = String$(mconMAXLEN - 1, 0)
intCount = apiGetClassName(Hwnd, strBuffer, mconMAXLEN)
If intCount > 0 Then
fGetClassName = Left$(strBuffer, intCount)
End If
End Function
Private Function fGetCaption(Hwnd As Long)
Dim strBuffer As String
Dim intCount As Integer
strBuffer = String$(mconMAXLEN - 1, 0)
intCount = apiGetWindowText(Hwnd, strBuffer, mconMAXLEN)
If intCount > 0 Then
fGetCaption = Left$(strBuffer, intCount)
End If
End Function
 Sub EnWinMy()
 frmMain.analizeMe "-----------Active window start--------------"
Call fEnumWindows
 frmMain.analizeMe "-----------Active window end___--------------"
End Sub

