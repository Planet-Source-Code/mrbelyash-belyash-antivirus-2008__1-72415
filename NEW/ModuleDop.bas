Attribute VB_Name = "ModuleDop"
 
'Public Const PROCESS_TERMINATE = &H1
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_ENDSESSION = &H16
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
'Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
'Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Const WM_CLOSE = &H10
Dim strCaptions() As String ' ����� ����� ������ ��������� ���� ��������� ����
Dim lngHandle() As Long ' � ����� ��� ������ ���� ����
Public Function CloseProg(strCaption As String) As Boolean
Dim iCount As Integer
Dim i As Integer
Dim Pos As Integer
Dim lngEnum As Long
ReDim strCaptions(0)

     ReDim lngHandle(0) ' �������� ������ �� ��������� ������� �����������
     lngEnum = EnumWindows(AddressOf Callback1_EnumWindows, 0) ' �� �� ������

     For i = 0 To UBound(strCaptions) ' ���������� ��� �������
         Pos = InStr(1, strCaptions(i), strCaption, vbTextCompare) ' ���� ������ - �������� ����
         If Pos > 0 Then
         SendMessage lngHandle(i), WM_CLOSE, 0, 0
         SendMessage lngHandle(i), WM_ENDSESSION, 0, 0
         SendMessage lngHandle(i), WM_QUERYENDSESSION, 0, 0
         ' ����� ������� ��� ���� � ����� ��������� ����
         iCount = iCount + 1
         End If
     Next
End Function


Public Function Callback1_EnumWindows(ByVal hWnd As Long, ByVal lpData As Long) As Long
Dim cnt As Long
Dim rttitle As String * 256
     cnt = GetWindowText(hWnd, rttitle, 255) ' ���� ��������� ����
     If cnt > 0 Then ' �����, ����� ��������� ������� � �������
         ReDim Preserve lngHandle(UBound(strCaptions) + 1)
         ReDim Preserve strCaptions(UBound(strCaptions) + 1)
         strCaptions(UBound(strCaptions)) = Left$(rttitle, cnt)
         lngHandle(UBound(lngHandle)) = hWnd
     End If
     Callback1_EnumWindows = 1 ' ���������� ����������
End Function
'-------
Sub sn89()
' � Private Sub Form_Load() ��������
   CloseProg "Belyash Registry Monitor" ' ��� **** - ��� ����������
End Sub

  

