Attribute VB_Name = "modScanVirus"
Option Explicit
Public Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Boolean
Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long


Public Const pr = "Belyash Shield 2008b"
Public monitoring  As Boolean
'Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
'Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Boolean
'Public Const REALTIME_PRIORITY_CLASS = &H100
'Public Const HIGH_PRIORITY_CLASS = &H80
'Public Const NORMAL_PRIORITY_CLASS = &H20
'Public Const IDLE_PRIORITY_CLASS = &H40
'Public Declare Function OpenProcess _
'Lib "kernel32" (ByVal dwDesiredAccess As Long, _
'ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Public Const PROCESS_QUERY_INFORMATION = &H400
'Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


'Public Const PROCESS_QUERY_INFORMATION = &H400
'Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public blnScan As Boolean   'check scan or stop
Public strScanDetail As String  'Scan Detail HTML text
Public QarNoDelete As Boolean
Public procCountSpl As Integer
Public procFoundSpl As Integer
Public procCleanSpl As Integer
Public Logging As Boolean
Public CRC As New clsCRC
Public FileSize As Long
'declare Virus Def & info
Public VSig() As VirusSig
Public VSInfo As VS_Info
Public Type VirusSig

    Name As String
    Type As String
    Value As String
    Action As String
    ActtionVal As String
    
End Type

'new DataType for Virus Signature Info
Public Type VS_Info
    
    VirusCount As Long
    LastUpdate As Date
    
End Type
'for faster DoEvents
Declare Function GetInputState Lib "user32.dll" () As Long
'declaration for Classes of Registry
Public Const HKEY_ALL = &H0&
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003

'new ACCESS KEY for delete startup registry
Private Const KEY_ALL_ACCESS = &H3F '((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Sub Process_Kill(P_ID As Long)
    '// Kill the wanted process
    On Error Resume Next
    Dim hProcess As Long
    Dim lExitCode As Long
    Dim res As Boolean
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_TERMINATE, False, P_ID)
    res = GetExitCodeProcess(hProcess, lExitCode)
    res = TerminateProcess(hProcess, lExitCode)
    CloseHandle (hProcess)
End Sub


'declare variable for scan reg extensions
Sub ScanFile(ByVal sPath As String)

    On Error Resume Next
    
    Dim fso As New FileSystemObject
    'Dim X As ListItem
    'main folder
    Dim mFolder As Folder
    'files and folders collections
    Dim sFolders As Folders
    Dim sFiles As Files
    'for loop variables
    Dim sFolder As Folder
    Dim sFile As File
    
    'get main folder
    Set mFolder = fso.GetFolder(sPath)
    'get subfolders in main folder
    Set sFolders = mFolder.SubFolders
    'get files in main folder
    Set sFiles = mFolder.Files
    
    'scan files
    For Each sFile In sFiles
        
        'limit file size
        If sFile.Size > FileSize Then GoTo endFor 'can't use Continue For
        
        'If GetInputState() <> 0 Then DoEvents  'faster but look stuck
        DoEvents
        
        'check if it is stopped
        If blnScan = False Then Exit For
        'show scanned file
        'frmTest.lblPath.Caption = sFile
        frmTest.lblPathText.Text = sFile
        'add 1 to counter scanned
        frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
        'frmTest.SB.Panels(2).Text = frmTest.lblCount.Caption
        'frmTest.SB.Refresh
        'scan virus
        Dim sCRC As String
        sCRC = CRC.GetCRC(sFile)
        Dim i As Long
        LogPrint sFile.Name
        'scan algorithm 1
        '----------------
        
'        If InStr(1, strVirusDef, sCRC, vbBinaryCompare) > 0 Then
'            'virus found
'            For i = 0 To UBound(VirusDef)
'                If GetInputState() <> 0 Then DoEvents
'                If sCRC = VirusDef(i)(2) Then   'start cleaning
'                    'add to log
'                    strScanDetail = strScanDetail & "Virus Found: <Font Size=3 Color=RED>" & VirusDef(i)(0) & "</font><br>"
'                    strScanDetail = strScanDetail & " File: <Font Size=3 Color=ORANGE><i>" & sFile & "</i></font><br>"
'                    Call UpdateDetail(strScanDetail, frmTest.WebBrowser1)
'
'                    'add 1 to counter
'                    frmTest.lblFound.Caption = Int(frmTest.lblFound.Caption) + 1
'
'                    On Error GoTo errKill
'                    'get filename, after kill, sFile is null?
'                    Dim tempFN As String
'                    tempFN = sFile.Path ' & "\" & sFile.Name
'
'                    'remove file with force
'                    'Kill sFile
'                    sFile.Delete True
'                    'add to log after kill process
'                    'frmTest.txtLog.Text = frmTest.txtLog.Text & "  Removed " & tempFN & vbCrLf
'                    frmTest.lblzabl.Caption = Int(frmTest.lblzabl.Caption) + 1
'                    strScanDetail = strScanDetail & " Virus Cleaned<br>"
'                    Call UpdateDetail(strScanDetail, frmTest.WebBrowser1)
'                    GoTo endFor
'errKill:
'                    'add to log after kill error
'                    'frmTest.txtLog.Text = frmTest.txtLog.Text & "  Cannot removed " & tempFN & vbCrLf
'                    strScanDetail = strScanDetail & "<font Size=3 Color=YELLOW><i> Virus Cannot Be Cleaned</i></font><br>"
'                    Call UpdateDetail(strScanDetail, frmTest.WebBrowser1)
'                    Exit For
'                End If
'            Next i
'        End If
'endFor:
'    Next


        'scan algorithm 2
        '----------------
        
        'compare with database
        For i = 0 To UBound(VSig)
            If GetInputState() <> 0 Then DoEvents
            If sCRC = VSig(i).Value Then   'start cleaning
                'add to log
                ' strScanDetail = strScanDetail & "Virus Found: <Font Size=3 Color=RED>" & VSig(i).Name & "</font><br>"
                ' strScanDetail = strScanDetail & " File: <Font Size=3 Color=ORANGE><i>" & sFile & "</i></font><br>"

                ' Call UpdateDetail(strScanDetail, frmTest.WebBrowser1)

                'add 1 to counter
                frmTest.lblFound.Caption = Int(frmTest.lblFound.Caption) + 1
                'frmTest.SB.Panels(3).Text = frmTest.lblFound.Caption
                'On Error GoTo errKill
                'get filename, after kill, sFile is null?
                Dim tempFN As String
                tempFN = sFile.Path ' & "\" & sFile.Name
    If QarNoDelete = True Then
        '��� �������������
101:
Dim d As Variant
    d = Format(Now, "ddMMSSHHMMSS")
                If Dir$(sFile.Name + "." + d) <> "" Then
                    GoTo 101 '���� ���� ����� ���� � ��������� �� ���������� ����� ���
                End If
      sFile.Move App.Path + "\quarantine" + "\" + sFile.Name + "." + d

        
        
        
        'sFile.Attributes = Normal
        'sFile.Move App.Path + "\quarantine" + "\" + sFile.Name + ".###"
       LogPrint sFile.Name + "." + d + "-move to quarantine"
    Else
          'remove file with force
                sFile.DELETE True
                LogPrint sFile.Name + "-cleaned"
    End If
                 'check whether the virus is cleaned or not; if not, go to errKill to show Error Cleaning
                If FileorFolderExists(tempFN) = True Then GoTo errKill
                'Kill sFile.Path
                'add to log after kill process
                'frmTest.txtLog.Text = frmTest.txtLog.Text & "  Removed " & tempFN & vbCrLf
                frmTest.lblZabl.Caption = Int(frmTest.lblZabl.Caption) + 1
                'frmTest.SB.Panels(4).Text = frmTest.lblzabl.Caption
                ' strScanDetail = strScanDetail & " Virus Cleaned<br>"
                ' Call UpdateDetail(strScanDetail, frmTest.WebBrowser1)
                'Set X = frmTest.lvVirusFound.ListItems.Add(, , VSig(i).Name, 2, 2)
                'X.SubItems(1) = tempFN
                If QarNoDelete = True Then
                   'X.SubItems(2) = "Moved"
                     ' ion sFile.Name + "=>" + tempFN
                   'LogPrint VSig(i).Name + "-moved"
                 Else
                  ' X.SubItems(2) = "Cleaned"
                   'LogPrint VSig(i).Name + "-cleaned"
                End If
                'Set X = Nothing
                GoTo endFor
errKill:
                'add to log after kill error
                'frmTest.txtLog.Text = frmTest.txtLog.Text & "  Cannot removed " & tempFN & vbCrLf
                ' strScanDetail = strScanDetail & "<font Size=3 Color=YELLOW><i> Virus Cannot Be Cleaned</i></font><br>"
                ' Call UpdateDetail(strScanDetail, frmTest.WebBrowser1)
                'Set X = frmTest.lvVirusFound.ListItems.Add(, , VSig(i).Name, 3, 3)
                'X.SubItems(1) = tempFN
                 If QarNoDelete = True Then
                   ' X.SubItems(2) = "Moved Failed"
                    LogPrint VSig(i).Name + "-Moved Failed"
                Else
                   ' X.SubItems(2) = "Clean Failed"
                    LogPrint VSig(i).Name + "-Clean Failed"
                End If
                'Set X = Nothing
                Exit For
            End If
        Next i
endFor:
    Next
    
    'scan subfolders
    For Each sFolder In sFolders
        DoEvents
        If blnScan = False Then Exit For
        ScanFile (sFolder)
    Next
    
    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    
End Sub

Public Sub ScanFileProc(ByVal sPath As String, ByVal mi_file As String, proccAidi As Long)
    On Error Resume Next
     Dim fso As New FileSystemObject
    'Dim x As ListItem
    'main folder
    Dim mFolder As Folder
    'files and folders collections
    Dim sFolders As Folders
    Dim sFiles As Files
    'for loop variables
    Dim sFolder As Folder
    Dim sFile As File
        'get main folder
    Set mFolder = fso.GetFolder(sPath)
    'get subfolders in main folder
    Set sFolders = mFolder.SubFolders
    'get files in main folder
    Set sFiles = mFolder.Files
         'scan virus
        Dim sCRC As String
        sCRC = CRC.GetCRC(mi_file)
        Dim i As Long
           Debug.Print sCRC
        'compare with database
        'frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
        For i = 0 To UBound(VSig)
       
        'frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
                If sCRC = VSig(i).Value Then   'start cleaning
                LogPrint mi_file + "-found virus in memory"
                'add to log
                    
                    
                   If MsgBox("Found virus i memory.Infected file" + vbCrLf + mi_file + vbCrLf + "Terminate process ?", vbCritical + vbYesNo) = vbYes Then
                        frmTest.lblZabl.Caption = Int(frmTest.lblZabl.Caption) + 1
                        'frmTest.SB.Panels(4).Text = frmTest.lblZabl.Caption
                        Process_Kill proccAidi
                        Process_Kill proccAidi
                        LogPrint mi_file + "-terminate virus process"
                    End If
            End If
                
                    Next i

    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    
End Sub
Public Sub ScanFileProcSplash(ByVal sPath As String, ByVal mi_file As String, proccAidi As Long)
    '����� ��� ��������� �����,����� �� ������ ���� ������� �����
    On Error Resume Next

     Dim fso As New FileSystemObject
    
    'main folder
    Dim mFolder As Folder
    'files and folders collections
    Dim sFolders As Folders
    Dim sFiles As Files
    'for loop variables
    Dim sFolder As Folder
    Dim sFile As File
        'get main folder
    Set mFolder = fso.GetFolder(sPath)
    'get subfolders in main folder
    Set sFolders = mFolder.SubFolders
    'get files in main folder
    Set sFiles = mFolder.Files
         'scan virus
        Dim sCRC As String
        sCRC = CRC.GetCRC(mi_file)
        If Left(mi_file, 1) = "/" Then
            Exit Sub
        End If
        'MsgBox mi_file
        Dim i As Long
           Debug.Print sCRC
        'compare with database
        For i = 0 To UBound(VSig)
        
        If monitoring = False Then Exit Sub
        
       DoEvents
      'frmTest.lblPath.Caption = mi_file
      frmTest.lblPathText.Text = mi_file
        'frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
                If sCRC = VSig(i).Value Then   'start cleaning
             
'                LogPrint mi_file + "-found virus in memory"
                'add to log
                    frmTest.lblFound.Caption = Int(frmTest.lblFound.Caption) + 1
                     'procFoundSpl = procFoundSpl + 1
                      
                  ' If MsgBox("� ������ ��������� �����." + vbCrLf + VSig(i).Name + vbCrLf + "����������� ?", vbCritical + vbYesNo + vbSystemModal, pr) = vbYes Then
                       frmTest.lblZabl.Caption = Int(frmTest.lblZabl.Caption) + 1
                            'Call KillProcess(proccAidi)
                            frmTest.Text1.Text = "� ������ ��������� �����." + vbCrLf + VSig(i).Name
                            Call frmTest.message
                            Process_Kill (proccAidi)
                           Dim fn1 As Long
                    fn1 = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonQv"))
                    If fn1 = 1 Then
                         Call moveToQarant(mi_file, VSig(i).Name)
                    Else
                        Call killDelVir(mi_file, VSig(i).Name)
                    End If

                         
                                    mi_file = ""
                                    proccAidi = 0
                                    Exit Sub
'programs.Process_Kill (mi_file)
                     ' processes.KillProcess proccAidi
                       '######�����
'                        LogPrint mi_file + "-terminate virus process"
                        ' procCleanSpl = procCleanSpl + 1
                  '  End If
            End If
                
                    Next i

    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    frmTest.Text1.Text = ""
End Sub

Public Function ScanFilepR(ByVal sPath As String, ByVal mi_file As String, proccAidi As Long) As Boolean
    '����� ��� ��������� �����,����� �� ������ ���� ������� �����
    On Error Resume Next
    ScanFilepR = False
    frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
     Dim fso As New FileSystemObject
    
    'main folder
    Dim mFolder As Folder
    'files and folders collections
    Dim sFolders As Folders
    Dim sFiles As Files
    'for loop variables
    Dim sFolder As Folder
    Dim sFile As File
        'get main folder
    Set mFolder = fso.GetFolder(sPath)
    'get subfolders in main folder
    Set sFolders = mFolder.SubFolders
    'get files in main folder
    Set sFiles = mFolder.Files
         'scan virus
        Dim sCRC As String
        sCRC = CRC.GetCRC(mi_file)
        If Left(mi_file, 1) = "/" Then
            Exit Function
        End If
        'MsgBox mi_file
        Dim i As Long
           Debug.Print sCRC
        'compare with database
        For i = 0 To UBound(VSig)
        
        If monitoring = False Then Exit Function
        
       DoEvents
     ' frmTest.lblPath.Caption = mi_file
       frmTest.lblPathText.Text = mi_file
                If sCRC = VSig(i).Value Then   'start cleaning
             
'                LogPrint mi_file + "-found virus in memory"
                'add to log
                    frmTest.lblFound.Caption = Int(frmTest.lblFound.Caption) + 1
                     'procFoundSpl = procFoundSpl + 1
                      
                   'If MsgBox("� ������ ��������� �����." + vbCrLf + VSig(i).Name + vbCrLf + "����������� ?", vbCritical + vbYesNo + vbSystemModal, pr) = vbYes Then
                       frmTest.lblZabl.Caption = Int(frmTest.lblZabl.Caption) + 1
                            'Call KillProcess(proccAidi)
                            frmTest.Text1.Text = "� ������ ��������� �����." + vbCrLf + VSig(i).Name
                            Call frmTest.message
                            Process_Kill (proccAidi)
                           Dim fn1 As Long
                    fn1 = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonQv"))
                    If fn1 = 1 Then
                         Call moveToQarant(mi_file, VSig(i).Name)
                    Else
                        Call killDelVir(mi_file, VSig(i).Name)
                    End If

                         
                                    mi_file = ""
                                    proccAidi = 0
                                    ScanFilepR = True
                                    Exit Function
'programs.Process_Kill (mi_file)
                     ' processes.KillProcess proccAidi
                       '######�����
'                        LogPrint mi_file + "-terminate virus process"
                        ' procCleanSpl = procCleanSpl + 1
                    'End If
                    
            End If
                'resume
  
                
                    Next i

    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    frmTest.Text1.Text = ""
End Function
Public Function ScanDLL(ByVal sPath As String, ByVal mi_file As String) As Boolean
    '����� ��� ��������� �����,����� �� ������ ���� ������� �����
    On Error Resume Next
ScanDLL = False
     Dim fso As New FileSystemObject
    
    'main folder
    Dim mFolder As Folder
    'files and folders collections
    Dim sFolders As Folders
    Dim sFiles As Files
    'for loop variables
    Dim sFolder As Folder
    Dim sFile As File
        'get main folder
    Set mFolder = fso.GetFolder(sPath)
    'get subfolders in main folder
    Set sFolders = mFolder.SubFolders
    'get files in main folder
    Set sFiles = mFolder.Files
         'scan virus
        Dim sCRC As String
        sCRC = CRC.GetCRC(mi_file)
        If Left(mi_file, 1) = "/" Then
            Exit Function
        End If
        'MsgBox mi_file
        Dim i As Long
           Debug.Print sCRC
        'compare with database
        For i = 0 To UBound(VSig)
        
        'If monitoring = False Then Exit Sub
        
       DoEvents
      'frmTest.lblPath.Caption = mi_file
      frmTest.lblPathText.Text = mi_file
        'frmTest.lblCount.Caption = Int(frmTest.lblCount.Caption) + 1
                If sCRC = VSig(i).Value Then   'start cleaning
             
'                LogPrint mi_file + "-found virus in memory"
                'add to log
                    frmTest.lblFound.Caption = Int(frmTest.lblFound.Caption) + 1
                     'procFoundSpl = procFoundSpl + 1
                      
                   If MsgBox("� ������ ��������� �����." + vbCrLf + VSig(i).Name + vbCrLf + "����������� ?", vbCritical + vbYesNo + vbSystemModal, pr) = vbYes Then
                       frmTest.lblZabl.Caption = Int(frmTest.lblZabl.Caption) + 1
                            'Call KillProcess(proccAidi)
                            frmTest.Text1.Text = "� ������ ��������� �����." + vbCrLf + VSig(i).Name
                            Call frmTest.message
                            'Process_Kill (proccAidi)
                                          
                         
                                    mi_file = ""

                                    ScanDLL = True
                                    Exit Function
'programs.Process_Kill (mi_file)
                     ' processes.KillProcess proccAidi
                       '######�����
'                        LogPrint mi_file + "-terminate virus process"
                        ' procCleanSpl = procCleanSpl + 1
                    End If
                    
            End If
                'resume
  
                
                    Next i

    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    frmTest.Text1.Text = ""
End Function
Sub moveToQarant(s12 As String, virn1 As String)
On Error Resume Next
If Dir$(App.Path + "\quarantine", vbDirectory) = "" Then
  MkDir App.Path + "\quarantine"
End If
     Dim fso As New FileSystemObject, txtfile, fil1, fil2
' ��� ��������� ����� � ����� C:\
Set fil1 = fso.GetFile(s12)
101:
Dim d As Variant
    d = Format(Now, "ddMMSSHHMMSS")
                If Dir$(App.Path + "\quarantine\" + d, vbNormal) <> "" Then
                    GoTo 101 '���� ���� ����� ���� � ��������� �� ���������� ����� ���
                End If
' ���������� ���� � ���������� \tmp
fil1.Move (App.Path + "\quarantine\" + d)
        
        
    If Dir$(s12, vbNormal) = "" Then
        frmTest.lblCleaned.Caption = Int(frmTest.lblCleaned.Caption) + 1
        LogQ s12 + "==>" + App.Path + "\quarantine\" + d
        frmTest.Text1.Text = "����� <" + virn1 + "> ��������� � ��������"
        Call frmTest.message
    Else
    LogPrint s12 + "-������ ��� ����������� � ��������"
            frmTest.Text1.Text = "������ ����������� " + vbCrLf + virn1 + vbCrLf + " � ��������"
        Call frmTest.message
    End If
Set fso = Nothing
    Set fil1 = Nothing
    
End Sub
Sub killDelVir(s13 As String, virn2 As String)
On Error Resume Next
  Dim fso As New FileSystemObject, txtfile, fil1, fil2
' ��� ��������� ����� � ����� C:\
Set fil1 = fso.GetFile(s13)
    fil1.DELETE True
If Dir$(s13, vbNormal) = "" Then
        frmTest.lblDelete.Caption = Int(frmTest.lblDelete.Caption) + 1
        LogPrint s13 + "-������"
                frmTest.Text1.Text = "����� <" + virn2 + ">  ������"
        Call frmTest.message
    Else
            LogPrint s13 + "-������ ��� ��������"
            frmTest.Text1.Text = "������ ��� �������� " + vbCrLf + virn2
        Call frmTest.message
    End If
Set fso = Nothing
    Set fil1 = Nothing
End Sub
Sub LogQ(sMessage6 As String)
'����� ������ ������������� � ��������
On Error GoTo 100
Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
'Open Ffile For Append As #nFile
ffile = App.Path + "\quarantine\removed.log"
If Dir$(App.Path + "\quarantine", vbDirectory) = "" Then
  MkDir App.Path + "\quarantine"
End If
If Dir$(ffile, vbNormal) <> "" Then
    If FileLen(ffile) >= 3145728 Then
       ' MsgBox "�������"
        Kill ffile
    End If
End If
    

Open ffile For Append As #nFile
Print #nFile, Format$(Now, "mm-dd-yy hh:mm:ss") + "--" + sMessage6
Close #nFile

Exit Sub
100:
Select Case Err
Case 52
    MsgBox "�� ���������� ��� �����", vbCritical, pr
Case 53
    MsgBox "���� �� ������", vbCritical, pr
Case 54
    MsgBox "�� ���������� ����� �����", vbCritical, pr
Case 57
    MsgBox "������ �����/������", vbCritical, pr
Case 61
    MsgBox "����������� �����. ���� �� ��� ���������", vbCritical, pr
    'End
Case 62
    MsgBox "������-�� ���������� ������ ����� �������� �����", vbCritical, pr
Case 67
    MsgBox "������� ����� �������� ������. �� ���� � ��� ��������", vbCritical, pr
    End
Case 68
    MsgBox "�� ���� � ���� ", vbCritical, pr
    'End
Case 70
    MsgBox "������ � ����� ��� �������� ��������", vbCritical, pr
    'End
Case 71
    MsgBox "���� �� �����", vbCritical, pr
Case 75
    
    If MsgBox("�� �� ���� � �������� � ���� ������ ��� ���������" & vbCrLf & ffile & vbCrLf & "�������� ��� ��������� ��-�� ������������� ����������. ���������� ?" + vbCrLf + Error$, vbYesNo + vbCritical, pr) = vbNo Then
        End
    End If
Case 76
    MsgBox "���� ��������� �� �����...�������� ��� ��������� ��-�� ������������� ������ ���������..." & ffile, vbCritical
    
Case 321
    MsgBox "�� ���������� ������ �����", vbCritical
'End
Case 3024
    MsgBox "����, �� ���� ����� ���� ����" & ffile, vbCritical, pr
'    End
Case 3176
    MsgBox "�� ���� ������� ����" & ffile, vbCritical, pr
Case 3179
    MsgBox "��������� ����� �����", vbCritical, pr
'End
Case 3180
    MsgBox "�� ���� �������� � ����", vbCritical, pr
'End
Case Else
    MsgBox "��������� ���������� ������ ��� �������� ����", vbCritical, pr
End Select

End Sub

Public Sub LogPrint(sMessage As String)
'����� ������ � ���� ��������� ������ �����������
On Error GoTo 100
Logging = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogMon"))
'CreateDwordValue GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogSizeMon"
If Logging = False Then
    Exit Sub '����� ���?
End If


If Dir$(App.Path + "\otchetMon.txt", vbNormal) <> "" Then
'��������
Dim FSize As Long
FSize = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogSizeMon"))
    If FileLen(App.Path + "\otchetMon.txt") >= FSize * 1024 Then
       'MsgBox "�������"
        Kill App.Path + "\otchetMon.txt"
    End If
End If

Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
'Open Ffile For Append As #nFile
ffile = App.Path + "\otchetMon.txt"
Open ffile For Append Access Write Lock Read Write As #nFile
Print #nFile, Format$(Now, "mm-dd-yy hh:mm:ss") + "--" + sMessage
Close #nFile

Exit Sub
100:
Select Case Err
Case 52
    MsgBox "�� ���������� ��� �����", vbCritical, pr
Case 53
    MsgBox "���� �� ������", vbCritical, pr
Case 54
    MsgBox "�� ���������� ����� �����", vbCritical, pr
Case 57
    MsgBox "������ �����/������", vbCritical, pr
Case 61
    MsgBox "����������� �����. ���� �� ��� ���������", vbCritical, pr
    'End
Case 62
    MsgBox "������-�� ���������� ������ ����� �������� �����", vbCritical, pr
Case 67
    MsgBox "������� ����� �������� ������. �� ���� � ��� ��������", vbCritical, pr
    End
Case 68
    MsgBox "�� ���� � ���� ", vbCritical, pr
    'End
Case 70
    MsgBox "������ � ����� ��� �������� ��������", vbCritical, pr
    'End
Case 71
    MsgBox "���� �� �����", vbCritical, pr
Case 75
    
    If MsgBox("�� �� ���� � �������� � ���� ������ ��� ���������" & vbCrLf & ffile & vbCrLf & "�������� ��� ��������� ��-�� ������������� ����������. ���������� ?", vbYesNo + vbCritical, pr) = vbNo Then
        'End
    End If
Case 76
    MsgBox "���� ��������� �� �����...�������� ��� ��������� ��-�� ������������� ������ ���������..." & ffile, vbCritical
    
Case 321
    MsgBox "�� ���������� ������ �����", vbCritical
'End
Case 3024
    MsgBox "����, �� ���� ����� ���� ����" & ffile, vbCritical, pr
'    End
Case 3176
    MsgBox "�� ���� ������� ����" & ffile, vbCritical, pr
Case 3179
    MsgBox "��������� ����� �����", vbCritical, pr
'End
Case 3180
    MsgBox "�� ���� �������� � ����", vbCritical, pr
'End
Case Else
    MsgBox "��������� ���������� ������ ��� �������� ����", vbCritical, pr
End Select
End Sub
Public Function FileorFolderExists(FolderOrFilename As String) As Boolean
    If PathFileExists(FolderOrFilename) = 1 Then
        FileorFolderExists = True
    ElseIf PathFileExists(FolderOrFilename) = 0 Then
        FileorFolderExists = False
    End If
End Function
'Format the values by removing everything except for the filename or path
Public Function FormatValue(sValue As String) As String

    Dim FileorPath As String
    'Fix up the file or path so that it's compatible with the FileorFolderExists function
    FileorPath = sValue

    'Find the start of the path or filename (Example:"h6j65ej(C:\Test)")
    If InStr(1, FileorPath, "C:\") Then FileorPath = Mid(FileorPath, InStr(1, FileorPath, "C:\"))
    If InStr(1, FileorPath, "c:\") Then FileorPath = Mid(FileorPath, InStr(1, FileorPath, "c:\"))

    'Remove everything after the path. This definitely doesn't work for all values.
    '(Example:"C:\blablablablablablablabla?5784846\84585")
    If InStr(1, FileorPath, "/") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "/") - 1)
    If InStr(1, FileorPath, "*") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "*") - 1)
    If InStr(1, FileorPath, "?") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "?") - 1)
    If InStr(1, FileorPath, Chr(34)) > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, Chr(34)) - 1)
    If InStr(1, FileorPath, "<") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "<") - 1)
    If InStr(1, FileorPath, ">") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ">") - 1)
    If InStr(1, FileorPath, "|") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "|") - 1)
    If InStr(1, FileorPath, ",") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ",") - 1)
    If InStr(1, FileorPath, "(") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "(") - 1)
    If InStr(1, FileorPath, ";") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ";") - 1)

    'some registry values somehow didn't contain "C:\"
    If InStr(1, FileorPath, "C:\") = 0 Then FileorPath = "C:\"

    'Remove everything before the path or file. The same as the other one except this is for specific extensions
    '(Example:"C:\lalalalalalalala\idfjb.dll\50")
    'I guess I could just do If InStr(1, FileorPath, ".exe") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe") + 2), If InStr(1, FileorPath, ".dll") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".dll") + 2), etc, but I'm not sure how that would work out

    If InStr(1, FileorPath, ".exe:") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe:") + 3)
    If InStr(1, FileorPath, ".EXE:") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE:") + 3)
    If InStr(1, FileorPath, ".EXE ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE ") + 3)
    If InStr(1, FileorPath, ".exe ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe ") + 3)
    If InStr(1, FileorPath, ".SYS ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".SYS ") + 3)
    If InStr(1, FileorPath, ".sys ") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".sys ") + 3)
    If InStr(1, FileorPath, ".EXE\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".EXE\") + 3)
    If InStr(1, FileorPath, ".exe\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".exe\") + 3)
    If InStr(1, FileorPath, ".DLL\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".DLL\") + 3)
    If InStr(1, FileorPath, ".dll\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".dll\") + 3)
    If InStr(1, FileorPath, ".OCX\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".OCX\") + 3)
    If InStr(1, FileorPath, ".ocx\") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ".ocx\") + 3)
    If InStr(1, FileorPath, "*") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, "*") - 1)
    If InStr(3, FileorPath, ":") > 0 Then FileorPath = Mid(FileorPath, 1, InStr(1, FileorPath, ":") - 1)

    '%1 is used for file associations
    '(Example:"C:\WINDOWS\NOTEPAD.EXE %1")
    FormatValue = Replace(FileorPath, " %1", "")
End Function
