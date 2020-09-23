Attribute VB_Name = "modScanVirus"
Option Explicit
Public maimCRC As String
Public WS As Workspace
Public db As Database
Public rs As Recordset
Public strDBPath As String
Public nameVir_old As String
Public nameVir As String
Public zp1 As String
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
'Public VSig() As VirusSig
'Public VSInfo As VS_Info
'Public Type VirusSig

  '  Name As String
  '  Type As String
  '  Value As String
  '  Action As String
  '  ActtionVal As String
  '
'End Type

'new DataType for Virus Signature Info
Public Type VS_Info
    
    VirusCount As Long
    LastUpdate As Date
    
End Type
'for faster DoEvents
Declare Function GetInputState Lib "USER32.dll" () As Long
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
     Dim CRC As New clsCRC
    CRC.BuildTable
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
        If sFile.Size > 2478208 Then GoTo endFor 'can't use Continue For
        
        'If GetInputState() <> 0 Then DoEvents  'faster but look stuck
        DoEvents
       
        'check if it is stopped
        If blnScan = False Then Exit For
        'show scanned file
        'FrmNet.lblPath.Caption = sFile
      '  FrmNet.lblPathText.Text = sFile
        'add 1 to counter scanned
       ' FrmNet.lblCount.Caption = Int(FrmNet.lblCount.Caption) + 1
        'FrmNet.SB.Panels(2).Text = FrmNet.lblCount.Caption
        'FrmNet.SB.Refresh
        'scan virus

        Dim sCRC As String
        sCRC = modFileManipulation.GetMD5(CStr(sFile))
        Debug.Print sCRC
                'Çäåñü âîçüìåì ñòàðóþ MD5
        Dim Old_CRC As String
        Old_CRC = CRC.GetCRC(CStr(sFile))
          Debug.Print Old_CRC
          Dim i As Long
        LogPrint sFile.Name
        'scan algorithm 1
        '----------------
'        For i = 0 To UBound(VSig)
            If GetInputState() <> 0 Then DoEvents
              If get_base(sCRC) = True Then
            'If sCRC = VSig(i).Value Then   'start cleaning
                'add to log
                ' strScanDetail = strScanDetail & "Virus Found: <Font Size=3 Color=RED>" & nameVir & "</font><br>"
                ' strScanDetail = strScanDetail & " File: <Font Size=3 Color=ORANGE><i>" & sFile & "</i></font><br>"

                ' Call UpdateDetail(strScanDetail, FrmNet.WebBrowser1)

                'add 1 to counter
               ' FrmNet.lblFound.Caption = Int(FrmNet.lblFound.Caption) + 1
                'FrmNet.SB.Panels(3).Text = FrmNet.lblFound.Caption
                'On Error GoTo errKill
                'get filename, after kill, sFile is null?
                Dim tempFN As String
                tempFN = sFile.path ' & "\" & sFile.Name
    If QarNoDelete = True Then
        'âàõ ïåðåèìåíîâàòü
101:
Dim d As Variant
    d = Format(Now, "ddMMSSHHMMSS")
                If Dir$(sFile.Name + "." + d) <> "" Then
                    GoTo 101 'åñëè åñòü òàêîé ôàéë â êàðàíòèíå òî ãåíåðèðóåì íîâîå èìÿ
                End If
      sFile.Move App.path + "\quarantine" + "\" + sFile.Name + "." + d

        
        
        
        'sFile.Attributes = Normal
        'sFile.Move App.Path + "\quarantine" + "\" + sFile.Name + ".###"
       LogPrint sFile.Name + "." + d + "-move to quarantine"
    Else
          'remove file with force
                sFile.Delete True
                LogPrint sFile.Name + "-cleaned"
    End If
                 'check whether the virus is cleaned or not; if not, go to errKill to show Error Cleaning
                If FileorFolderExists(tempFN) = True Then GoTo errKill
                'Kill sFile.Path
                'add to log after kill process
                'FrmNet.txtLog.Text = FrmNet.txtLog.Text & "  Removed " & tempFN & vbCrLf
                'FrmNet.lblZabl.Caption = Int(FrmNet.lblZabl.Caption) + 1
                'FrmNet.SB.Panels(4).Text = FrmNet.lblzabl.Caption
                ' strScanDetail = strScanDetail & " Virus Cleaned<br>"
                ' Call UpdateDetail(strScanDetail, FrmNet.WebBrowser1)
                'Set X = FrmNet.lvVirusFound.ListItems.Add(, , nameVir, 2, 2)
                'X.SubItems(1) = tempFN
                If QarNoDelete = True Then
                   'X.SubItems(2) = "Moved"
                     ' ion sFile.Name + "=>" + tempFN
                   'LogPrint nameVir + "-moved"
                 Else
                  ' X.SubItems(2) = "Cleaned"
                   'LogPrint nameVir + "-cleaned"
                End If
                'Set X = Nothing
                GoTo endFor
errKill:
                'add to log after kill error
                'FrmNet.txtLog.Text = FrmNet.txtLog.Text & "  Cannot removed " & tempFN & vbCrLf
                ' strScanDetail = strScanDetail & "<font Size=3 Color=YELLOW><i> Virus Cannot Be Cleaned</i></font><br>"
                ' Call UpdateDetail(strScanDetail, FrmNet.WebBrowser1)
                'Set X = FrmNet.lvVirusFound.ListItems.Add(, , nameVir, 3, 3)
                'X.SubItems(1) = tempFN
                 If QarNoDelete = True Then
                   ' X.SubItems(2) = "Moved Failed"
                    LogPrint nameVir + "-Moved Failed"
                Else
                   ' X.SubItems(2) = "Clean Failed"
                    LogPrint nameVir + "-Clean Failed"
                End If
                'Set X = Nothing
                Exit For
               Else

         If yesVir_old(Old_CRC) Then
       
             If MsgBox("Ïîäîçðåíèå íà âèðóñíóþ àêòèâíîñòü, õàðàêòåðíóþ äëÿ " + vbCrLf + nameVir_old + vbCrLf + "Áëîêèðîâàòü ?", vbCritical + vbYesNo, pr) = vbYes Then
                'Process_Kill (proccAidi)
                FrmNet.message 2, "Çàáëîêèðîâàí ïîäîçðèòåëüíûé ïðîöåññ,õàðàêòåðíûé äëÿ" + vbCrLf + nameVir_old + vbCrLf + sFile
                LogPrint "Çàáëîêèðîâàí ïîäîçðèòåëüíûé ïðîöåññ,õàðàêòåðíûé äëÿ-" + nameVir_old + vbCrLf + sFile
            Else
                FrmNet.message 2, nameVir_old + vbCrLf + "Ïðîèãíîðèðîâàíî" + vbCrLf + sFile
                LogPrint nameVir_old + vbCrLf + "-ïðîèãíîðèðîâàíî" + vbCrLf + sFile
            End If

            End If
End If
 
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
Function yesVir_old(sim As String) As Boolean
   
On Error GoTo 10
  yesVir_old = False
If Dir$(App.path + "\old.bmd") = "" Then
    Exit Function
End If
    Dim miNumBase As Integer
    Dim sMD5 As String
    Dim zb As String
    Dim bm As Integer
    miNumBase = FreeFile
Open App.path + "\old.bmd" For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD5
            bm = InStr(1, sMD5, sim, vbTextCompare)
    If bm <> 0 Then
        zb = Right$(sMD5, Len(sMD5) - 9)
        yesVir_old = True
        nameVir_old = CStr(zb)
       'MsgBox "" + CStr(zb)
        Close #miNumBase
        Exit Function
     End If
     Wend
10:

   yesVir_old = False
  Close #miNumBase

End Function
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
        sCRC = modFileManipulation.GetMD5(CStr(mi_file))
        Dim i As Long
           Debug.Print sCRC
           MsgBox ""
        'compare with database
        'FrmNet.lblCount.Caption = Int(FrmNet.lblCount.Caption) + 1
        'For i = 0 To UBound(VSig)
       
        'FrmNet.lblCount.Caption = Int(FrmNet.lblCount.Caption) + 1
          If get_base(sCRC) = True Then
              '  If sCRC = VSig(i).Value Then   'start cleaning
                LogPrint mi_file + "-found virus in memory"
                'add to log
                    
                    
                   If MsgBox("Found virus i memory.Infected file" + vbCrLf + mi_file + vbCrLf + "Terminate process ?", vbCritical + vbYesNo) = vbYes Then
                        'FrmNet.lblZabl.Caption = Int(FrmNet.lblZabl.Caption) + 1
                        'FrmNet.SB.Panels(4).Text = FrmNet.lblZabl.Caption
                        'Process_Kill proccAidi
                        Process_Kill proccAidi
                        LogPrint mi_file + "-terminate virus process"
                    End If
            End If
                
             '       Next i

    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    
End Sub

Public Function ScanFilepR(ByVal sPath As String, ByVal mi_file As String, proccAidi As Long) As Boolean
    '÷èñòî äëÿ ñòàðòîâîé ôîðìû,÷òîáû íå ìåøàòü ñþäà ãëàâíóþ ôîðìó
    On Error Resume Next
    ScanFilepR = False
    'FrmNet.lblCount.Caption = Int(FrmNet.lblCount.Caption) + 1
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
        sCRC = modFileManipulation.GetMD5(CStr(mi_file))
        Debug.Print sCRC
               'Çäåñü âîçüìåì ñòàðóþ MD5
          Dim Old_CRC As String
        Old_CRC = CRC.GetCRC(CStr(mi_file))
          Debug.Print Old_CRC + "-ñòàðàÿ2"
        If Left(mi_file, 1) = "/" Then
            Exit Function
        End If
        'MsgBox mi_file
        Dim i As Long
       DoEvents

             
              If get_base(sCRC) = True Then
'                LogPrint mi_file + "-found virus in memory"
'MsgBox ""
                'add to log
                    'FrmNet.lblFound.Caption = Int(FrmNet.lblFound.Caption) + 1
                     'procFoundSpl = procFoundSpl + 1
                      
                 ' If MsgBox("Â ïàìÿòè îáíàðóæåí âèðóñ." + vbCrLf + nameVir + vbCrLf + "Áëîêèðîâàòü ?", vbCritical + vbYesNo + vbSystemModal, pr) = vbYes Then
                  'End If
                      ' FrmNet.lblZabl.Caption = Int(FrmNet.lblZabl.Caption) + 1
                            'Call KillProcess(proccAidi)
                          '  FrmNet.Text1.Text = "Â ïàìÿòè îáíàðóæåí âèðóñ." + vbCrLf + nameVir
                            'Call FrmNet.message
                            'ResumeThreads (proccAidi)
                           'Thread_Resume (proccAidi)
                            Process_Kill (proccAidi)
                           Dim fn1 As Long
                    fn1 = CStr(GetSetting("Belyash AV", "Options", "MonQv", "True"))
                    
                    If fn1 = "True" Then
                         Call moveToQarant(mi_file, nameVir)
                   Else
                        Call killDelVir(mi_file, nameVir)
                   End If

                         
                                    mi_file = ""
                                    
                                    ScanFilepR = True
                                    Exit Function
'programs.Process_Kill (mi_file)
                     ' processes.KillProcess proccAidi
                       '######áýýýý
'                        LogPrint mi_file + "-terminate virus process"
                        ' procCleanSpl = procCleanSpl + 1
                    'End If
                    Else
 
         If yesVir_old(Old_CRC) Then
       
             If MsgBox("Ïîäîçðåíèå íà âèðóñíóþ àêòèâíîñòü, õàðàêòåðíóþ äëÿ " + vbCrLf + nameVir_old + vbCrLf + "Áëîêèðîâàòü ?", vbCritical + vbSystemModal + vbYesNo, pr) = vbYes Then
                Process_Kill (proccAidi)
                FrmNet.message 2, "Çàáëîêèðîâàí ïîäîçðèòåëüíûé ïðîöåññ,õàðàêòåðíûé äëÿ" + vbCrLf + nameVir_old
                LogPrint "Çàáëîêèðîâàí ïîäîçðèòåëüíûé ïðîöåññ,õàðàêòåðíûé äëÿ-" + nameVir_old
            Else
                FrmNet.message 2, nameVir_old + vbCrLf + "Ïðîèãíîðèðîâàíî"
                LogPrint nameVir_old + vbCrLf + "-ïðîèãíîðèðîâàíî"
            End If

            End If
            
End If
                'resume
  
                
                '    Next i

    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
   ' FrmNet.Text1.Text = ""
End Function

Function get_base(CRC_checkSUM As String) As Boolean
'ïåðåáîð âñåõ âèðóñíûõ áàç
On Error GoTo 100
get_base = False
Dim s As String
s = Dir(App.path + "\*.bVb")
' Êîä îáðàáîòêè áàç
Do While s <> ""
   Debug.Print s
         If yesVir(CRC_checkSUM, s) = True Then
           get_base = True
           Exit Function
         End If
         
    s = Dir
Loop

Exit Function
100:
MsgBox "" + Error$
End Function
Public Function ScanDLL(ByVal sPath As String, ByVal mi_file As String) As Boolean
    '÷èñòî äëÿ ñòàðòîâîé ôîðìû,÷òîáû íå ìåøàòü ñþäà ãëàâíóþ ôîðìó
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
        'Dim sCRC As String
       ' sCRC = modFileManipulation.GetMD5(CStr(mi_file)
       ' If Left(mi_file, 1) = "/" Then
            Exit Function
      '  End If
        'MsgBox mi_file
        'Dim i As Long
         '  Debug.Print sCRC
        'compare with database
      '  For i = 0 To UBound(nameVir)
        
        'If monitoring = False Then Exit Sub
        
       DoEvents
      'FrmNet.lblPath.Caption = mi_file
      
        'FrmNet.lblCount.Caption = Int(FrmNet.lblCount.Caption) + 1
           '     If sCRC = VSig(i).Value Then   'start cleaning
             
'                LogPrint mi_file + "-found virus in memory"
                'add to log
                   ' FrmNet.lblFound.Caption = Int(FrmNet.lblFound.Caption) + 1
                     'procFoundSpl = procFoundSpl + 1
                      
                   'If MsgBox("Â ïàìÿòè îáíàðóæåí âèðóñ." + vbCrLf + nameVir + vbCrLf + "Áëîêèðîâàòü ?", vbCritical + vbYesNo + vbSystemModal, pr) = vbYes Then
               '        FrmNet.lblZabl.Caption = Int(FrmNet.lblZabl.Caption) + 1
                            'Call KillProcess(proccAidi)
                     '       FrmNet.Text1.Text = "Â ïàìÿòè îáíàðóæåí âèðóñ." + vbCrLf + nameVir
                   '         Call FrmNet.message
                            'Process_Kill (proccAidi)
                                          
                         
                           '         mi_file = ""

                         ''           ScanDLL = True
                                 '   Exit Function
'programs.Process_Kill (mi_file)
                     ' processes.KillProcess proccAidi
                       '######áýýýý
'                        LogPrint mi_file + "-terminate virus process"
                        ' procCleanSpl = procCleanSpl + 1
                 '  End If
                    
        '    End If
                'resume
  
                
                  '  Next i

    'clear variables
 '   Set fso = Nothing
'    Set mFolder = Nothing
'    Set sFolders = Nothing
 '   Set sFiles = Nothing
 '   Set sFolder = Nothing
 '   Set sFile = Nothing
 '   FrmNet.Text1.Text = ""
End Function
Sub moveToQarant(s12 As String, virn1 As String)
On Error Resume Next
If Dir$(App.path + "\quarantine", vbDirectory) = "" Then
  MkDir App.path + "\quarantine"
End If
     Dim fso As New FileSystemObject, txtfile, fil1, fil2
' Êîä îáðàáîòêè ôàéëà â êîðíå C:\
Set fil1 = fso.GetFile(s12)
101:
Dim d As Variant
    d = Format(Now, "ddMMSSHHMMSS")
                If Dir$(App.path + "\quarantine\" + d, vbNormal) <> "" Then
                    GoTo 101 'åñëè åñòü òàêîé ôàéë â êàðàíòèíå òî ãåíåðèðóåì íîâîå èìÿ
                End If
' Ïåðåìåùàåì ôàéë â äèðåêòîðèþ \tmp
fil1.Move (App.path + "\quarantine\" + d)
        
        
    If Dir$(s12, vbNormal) = "" Or Dir$(s12, vbArchive) = "" Then
        
        LogQ s12 + "==>" + App.path + "\quarantine\" + d
        FrmNet.message 1, "Âèðóñ <" + virn1 + "> ïåðåìåùåí â êàðàíòèí"
        
    Else
    LogPrint s12 + "-îøèáêà ïðè ïåðåìåùåíèè â êàðàíòèí"
            FrmNet.message 1, "Îøèáêà ïåðåìåùåíèÿ " + vbCrLf + virn1 + vbCrLf + " â êàðàíòèí"
        
    End If
Set fso = Nothing
    Set fil1 = Nothing
    
End Sub
Sub killDelVir(s13 As String, virn2 As String)
On Error Resume Next
  Dim fso As New FileSystemObject, txtfile, fil1, fil2
' Êîä îáðàáîòêè ôàéëà â êîðíå C:\
Set fil1 = fso.GetFile(s13)
    fil1.Delete True
If Dir$(s13, vbNormal) = "" Or Dir$(s13, vbNormal + vbArchive) = "" Then
       ' FrmNet.lblDelete.Caption = Int(FrmNet.lblDelete.Caption) + 1
        LogPrint s13 + "-óäàëåí"
               ' FrmNet.Text1.Text = "Âèðóñ <" + virn2 + ">  óäàëåí"
         FrmNet.message 1, "Âèðóñ <" + virn2 + ">  óäàëåí"
    Else
            LogPrint s13 + "-îøèáêà ïðè óäàëåíèè"
          '  FrmNet.Text1.Text = "Îøèáêà ïðè óäàëåíèè " + vbCrLf + virn2
         FrmNet.message 3, "Îøèáêà ïðè óäàëåíèè " + vbCrLf + virn2
         Debug.Print Error$
    End If
Set fso = Nothing
    Set fil1 = Nothing
End Sub
Sub LogQ(sMessage6 As String)
'ïðîöà çàïèñè ïåðåìåùàåìîãî â êàðàíòèí
On Error GoTo 100
Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
'Open Ffile For Append As #nFile
ffile = App.path + "\quarantine\removed.log"
If Dir$(App.path + "\quarantine", vbDirectory) = "" Then
  MkDir App.path + "\quarantine"
End If
If Dir$(ffile, vbNormal) <> "" Then
    If FileLen(ffile) >= 3145728 Then
       ' MsgBox "îáðåçàë"
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
    MsgBox "Íå ïðàâèëüíîå èìÿ ôàéëà", vbCritical, pr
Case 53
    MsgBox "Ôàéë ëîãà íå íàéäåí", vbCritical, pr
Case 54
    MsgBox "Íå ïðàâèëüíûé ðåæèì ôàéëà", vbCritical, pr
Case 57
    MsgBox "Îøèáêà ââîäà/âûâîäà", vbCritical, pr
Case 61
    MsgBox "Ïðåïîëíåíèå äèñêà. Ïîðà áû åãî ïî÷èñòèòü", vbCritical, pr
    'End
Case 62
    MsgBox "Ïî÷åìó-òî ïðîèçîøåëà çàïèñü ïîñëå çàêðûòèÿ ôàéëà", vbCritical, pr
Case 67
    MsgBox "Ñëèøêîì ìíîãî îòêðûòûõ ôàéëîâ. Íå ìîãó ÿ òàê ðàáîòàòü", vbCritical, pr
    End
Case 68
    MsgBox "Íå âèæó ÿ äèñê ", vbCritical, pr
    'End
Case 70
    MsgBox "Äîñòóï ê äèñêó èëè êàòàëîãó çàïðåùåí", vbCritical, pr
    'End
Case 71
    MsgBox "Äèñê íå ãîòîâ", vbCritical, pr
Case 75
    
    If MsgBox("Íó íå ìîãó ÿ ðàáîòàòü ñ ýòèì ôàéëîì èëè êàòàëîãîì" & vbCrLf & ffile & vbCrLf & "Âîçìîæíî ýòî ïðîèçîøëî èç-çà âìåøàòåëüñòâà àíòèâèðóñà. Ïðîäîëæèòü ?" + vbCrLf + Error$, vbYesNo + vbCritical, pr) = vbNo Then
        End
    End If
Case 76
    MsgBox "Ïóòü îïðåäåëåí íå âåðíî...Âîçìîæíî ýòî ïðîèçîøëî èç-çà âìåøàòåëüñòâà äðóãîé ïðîãðàììû..." & ffile, vbCritical
    
Case 321
    MsgBox "Íå ïðàâèëüíûé ôîðìàò ôàéëà", vbCritical
'End
Case 3024
    MsgBox "×åðò, íå ìîãó íàéòè ýòîò ôàéë" & ffile, vbCritical, pr
'    End
Case 3176
    MsgBox "Íå ìîãó îòêðûòü ôàéë" & ffile, vbCritical, pr
Case 3179
    MsgBox "Îøèáî÷íûé êîíåö ôàéëà", vbCritical, pr
'End
Case 3180
    MsgBox "Íå ìîãó çàïèñàòü â ôàéë", vbCritical, pr
'End
Case Else
    MsgBox "Ïðîèçîøëà âíóòðåííÿÿ îøèáêà ïðè îòêðûòèè ôàëà", vbCritical, pr
End Select

End Sub

Public Sub LogPrint(sMessage As String)
'ïðîöà çàïèñè â ôàéë èñõîäíîãî òåêñòà ìàêðîâèðóñà
On Error GoTo 100
'Logging = Val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "LogMon"))

Dim a4 As String
a4 = CStr(GetSetting("Belyash AV", "Options", "LogMon", "True"))
If a4 = "false" Then
   Exit Sub 'âåñòè ëîã?
End If


If Dir$(App.path + "\otchetMon.txt", vbNormal) <> "" Then
'îáðåçàåì
Dim FSize As Long
FSize = CLng(GetSetting("Belyash AV", "Options", "LogSizeMon", "1024496"))
    If FileLen(App.path + "\otchetMon.txt") >= FSize Then
       'MsgBox "îáðåçàë"
        Kill App.path + "\otchetMon.txt"
    End If
End If

Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
'Open Ffile For Append As #nFile
ffile = App.path + "\otchetMon.txt"
Open ffile For Append Access Write Lock Read Write As #nFile
Print #nFile, Format$(Now, "mm-dd-yy hh:mm:ss") + "--" + sMessage
Close #nFile

Exit Sub
100:
Select Case Err
Case 52
    MsgBox "Íå ïðàâèëüíîå èìÿ ôàéëà", vbCritical, pr
Case 53
    MsgBox "Ôàéë ëîãà íå íàéäåí", vbCritical, pr
Case 54
    MsgBox "Íå ïðàâèëüíûé ðåæèì ôàéëà", vbCritical, pr
Case 57
    MsgBox "Îøèáêà ââîäà/âûâîäà", vbCritical, pr
Case 61
    MsgBox "Ïðåïîëíåíèå äèñêà. Ïîðà áû åãî ïî÷èñòèòü", vbCritical, pr
    'End
Case 62
    MsgBox "Ïî÷åìó-òî ïðîèçîøåëà çàïèñü ïîñëå çàêðûòèÿ ôàéëà", vbCritical, pr
Case 67
    MsgBox "Ñëèøêîì ìíîãî îòêðûòûõ ôàéëîâ. Íå ìîãó ÿ òàê ðàáîòàòü", vbCritical, pr
    End
Case 68
    MsgBox "Íå âèæó ÿ äèñê ", vbCritical, pr
    'End
Case 70
    MsgBox "Äîñòóï ê äèñêó èëè êàòàëîãó çàïðåùåí", vbCritical, pr
    'End
Case 71
    MsgBox "Äèñê íå ãîòîâ", vbCritical, pr
Case 75
    
    If MsgBox("Íó íå ìîãó ÿ ðàáîòàòü ñ ýòèì ôàéëîì èëè êàòàëîãîì" & vbCrLf & ffile & vbCrLf & "Âîçìîæíî ýòî ïðîèçîøëî èç-çà âìåøàòåëüñòâà àíòèâèðóñà. Ïðîäîëæèòü ?", vbYesNo + vbCritical, pr) = vbNo Then
        'End
    End If
Case 76
    MsgBox "Ïóòü îïðåäåëåí íå âåðíî...Âîçìîæíî ýòî ïðîèçîøëî èç-çà âìåøàòåëüñòâà äðóãîé ïðîãðàììû..." & ffile, vbCritical
    
Case 321
    MsgBox "Íå ïðàâèëüíûé ôîðìàò ôàéëà", vbCritical
'End
Case 3024
    MsgBox "×åðò, íå ìîãó íàéòè ýòîò ôàéë" & ffile, vbCritical, pr
'    End
Case 3176
    MsgBox "Íå ìîãó îòêðûòü ôàéë" & ffile, vbCritical, pr
Case 3179
    MsgBox "Îøèáî÷íûé êîíåö ôàéëà", vbCritical, pr
'End
Case 3180
    MsgBox "Íå ìîãó çàïèñàòü â ôàéë", vbCritical, pr
'End
Case Else
    MsgBox "Ïðîèçîøëà âíóòðåííÿÿ îøèáêà ïðè îòêðûòèè ôàëà", vbCritical, pr
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
Function yesVir(CFc As String, my_base As String) As Boolean
'On Error GoTo 10

'On Error GoTo 10
yesVir = False
    Dim miNumBase As Integer
    Dim sMD5 As String
    Dim zb As String
    Dim bm As Integer
    miNumBase = FreeFile
     Open my_base For Input As #miNumBase
        While Not EOF(miNumBase)
             Line Input #miNumBase, sMD5
             
            bm = InStr(1, sMD5, CFc, vbTextCompare)
    If bm <> 0 Then
        zb = Right$(sMD5, Len(sMD5) - 33)
        yesVir = True
        nameVir = CStr(zb)
       'MsgBox "" + CStr(zb)
        Close #miNumBase
        Exit Function
     End If
     Wend
10:

   yesVir = False
  Close #miNumBase

End Function

Sub selfTesting()
On Error GoTo 200
If Dir$(App.path + "\mon.sec", vbNormal) = "" Then
    If MsgBox("Â êàòàëîãå ñ ïðîãðàììîé îòñóòñòâóåò ôàéë ñàìîïðîâåðêè.Âîçìîæíî çàðàæåíèå ïðîãðàììû." + vbCrLf + "Ïðîäîëæèòü ðàáîòó?", vbCritical + vbYesNo, "Îïàñíîñòü") = vbNo Then
        End
    Else
        GoTo 200
    End If
End If



Dim sCRC1 As String
sCRC1 = Trim$(UCase$(modFileManipulation.GetMD5(App.path & "\" & App.exeName & ".exe")))
'=============
samozaxist
Dim was As String
Dim encFinal As String
was = zp1 + Chr(83) + Chr(75) + Chr(89)
encFinal = UCase$(RC4(sCRC1, was))


If Trim$(encFinal) <> Trim$(UCase$(maimCRC)) Then
    If MsgBox("Íå ñîâïàäåíèå êîíòðîëüíûõ ñóì. Âîçìîæíî çàðàæåíèå." + vbCrLf + "Ïðîäîëæèòü?", vbYesNo + vbCritical, "Îïàñíîñòü") = vbNo Then
        End
    End If
End If
200:
End Sub

Sub samozaxist()
   On Error GoTo 10
  zp1 = Chr(83) + Chr(79) + Chr(75) + Chr(85) + Chr(76)
    Dim miNumBase As Integer
    Dim sMD5 As String
    Dim zb As String
    Dim bm As Integer
    miNumBase = FreeFile
Open App.path + "\mon.sec" For Input As #miNumBase
        Line Input #miNumBase, sMD5
        maimCRC = Trim$(sMD5)
        Close #miNumBase
10:
  Close #miNumBase
End Sub



Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
On Error Resume Next
Dim RB(0 To 255) As Integer, X As Long, Y As Long, Z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
If Len(Password) = 0 Then
    Exit Function
End If
If Len(Expression) = 0 Then
    Exit Function
End If
If Len(Password) > 256 Then
    Key() = StrConv(Left$(Password, 256), vbFromUnicode)
Else
    Key() = StrConv(Password, vbFromUnicode)
End If
For X = 0 To 255
    RB(X) = X
Next X
X = 0
Y = 0
Z = 0
For X = 0 To 255
    Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
    Temp = RB(X)
    RB(X) = RB(Y)
    RB(Y) = Temp
Next X
X = 0
Y = 0
Z = 0
ByteArray() = StrConv(Expression, vbFromUnicode)
For X = 0 To Len(Expression)
    Y = (Y + 1) Mod 256
    Z = (Z + RB(Y)) Mod 256
    Temp = RB(Y)
    RB(Y) = RB(Z)
    RB(Z) = Temp
    ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(Z)) Mod 256))
Next X
RC4 = StrConv(ByteArray, vbUnicode)
End Function

