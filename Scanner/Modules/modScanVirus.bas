Attribute VB_Name = "modScanVirus"
'поиск вирусов
'Option Explicit
Public nameVir_old As String
Public countZap As Long
Public zp1 As String
Public maimCRC As String
Dim a(100) As Long
Public countZapOLD As Long
Public allZapis As Long
Public baze_File As String
Public nameVir As String
Declare Function MoveFileEx& Lib "kernel32" Alias "MoveFileExA" (ByVal _
lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As _
Long)
'MOVEFILE_REPLACE_EXISTING = -1&
'MOVEFILE_COPY_ALLOWED = 0
'MOVEFILE_DELAY_UNTIL_REBOOT = 1&
Public Const pr = "Belyash AntiTrojan 2008 beta"
Public blnScan As Boolean   'check scan or stop
Public strScanDetail As String  'Scan Detail HTML text
Public QarNoDelete As Boolean
Public procCountSpl As Integer
Public procFoundSpl As Integer
Public procCleanSpl As Integer
Public chkFiles As Long
Public Logging As Boolean
Public mRead As Long
Public mhiden As Long
Public mNormal As Long
Public mSystem As Long
Public mArch As Long
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
Sub ScanFile(ByVal sPath As String)

    On Error Resume Next
    Dim mi_rash As String
    Dim fso As New FileSystemObject
    Dim X As ListItem
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
        DoEvents

    If sFile.Size >= val(frmMain.txtSizeChk.Text) Or sFile.Size = 0 Then
     '  If sFile.Size > 22478208 Or sFile.Size = 0 Then
       mNormal = mNormal + 1
       frmMain.lblNormal.Caption = "UnScaning:" & CStr(mNormal)
      'If sFile.Size > val(frmMain.Text1.Text) Or sFile.Size = 0 Then
            GoTo endFor   'can't use Continue For
    End If
        If blnScan = False Then
            Exit For
        End If
        mi_rash = Right$(sFile.Name, 4)
If frmMain.xpcheckbox71.Value = Unchecked Then
    If checkExten(Trim$(mi_rash)) = False Then
        GoTo endFor
    End If
End If
        r1 = Timer
       attr1 (sFile)
        'show scanned file
        frmMain.lblPath.Caption = sFile
   frmMain.SB.Caption = sFile
    frmMain.SB.Refresh
    If GetInputState() <> 0 Then DoEvents

             'Здесь возьмем старую MD5
           
        Dim Old_CRC As String
        Old_CRC = CRC.GetCRC(sFile)
If yesVir_old(Old_CRC) = True Then
    
        Set X = frmMain.lvVirusFound.ListItems.Add(, , Spliting(CStr(sFile), "\"), 2, 4)
        X.SubItems(1) = nameVir_old
        X.SubItems(2) = "Игнорирован"
        Set X = Nothing
        GoTo endFor
Else
    Dim sCRC As String
    sCRC = modFileManipulation.GetMD5(CStr(sFile))
    Dim i As Long
    If get_base(sCRC) = True Then
                LogPrint "" + sFile + "-инфицирован (" + nameVir + ")"
                frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + 1
                Dim tempFN As String
                tempFN = sFile.Path ' & "\" & sFile.Name
If QarNoDelete = True Then
                 'вах переименовать
101:
                Dim d As Variant
                 d = Format(Now, "ddMMSSHHMMSS")
                If Dir$(sFile.Name + "." + d) <> "" Then
                    GoTo 101 'если есть такой файл в карантине то генерируем новое имя
                End If
  
            moveToQarant sFile.Name, sFile.Name + "." + d
  
        Else
                sFile.Delete True
                LogPrint sFile.Name + "-инфицированный (" + Trim$(nameVir) + ") -удален"
     End If
                 'check whether the virus is cleaned or not; if not, go to errKill to show Error Cleaning
                If FileorFolderExists(tempFN) = True Then GoTo errKill
                'Kill sFile.Path
                'add to log after kill process
                'frmMain.txtLog.Text = frmMain.txtLog.Text & "  Removed " & tempFN & vbCrLf
                frmMain.lblCleaned.Caption = Int(frmMain.lblCleaned.Caption) + 1
               
                ' strScanDetail = strScanDetail & " Virus Cleaned<br>"
                ' Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)
                Set X = frmMain.lvVirusFound.ListItems.Add(, , Trim$(nameVir), 2, 4)
                X.SubItems(1) = tempFN
                If QarNoDelete = True Then
                   X.SubItems(2) = "Перемещен"
                      ion sFile.Name + "=>" + tempFN
                   'LogPrint nameVir + "-moved"
                 Else
                   X.SubItems(2) = "Удалён"
                   'LogPrint nameVir + "-cleaned"
                End If
                
                GoTo endFor
errKill:
                'add to log after kill error
                'frmMain.txtLog.Text = frmMain.txtLog.Text & "  Cannot removed " & tempFN & vbCrLf
                ' strScanDetail = strScanDetail & "<font Size=3 Color=YELLOW><i> Virus Cannot Be Cleaned</i></font><br>"
                ' Call UpdateDetail(strScanDetail, frmMain.WebBrowser1)
                Set X = frmMain.lvVirusFound.ListItems.Add(, , nameVir, 3, 4)
                X.SubItems(1) = tempFN
                 If QarNoDelete = True Then
                    X.SubItems(2) = "Ошибка перемещения"
                    LogPrint nameVir + "-Ошибка перемещения"
                Else
                    X.SubItems(2) = "Ошибка удаления"
                    LogPrint nameVir + "-Ошибка удаления"
                End If
                Set X = Nothing
                Exit For
    End If
End If
    If mast_md5 = True Then
    'писать в лог md5?
        If Right(sPath, 1) = "\" Then
        LogPrint sPath + sFile.Name + "---" + CStr(sCRC)
        Else
         LogPrint sPath + "\" + sFile.Name + "---" + CStr(sCRC)
        End If
    Else
     If Right(sPath, 1) = "\" Then
            LogPrint sPath + sFile.Name + "-OK"
        Else
            LogPrint sPath + "\" + sFile.Name + "-OK"
        End If
    End If
        

    r2 = Timer
    frmMain.lblNormal09.Caption = "Time:" & CStr(r2 - r1)
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
Function checkExten(filext1 As String) As Boolean
checkExten = False

Select Case LCase$(filext1)
Case ".doc"
    checkExten = True
    Exit Function
Case ".xls"
    checkExten = True
    Exit Function

Case ".dot"
    checkExten = True
    Exit Function

Case ".htm"
    checkExten = True
    Exit Function

Case "html"
    checkExten = True
    Exit Function

Case ".vbs"
    checkExten = True
    Exit Function

Case ".wsh"
    checkExten = True
    Exit Function

Case ".com"
    checkExten = True
    Exit Function
Case ".exe"
    checkExten = True
    Exit Function

Case ".bat"
    checkExten = True
    Exit Function

Case ".mht"
    checkExten = True
    Exit Function
Case ".asp"
    checkExten = True
    Exit Function
Case ".bin"
    checkExten = True
    Exit Function
Case ".chm"
    checkExten = True
    Exit Function
Case ".dll"
    checkExten = True
    Exit Function

Case ".eml"
    checkExten = True
    Exit Function
Case ".hta"
    checkExten = True
    Exit Function
Case ".ocx"
    checkExten = True
    Exit Function
Case ".php"
    checkExten = True
    Exit Function
Case "lass"
    checkExten = True
    Exit Function
End Select

If Right$(filext1, 3) = ".js" Then
    checkExten = True
    Exit Function
End If


End Function


Function yesVir_old(sim As String) As Boolean
   On Error GoTo 10
  
  yesVir_old = False
If Dir$(App.Path + "\old.bmd") = "" Then
    Exit Function
End If
    Dim miNumBase As Integer
    Dim sMD5 As String
    Dim zb As String
    Dim fg As String
    Dim bm As Integer
    miNumBase = FreeFile
Open App.Path + "\old.bmd" For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD5
        bm = InStr(1, sMD5, sim, vbTextCompare)
    If bm <> 0 Then
        zb = Right$(sMD5, Len(sMD5) - 9)
        yesVir_old = True
        nameVir_old = "возможно " + CStr(zb)
       'MsgBox "" + CStr(zb)
        Close #miNumBase
        Exit Function
     End If
     Wend
10:

   yesVir_old = False
  Close #miNumBase

End Function
Function get_base(CRC_checkSUM As String) As Boolean
'перебор всех вирусных баз
On Error GoTo 100
'add 1 to counter scanned
chkFiles = chkFiles + 1
 frmMain.lblCount.Caption = chkFiles
 frmMain.lblNormal112.Caption = "Total:" & CStr(chkFiles)
get_base = False
Dim S As String
S = Dir(App.Path + "\*.bVb")
' Код обработки баз
Do While S <> ""
   'Debug.Print S
         If yesVir(CRC_checkSUM, S) = True Then
           get_base = True
           Exit Function
         End If
         
    S = Dir
Loop

Exit Function
100:
MsgBox "" + Error$
End Function
Function yesVir(sim As String, my_base As String) As Boolean
 'On Error GoTo 10

yesVir = False
    Dim miNumBase As Integer
    Dim sMD5 As String
    Dim zb As String
    Dim bm As Integer
    Dim hj As String
    miNumBase = FreeFile
Open my_base For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD5
        
            bm = InStr(1, sMD5, sim, vbTextCompare)
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

Function yesVirLogoTWO(CFc8 As String, b1_file As String) As Boolean

On Error GoTo 10
    yesVirLogoTWO = False
    Dim miNumBase As Integer
     Dim sMD5 As String
    Dim zb As String
    Dim hj As String
    miNumBase = FreeFile
     Open b1_file For Input As #miNumBase
        While Not EOF(miNumBase)
             Line Input #miNumBase, sMD5
               Dim bm As Integer
            bm = InStr(1, sMD5, CFc8, vbTextCompare)
    If bm <> 0 Then
        zb = Right$(sMD5, Len(sMD5) - 33)
        yesVirLogoTWO = True
        nameVir = CStr(zb)
       'MsgBox "" + CStr(zb)
        Close #miNumBase
        Exit Function
     End If
     Wend
10:

   yesVirLogoTWO = False
  Close #miNumBase

End Function
Function get_base_yesVirLogoTWO(CRC_checkSUM As String) As Boolean
'перебор всех вирусных баз
On Error GoTo 100
get_base_yesVirLogoTWO = False
Dim S As String
S = Dir(App.Path + "\*.bVb")
' Код обработки баз
Do While S <> ""
  ' Debug.Print S
         If yesVirLogoTWO(CRC_checkSUM, S) = True Then
           get_base_yesVirLogoTWO = True
           Exit Function
         End If
         
    S = Dir
Loop

Exit Function
100:
MsgBox "" + Error$
End Function
Function yesVirLogo(CFc As String, b2_file As String) As Boolean

On Error GoTo 10
    yesVirLogo = False
    Dim miNumBase As Integer
Dim hj As String
    Dim sMD5 As String
      Dim zb As String
    miNumBase = FreeFile
     Open b2_file For Input As #miNumBase
        While Not EOF(miNumBase)
             Line Input #miNumBase, sMD5
             Dim bm As Integer
            bm = InStr(1, sMD5, CFc, vbTextCompare)
    If bm <> 0 Then
        zb = Right$(sMD5, Len(sMD5) - 33)
        yesVirLogo = True
        nameVir = CStr(zb)
       'MsgBox "" + CStr(zb)
        Close #miNumBase
        Exit Function
     End If
     Wend
10:

   yesVirLogo = False
  Close #miNumBase

End Function
Function get_base_yesVirLogo(CRC_checkSUM As String) As Boolean
'перебор всех вирусных баз
On Error GoTo 100
get_base_yesVirLogo = False
Dim S As String
S = Dir(App.Path + "\*.bVb")

' Код обработки баз
Do While S <> ""
   'Debug.Print S
         If yesVirLogo(CRC_checkSUM, S) = True Then
           get_base_yesVirLogo = True
           Exit Function
         End If
         
    S = Dir
Loop

Exit Function
100:
MsgBox "" + Error$
End Function
Public Sub LogPrint(sMessage As String)
'проца записи в файл исходного текста макровируса
On Error GoTo 100
Logging = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "Log"))
If Logging = False Then
    Exit Sub 'вести лог?
End If
If Dir$(App.Path + "\otchetScan.txt", vbNormal) <> "" Then
'обрезаем
Dim a95 As Long
a95 = CLng(GetSetting(App.exeName, "Options", "logSize", "100024"))

    If FileLen(App.Path + "\otchetScan.txt") >= a95 Then
       ' MsgBox "обрезал"
        Kill App.Path + "\otchetScan.txt"
    End If
End If

Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
'Open Ffile For Append As #nFile
ffile = App.Path + "\otchetScan.txt"
Open ffile For Append Access Write Lock Read Write As #nFile
Print #nFile, Format$(Now, "mm-dd-yy hh:mm:ss") + "--" + sMessage
Close #nFile

Exit Sub
100:
Select Case Err
Case 52
    MsgBox "Не правильное имя файла", vbCritical, pr
Case 53
    MsgBox "Файл не найден", vbCritical, pr
Case 54
    MsgBox "Не правильный режим файла", vbCritical, pr
Case 57
    MsgBox "Ошибка ввода/вывода", vbCritical, pr
Case 61
    MsgBox "Преполнение диска. Пора бы его почистить", vbCritical, pr
    'End
Case 62
    MsgBox "Почему-то произошела запись после закрытия файла", vbCritical, pr
Case 67
    MsgBox "Слишком много открытых файлов. Не могу я так работать", vbCritical, pr
    End
Case 68
    MsgBox "Не вижу я диск ", vbCritical, pr
    'End
Case 70
    MsgBox "Доступ к диску или каталогу запрещен.Возможно указан очень маленький размер отчета.", vbCritical, pr
    'End
Case 71
    MsgBox "Диск не готов", vbCritical, pr
Case 75
    
    If MsgBox("Ну не могу я работать с этим файлом или каталогом" & vbCrLf & ffile & vbCrLf & "Возможно это произошло из-за вмешательства антивируса. Продолжить ?", vbYesNo + vbCritical, pr) = vbNo Then
        'End
    End If
Case 76
    MsgBox "Путь определен не верно...Возможно это произошло из-за вмешательства другой программы..." & ffile, vbCritical
    
Case 321
    MsgBox "Не правильный формат файла", vbCritical
'End
Case 3024
    MsgBox "Черт, не могу найти этот файл" & ffile, vbCritical, pr
'    End
Case 3176
    MsgBox "Не могу открыть файл" & ffile, vbCritical, pr
Case 3179
    MsgBox "Ошибочный конец файла", vbCritical, pr
'End
Case 3180
    MsgBox "Не могу записать в файл", vbCritical, pr
'End
Case Else
    MsgBox "Произошла внутренняя ошибка при открытии фала", vbCritical, pr
End Select
End Sub
Public Sub ion(sMessage6 As String)
'проца записи перемещаемого в карантин
On Error GoTo 100
If Dir$(App.Path + "\quarantine", vbDirectory) = "" Then
  MkDir App.Path + "\quarantine"
End If

Dim nFile As Integer
Dim ffile As String
nFile = FreeFile
'Open Ffile For Append As #nFile
ffile = App.Path + "\removed.log"
Open ffile For Append As #nFile
Print #nFile, Format$(Now, "mm-dd-yy hh:mm:ss") + "--" + sMessage6
Close #nFile

Exit Sub
100:
Select Case Err
Case 52
    MsgBox "Не правильное имя файла", vbCritical, pr
Case 53
    MsgBox "Файл не найден", vbCritical, pr
Case 54
    MsgBox "Не правильный режим файла", vbCritical, pr
Case 57
    MsgBox "Ошибка ввода/вывода", vbCritical, pr
Case 61
    MsgBox "Преполнение диска. Пора бы его почистить", vbCritical, pr
    'End
Case 62
    MsgBox "Почему-то произошела запись после закрытия файла", vbCritical, pr
Case 67
    MsgBox "Слишком много открытых файлов. Не могу я так работать", vbCritical, pr
    End
Case 68
    MsgBox "Не вижу я диск ", vbCritical, pr
    'End
Case 70
    MsgBox "Доступ к диску или каталогу запрещен", vbCritical, pr
    'End
Case 71
    MsgBox "Диск не готов", vbCritical, pr
Case 75
    
    If MsgBox("Ну не могу я работать с этим файлом или каталогом" & vbCrLf & ffile & vbCrLf & "Возможно это произошло из-за вмешательства антивируса. Продолжить ?" + vbCrLf + Error$, vbYesNo + vbCritical, pr) = vbNo Then
        End
    End If
Case 76
    MsgBox "Путь определен не верно...Возможно это произошло из-за вмешательства другой программы..." & ffile, vbCritical
    
Case 321
    MsgBox "Не правильный формат файла", vbCritical
'End
Case 3024
    MsgBox "Черт, не могу найти этот файл" & ffile, vbCritical, pr
'    End
Case 3176
    MsgBox "Не могу открыть файл" & ffile, vbCritical, pr
Case 3179
    MsgBox "Ошибочный конец файла", vbCritical, pr
'End
Case 3180
    MsgBox "Не могу записать в файл", vbCritical, pr
'End
Case Else
    MsgBox "Произошла внутренняя ошибка при открытии фала", vbCritical, pr
End Select
End Sub


Public Function ScanFilepR(ByVal sPath As String, ByVal mi_file As String, proccAidi As Long) As Boolean
    'чисто для стартовой формы,чтобы не мешать сюда главную форму
    
    
    On Error Resume Next
    
    ScanFilepR = False
    'frmMain.lblCount.Caption = Int(frmSplash.lblCount.Caption) + 1
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
        If Left(mi_file, 1) = "/" Then
            Exit Function
        End If
        'MsgBox mi_file
        Dim i As Long
           'Debug.Print sCRC
        'compare with database
       ' For i = 0 To UBound(base)
        
       
       DoEvents
      
        If get_base_yesVirLogo(sCRC) = True Then
                'If sCRC = base(i).Value Then   'start cleaning
             
'                LogPrint mi_file + "-found virus in memory"
                'add to log
                    frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + 1
                     procFoundSpl = procFoundSpl + 1



                   If MsgBox("В памяти обнаружен вирус." + vbCrLf + nameVir + vbCrLf + "Уничтожить ?", vbCritical + vbYesNo + vbSystemModal, pr) = vbYes Then
                       'frmMain.lblZabl.Caption = Int(frmMain.lblZabl.Caption) + 1
                       
                            Call Process_Kill(proccAidi)
                            'frmSplash.Text1.Text = "В памяти обнаружен вирус." + vbCrLf + nameVir
                            'Call frmMain.message
                            'Process_Kill (proccAidi)
                           Dim fn1 As Long
                    fn1 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonQv"))
                    If fn1 = 1 Then
                           'вах переименовать
                         Call moveToQarant(mi_file, nameVir)
                      '  MsgBox "карантин"
                    Else
                        Call killDelVir(mi_file, nameVir)
                      '  MsgBox "удалить"
                    End If

                         
                                    mi_file = ""
                                    proccAidi = 0
                                    ScanFilepR = True
                                    Exit Function
'programs.Process_Kill (mi_file)
                     ' processes.KillProcess proccAidi
                       '######бээээ
'                        LogPrint mi_file + "-terminate virus process"
                        ' procCleanSpl = procCleanSpl + 1
                   End If
                    
            End If
                'resume
  
                
              '      Next i

    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    'frmMain.Text1.Text = ""
End Function
Public Function ScanFileCMD(ByVal sPath As String, ByVal mi_file As String, proccAidi As Long) As Boolean
    'чисто для стартовой формы,чтобы не мешать сюда главную форму
    
    
    On Error Resume Next
    
    ScanFileCMD = False
    frmMain.lblCount.Caption = Int(frmMain.lblCount.Caption) + 1
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
        If Left(mi_file, 1) = "/" Then
            Exit Function
        End If
        'MsgBox mi_file
        Dim i As Long
          ' Debug.Print sCRC
        'compare with database
       ' For i = 0 To UBound(base)
        
       
       DoEvents
      
 If get_base_yesVirLogo(sCRC) = True Then
                'If sCRC = base(i).Value Then   'start cleaning
             
'                LogPrint mi_file + "-found virus in memory"
                'add to log
                    frmMain.lblFound.Caption = Int(frmMain.lblFound.Caption) + 1
                     procFoundSpl = procFoundSpl + 1
            Dim X As ListItem
            Set X = frmMain.lvVirusFound.ListItems.Add(, , nameVir, 1, 4)
            X.SubItems(1) = mi_file


                   'If MsgBox("В памяти обнаружен вирус." + vbCrLf + nameVir + vbCrLf + "Уничтожить ?", vbCritical + vbYesNo + vbSystemModal, pr) = vbYes Then
                       'frmMain.lblZabl.Caption = Int(frmMain.lblZabl.Caption) + 1
                       
                            Call Process_Kill(proccAidi)
                            'frmSplash.Text1.Text = "В памяти обнаружен вирус." + vbCrLf + nameVir
                            'Call frmMain.message
                            'Process_Kill (proccAidi)
                           Dim fn1 As Long
                    fn1 = val(getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "MonQv"))
                    If fn1 = 1 Then
                           'вах переименовать
                            X.SubItems(2) = "Перемещен"
                         Call moveToQarant(mi_file, nameVir)
                      '  MsgBox "карантин"
                    Else
                    X.SubItems(2) = "Удалён"
                        Call killDelVir(mi_file, nameVir)
                      '  MsgBox "удалить"
                    'End If

                         
                                    mi_file = ""
                                    proccAidi = 0
                                    ScanFileCMD = True
                                    GoTo 12
'programs.Process_Kill (mi_file)
                     ' processes.KillProcess proccAidi
                       '######бээээ
'                        LogPrint mi_file + "-terminate virus process"
                        ' procCleanSpl = procCleanSpl + 1
                   End If
 Else
     
    'Здесь возьмем старую MD5
        Dim Old_CRC As String
        Old_CRC = CRC.GetCRC(mi_file)
    If yesVir_old(Old_CRC) Then
            'MsgBox "" + mi_file
'    Stop
             Call Process_Kill(proccAidi)
        Set X = frmMain.lvVirusFound.ListItems.Add(, , Spliting(CStr(mi_file), "\"), 2, 4)
        X.SubItems(1) = nameVir_old
        X.SubItems(2) = "Заблокирован"
        Set X = Nothing
        GoTo 12
    End If
        
     
 End If
                'resume
  
                
              '      Next i
12:
    'clear variables
    Set fso = Nothing
    Set mFolder = Nothing
    Set sFolders = Nothing
    Set sFiles = Nothing
    Set sFolder = Nothing
    Set sFile = Nothing
    'frmMain.Text1.Text = ""
End Function
Sub moveToQarant(s12 As String, virn1 As String)
On Error Resume Next


If Dir$(App.Path + "\quarantine", vbDirectory) = "" Then
  MkDir App.Path + "\quarantine"
End If
     Dim fso As New FileSystemObject, txtfile, fil1, fil2
' Код обработки файла в корне C:\
Set fil1 = fso.GetFile(s12)
101:
Dim d As Variant
    d = Format(Now, "ddMMSSHHMMSS")
                If Dir$(App.Path + "\quarantine\" + d, vbNormal) <> "" Then
                    GoTo 101 'если есть такой файл в карантине то генерируем новое имя
                End If
' Перемещаем файл в директорию \tmp
fil1.Move (App.Path + "\quarantine\" + d)
        
        
    If Dir$(s12, vbNormal) = "" Then
    
       ' frmMain.lblCleaned.Caption = Int(frmMain.lblCleaned.Caption) + 1
        LogQ s12 + "==>" + App.Path + "\quarantine\" + d
'        frmMain.Text1.Text = "Вирус <" + virn1 + "> перемещен в карантин"
'        Call frmMain.message
    Else
    LogPrint s12 + "-ошибка при перемещении в карантин"
'            frmMain.Text1.Text = "Ошибка перемещения " + vbCrLf + virn1 + vbCrLf + " в карантин"
'        Call frmMain.message
    End If
    procCleanSpl = procCleanSpl + 1
Set fso = Nothing
    Set fil1 = Nothing
    
End Sub
Sub killDelVir(s13 As String, virn2 As String)
On Error Resume Next

  Dim fso As New FileSystemObject, txtfile, fil1, fil2
' Код обработки файла в корне C:\
Set fil1 = fso.GetFile(s13)
    fil1.Delete True
If Dir$(s13, vbNormal) = "" Then
        'frmMain.lblDelete.Caption = Int(frmMain.lblDelete.Caption) + 1
        LogPrint s13 + "-удален"
        '        frmMain.Text1.Text = "Вирус <" + virn2 + ">  удален"
        'Call frmMain.message
    Else
            LogPrint s13 + "-ошибка при удалении"
            'frmMain.Text1.Text = "Ошибка при удалении " + vbCrLf + virn2
        'Call frmMain.message
    End If
    procCleanSpl = procCleanSpl + 1
Set fso = Nothing
    Set fil1 = Nothing
End Sub
Sub LogQ(sMessage6 As String)
'проца записи перемещаемого в карантин
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
       ' MsgBox "обрезал"
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
    MsgBox "Не правильное имя файла", vbCritical, pr
Case 53
    MsgBox "Файл не найден", vbCritical, pr
Case 54
    MsgBox "Не правильный режим файла", vbCritical, pr
Case 57
    MsgBox "Ошибка ввода/вывода", vbCritical, pr
Case 61
    MsgBox "Преполнение диска. Пора бы его почистить", vbCritical, pr
    'End
Case 62
    MsgBox "Почему-то произошела запись после закрытия файла", vbCritical, pr
Case 67
    MsgBox "Слишком много открытых файлов. Не могу я так работать", vbCritical, pr
    End
Case 68
    MsgBox "Не вижу я диск ", vbCritical, pr
    'End
Case 70
    MsgBox "Доступ к диску или каталогу запрещен", vbCritical, pr
    'End
Case 71
    MsgBox "Диск не готов", vbCritical, pr
Case 75
    
    If MsgBox("Ну не могу я работать с этим файлом или каталогом" & vbCrLf & ffile & vbCrLf & "Возможно это произошло из-за вмешательства антивируса. Продолжить ?" + vbCrLf + Error$, vbYesNo + vbCritical, pr) = vbNo Then
        End
    End If
Case 76
    MsgBox "Путь определен не верно...Возможно это произошло из-за вмешательства другой программы..." & ffile, vbCritical
    
Case 321
    MsgBox "Не правильный формат файла", vbCritical
'End
Case 3024
    MsgBox "Черт, не могу найти этот файл" & ffile, vbCritical, pr
'    End
Case 3176
    MsgBox "Не могу открыть файл" & ffile, vbCritical, pr
Case 3179
    MsgBox "Ошибочный конец файла", vbCritical, pr
'End
Case 3180
    MsgBox "Не могу записать в файл", vbCritical, pr
'End
Case Else
    MsgBox "Произошла внутренняя ошибка при открытии фала", vbCritical, pr
End Select

End Sub
Private Function Spliting(sFullPath As String, point As String)
On Error GoTo 10
Dim str1() As String
str1 = Split(sFullPath, point)
Spliting = str1(UBound(str1))
Exit Function
10:
Spliting = "Неизвестное"
End Function

Sub attr1(pnpFile As String)
On Error GoTo 100
 Dim fattr As Integer
        fattr = GetAttr(pnpFile)
        
If (fattr And vbReadOnly) Then
mRead = mRead + 1
frmMain.lblRead.Caption = "ReadOnly:" & CStr(mRead)
End If
 If (fattr And vbHidden) Then
mhiden = mhiden + 1
frmMain.LblHiden.Caption = "Hidden:" & CStr(mhiden)
End If
If (fattr And vbSystem) Then
mSystem = mSystem + 1
frmMain.lblSyst.Caption = "System:" & CStr(mSystem)
End If
If (fattr And vbArchive) Then
    mArch = mArch + 1
    frmMain.lblArch.Caption = "Archive:" & CStr(mArch)
End If
frmMain.lblSize.Caption = "Size:" & CStr(FileLen(pnpFile)) & "byte"

Exit Sub
LogPrint "" + Error$
100:
End Sub

Sub ochistStatistic()
chkFiles = 0
mRead = 0
mhiden = 0
mSystem = 0
mArch = 0
frmMain.lblNormal09.Caption = "Time:0"
frmMain.lblSize.Caption = "Size:0"
End Sub

Sub kolwoVirusow()
On Error GoTo 100
If Dir$(App.Path + "\old.bmd", vbNormal) = "" Then
    Exit Sub
End If
Dim miNumBase As Integer
miNumBase = FreeFile
Dim sMD15 As String
Open App.Path + "\old.bmd" For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD15
        countZap = countZap + 1
     Wend
 Close #miNumBase
 Exit Sub
100:
 LogPrint "" + Error$
End Sub

Sub kolwoVirusowAll(baZE As String)
On Error GoTo 100
Static i As Integer
i = i + 1
Dim miNumBase As Integer
miNumBase = FreeFile
Dim sMD15 As String
Open App.Path + "\" + baZE For Input As #miNumBase
  While Not EOF(miNumBase)
        Line Input #miNumBase, sMD15
        a(i) = a(i) + 1
     Wend
 Close #miNumBase

 Exit Sub
100:
 LogPrint "" + Error$
End Sub
Sub sbor()
Dim sNext1 As String
sNext1 = Dir$(App.Path + "\*.bVb")
While sNext1 <> ""

    kolwoVirusowAll (sNext1)
    sNext1 = Dir$
Wend
Dim baZE As String
kolwoVirusow
Dim nm As Long
Dim z As Byte

For z = 0 To 100
nm = nm + a(z)
Next z
countZapOLD = nm + countZap
End Sub
Sub selfTesting()
On Error GoTo 200
If Dir$(App.Path + "\scan.sec", vbNormal) = "" Then
    If MsgBox("В каталоге с программой отсутствует файл самопроверки.Возможно заражение программы." + vbCrLf + "Продолжить работу?", vbCritical + vbYesNo, "Опасность") = vbNo Then
        End
    Else
        GoTo 200
    End If
End If



Dim sCRC1 As String
sCRC1 = Trim$(UCase$(modFileManipulation.GetMD5(App.Path & "\" & App.exeName & ".exe")))
'=============
samozaxist
Dim was As String
Dim encFinal As String
was = zp1 + Chr(83) + Chr(75) + Chr(89)
encFinal = UCase$(RC4(sCRC1, was))


If Trim$(encFinal) <> Trim$(UCase$(maimCRC)) Then
    If MsgBox("Не совпадение контрольных сум. Возможно заражение." + vbCrLf + "Продолжить?", vbYesNo + vbCritical, "Опасность") = vbNo Then
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
Open App.Path + "\scan.sec" For Input As #miNumBase
        Line Input #miNumBase, sMD5
        maimCRC = Trim$(sMD5)
        Close #miNumBase
10:
  Close #miNumBase
End Sub



Public Function RC4(ByVal Expression As String, ByVal Password As String) As String
On Error Resume Next
Dim RB(0 To 255) As Integer, X As Long, Y As Long, z As Long, Key() As Byte, ByteArray() As Byte, Temp As Byte
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
z = 0
For X = 0 To 255
    Y = (Y + RB(X) + Key(X Mod Len(Password))) Mod 256
    Temp = RB(X)
    RB(X) = RB(Y)
    RB(Y) = Temp
Next X
X = 0
Y = 0
z = 0
ByteArray() = StrConv(Expression, vbFromUnicode)
For X = 0 To Len(Expression)
    Y = (Y + 1) Mod 256
    z = (z + RB(Y)) Mod 256
    Temp = RB(Y)
    RB(Y) = RB(z)
    RB(z) = Temp
    ByteArray(X) = ByteArray(X) Xor (RB((RB(Y) + RB(z)) Mod 256))
Next X
RC4 = StrConv(ByteArray, vbUnicode)
End Function


