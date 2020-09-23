Attribute VB_Name = "mScaning"
Option Explicit
'****************************************************************
'Äàííûé ìîäóëü ïðåäíàçíà÷åí äëÿ ïðîâåäåíèÿ ïîèñêà ôàéëîâ ïî ìàñêå
'â óêàçàííîì êàòàëîãå è åãî ïîäêàòàëîãàõ
'****************************************************************
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Enum AtributConst
    FILE_ATTRIBUTE_ARCHIVE = &H20
    FILE_ATTRIBUTE_COMPRESSED = &H800
    FILE_ATTRIBUTE_DIRECTORY = &H10
    'FILE_ATTRIBUTE_ENCRYPTED=???
    FILE_ATTRIBUTE_HIDDEN = &H2
    FILE_ATTRIBUTE_NORMAL = &H80
    'FILE_ATTRIBUTE_OFFLINE=???
    FILE_ATTRIBUTE_READONLY = &H1
    'FILE_ATTRIBUTE_REPARSE_POINT=???
    'FILE_ATTRIBUTE_SPARSE_FILE=???
    FILE_ATTRIBUTE_SYSTEM = &H4
    FILE_ATTRIBUTE_TEMPORARY = &H100
End Enum
Public Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes    As AtributConst
    ftCreationTime      As FILETIME
    ftLastAccessTime    As FILETIME
    ftLastWriteTime     As FILETIME
    nFileSizeHigh       As Long
    nFileSizeLow        As Long
    dwReserved0         As Long
    dwReserved1         As Long
    cFileName           As String * 260
    cAlternate          As String * 14
End Type
Private Const INVALID_HANDLE_VALUE = -1                     'Âîçíèêàåò ïðè îòñóòñòâèè ôàéëà

Public Function StartFind(ByVal sFilter As String, ByRef FindArray() As WIN32_FIND_DATA)
    Dim lngHandle As Long                                   'Õàíäë îòêðûòîãî ôàéëà
    Dim TempData As WIN32_FIND_DATA                         'Âðåìåííàÿ ïåðåìåííàÿ
    
    lngHandle = FindFirstFile(sFilter & Chr(0), TempData)   'Íàõîäèì õàíäë
    If lngHandle = INVALID_HANDLE_VALUE Then Exit Function  'Åñëè íè÷åãî íåò, óõîäèì
    ReDim FindArray(0)                                      'Ðåñàéçèì ìàññèâ ïîä 1-é ôàéë (äèðåêòîðèþ)
    FindArray(0) = TempData                                 'Ïðèñâàåâàåì èíôîðìàöèþ î ôàéëå
    
    Do Until FindNextFile(lngHandle, TempData) = 0          'Ïðîäîëæàåì ïîèñê, ïîêà íå îøèáåìñÿ
        DoEvents                                            'Íå çàâèñàåì
        ReDim Preserve FindArray(UBound(FindArray) + 1)     'Ðåñàéçèì ìàññèâ íà 1 áîëüøå ïðåæíåãî
        FindArray(UBound(FindArray)) = TempData             'Ïðèñâîåíèå î÷åðåäíîé ïîðöèè èíôîðìàöèè
    Loop

    Call FindClose(lngHandle)                               'Ïîáàëîâàëèñü äà è õâàòèò...
End Function

'Ïîèñê ôàéëîâ ïî ìàñêå â ïîäêàòàëîãàõ
Public Sub ScanForFiles(StartPath As String, Pattern As String, Files() As String)
    Dim i As Long, Dir1() As String, File1() As String, File As Long
    File = UboundS(Files())
    File = File + 1
    If Mid(StartPath, Len(StartPath), 1) = "\" Then
    Else
        StartPath = StartPath + "\"
    End If
    DirBoxEmu StartPath, Dir1()
    FileBoxEmu StartPath, Pattern, File1()
    For i = 0 To UboundS(File1())
        DoEvents
        ReDim Preserve Files(i + File)
        Files(i + File) = StartPath + File1(i)
    Next
    'For i = 0 To UboundS(Dir1())
     '   DoEvents
      '  ScanForFiles StartPath + Dir1(i), Pattern, Files()
    'Next
End Sub

'Ýìóëÿöèÿ DirBox
Private Sub DirBoxEmu(Path As String, Dirs() As String)
    Dim Data() As WIN32_FIND_DATA, i As Long, Counter As Long, TempName As String
    Counter = -1
    Erase Dirs()
    If Mid(Path, Len(Path), 1) = "\" Then
    Else
        Path = Path + "\"
    End If
    Erase Data()
    StartFind Path + "*", Data()
    For i = 0 To UboundFind(Data())
        DoEvents
        TempName = Left(Data(i).cFileName, InStr(1, Data(i).cFileName, Chr(0)) - 1)
        If TempName <> "." And TempName <> ".." Then
            If GetEntryType(Data(i).dwFileAttributes) = 1 Then
                Counter = Counter + 1
                ReDim Preserve Dirs(Counter)
                Dirs(Counter) = TempName
            End If
        End If
    Next
End Sub

'Ýìóëÿöèÿ FileBox
Private Sub FileBoxEmu(Path As String, Filter As String, Files() As String)
    Dim Data() As WIN32_FIND_DATA, Repeat As Long, i As Long, j As Long, Counter As Long, Filters() As String, TempName As String
    Counter = -1
    Erase Files()
    Erase Filters()
    If Mid(Path, Len(Path), 1) = "\" Then
    Else
        Path = Path + "\"
    End If
    Filters() = Split(Filter, ";")
    Counter = UboundS(Filters())
    If Counter = -1 Then
        Repeat = 0
        ReDim Filters(0)
        Filters(0) = "*"
    Else
        Repeat = Counter
    End If
    Counter = -1
    For j = 0 To Repeat
        DoEvents
        Erase Data()
        StartFind Path + Filters(j), Data()
        For i = 0 To UboundFind(Data())
            DoEvents
            TempName = Left(Data(i).cFileName, InStr(1, Data(i).cFileName, Chr(0)) - 1)
            If TempName <> "." And TempName <> ".." Then
                If GetEntryType(Data(i).dwFileAttributes) = 0 Then
                    Counter = Counter + 1
                    ReDim Preserve Files(Counter)
                    Files(Counter) = TempName
                End If
            End If
        Next
    Next
End Sub

Private Function GetEntryType(dwAttributes As Long) As Long
    GetEntryType = 0 'ôàéë ïî óìîë÷àíèÿ
    If (dwAttributes And &H10) = &H10 Then
        GetEntryType = 1 ' êàòàëîã
    Else
        GetEntryType = 0 'ôàéë
    End If
End Function

'Ñòðàííîñòè â ðàáîòè ôóíêöèè UBound() âûíóäèëè íàïèñàòü ñâîè
Public Function UboundFind(Data() As WIN32_FIND_DATA) As Long
    On Error GoTo Handler:
    UboundFind = UBound(Data())
    Exit Function
Handler:
    UboundFind = -1
End Function

Public Function UboundS(Data() As String) As Long
    On Error GoTo Handler:
    UboundS = UBound(Data())
    Exit Function
Handler:
    UboundS = -1
End Function

