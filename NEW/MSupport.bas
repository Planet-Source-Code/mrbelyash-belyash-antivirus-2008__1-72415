Attribute VB_Name = "MSupport"
  Option Explicit
  ' demo project showing how to use the FindFirstChangeNotification and
  ' FindNextChangeNotification API functions to watch directories for
  ' changes.  by Bryan Stafford of New Vision Software - newvision@mvps.org
  ' this demo is released into the public domain "as is" without warranty or
  ' guaranty of any kind.  In other words, use at your own risk.
  '
  '
  ' NOTE:  there are other routines that can detect changes in a directory.
  ' structure, most notably registering a change notification routine using
  ' the SHChangeNotifyRegister API function.  However, unless the application
  ' making changes to the file system is sending the proper notifications via
  ' the SHChangeNotify API function, only a minimal amount of information is
  ' returned.


  Private Const vbZLString As String = ""
  Private Const MAX_PATH As Long = 260&
  
  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
  End Type

  Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
  End Type

  ' DrawText() Format Flags
  Private Const DT_TOP As Long = &H0&
  Private Const DT_LEFT As Long = &H0&
  Private Const DT_CENTER As Long = &H1&
  Private Const DT_RIGHT As Long = &H2&
  Private Const DT_VCENTER As Long = &H4&
  Private Const DT_BOTTOM As Long = &H8&
  Private Const DT_WORDBREAK As Long = &H10&
  Private Const DT_SINGLELINE As Long = &H20&
  Private Const DT_EXPANDTABS As Long = &H40&
  Private Const DT_TABSTOP As Long = &H80&
  Private Const DT_NOCLIP As Long = &H100&
  Private Const DT_EXTERNALLEADING As Long = &H200&
  Private Const DT_CALCRECT As Long = &H400&
  Private Const DT_NOPREFIX As Long = &H800&
  Private Const DT_INTERNAL As Long = &H1000&
  Private Const DT_EDITCONTROL As Long = &H2000&
  Private Const DT_END_ELLIPSIS As Long = &H8000&
  Private Const DT_MODIFYSTRING As Long = &H10000
  Private Const DT_PATH_ELLIPSIS As Long = &H4000&
  Private Const DT_RTLREADING As Long = &H20000
  Private Const DT_WORD_ELLIPSIS As Long = &H40000
  
  Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc&, ByVal lpStr$, ByVal nCount&, lpRect As RECT, ByVal wFormat&) As Long

  Private Const INVALID_HANDLE_VALUE As Long = (-1&)
  Public Const ERROR_NO_MORE_FILES As Long = 18&

  Public Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10&
  
  Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName$, lpFindFileData As WIN32_FIND_DATA) As Long
  Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile&, lpFindFileData As WIN32_FIND_DATA) As Long
  Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName$) As Long
  Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile&) As Long
  
  Private Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
  Private Declare Sub GetSystemTimeAsFileTime Lib "kernel32" (lpFileTime As FILETIME)
  
  
  Public Enum BROWSFORFOLDERFLAGS
    BIF_UNDIFINED = &H0&
    BIF_RETURNONLYFSDIRS = &H1&
    BIF_DONTGOBELOWDOMAIN = &H2&
    BIF_STATUSTEXT = &H4&
    BIF_RETURNFSANCESTORS = &H8&
    BIF_EDITBOX = &H10&
    BIF_VALIDATE = &H20&
    BIF_BROWSEFORCOMPUTER = &H1000&
    BIF_BROWSEFORPRINTER = &H2000&
    BIF_BROWSEINCLUDEFILES = &H4000&
  End Enum

  Private Type BrowseInfo
    hOwner As Long
    pIDLRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
  End Type
  
  Private Const WM_USER As Long = &H400&
  
  Private Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100&)
  Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102&)
    
  Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
  Private Declare Function ScreenToClientLong& Lib "user32" Alias "ScreenToClient" (ByVal hWnd&, lpPoint&)
  Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent&, ByVal hWndChildAfter&, ByVal lpClassName$, ByVal lpWindowName$) As Long
  Private Declare Function SHBrowseForFolder Lib "shell32" (lpBrowseInfo As BrowseInfo) As Long
  Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidl&, ByVal pszPath$) As Long

Public Function GetFolder(ByVal hWnd&, Optional ByVal sPrompt$) As String

  Dim bi As BrowseInfo
  Dim pidl&, sPath As String * MAX_PATH
    
  With bi
    .hOwner = hWnd
    .pIDLRoot = 0
    .lpszTitle = sPrompt
  End With
    
  pidl = SHBrowseForFolder(bi)
   
  If pidl Then
    If SHGetPathFromIDList(pidl, sPath) Then
      GetFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1)
    End If
     
    Call CoTaskMemFree(pidl)
  End If
  
End Function


Public Function StripNulls(ByVal sText As String) As String
  ' strips any nulls from the end of a string
  Dim nPosition&
  
  StripNulls = sText
  
  nPosition = InStr(sText, vbNullChar)
  If nPosition Then StripNulls = Left$(sText, nPosition - 1)
  If Len(sText) Then If Left$(sText, 1) = vbNullChar Then StripNulls = vbZLString
End Function

Public Function FileExists(ByVal sFileName$) As Boolean
  'checks if a file or dir exists
  'returns true if it does, returns false otherwise
  Dim HFile&, Win32FindData As WIN32_FIND_DATA
  
  sFileName = Trim$(sFileName)
  HFile = FindFirstFile(sFileName, Win32FindData)
  If (HFile <> INVALID_HANDLE_VALUE) And (HFile <> ERROR_NO_MORE_FILES) Then
    FileExists = True
  ElseIf GetFileAttributes(sFileName) <> (-1) Then
    ' FindFirstFile will not return the root dir of a drive so we check the attributes
    ' of sFileName in case it is the root
    FileExists = True
  End If
  
  Call FindClose(HFile)

End Function

