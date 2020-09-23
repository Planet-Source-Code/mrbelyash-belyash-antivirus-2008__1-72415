Attribute VB_Name = "modHelp"
'******************************************************************************
'----- Modul HTMLHelp API for Visual Basic 6 - (c) Ulrich Kulle
'----- 2002-08-26 Version 1.0 first release
'----- 2005-07-17 Version 1.1 updated for Pop-Up help
'----- 2006-08-27 Version 1.2 minor changes
'----- 2007-05-14 Version 1.3 added HH_AKLINK structure and call
'******************************************************************************
'----- Portions of this code courtesy of David Liske.
'----- Thanks to:
'----- David Liske, Don Lammers, Matthew Brown, Thomas Schulz and Dimitry Karpezo
'------------------------------------------------------------------------------

Type HH_IDPAIR
  dwControlId As Long
  dwTopicId As Long
End Type

' This array should contain the number of controls that have
' context-sensitive help, plus one more for a zero-terminating pair.
'------------------------------------------------------------------------------
Public ids(2) As HH_IDPAIR

Declare Function GetDlgCtrlID Lib "user32" _
                (ByVal hwnd As Long) As Long

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
                (ByVal hWndCaller As Long, ByVal pszFile As String, _
                ByVal uCommand As Long, ByVal dwData As Long) As Long
                
Private Declare Function HTMLHelpTopic Lib "hhctrl.ocx" Alias "HtmlHelpA" _
                 (ByVal hWndCaller As Long, ByVal pszFile As String, _
                ByVal uCommand As Long, ByVal dwData As String) As Long
         
Private Declare Function HtmlHelpSearch Lib "hhctrl.ocx" Alias "HtmlHelpA" _
                (ByVal hWndCaller As Long, ByVal pszFile As String, _
                ByVal uCommand As Long, dwData As HH_FTS_QUERY) As Long
                
Private Declare Function HtmlHelpIndex Lib "hhctrl.ocx" Alias "HtmlHelpA" _
                (ByVal hWndCaller As Long, ByVal pszFile As String, _
                ByVal uCommand As Long, dwData As HH_AKLINK) As Long
         
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                (ByVal hwnd As Long, ByVal wMsg As Long, _
                ByVal wParam As Long, ByVal lParam As Any) As Long
                               
' HH API const
'------------------------------------------------------------------------------
Public Const HH_DISPLAY_TOPIC = &H0         ' select last opened tab, [display a specified topic]
Public Const HH_DISPLAY_TOC = &H1           ' select contents tab, [display a specified topic]
Public Const HH_DISPLAY_INDEX = &H2         ' select index tab and searches for a keyword
Public Const HH_DISPLAY_SEARCH = &H3        ' select search tab (perform a search is not working)
      
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_KEYWORD_LOOKUP = &HD       ' Searches for one or more keywords
Private Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or
  
Private Const HH_HELP_CONTEXT = &HF         ' display mapped numeric value in dwData
        
Private Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Private Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP.

' HH_DISPLAY_SEARCH Command Related Structures and Constants
'-----------------------------------------------------------
Public Const HH_FTS_DEFAULT_PROXIMITY = -1

Public Type HH_FTS_QUERY                ' UDT for accessing the Search tab
  cbStruct          As Long             ' Sizeof structure in bytes.
  fUniCodeStrings   As Long             ' TRUE if all strings are unicode.
  pszSearchQuery    As String           ' String containing the search query.
  iProximity        As Long             ' Word proximity.
  fStemmedSearch    As Long             ' TRUE for StemmedSearch only.
  fTitleOnly        As Long             ' TRUE for Title search only.
  fExecute          As Long             ' TRUE to initiate the search.
  pszWindow         As String           ' Window to display in
End Type

' HH_ALINK Command Related Structures and Constants
'-----------------------------------------------------------
Public Type HH_AKLINK
  cbStruct          As Long
  fReserved         As Boolean
  pszKeywords       As String
  pszUrl            As String
  pszMsgText        As String
  pszMsgTitle       As String
  pszWindow         As String
  fIndexOnFail      As Boolean
End Type



Public Function HFile(ByVal i_HFile As Integer) As String
'----- Set the string variable to include the application path of helpfile
  Select Case i_HFile
  Case 1
    '--- Applications main help file
    HFile = App.Path & "\AVHlp.chm"
  Case 2
'----- [optional] special help window definition
    HFile = App.Path & "\AVHlp.chm"
  Case 3
'----- Place other Help file paths in successive case statements
    HFile = App.Path & "\AVHlp.chm"
  End Select
End Function

Public Sub ShowContents(ByVal intHelpFile As Integer)
   HtmlHelp hwnd, HFile(intHelpFile), HH_DISPLAY_TOC, 0
End Sub

Public Sub ShowIndex(ByVal intHelpFile As Integer)
    HtmlHelp hwnd, HFile(intHelpFile), HH_DISPLAY_INDEX, 0
End Sub

Public Sub ShowIndexKeyword(ByVal intHelpFile As Integer, ByVal sKeyword As String)
' ---------------------------------------------------------------------------------
' from http://msdn.microsoft.com/library/default.asp?url=/library/en-us/htmlhelp/html/vsconStrhhaklink.asp
' ---------------------------------------------------------------------------------
' cbStruct      Specifies the size of the structure. This value must always be filled in before passing the structure to the HTML Help API.
' fReserved     This parameter must be set to FALSE.
' pszKeywords   Specifies one or more ALink names or KLink keywords to look up. Multiple entries are delimited by a semicolon.
' pszUrl        Specifies the topic file to navigate to if the lookup fails. pszURL refers to a valid topic within the specified compiled help (.chm) file and does not support Internet protocols that point to an HTML file.
' pszMsgText    Specifies the text to display in a message box if the lookup fails and fIndexOnFail is FALSE and pszURL is NULL.
' pszMsgTitle   Specifies the caption of the message box in which the pszMsgText parameter appears.
' pszWindow     Specifies the name of the window type in which to display one of the following:
'               - The selected topic, if the lookup yields one or more matching topics.
'               - The topic specified in pszURL, if the lookup fails and a topic is specified in pszURL.
'               The Index tab, if the lookup fails and fIndexOnFail is specified as TRUE.
 
' fIndexOnFail  Specifies whether to display the keyword in the Index tab of the HTML Help Viewer if the lookup fails. The value of pszWindow specifies the Help Viewer.

' ALink name and KLink keyword lookups are case sensitive, and multiple keywords are delimited by a semicolon.
' If the lookup yields one or more matching topics, the topic titles appear in the Topics Found dialog box.
' -----------------------------------------------------------------------------------
Dim keyData As HH_AKLINK
  With keyData
    .cbStruct = Len(keyData)                          'sizeof this structure
    .fReserved = 0&                                   'must be FALSE (really!)
    .pszKeywords = sKeyword
    .pszUrl = vbNullString
    .pszMsgText = "keyword found: " + sKeyword
    .pszMsgTitle = "Keyword Not Found"
    .pszWindow = ""
    .fIndexOnFail = 0&        'FALSE
  End With

'--- use HH_DISPLAY_INDEX to open the help window  ---------
  HTMLHelpTopic hwnd, HFile(intHelpFile), HH_DISPLAY_INDEX, sKeyword
'--- use HH_KEYWORD_LOOKUP to open the topic in the rigt pane  ---------
  HtmlHelpIndex hwnd, HFile(intHelpFile), HH_KEYWORD_LOOKUP, keyData
End Sub

Public Sub ShowTopic(ByVal intHelpFile As Integer, strTopic As String)
    HTMLHelpTopic hwnd, HFile(intHelpFile), HH_DISPLAY_TOPIC, strTopic
End Sub

Public Sub ShowTopicID(ByVal intHelpFile As Integer, IdTopic As Long)
  HtmlHelp hwnd, HFile(intHelpFile), HH_HELP_CONTEXT, IdTopic
End Sub

Public Sub ShowSearch(ByVal intHelpFile As Integer)
'------------------------------------------------------------------------------
' show search tab
' ** BUG: ** HTML Help: HH_DISPLAY_SEARCH API Command Does Not Perform a Search
' http://support.microsoft.com/kb/241381/en-us
'------------------------------------------------------------------------------
Dim searchIt As HH_FTS_QUERY
  With searchIt
    .cbStruct = Len(searchIt)                   ' Size of structute in bytes
    .fUniCodeStrings = 0&                       ' TRUE if all strings are unicode
    .pszSearchQuery = "help-1"                  ' String containing the search query
    .iProximity = HH_FTS_DEFAULT_PROXIMITY      ' word proximity
    .fStemmedSearch = 0&                        ' TRUE for StemmedSearch only
    .fTitleOnly = 1&                            ' TRUE for Title Search only
    .fExecute = 0&                              ' TRUE to initiate the search
    .pszWindow = ""                             ' Window to display in
  End With
  HtmlHelpSearch 0&, HFile(intHelpFile), HH_DISPLAY_SEARCH, searchIt
End Sub
