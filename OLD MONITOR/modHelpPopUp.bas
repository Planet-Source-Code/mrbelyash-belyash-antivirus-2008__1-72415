Attribute VB_Name = "modHelpPopUp"
'******************************************************************************
'----- Modul HH API for Visual Basic 6 - (c) Ulrich Kulle
'----- 2007-07-19 not finished - Version 0.9 test for PopUp only
'******************************************************************************
               
Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hWndCaller As Long, ByVal pszFile As String, _
    ByVal uCommand As Long, dwData As Any) As Long
               
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                (ByVal hwnd As Long, ByVal wMsg As Long, _
                ByVal wParam As Long, ByVal lParam As Any) As Long
                
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
                (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
                ByVal lpClassName As String, ByVal lpWindowName As String) As Long
                
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" _
                (ByVal hwnd As Long, ByVal lpClassName As String, _
                ByVal nMaxCount As Long) As Long
        
Public Declare Function GetDlgItem Lib "user32" _
                (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
                
' API const
'------------------------------------------------------------------------------
Public Const WM_COMMAND As Long = &H111     ' WM_COMMAND = 0x0111
Public Const WM_SETTEXT As Long = &HC       ' WM_SETTEXT = 0x000C

                
' HH API const
'------------------------------------------------------------------------------
Public Const HH_DISPLAY_TOPIC = &H0         ' select last opened tab, [display a specified topic]
Public Const HH_DISPLAY_TOC = &H1           ' select contents tab, [display a specified topic]
Public Const HH_DISPLAY_INDEX = &H2         ' select index tab and searches for a keyword
Public Const HH_DISPLAY_SEARCH = &H3        ' select search tab (perform a search is not working)
      
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or
  
Private Const HH_HELP_CONTEXT = &HF          ' display mapped numeric value in dwData
     
Private Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Private Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP.

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type

Private Type POINTAPI
    X   As Long
    Y   As Long
End Type

' HH_DISPLAY_SEARCH Command Related Structures and Constants
'-----------------------------------------------------------
Public Const HH_FTS_DEFAULT_PROXIMITY = -1
Public Const HH_MAX_TABS = 19               ' maximum number of tabs

'Propiedades públicas
Public hwnd As Long         'El hWnd del form que llama a la ayuda
Public HelpFile As String   'El fichero de ayuda con extensión chm


Public Type tagHH_POPUP                        ' UDT for text popups
  cbStruct          As Long                 ' sizeof this structure
  hinst             As Long                 ' instance handle for string resource
  idString          As Long                 ' string resource id, or text id if pszFile is specified in HtmlHelp call
  pszText           As String               ' used if idString is zero
  pt                As POINTAPI             ' top center of popup window
  clrForeground     As ColorConstants       ' either use VB constant or &HBBGGRR
  clrBackground     As ColorConstants       ' either use VB constant or &HBBGGRR
  rcMargins         As RECT                 ' amount of space between edges of window and text, -1 for each member to ignore
  pszFont           As String               ' facename, point size, char set, BOLD ITALIC UNDERLINE
End Type

Public Type HH_FTS_QUERY                    ' UDT for accessing the Search tab
  cbStruct          As Long                 ' Sizeof structure in bytes.
  fUniCodeStrings   As Long                 ' TRUE if all strings are unicode.
  pszSearchQuery    As String               ' String containing the search query.
  iProximity        As Long                 ' Word proximity.
  fStemmedSearch    As Long                 ' TRUE for StemmedSearch only.
  fTitleOnly        As Long                 ' TRUE for Title search only.
  fExecute          As Long                 ' TRUE to initiate the search.
  pszWindow         As String               ' Window to display in
End Type

Public Function PopUp(ByVal sHelpFile As String, ByVal Text As String) As Long
    'Para mostrar una ventana PopUp con el texto indicado
    Dim HH_POPUP As tagHH_POPUP
    Dim elForm As Form
    'Asignar el form activo
    Set elForm = Screen.ActiveForm
    
    On Local Error Resume Next
    MsgBox (Text)
    Text = "(c) 2007 HelpInformation Ulrich Kulle"
  
    With HH_POPUP
        .cbStruct = Len(HH_POPUP)
        .clrBackground = -1
        .clrForeground = -1
        .pszFont = "Verdana,8"
        .hinst = vbNullString
        .idString = 100  '---topic number in a text file.
        .pszText = vbNullString
        'posicionar la ventana de PopUp
        .pt.X = (elForm.Left + 360) \ Screen.TwipsPerPixelX
        .pt.Y = (elForm.Top + (elForm.Height \ 2) + 240) \ Screen.TwipsPerPixelY
        .rcMargins.Bottom = -1
        .rcMargins.Left = -1
        .rcMargins.Right = -1
        .rcMargins.Top = -1
    End With
    
    MsgBox ("Before: " & HH_POPUP.pszText)
    
    PopUp = HtmlHelp(hwnd, sHelpFile, HH_DISPLAY_TEXT_POPUP, HH_POPUP)

    MsgBox ("After: " & HH_POPUP.pszText)
    
    #If ES_DEBUG Then
        Debug.Print "PopUp= " & PopUp
    #End If
    
    Err = 0
End Function

Public Function HelpContextPop(Optional ByVal elControl As Control) As Long
    'Se pasará el control en el que se ha pulsado F1
    'Deberá tener asignado el valor del HelpContextID
    'Esta función es para usar con VB5 o con un formulario en el que no se
    'ha especificado el WhatThisHelp...
    Dim vControl As Control
    
    On Local Error Resume Next
    
    If elControl Is Nothing Then
        Set vControl = Screen.ActiveControl
    Else
        Set vControl = elControl
    End If
    
    ids(0).dwTopicId = CLng(vControl.HelpContextID)
    ids(0).dwControlId = GetDlgCtrlID(vControl.hwnd)
    ' The last pair in the array must contain zeros (0)
    ids(1).dwControlId = 0
    ids(1).dwTopicId = 0
    
    If Err = 0 Then
        HelpContextPop = HtmlHelp(vControl.hwnd, HelpFile, HH_TP_HELP_WM_HELP, ids(0))
    End If
    
    Err = 0
End Function



