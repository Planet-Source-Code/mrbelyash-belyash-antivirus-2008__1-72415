Attribute VB_Name = "basRegistry"
'Option Explicit

' --------------------------------------------------------------
' Update the Windows registry.
' Written by Kenneth Ives                        kenaso@home.com
' NT tested by Brett Gerhardi       Brett.Gerhardi@trinite.co.uk
'
' Perform the four basic functions on the Windows registry.
'           Add
'           Change
'           Delete
'           Query
'
' Important:   If you treat all key data strings as being
'              case sensitive, you should never have a problem.
'              Always backup your registry files (System.dat
'              and User.dat) before performing any type of
'              modifications
'
' Software developers vary on where they want to update the
' registry with their particular information.  The most common
' are in HKEY_lOCAL_MACHINE or HKEY_CURRENT_USER.
'
' This BAS module handles all of my needs for string and
' basic numeric updates in the Windows registry.
'
' Brett found that NT users must delete each major key
' separately.  See bottom of TEST routine for an example.
' --------------------------------------------------------------

' --------------------------------------------------------------
' Private variables
' --------------------------------------------------------------
  Private m_lngRetVal As Long


Private Type HUFFMANTREE
  ParentNode As Integer
  RightNode As Integer
  LeftNode As Integer
  Value As Integer
  Weight As Long
End Type

Private Type ByteArray
  Count As Byte
  Data() As Byte
End Type

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

  
' --------------------------------------------------------------
' Constants required for values in the keys
' --------------------------------------------------------------
  Private Const REG_NONE As Long = 0                  ' No value type
  Private Const REG_SZ As Long = 1                    ' nul terminated string
  Private Const REG_EXPAND_SZ As Long = 2             ' nul terminated string w/enviornment var
  Private Const REG_BINARY As Long = 3                ' Free form binary
  Private Const REG_DWORD As Long = 4                 ' 32-bit number
  Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4   ' 32-bit number (same as REG_DWORD)
  Private Const REG_DWORD_BIG_ENDIAN As Long = 5      ' 32-bit number
  Private Const REG_LINK As Long = 6                  ' Symbolic Link (unicode)
  Private Const REG_MULTI_SZ As Long = 7              ' Multiple Unicode strings
  Private Const REG_RESOURCE_LIST As Long = 8         ' Resource list in the resource map
  Private Const REG_FULL_RESOURCE_DESCRIPTOR As Long = 9 ' Resource list in the hardware description
  Private Const REG_RESOURCE_REQUIREMENTS_LIST As Long = 10

' --------------------------------------------------------------
' Registry Specific Access Rights
' --------------------------------------------------------------
  Private Const KEY_QUERY_VALUE As Long = &H1
  Private Const KEY_SET_VALUE As Long = &H2
  Private Const KEY_CREATE_SUB_KEY As Long = &H4
  Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
  Private Const KEY_NOTIFY As Long = &H10
  Private Const KEY_CREATE_LINK As Long = &H20
  Private Const KEY_ALL_ACCESS As Long = &H3F

' --------------------------------------------------------------
' Constants required for key locations in the registry
' --------------------------------------------------------------
  Private Const HKEY_CLASSES_ROOT As Long = &H80000000
  Private Const HKEY_CURRENT_USER As Long = &H80000001
  Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
  Private Const HKEY_USERS As Long = &H80000003
  Private Const HKEY_PERFORMANCE_DATA As Long = &H80000004
  Private Const HKEY_CURRENT_CONFIG As Long = &H80000005
  Private Const HKEY_DYN_DATA As Long = &H80000006

' --------------------------------------------------------------
' Constants required for return values (Error code checking)
' --------------------------------------------------------------
  Private Const ERROR_SUCCESS As Long = 0
  Private Const ERROR_ACCESS_DENIED As Long = 5
  Private Const ERROR_NO_MORE_ITEMS As Long = 259

' --------------------------------------------------------------
' Open/Create constants
' --------------------------------------------------------------
  Private Const REG_OPTION_NON_VOLATILE As Long = 0
  Private Const REG_OPTION_VOLATILE As Long = &H1

' --------------------------------------------------------------
' Declarations required to access the Windows registry
' --------------------------------------------------------------
  Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal lngRootKey As Long) As Long
  
  Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String) As Long
  
  Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String) As Long
  
  Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
            (ByVal lngRootKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
             lpType As Long, lpData As Any, lpcbData As Long) As Long
  
  Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
            (ByVal lngRootKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
             ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Type STARTUPINFO
cb As Long
lpReserved As String
lpDesktop As String
lpTitle As String
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Long
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type
Private Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessID As Long
dwThreadId As Long
End Type
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey&, ByVal lpClass$, lpcbClass&, ByVal lpReserved&, lpcSubKeys&, lpcbMaxSubKeyLen&, lpcbMaxClassLen&, lpcValues&, lpcbMaxValueNameLen&, lpcbMaxValueLen&, lpcbSecurityDescriptor&, lpftLastWriteTime As Any) As Long
Private Declare Function RegNotifyChangeKeyValue Lib "advapi32" (ByVal hKey As Long, ByVal bWatchSubTree As Boolean, ByVal dwNotifyFilter As Long, ByVal hEvent As Long, ByVal fAsynchronous As Boolean) As Long

Public Sub ExecCmd(cmdline$)
Dim proc As PROCESS_INFORMATION
Dim start As STARTUPINFO
' Èíèöèàëèçèðóåì ñòðóêòóðó STARTUPINFO:
start.cb = Len(start)
' Çàïóñêàåì ïðèëîæåíèå:
ret& = CreateProcessA(0&, cmdline$, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)
' Æäåì çàâåðøåíèÿ çàïóùåííîãî ïðèëîæåíèÿ:
ret& = WaitForSingleObject(proc.hProcess, INFINITE)
ret& = CloseHandle(proc.hProcess)
End Sub
Private Function regDelete_Sub_Key(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String, _
                                  ByVal strRegSubKey As String)
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for removing a sub key.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be removed.
'
' Syntax:
'    regDelete_Sub_Key HKEY_CURRENT_USER, _
                  "Software\AAA-Registry Test\Products", "StringTestData"
'
' Removes the sub key "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the sub key.  If it does not exist, then ignore it.
      m_lngRetVal = RegDeleteValue(lngKeyHandle, strRegSubKey)
  
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

Private Function regDoes_Key_Exist(ByVal lngRootKey As Long, _
                                  ByVal strRegKeyPath As String) As Boolean
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function to see if a key does exist
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you want to test
'
' Syntax:
'    strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, _
'                       "Software\AAA-Registry Test\Products")
'
' Returns the value of TRUE or FALSE
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long

' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  lngKeyHandle = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regDoes_Key_Exist = False
  Else
      regDoes_Key_Exist = True
  End If
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function

Private Function regQuery_A_Key(ByVal lngRootKey As Long, _
                               ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String) As Variant
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for querying a sub key value.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be queryed.
'
' Syntax:
'    strKeyQuery = regQuery_A_Key(HKEY_CURRENT_USER, _
'                       "Software\AAA-Registry Test\Products", _
                        "StringTestData")
'
' Returns the key value of "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim intPosition As Integer
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngBufferSize As Long
  Dim lngBuffer As Long
  Dim strBuffer As String

' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  lngKeyHandle = 0
  lngBufferSize = 0
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave here.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Query the registry and determine the data type.
' --------------------------------------------------------------
  m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, _
                         lngDataType, ByVal 0&, lngBufferSize)
  
' --------------------------------------------------------------
' If no key handle was found then there is no key.  Leave.
' --------------------------------------------------------------
  If lngKeyHandle = 0 Then
      regQuery_A_Key = ""
      m_lngRetVal = RegCloseKey(lngKeyHandle)   ' always close the handle
      Exit Function
  End If
  
' --------------------------------------------------------------
' Make the API call to query the registry based on the type
' of data.
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data (most common)
              ' Preload the receiving buffer area
              strBuffer = Space(lngBufferSize)
      
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, _
                                     ByVal strBuffer, lngBufferSize)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Strip out the string data
                  intPosition = InStr(1, strBuffer, Chr(0))  ' look for the first null char
                  If intPosition > 0 Then
                      ' if we found one, then save everything up to that point
                      regQuery_A_Key = Left(strBuffer, intPosition - 1)
                  Else
                      ' did not find one.  Save everything.
                      regQuery_A_Key = strBuffer
                  End If
              End If
              
         Case REG_DWORD:    ' Numeric data (Integer)
              m_lngRetVal = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                     lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
              
              ' If NOT a successful call then leave
              If m_lngRetVal <> ERROR_SUCCESS Then
                  regQuery_A_Key = ""
              Else
                  ' Save the captured data
                  regQuery_A_Key = lngBuffer
              End If
         
         Case Else:    ' unknown
              regQuery_A_Key = ""
  End Select
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function
Private Sub regCreate_Key_Value(ByVal lngRootKey As Long, ByVal strRegKeyPath As String, _
                               ByVal strRegSubKey As String, varRegData As Variant)
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for saving string data.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'     strRegSubKey - is the name of the key which will be updated.
'       varRegData - Update data.
'
' Syntax:
'    regCreate_Key_Value HKEY_CURRENT_USER, _
'                      "Software\AAA-Registry Test\Products", _
'                      "StringTestData", "22 Jun 1999"
'
' Saves the key value of "22 Jun 1999" to sub key "StringTestData"
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  Dim lngDataType As Long
  Dim lngKeyValue As Long
  Dim strKeyValue As String
  
' --------------------------------------------------------------
' Determine the type of data to be updated
' --------------------------------------------------------------
  If IsNumeric(varRegData) Then
      lngDataType = REG_DWORD
  Else
      lngDataType = REG_SZ
  End If
  
' --------------------------------------------------------------
' Query the key path
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)
    
' --------------------------------------------------------------
' Update the sub key based on the data type
' --------------------------------------------------------------
  Select Case lngDataType
         Case REG_SZ:       ' String data
              strKeyValue = Trim(varRegData) & Chr(0)     ' null terminated
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          ByVal strKeyValue, Len(strKeyValue))
                                   
         Case REG_DWORD:    ' numeric data
              lngKeyValue = CLng(varRegData)
              m_lngRetVal = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                          lngKeyValue, 4&)  ' 4& = 4-byte word (long integer)
                                   
  End Select
  
' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Sub
Private Function regCreate_A_Key(ByVal lngRootKey As Long, ByVal strRegKeyPath As String)

' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   This function will create a new key
'
' Parameters:
'          lngRootKey  - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                  HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'   strRegKeyPath  - is name of the key you wish to create.
'                  to make sub keys, continue to make this
'                  call with each new level.  MS says you
'                  can do this in one call; however, the
'                  best laid plans of mice and men ...
'
' Syntax:
'   regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test"
'   regCreate_A_Key HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products"
' --------------------------------------------------------------

' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Create the key.  If it already exist, ignore it.
' --------------------------------------------------------------
  m_lngRetVal = RegCreateKey(lngRootKey, strRegKeyPath, lngKeyHandle)

' --------------------------------------------------------------
' Always close the handle in the registry.  We do not want to
' corrupt these files.
' --------------------------------------------------------------
  m_lngRetVal = RegCloseKey(lngKeyHandle)
  
End Function
Private Function regDelete_A_Key(ByVal lngRootKey As Long, _
                                ByVal strRegKeyPath As String, _
                                ByVal strRegKeyName As String) As Boolean
    
' --------------------------------------------------------------
' Written by Kenneth Ives                     kenaso@home.com
'
' Important:     If you treat all key data strings as being
'                case sensitive, you should never have a problem.
'                Always backup your registry files (System.dat
'                and User.dat) before performing any type of
'                modifications
'
' Description:   Function for removing a complete key.
'
' Parameters:
'           lngRootKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER,
'                        HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'    strRegKeyPath - is name of the key path you wish to traverse.
'   strRegKeyValue - is the name of the key which will be removed.
'
' Returns a True or False on completion.
'
' Syntax:
'    regDelete_A_Key HKEY_CURRENT_USER, "Software", "AAA-Registry Test"
'
' Removes the key "AAA-Registry Test" and all of its sub keys.
' --------------------------------------------------------------
    
' --------------------------------------------------------------
' Define variables
' --------------------------------------------------------------
  Dim lngKeyHandle As Long
  
' --------------------------------------------------------------
' Preset to a failed delete
' --------------------------------------------------------------
  regDelete_A_Key = False
  
' --------------------------------------------------------------
' Make sure the key exist before trying to delete it
' --------------------------------------------------------------
  If regDoes_Key_Exist(lngRootKey, strRegKeyPath) Then
  
      ' Get the key handle
      m_lngRetVal = RegOpenKey(lngRootKey, strRegKeyPath, lngKeyHandle)
      
      ' Delete the key
      m_lngRetVal = RegDeleteKey(lngKeyHandle, strRegKeyName)
      
      ' If the value returned is equal zero then we have succeeded
      If m_lngRetVal = 0 Then regDelete_A_Key = True
      
      ' Always close the handle in the registry.  We do not want to
      ' corrupt the registry.
      m_lngRetVal = RegCloseKey(lngKeyHandle)
  End If
  
End Function

Function regPiS(lng1 As String, softKey As String, softName As String, softZn As String) As Boolean
  regPiS = False
  On Error GoTo 10
  Dim lngRootKey As Long
  Dim strKeyQuery As Variant    ' we are not sure what type of
                              ' data will be returned
  
' --------------------------------------------------------------
' Initialize variables
' --------------------------------------------------------------
  strKeyQuery = vbNullString
  'Dim softKey As String
  'Dim softName As String
  'Dim softZn As Variant
  Select Case lng1
  Case "1"
    lngRootKey = HKEY_CLASSES_ROOT
  Case "2"
    lngRootKey = HKEY_CURRENT_USER
  Case "3"
    lngRootKey = HKEY_LOCAL_MACHINE
                 
  Case "4"
    lngRootKey = HKEY_USERS
  Case "5"
    lngRootKey = HKEY_CURRENT_CONFIG
  Case Else
    regPiS = False
    Exit Function
  End Select
  'softKey = ""
  'softName = ""
  'softZn = 1
  
  If Not regDoes_Key_Exist(lngRootKey, softKey) Then
      ' create the main key and the first sub key
     regCreate_A_Key lngRootKey, softKey
  End If
  If Not regDoes_Key_Exist(lngRootKey, softKey) Then
      ' create the first sub key
      regCreate_A_Key lngRootKey, softKey
  End If
  strKeyQuery = regQuery_A_Key(lngRootKey, softKey, softName)
  'MsgBox "÷èòàþ " + CStr(strKeyQuery)
' Change the value of the sub key, "StringTestData", from
regCreate_Key_Value lngRootKey, softKey, _
                      softName, softZn
' See if the sub key has been updated
  strKeyQuery = regQuery_A_Key(lngRootKey, softKey, softName)
  'MsgBox "ïðîâåðÿþ " + CStr(strKeyQuery)
  regPiS = True
  Exit Function
10:
regPiS = False
End Function
'Huffman Encoding/Decoding Class
'-------------------------------
'
'(c) 2000, Fredrik Qvarfort
'

Public Sub CompressFile(SourceFile As String, DestFile As String)

  Dim ByteArray() As Byte
  Dim Filenr As Integer
  
  'Make sure the source file exists
  If (Not FileExist(SourceFile)) Then Exit Sub
  
  'Read the data from the sourcefile
  Open SourceFile For Binary As #Filenr
  ReDim ByteArray(LOF(Filenr) - 1)
  Get #Filenr, , ByteArray()
  Close #Filenr
  
  'Compress the data
  Call EncodeByte(ByteArray(), UBound(ByteArray) + 1)
  
  'If the destination file exist we need to
  'destroy it because opening it as binary
  'will not clear the old data
  If (FileExist(DestFile)) Then Kill DestFile
  
  'Save the destination string
  Open DestFile For Binary As #Filenr
  Put #Filenr, , ByteArray()
  Close #Filenr

End Sub
Public Sub UncompressFile(SourceFile As String, DestFile As String)

  Dim ByteArray() As Byte
  Dim Filenr As Integer
  
  'Make sure the source file exists
  If (Not FileExist(SourceFile)) Then Exit Sub
  
  'Read the data from the sourcefile
  Open SourceFile For Binary As #Filenr
  ReDim ByteArray(LOF(Filenr) - 1)
  Get #Filenr, , ByteArray()
  Close #Filenr
  
  'Uncompress the data
  Call DecodeByte(ByteArray(), UBound(ByteArray) + 1)
  
  'If the destination file exist we need to
  'destroy it because opening it as binary
  'will not clear the old data
  If (FileExist(DestFile)) Then Kill DestFile
  
  'Save the destination string
  Open DestFile For Binary As #Filenr
  Put #Filenr, , ByteArray()
  Close #Filenr

End Sub
Private Sub CreateTree(Nodes() As HUFFMANTREE, NodesCount As Long, Char As Long, Bytes As ByteArray)

  Dim a As Integer
  Dim NodeIndex As Long
  
  NodeIndex = 0
  For a = 0 To (Bytes.Count - 1)
    If (Bytes.Data(a) = 0) Then
      'Left node
      If (Nodes(NodeIndex).LeftNode = -1) Then
        Nodes(NodeIndex).LeftNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).LeftNode
    ElseIf (Bytes.Data(a) = 1) Then
      'Right node
      If (Nodes(NodeIndex).RightNode = -1) Then
        Nodes(NodeIndex).RightNode = NodesCount
        Nodes(NodesCount).ParentNode = NodeIndex
        Nodes(NodesCount).LeftNode = -1
        Nodes(NodesCount).RightNode = -1
        Nodes(NodesCount).Value = -1
        NodesCount = NodesCount + 1
      End If
      NodeIndex = Nodes(NodeIndex).RightNode
    Else
      
    End If
  Next
  
  Nodes(NodeIndex).Value = Char

End Sub
Public Sub EncodeByte(ByteArray() As Byte, ByteLen As Long)
  
  Dim i As Long
  Dim j As Long
  Dim Char As Byte
  Dim BitPos As Byte
  Dim lNode1 As Long
  Dim lNode2 As Long
  Dim lNodes As Long
  Dim lLength As Long
  Dim Count As Integer
  Dim lWeight1 As Long
  Dim lWeight2 As Long
  Dim Result() As Byte
  Dim ByteValue As Byte
  Dim ResultLen As Long
  Dim Bytes As ByteArray
  Dim NodesCount As Integer
  Dim BitValue(0 To 7) As Byte
  Dim CharCount(0 To 255) As Long
  Dim Nodes(0 To 511) As HUFFMANTREE
  Dim CharValue(0 To 255) As ByteArray
  
  'If the source string is empty or contains
  'only one character we return it uncompressed
  'with the prefix string "HEO" & vbCr
  If (ByteLen = 0) Then
    ReDim Preserve ByteArray(ByteLen + 3)
    If (ByteLen > 0) Then
      Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
    End If
   ' ByteArray(0) = 72 '"H"
  '  ByteArray(1) = 69 '"E"
 '   ByteArray(2) = 48 '"0"
' ByteArray(3) = 13    'vbCr
ByteArray(0) = 112 '"P"
ByteArray(1) = 114 '"R"
ByteArray(2) = 103 '"G"
ByteArray(3) = 13 'vbCr
    Exit Sub
  End If
  
  'Create the temporary result array and make
  'space for identifier, checksum, textlen and
  'the ASCII values inside the Huffman Tree
  ReDim Result(0 To 522)
  
  'Prefix the destination string with the
  '"HE3" & vbCr identification string

  Result(0) = 112
  Result(1) = 114
  Result(2) = 103
  Result(3) = 13
  ResultLen = 4
  
  'Count the frequency of each ASCII code
  For i = 0 To (ByteLen - 1)
    CharCount(ByteArray(i)) = CharCount(ByteArray(i)) + 1
  Next
  
  'Create a leaf for each character
  For i = 0 To 255
    If (CharCount(i) > 0) Then
      With Nodes(NodesCount)
        .Weight = CharCount(i)
        .Value = i
        .LeftNode = -1
        .RightNode = -1
        .ParentNode = -1
      End With
      NodesCount = NodesCount + 1
    End If
  Next
  
  'Create the Huffman Tree
  For lNodes = NodesCount To 2 Step -1
    'Get the two leafs with the smallest weights
    lNode1 = -1: lNode2 = -1
    For i = 0 To (NodesCount - 1)
      If (Nodes(i).ParentNode = -1) Then
        If (lNode1 = -1) Then
          lWeight1 = Nodes(i).Weight
          lNode1 = i
        ElseIf (lNode2 = -1) Then
          lWeight2 = Nodes(i).Weight
          lNode2 = i
        ElseIf (Nodes(i).Weight < lWeight1) Then
          If (Nodes(i).Weight < lWeight2) Then
            If (lWeight1 < lWeight2) Then
              lWeight2 = Nodes(i).Weight
              lNode2 = i
            Else
              lWeight1 = Nodes(i).Weight
              lNode1 = i
            End If
          Else
            lWeight1 = Nodes(i).Weight
            lNode1 = i
          End If
        ElseIf (Nodes(i).Weight < lWeight2) Then
          lWeight2 = Nodes(i).Weight
          lNode2 = i
        End If
      End If
    Next
    
    'Create a new leaf
    With Nodes(NodesCount)
      .Weight = lWeight1 + lWeight2
      .LeftNode = lNode1
      .RightNode = lNode2
      .ParentNode = -1
      .Value = -1
    End With
    
    'Set the parentnodes of the two leafs
    Nodes(lNode1).ParentNode = NodesCount
    Nodes(lNode2).ParentNode = NodesCount
    
    'Increase the node counter
    NodesCount = NodesCount + 1
  Next

  'Traverse the tree to get the bit sequence
  'for each character, make temporary room in
  'the data array to hold max theoretical size
  ReDim Bytes.Data(0 To 255)
  Call CreateBitSequences(Nodes(), NodesCount - 1, Bytes, CharValue)
  
  'Calculate the length of the destination
  'string after encoding
  For i = 0 To 255
    If (CharCount(i) > 0) Then
      lLength = lLength + CharValue(i).Count * CharCount(i)
    End If
  Next
  lLength = IIf(lLength Mod 8 = 0, lLength \ 8, lLength \ 8 + 1)
  
  'If the destination is larger than the source
  'string we leave it uncompressed and prefix
  'it with a 4 byte header ("HE0" & vbCr)
  If ((lLength = 0) Or (lLength > ByteLen)) Then
    Call CopyMem(ByteArray(4), ByteArray(0), ByteLen)
ByteArray(0) = 112 '"P"
ByteArray(1) = 114 '"R"
ByteArray(2) = 103 '"G"
ByteArray(3) = 13 'vbCr
    Exit Sub
  End If
  
  'Add a simple checksum value to the result
  'header for corruption identification
  Char = 0
  For i = 0 To (ByteLen - 1)
    Char = Char Xor ByteArray(i)
  Next
  Result(ResultLen) = Char
  ResultLen = ResultLen + 1
  
  'Add the length of the source string to the
  'header for corruption identification
  Call CopyMem(Result(ResultLen), ByteLen, 4)
  ResultLen = ResultLen + 4
  
  'Create a small array to hold the bit values,
  'this is faster than calculating on-fly
  For i = 0 To 7
    BitValue(i) = 2 ^ i
  Next
  
  'Store the number of characters used
  Count = 0
  For i = 0 To 255
    If (CharValue(i).Count > 0) Then
      Count = Count + 1
    End If
  Next
  Call CopyMem(Result(ResultLen), Count, 2)
  ResultLen = ResultLen + 2
  
  'Store the used characters and the length
  'of their respective bit sequences
  Count = 0
  For i = 0 To 255
    If (CharValue(i).Count > 0) Then
      Result(ResultLen) = i
      ResultLen = ResultLen + 1
      Result(ResultLen) = CharValue(i).Count
      ResultLen = ResultLen + 1
      Count = Count + 16 + CharValue(i).Count
    End If
  Next
  
  'Make room for the Huffman Tree in the
  'destination byte array
  ReDim Preserve Result(ResultLen + Count \ 8)
  
  'Store the Huffman Tree into the result
  'converting the bit sequences into bytes
  BitPos = 0
  ByteValue = 0
  For i = 0 To 255
    With CharValue(i)
      If (.Count > 0) Then
        For j = 0 To (.Count - 1)
          If (.Data(j)) Then ByteValue = ByteValue + BitValue(BitPos)
          BitPos = BitPos + 1
          If (BitPos = 8) Then
            Result(ResultLen) = ByteValue
            ResultLen = ResultLen + 1
            ByteValue = 0
            BitPos = 0
          End If
        Next
      End If
    End With
  Next
  If (BitPos > 0) Then
    Result(ResultLen) = ByteValue
    ResultLen = ResultLen + 1
  End If
  
  'Resize the destination string to be able to
  'contain the encoded string
  ReDim Preserve Result(ResultLen - 1 + lLength)
  
  'Now we can encode the data by exchanging each
  'ASCII byte for its appropriate bit string.
  Char = 0
  BitPos = 0
  For i = 0 To (ByteLen - 1)
    With CharValue(ByteArray(i))
      For j = 0 To (.Count - 1)
        If (.Data(j) = 1) Then Char = Char + BitValue(BitPos)
        BitPos = BitPos + 1
        If (BitPos = 8) Then
          Result(ResultLen) = Char
          ResultLen = ResultLen + 1
          BitPos = 0
          Char = 0
        End If
      Next
    End With
  Next

  'Add the last byte
  If (BitPos > 0) Then
    Result(ResultLen) = Char
    ResultLen = ResultLen + 1
  End If
  
  'Return the destination in string format
  ReDim ByteArray(ResultLen - 1)
  Call CopyMem(ByteArray(0), Result(0), ResultLen)

End Sub
Public Function DecodeString(Text As String) As String
  
  Dim ByteArray() As Byte
  
  'Convert the string to a byte array
  ByteArray() = StrConv(Text, vbFromUnicode)
  
  'Compress the byte array
  Call DecodeByte(ByteArray, Len(Text))
  
  'Convert the compressed byte array to a string
  Text = StrConv(ByteArray(), vbUnicode)
  
End Function
Public Function EncodeString(Text As String) As String
  
  Dim ByteArray() As Byte
  
  'Convert the string to a byte array
  ByteArray() = StrConv(Text, vbFromUnicode)
  
  'Compress the byte array
  Call EncodeByte(ByteArray, Len(Text))
  
  'Convert the compressed byte array to a string
  Text = StrConv(ByteArray(), vbUnicode)
  
End Function

Public Function DecodeByte(ByteArray() As Byte, ByteLen As Long)
  
  Dim i As Long
  Dim j As Long
  Dim Pos As Long
  Dim Char As Byte
  Dim CurrPos As Long
  Dim Count As Integer
  Dim CheckSum As Byte
  Dim Result() As Byte
  Dim BitPos As Integer
  Dim NodeIndex As Long
  Dim ByteValue As Byte
  Dim ResultLen As Long
  Dim NodesCount As Long
  Dim lResultLen As Long
  Dim BitValue(0 To 7) As Byte
  Dim Nodes(0 To 511) As HUFFMANTREE
  Dim CharValue(0 To 255) As ByteArray


  If (ByteArray(0) <> 112) Or (ByteArray(1) <> 114) Or (ByteArray(3) <> 103) Then
    'The source did not contain the identification
    'string "HE?" & vbCr where ? is undefined at
    'the moment (does not matter)
  ElseIf (ByteArray(2) = 103) Then
    'The text is uncompressed, return the substring
    'Decode = Mid$(Text, 5)
    Call CopyMem(ByteArray(0), ByteArray(4), ByteLen - 4)
    ReDim Preserve ByteArray(ByteLen - 5)
    Exit Function
  ElseIf (ByteArray(2) <> 51) Then
    'This is not a Huffman encoded string
    Err.Raise vbObjectError, "HuffmanDecode()", "The data either was not compressed with HE3 or is corrupt (identification string not found)"
    Exit Function
  End If
  
  CurrPos = 5
    
  'Extract the checksum
  CheckSum = ByteArray(CurrPos - 1)
  CurrPos = CurrPos + 1
  
  'Extract the length of the original string
  Call CopyMem(ResultLen, ByteArray(CurrPos - 1), 4)
  CurrPos = CurrPos + 4
  lResultLen = ResultLen
  
  'If the compressed string is empty we can
  'skip the function right here
  If (ResultLen = 0) Then Exit Function
  
  'Create the result array
  ReDim Result(ResultLen - 1)
  
  'Get the number of characters used
  Call CopyMem(Count, ByteArray(CurrPos - 1), 2)
  CurrPos = CurrPos + 2
  
  'Get the used characters and their
  'respective bit sequence lengths
  For i = 1 To Count
    With CharValue(ByteArray(CurrPos - 1))
      CurrPos = CurrPos + 1
      .Count = ByteArray(CurrPos - 1)
      CurrPos = CurrPos + 1
      ReDim .Data(.Count - 1)
    End With
  Next
  
  'Create a small array to hold the bit values,
  'this is (still) faster than calculating on-fly
  For i = 0 To 7
    BitValue(i) = 2 ^ i
  Next
  
  'Extract the Huffman Tree, converting the
  'byte sequence to bit sequences
  ByteValue = ByteArray(CurrPos - 1)
  CurrPos = CurrPos + 1
  BitPos = 0
  For i = 0 To 255
    With CharValue(i)
      If (.Count > 0) Then
        For j = 0 To (.Count - 1)
          If (ByteValue And BitValue(BitPos)) Then .Data(j) = 1
          BitPos = BitPos + 1
          If (BitPos = 8) Then
            ByteValue = ByteArray(CurrPos - 1)
            CurrPos = CurrPos + 1
            BitPos = 0
          End If
        Next
      End If
    End With
  Next
  If (BitPos = 0) Then CurrPos = CurrPos - 1
  
  'Create the Huffman Tree
  NodesCount = 1
  Nodes(0).LeftNode = -1
  Nodes(0).RightNode = -1
  Nodes(0).ParentNode = -1
  Nodes(0).Value = -1
  For i = 0 To 255
    Call CreateTree(Nodes(), NodesCount, i, CharValue(i))
  Next
  
  'Decode the actual data
  ResultLen = 0
  For CurrPos = CurrPos To ByteLen
    ByteValue = ByteArray(CurrPos - 1)
    For BitPos = 0 To 7
      If (ByteValue And BitValue(BitPos)) Then
        NodeIndex = Nodes(NodeIndex).RightNode
      Else
        NodeIndex = Nodes(NodeIndex).LeftNode
      End If
      If (Nodes(NodeIndex).Value > -1) Then
        Result(ResultLen) = Nodes(NodeIndex).Value
        ResultLen = ResultLen + 1
        If (ResultLen = lResultLen) Then GoTo DecodeFinished
        NodeIndex = 0
      End If
    Next
  Next
DecodeFinished:

  'Verify data to check for corruption.
  Char = 0
  For i = 0 To (ResultLen - 1)
    Char = Char Xor Result(i)
  Next
  If (Char <> CheckSum) Then
    Err.Raise vbObjectError, "clsHuffman.Decode()", "The data might be corrupted (checksum did not match expected value)"
  End If

  'Return the uncompressed string
  ReDim ByteArray(ResultLen - 1)
  Call CopyMem(ByteArray(0), Result(0), ResultLen - 1)
  
End Function
Private Sub CreateBitSequences(Nodes() As HUFFMANTREE, ByVal NodeIndex As Integer, Bytes As ByteArray, CharValue() As ByteArray)

  Dim NewBytes As ByteArray
  
  'If this is a leaf we set the characters bit
  'sequence in the CharValue array
  If (Nodes(NodeIndex).Value > -1) Then
    CharValue(Nodes(NodeIndex).Value) = Bytes
    Exit Sub
  End If
  
  'Traverse the left child
  If (Nodes(NodeIndex).LeftNode > -1) Then
    NewBytes = Bytes
    NewBytes.Data(NewBytes.Count) = 0
    NewBytes.Count = NewBytes.Count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).LeftNode, NewBytes, CharValue)
  End If
  
  'Traverse the right child
  If (Nodes(NodeIndex).RightNode > -1) Then
    NewBytes = Bytes
    NewBytes.Data(NewBytes.Count) = 1
    NewBytes.Count = NewBytes.Count + 1
    Call CreateBitSequences(Nodes(), Nodes(NodeIndex).RightNode, NewBytes, CharValue)
  End If
  
End Sub

Private Function FileExist(Filename As String) As Boolean

  On Error GoTo FileDoesNotExist
  
  Call FileLen(Filename)
  FileExist = True
  Exit Function
  
FileDoesNotExist:
  FileExist = False
  
End Function
Function GetRegKeys(hKey As String, strSubKey As String, list As Object)
Dim hChildKey As Long, lngSubKeys As Long, lngMaxKeySize As Long, _
lngDataRetBytes As Long, i As Integer, strRetArray() As String
rtn = RegOpenKeyEx(hKey, strSubKey, 0, KEY_ALL_ACCESS, hChildKey)
rtn = QueryRegInfoKey(hChildKey, lngSubKeys, lngMaxKeySize)
 lngSubKeys = lngSubKeys - 1
 On Error Resume Next
 ReDim strRetArray(lngSubKeys) As String
    For i = 0 To lngSubKeys
        lngDataRetBytes = lngMaxKeySize
        strRetArray(i) = Space(lngMaxKeySize)
        RegEnumKeyEx hChildKey, i, strRetArray(i), lngDataRetBytes, 0&, vbNullString, ByVal 0&, ByVal 0&
        list.AddItem Left(strRetArray(i), lngDataRetBytes)
    Next i
    If Len(strSubKey) Then RegCloseKey hChildKey
    GetRegKeys = strRetArray
End Function
Function QueryRegInfoKey(hKey&, Optional lngSubKeys&, Optional lngMaxKeyLen&, Optional lngValues&, Optional lngMaxValNameLen&, Optional lngMaxValLen&)
QueryRegInfoKey = RegQueryInfoKey(hKey, vbNullString, ByVal 0&, 0&, lngSubKeys, lngMaxKeyLen, ByVal 0&, lngValues, lngMaxValNameLen, lngMaxValLen, ByVal 0&, ByVal 0&)
lngMaxKeyLen = lngMaxKeyLen + 1
lngMaxValNameLen = lngMaxValNameLen + 1
lngMaxValLen = lngMaxValLen + 1
End Function
Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
Dim lValueType As Long, strBuf As String, lDataBufSize As Long
rtn = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
 If rtn = 0 Then
  If lValueType = REG_EXPAND_SZ Then
    strBuf = String(lDataBufSize, Chr$(0))
    rtn = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
    If rtn = 0 Then
     RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
    End If
  End If
    If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, Chr$(0))
    rtn = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
    If rtn = 0 Then
     RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
    End If
  End If
 End If
End Function

