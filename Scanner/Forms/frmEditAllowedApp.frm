VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEditAppBlock 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Application Blocker"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   Icon            =   "frmEditAllowedApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   375
      Left            =   210
      TabIndex        =   10
      Top             =   5670
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   1410
      TabIndex        =   9
      Top             =   5670
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deny"
      Height          =   375
      Left            =   5790
      TabIndex        =   8
      Top             =   5670
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Allow"
      Height          =   375
      Left            =   4470
      TabIndex        =   7
      Top             =   5670
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   210
      TabIndex        =   6
      Top             =   240
      Width           =   7035
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   300
      TabIndex        =   5
      Top             =   3360
      Width           =   6975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1890
      Top             =   5730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAllow 
      Caption         =   "Allow"
      Height          =   375
      Left            =   4530
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeny 
      Caption         =   "Deny"
      Height          =   375
      Left            =   5850
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2940
      TabIndex        =   2
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   1470
      TabIndex        =   1
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   270
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frmEditAppBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()
'Word Documents ( *.doc )|*.doc |"
 Dim strFileType As String

strFileType = "All Files (*.*)|*.*|"
strFileType = strFileType & " Executable File (*.exe)|*.exe|"
'StrFileType = StrFileType & " Text Files (*.txt)|*.txt|"
    'Me.CommonDialog1.Filter = "(*.exe )| Executable File"
    CommonDialog1.Filter = strFileType

    Me.CommonDialog1.DialogTitle = "Browse for EXE File"
    Me.CommonDialog1.ShowOpen
    
    'file selected
    If Me.CommonDialog1.FileName <> "" Then
        'ask for action
        Dim X As ListItem
        If MsgBox("Deny This Application To Run", vbYesNo + vbDefaultButton1, "Add New Application") = vbYes Then
            'add to list
         Me.List1.AddItem (Me.CommonDialog1.FileName)
          '  X.SubItems(1) = "Allow"
          '  Set X = Nothing
        'Else
            'add to list
        '    Set X = Me.lvAppList.ListItems.Add(, , Me.CommonDialog1.FileName)
        '    X.SubItems(1) = "Deny"
        '    Set X = Nothing
        End If
    End If
End Sub

Private Sub Command4_Click()
    If MsgBox("Are you sure to remove the selected item?", vbYesNoCancel + vbDefaultButton3, "Remove Application") = vbYes Then
        Me.List1.RemoveItem (List1.ListIndex)
    End If
End Sub

'===============
' FORM LOAD
'===============
Private Sub Form_Load()
'    Me.lvAppList.ListItems.Clear
   ' Call LoadBanList
   ' Call LoadAllowList
    'GetAllKeys17

End Sub
Public Function GetAllKeys17()
Dim X As ListItem, hKey As Long, lCount As Long, i As Long
' "-------------Options programm start----------------"
Dim apppa As String
Dim RegExt As String


'apppa = GetString(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", "AppPath")
'analize "AppPath=" & apppa
    'Enumerate from HKEY_LOCAL_MACHINE , Run
   ' hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus")
  '  lCount = GetCount(hKey, Values)
   ' For i = 1 To lCount - 1
        'Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        'X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        'Set X = Nothing
        'analize EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus", EnumValue(hKey, i))
        
'       lvAppList.ListItems.Add GetKeyValue(hKey, EnumValue(hKey, i))
 '   Next i
    ' "-----Ban-----"
   

hKey = OpenKey(GetClassKey(&H80000001), "Software\BGAntivirus\Ban")
    lCount = GetCount(hKey, Values)
    For i = 1 To lCount - 1
        'Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        'X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        'Set X = Nothing
        'lvAppList.ListItems.Add EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", EnumValue(hKey, i))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i
   '"-----Allow-----"
hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow")
    lCount = GetCount(hKey, Values)
    For i = 1 To lCount - 1
        'Set X = lvStartUp.ListItems.Add(, , EnumValue(hKey, i))
        'X.SubItems(1) = GetKeyValue(hKey, EnumValue(hKey, i))
        'X.SubItems(2) = "HKEY_LOCAL_MACHINE"
        'Set X = Nothing
'        lvAppList.ListItems.Add EnumValue(hKey, i) & "=" & getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", EnumValue(hKey, i))
        
        'analize GetKeyValue(hKey, EnumValue(hKey, i))
    Next i


End Function
'===================
'FUNCTIONS
'===================
Sub LoadBanList()
    Dim bCount As Long
    Dim i As Long
    If getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count") <> "" Then
        bCount = getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", "Count")
    Else
        bCount = 0
    End If
    Dim X As ListItem
    For i = 1 To bCount
        'Set X = Me.lvAppList.ListItems.Add(, , getstring(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", CStr(i)))
        'X.SubItems(1) = "Deny"
        'Set X = Nothing
    Next
End Sub

Sub LoadAllowList()
     Dim aCount As Long
    Dim i As Long
    If getstring(&H80000001, "Software\BGAntivirus\Allow", "Count") <> "" Then
        aCount = getstring(&H80000001, "Software\BGAntivirus\Allow", "Count")
    Else
        aCount = 0
    End If
    Dim X As ListItem
    For i = 1 To aCount
        'Set X = Me.lvAppList.ListItems.Add(, , getstring(&H80000001, "Software\BGAntivirus\Allow", CStr(i)))
        'X.SubItems(1) = "Allow"
        'Set X = Nothing
    Next
End Sub

'=====================
' EVENTS
'=====================

Private Sub cmdAdd_Click()
 'Word Documents ( *.doc )|*.doc |"
 Dim strFileType As String

strFileType = "All Files (*.*)|*.*|"
strFileType = strFileType & " Executable File (*.exe)|*.exe|"
'StrFileType = StrFileType & " Text Files (*.txt)|*.txt|"
    'Me.CommonDialog1.Filter = "(*.exe )| Executable File"
    CommonDialog1.Filter = strFileType

    Me.CommonDialog1.DialogTitle = "Browse for EXE File"
    Me.CommonDialog1.ShowOpen
    
    'file selected
    If Me.CommonDialog1.FileName <> "" Then
        'ask for action
        Dim X As ListItem
        If MsgBox("Allow This Application To Run", vbYesNo + vbDefaultButton1, "Add New Application") = vbYes Then
            'add to list
         Me.List2.AddItem (Me.CommonDialog1.FileName)
          '  X.SubItems(1) = "Allow"
          '  Set X = Nothing
        'Else
            'add to list
        '    Set X = Me.lvAppList.ListItems.Add(, , Me.CommonDialog1.FileName)
        '    X.SubItems(1) = "Deny"
        '    Set X = Nothing
        End If
    End If
End Sub



Private Sub cmdRemove_Click()
    If MsgBox("Are you sure to remove the selected item?", vbYesNoCancel + vbDefaultButton3, "Remove Application") = vbYes Then
        Me.List2.RemoveItem (List2.ListIndex)
    End If
End Sub

Private Sub cmdSave_Click()
    Dim lCount As Long, bCount As Long
    Dim i As Long, hKey As Long
    lCount = 0
    'clear previous list in registry
    'from allow list
hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow")
    lCount = GetCount(hKey, Values)
    For i = 1 To lCount - 1
        'use delete registry (function in startup)
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", CStr(i)
    Next
      bCount = 0
    
    hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban")
    bCount = GetCount(hKey, Values)
    For i = 1 To bCount - 1
        'use delete registry (function in startup)
        DeleteStartup GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Ban", CStr(i)
    Next
       
    hKey = OpenKey(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow")
    lCount = GetCount(hKey, Values)
    'add from list to registry
    
    Dim aCount As Integer
    
    For i = 1 To Me.List2.ListCount - 1
                    aCount = aCount + 1
            'update in registry / count
            'add app to allow list
            Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", 1, CStr(aCount), Me.List2.List(i))
      Next i
                   Dim zCount As Integer
 For i = 1 To Me.List1.ListCount - 1
                     zCount = zCount + 1
            'update in registry / count
            'add app to allow list
            Call CreateStringValue(GetClassKey("HKEY_CURRENT_USER"), "Software\BGAntivirus\Allow", 1, CStr(zCount), Me.List1.List(i))
      Next i
    
    MsgBox "Save Complete", vbInformation, "Save Application List"
    Unload Me
End Sub
