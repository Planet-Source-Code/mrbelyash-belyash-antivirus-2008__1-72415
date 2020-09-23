Attribute VB_Name = "modBinaryFile"
Option Explicit
'Public base() As String
Public virCount As Long
Dim base() As virus
Private Type virus

    Name As String * 35
    Value As String * 8
   
End Type

Public Sub ReadSig()
ReDim base(0)
  Dim n As Integer
    n = FreeFile
    Dim a As String
    Dim B As String
     Dim C As String
    Dim k As Integer
    Dim f As Long
Open App.Path + "\1.go" For Input As #n
 While Not EOF(n)
     Line Input #n, a
    k = InStr(1, a, "|")
If k <> 0 Then
    B = Left$(a, k - 1)
    C = Right(a, Len(a) - (k + 1))
    If Len(C) > 65 Then
    'MsgBox "35"
    End If
    ReDim Preserve base(UBound(base) + 1)
    base(f).Name = C
    base(f).Value = B
    Debug.Print B + "===" + C + "====>" + CStr(f)
   ' Call APPbASE(f)
    f = f + 1
 End If
    
    
 Wend
    Close #n
    'MsgBox "done"










Exit Sub
'virCount = 0
'Dim a() As String
'Dim i As Integer
'ReDim base(1)
   ' ScanForFiles App.Path + "\", "*.sig", a()
  '  For i = 0 To UBound(a())
  '      Debug.Print a(i)
  '      readZ Trim$(a(i))
  '      Call main_base
   ' Next

'ReDim VSig(UBound(base))
'Dim g As Long
'For g = 0 To UBound(base)
'VSig(g) = base(g)

  '  base(g).Name = VSig(g).Name
  '  base(g).Value = VSig(g).Value
  '   base(g).Action = VSig(g).Action
   '  base(g).ActtionVal = VSig(g).ActtionVal
 '   base(g).Type = VSig(g).Type

'Next g
Erase base()
frmMain.lblFound.Caption = virCount
  
End Sub

Public Function CheckSig(mdcheck As String) As Boolean
'virCount = 0
CheckSig = False
Dim a() As String
Dim z As Integer
    ScanForFiles App.Path + "\", "*.sig", a()
For z = 0 To UBound(a())
        Debug.Print a(z)
        'readZ Trim$(a(i))
        Dim f As Long
    f = FreeFile
    Open Trim$(a(z)) For Binary Access Read As #f
        Get #f, , VSInfo
        ReDim VSig(VSInfo.VirusCount - 1) As VirusSig
        Dim i As Long
        For i = 0 To VSInfo.VirusCount - 1
            Get #f, , VSig(i)
            If mdcheck = VSig(i).Value Then
                CheckSig = True
                MsgBox "Íàøåë â îäíîé èç áàç âèðóñ.Áàçà-" + a(i) + "---" + VSig(i).Name
                Exit Function
            
            End If
            
            'virCount = virCount + 1
        Next i
        Close #f
        
Next z



   
End Function
Sub readZ(mybase As String)
'MsgBox mybase
Dim f As Long
    On Error GoTo Trap_Error
    f = FreeFile
    Open mybase For Binary Access Read As #f
        Get #f, , VSInfo
        ReDim VSig(VSInfo.VirusCount - 1) As VirusSig
        Dim i As Long
        For i = 0 To VSInfo.VirusCount - 1
            Get #f, , VSig(i)
            'Debug.Print VSig(i).Name + "===" + CStr(i)
            virCount = virCount + 1
        Next
    Close #f
'frmMain.SB.Panels(6).Text = virCount
   On Error GoTo 0
   Exit Sub

Trap_Error:
Debug.Print Error$
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetData of Form frmBinaAccess"
End Sub



Public Sub WriteSig(ByRef vs As VirusSig)
    
    Dim f As Long
    On Error GoTo Trap_Error
    f = FreeFile
    
    Dim i As Long
    
    'add 1 item into array
    ReDim Preserve VSig(UBound(VSig) + 1) As VirusSig
    VSig(UBound(VSig)).Name = vs.Name
    VSig(UBound(VSig)).Type = vs.Type
    VSig(UBound(VSig)).Value = vs.Value
    
    'add 1 for count
    VSInfo.VirusCount = UBound(VSig) + 1
    VSInfo.LastUpdate = Format(Date, "dd/mmmm/yyyy")
    
    'change virus last update
    'VSInfo.LastUpdate = Format("07 June 2007", "Short Date")
    
    Open App.Path & "\WDAV.sig" For Binary Access Write As #f
        Put #f, , VSInfo
        For i = 0 To UBound(VSig)
            Put #f, , VSig(i)
        Next
    Close #f

   On Error GoTo 0
   Exit Sub

Trap_Error:

    MsgBox "Îøèáêà " & Err.Number & " (" & Err.Description & ") in procedure PutData of Form frmBinaAccess"
End Sub

