Attribute VB_Name = "Module1"
Public Type PasswordSave
 Encyrpt As Integer
 EncryptString As String
End Type

Public Function EncryptTOFile(ByVal StrToEncrypt As String)
Dim iString, LString As Integer
Dim PasSav As PasswordSave
Randomize

iString = (Rnd * 4 + 1)
iString = CInt(iString)
LString = Len(StrToEncrypt)
tStr = iString + LString
tStr = CInt(tStr / 2)
PasSav.Encyrpt = tStr

For i = 1 To LString
x = Mid$(StrToEncrypt, i, 1)
xasc# = Asc(x)
newChar = Chr$((xasc# + tStr))
PasSav.EncryptString = PasSav.EncryptString + newChar
Next i

MsgBox PasSav.EncryptString

'The Decryptor

For i = 1 To LString
x = Mid$(PasSav.EncryptString, i, 1)
xasc = Asc(x)
newChar = Chr$(xasc# - tStr)
oldstring = oldstring + newChar
Next i

MsgBox oldstring


End Function


 
