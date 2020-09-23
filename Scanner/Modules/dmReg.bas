Attribute VB_Name = "dmReg"
'íå èñïîëüçóåìàÿ,äëÿ ãåíåðàöèè êëþ÷à è ðåãèñòðàöèè
Public Const VALID_CHARACTERS = "0123456789ABCDEFGHJKLMNPQRSTUVWXYZ"
Public Const DEFAULT_FORMAT = "&&&&-&&&&-&&&&-&&&&"
Public Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public rngready
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
'this function gets the color number under the mouse pointer.
Public Function VolumeSerialNumber(ByVal RootPath As String) As String
Dim VolLabel As String
Dim VolSize As Long
Dim Serial As Long
Dim MaxLen As Long
Dim flags As Long
Dim Name As String
Dim NameSize As Long
Dim S As String
Dim ret As Boolean
ret = GetVolumeSerialNumber(RootPath, VolLabel, VolSize, _
Serial, MaxLen, flags, Name, NameSize)
If ret Then
'Create an 8 character string
S = Format(Hex(Serial), "00000000")
'Adds the '-' between the first 4 characters and the last 4 characters
VolumeSerialNumber = Left(S, 4) + "-" + Right(S, 4)
Else
'If the call to API function fails the function returns a zero serial number
VolumeSerialNumber = "0000-0000"
End If
End Function


Public Function CreateKeyGad(ApplicationKey As String, UserName As String, Optional sFormat As String = DEFAULT_FORMAT, Optional ValidCharacters As String = VALID_CHARACTERS) As String
    ' for use in sFormat; use '&' to represent alpha-numeric characters
    Dim intTemp As Integer
    Dim strTextChar As String
    Dim strKeyChar As String
    Dim intEncryptedChar As String
    Dim strKey As String
    Dim I As Integer
    Dim strUserName As String
    
    
    strUserName = LCase(Trim(UserName))
    
    If Len(strUserName) = 0 Then
        Err.Raise vbError + 1001, , "Invalid Username"
        Exit Function
    End If
    
    'This is an altered simple encryption algorithm
    For I = 1 To CountAmpersands(sFormat)
        strTextChar = Mid(strUserName, (I Mod Len(strUserName)) + 1, 1)
        strKeyChar = Mid(ApplicationKey, (I Mod Len(ApplicationKey)) + 1, 1)
        intTemp = (((Asc(strKeyChar) * I) * Len(ApplicationKey) + 1) Mod Len(ValidCharacters) + 1)
        strTextChar = Chr(Asc(strTextChar) Xor intTemp)
        intTemp = (((Asc(strKeyChar) * I) * Len(UserName) + 1) Mod Len(ValidCharacters) + 1)
        strTextChar = Chr(Asc(strTextChar) Xor intTemp)
        intEncryptedChar = ((Asc(strTextChar) Xor Asc(strKeyChar)) Mod Len(ValidCharacters)) + 1
        strKey = strKey & Mid(ValidCharacters, intEncryptedChar, 1)
    Next I
    
    CreateKeyGad = Format(strKey, sFormat)
End Function

Public Function CountAmpersands(ByVal Format As String) As Integer
    'Counts the number of characters that need to be returned
    
    Dim I As Integer
    Dim intCount As Integer
    
    intCount = 0
    For I = 1 To Len(Format)
        If Mid(Format, I, 1) = "&" Then
            intCount = intCount + 1
        End If
    Next I
    
    CountAmpersands = intCount
End Function

Public Function IsGoodKey(ApplicationKey As String, UserName As String, Key As String, Optional sFormat As String = DEFAULT_FORMAT, Optional ValidCharacters As String = VALID_CHARACTERS) As Boolean
    'This function does not need to exist
    'It is here to make testing the key just a little simpler
    
    If LCase(Trim(Key)) = LCase(CreateKeyGad(ApplicationKey, UserName, sFormat, ValidCharacters)) Then
        IsGoodKey = True
    Else
        IsGoodKey = False
    End If
End Function

Public Function main_k() As Boolean
On Error GoTo 100
main_k = False
If Dir$(App.Path & "\gadyash.key") = "" Then
    MsgBox "Îòñóòñòâóåò êëþ÷" & vbCrLf & "Áîëüøèíñòâî ôóíêöèé çàáëîêèðîâàíî." & vbCrLf & "Çàðåãåñòðèðóéòå ïðîãðàììó"
    Exit Function
End If
Dim z As String
Dim Y As String
Dim gFile As Integer
gFile = FreeFile
Open App.Path + "\gadyash.key" For Input As #gFile
    Line Input #gFile, z
    Y = DecryptPassword("10", z)
    
Close #gFile
Dim GDriv As String
GDriv = Environ("homedrive")
Dim nUzwer As String
nUzwer = dmReg.VolumeSerialNumber(GDriv + "\")
Dim f As String
f = dmReg.CreateKeyGad(nUzwer, Y, DEFAULT_FORMAT, VALID_CHARACTERS)
If dmReg.IsGoodKey(nUzwer, Y, f, DEFAULT_FORMAT, VALID_CHARACTERS) = True Then
    main_k = True
    Exit Function
End If

Exit Function
100:
MsgBox "" + Error$
End Function
Function EncryptPassword(Number As Byte, DecryptedPassword As String)

Dim Password As String, Counter As Byte
Dim Temp As Integer

Counter = 1

Do Until Counter = Len(DecryptedPassword) + 1

Temp = Asc(Mid(DecryptedPassword, Counter, 1))

If Counter Mod 2 = 0 Then

        Temp = Temp–Number

Else
Temp = Temp + Number

End If

Temp = Temp Xor (10 - Number)
Password = Password & Chr$(Temp)
Counter = Counter + 1

Loop

EncryptPassword = Password

End Function
Function DecryptPassword(Number As Byte, EncryptedPassword As String)

Dim Password As String, Counter As Byte
Dim Temp As Integer

Counter = 1

Do Until Counter = Len(EncryptedPassword) + 1

Temp = Asc(Mid(EncryptedPassword, Counter, 1)) Xor (10 - Number)

If Counter Mod 2 = 0 Then

        Temp = Temp + Number

Else
Temp = Temp - Number

End If

Password = Password & Chr$(Temp)
Counter = Counter + 1

Loop

DecryptPassword = Password

End Function

