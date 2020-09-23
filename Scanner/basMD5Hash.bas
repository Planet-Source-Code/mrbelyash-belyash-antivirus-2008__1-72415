Attribute VB_Name = "basMD5Hash"
Option Explicit

Public Const MaxMd5FileLength As Long = 3145728     '3 mb maximum
Public CheckWholeFile As Boolean

Public Function GetFileMD5Hash(ByVal FileName As String) As String

  Dim LengthInBytes As Long
  Dim Nr As Integer, FileLength As Long
  Dim Data As String, FileInfo() As Byte
  Dim Kb As Single
  Dim lLengthInBytes As Long
lLengthInBytes = FileLen(FileName)
    On Error GoTo errors
    
    'save the length the programmer want, will almost always be the file length div 1024
    LengthInBytes = lLengthInBytes
    
    'keep file hash at maximum at MaxMd5FileLength
    If Not CheckWholeFile Then
        If LengthInBytes > MaxMd5FileLength Then
            LengthInBytes = MaxMd5FileLength
        End If
    Else
        'if the file length is to long take whole file, otherwise take amount that user entered
        If LengthInBytes > FileLen(FileName) Then
            LengthInBytes = FileLen(FileName)
        End If
    End If
    
    'if the file does exists
    If FileExists(FileName) Then

        Nr = FreeFile
        'get the filelength
        FileLength = FileLen(FileName)

        If Not CheckWholeFile Then
            'check how much we must hash
            FileLength = Minimum(LengthInBytes, FileLength)
        End If

        'open the file
        Open FileName For Binary As #Nr

        ReDim FileInfo(1 To FileLength)
        'LengthInKb * 1024)

        'get enought data from the file
        Get #Nr, , FileInfo

        'close file
        Close #Nr

        'MD5 the data
        GetFileMD5Hash = MD5String(StrConv(FileInfo, vbUnicode))
      
      Else 'FILEEXISTS(FileName) = FALSE/0
        
      End If
      
      On Error GoTo 0

      Exit Function

errors:
        MsgBox Err.Description, , Err.Number
        
        On Error Resume Next
        GetFileMD5Hash = -1
        Close #Nr

    On Error GoTo 0

End Function

