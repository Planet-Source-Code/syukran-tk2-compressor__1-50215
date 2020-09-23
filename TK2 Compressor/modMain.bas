Attribute VB_Name = "modMain"
Public PathToExtract As String
Public namafile As String
Public Position As Integer
Function FileExist(Filename As String) As Boolean
  On Error GoTo FileDoesNotExist
  Open Filename For Input As #10
  Close #10
  FileExist = True
  Exit Function
FileDoesNotExist:
  FileExist = False
End Function
Sub ToFile(NumFile As Integer, VarType As Integer, Content As Variant)
  Dim TmpInt As Integer
  Dim TmpStr As String
  Dim TmpLng As Long

  Select Case VarType
    Case 0 'Long
      TmpLng = Content
      Put #NumFile, , TmpLng
    Case 1 'Integer
      TmpInt = Content
      Put #NumFile, , TmpInt
    Case 2 'String
      TmpStr = Content
      Put #NumFile, , TmpStr
  End Select
End Sub
Function GetFileFromList(ByVal FileList As String, FileNumber As Integer) As String
  Dim Pos As Integer
  Dim Count As Integer
  Dim FNStart As Integer
  Dim FNLen As Integer
  Dim Path As String

  If InStr(FileList, Chr(0)) = 0 Then
    GetFileFromList = FileList
  Else
    Count = 0
    Path = Left(FileList, InStr(FileList, Chr(0)) - 1)
    If Right(Path, 1) <> "\" Then Path = Path + "\"
    FileList = FileList + Chr(0)
    For Pos = 1 To Len(FileList)
      If Mid$(FileList, Pos, 1) = Chr(0) Then
        Count = Count + 1
        If Count = FileNumber Then FNStart = Pos + 1
        If Count = (FileNumber + 1) Then
          FNLen = Pos - FNStart
          Exit For
        End If
      End If
    Next Pos
    GetFileFromList = Path + Mid(FileList, FNStart, FNLen)
  End If
End Function
Function CountFilesInList(ByVal FileList As String) As Integer
  Dim Count As Integer
  Dim Pos As Integer

  Count = 0
  For Pos = 1 To Len(FileList)
    If Mid$(FileList, Pos, 1) = Chr$(0) Then Count = Count + 1
  Next Pos
  If Count = 0 Then Count = 1
  CountFilesInList = Count
End Function

Function CheckFile(File As String) As Boolean
  Dim TmpInt As Integer

  On Error GoTo Check
  TmpInt = FreeFile
  If GetAttr(File) = vbReadOnly Or GetAttr(File) = vbHidden Then
    CheckFile = False
  ElseIf FileExist(File) = False Then
    CheckFile = False
  Else
    Open File For Binary Access Read Lock Read Write As #TmpInt
    Close #TmpInt
    CheckFile = True
  End If
  Exit Function
Check:
  CheckFile = False
  Resume Next
End Function
