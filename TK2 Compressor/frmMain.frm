VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TK2 COMPRESSOR"
   ClientHeight    =   6030
   ClientLeft      =   1455
   ClientTop       =   2055
   ClientWidth     =   4380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6030
   ScaleWidth      =   4380
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   -120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3255
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   11
         ToolTipText     =   "Add files to compress"
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdExtract 
         Caption         =   "&Extract..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   10
         ToolTipText     =   "Extract files"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Delete files"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Quit TK2 Compressor"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New..."
         Height          =   375
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Create new achive"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblfn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.ListBox lstFiles 
      Height          =   2595
      Left            =   2040
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   2520
      Width           =   2175
   End
   Begin VB.FileListBox flsFiles 
      Height          =   2235
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.DirListBox dirFolders 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1935
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Thanks for using this software [tk2_vb@yahoo.com]"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5760
      Width           =   3975
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00C0FFC0&
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   5280
      Width           =   2175
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuusing 
         Caption         =   "&Using TK2 Compressor"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About TK2 Compressor"
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuExtract 
         Caption         =   "Extract..."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete..."
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select all"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Temporal As String = "Compressed.tmp"
Dim SelectedFile As String, Here As String, NameOfFile As String, Errors As String, OldDrive As String
Dim LongOfFile As Long
Dim CantiFiles As Integer, FileToDelete As Integer
Private Compresion As clsCompresion

Function Dialog(Operation As String) As String
Start:
  On Error GoTo Verify
  With dlgDialog
    .Filename = ""
    .InitDir = dirFolders.Path
    .MaxFileSize = 10000
    Select Case LCase(Operation)
      Case "a"
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNExplorer + cdlOFNAllowMultiselect + cdlOFNLongNames
        .Filter = "All files (*.*)|*.*"
        .DialogTitle = "Select file(s) to compress"
        .ShowOpen
      Case "g"
        .Flags = cdlOFNOverwritePrompt + cdlOFNNoChangeDir + cdlOFNHideReadOnly
        .Filter = "All files (*.*)|*.*"
        .DialogTitle = "Save file as"
        .Filename = ReduceDot(lstFiles.List(Position), "(")
        .ShowSave
      Case "n"
        .Flags = cdlOFNOverwritePrompt + cdlOFNNoChangeDir + cdlOFNHideReadOnly
        .Filter = "tk2 compressed file (*.tk2)|*.tk2"
        .DialogTitle = "New archive"
        .ShowSave
    End Select
    Dialog = .Filename
  End With
  Exit Function
Verify:
  If Err.Number = 20476 Then
    If MsgBox("Too much files selected" + vbCrLf + "Try again?", vbInformation + vbYesNo) = vbYes Then
      Resume Start
    Else
      Dialog = ""
    End If
  End If
End Function
Function ReduceDot(Chain As String, StrToExtract As String, Optional DontRemove As Variant) As String
  ReduceDot = Chain
  If Not (InStr(ReduceDot, StrToExtract) = 0) Then
    If Not IsMissing(DontRemove) = False Then
      ReduceDot = Left$(ReduceDot, InStr(ReduceDot, StrToExtract) - 2)
    Else
      ReduceDot = Left$(ReduceDot, InStr(ReduceDot, StrToExtract) - 1)
    End If
  End If
End Function

Function ReduceSize(FileSize As Long) As String
  Select Case FileSize / 1024
    Case Is < 1
      ReduceSize = ReduceDot(Trim(Str(FileSize)), ".", True) + " bytes"
    Case Is < 1024
      ReduceSize = ReduceDot(Trim(Str(FileSize / 1024)), ".", True) + " KB"
    Case Is < 1024 ^ 2
      ReduceSize = ReduceDot(Trim(Str(FileSize / 1024 ^ 2)), ".", True) + " MB"
    Case Else
      ReduceSize = ReduceDot(Trim(Str(FileSize / 1024 ^ 3)), ".", True) + " GB"
  End Select
End Function

Function Simplificate(ByVal Chain As String) As String
  While InStr(Chain, "\") <> 0
    Chain = Mid$(Chain, InStr(Chain, "\") + 1)
  Wend
  Simplificate = Chain
End Function

Private Sub cmdAdd_Click()
  Dim File As String, File2 As String, TmpStr As String
  Dim i As Integer, ip As Integer
  Dim TmpLng As Long, TmpLng2 As Long

  If CheckFile(SelectedFile) = False Then
    MsgBox "Cannot open file", vbCritical
    Exit Sub
  End If
  Errors = ""
  File2 = Dialog("a")
  If File2 = "" Then Exit Sub
  frmProgress.Show
  For ip = 1 To CountFilesInList(File2)
    File = GetFileFromList(ByVal File2, ip)
    Open SelectedFile For Binary As #1
    TmpStr = Simplificate(File)
    For i = 1 To lstFiles.ListCount
      If LCase(ReduceDot(lstFiles.List(i - 1), "(")) = LCase(TmpStr) Then
        Progress 0
        Select Case MsgBox("File '" + TmpStr + "' already exists in archive." + vbCrLf + "Update it?", vbQuestion + vbYesNoCancel)
          Case vbCancel
            GoTo NextFile
          Case vbYes
            TmpLng = Loc(1)
            FileToDelete = i
            Close
            cmdDelete_Click
            Open SelectedFile For Binary As #1
            Seek #1, TmpLng + 1
            FileToDelete = 0
        End Select
        Exit For
      End If
    Next i
    Get #1, , CantiFiles
    Seek #1, 1
    ToFile 1, 1, CantiFiles + 1
    Seek #1, LOF(1) + 1
    ToFile 1, 2, Chr$(Len(TmpStr)) + TmpStr
    TmpLng = Loc(1)
    ToFile 1, 0, 0
    Close
    Compresion.EncodeFile File, SelectedFile
    Open SelectedFile For Binary As #1
    TmpLng2 = LOF(1) - TmpLng - 4
    Seek #1, TmpLng + 1
    ToFile 1, 0, TmpLng2
NextFile:
    Close
  Next ip
  Unload frmProgress
  flsFiles_Click
  dirFolders.Refresh
  If Errors <> "" Then
    If MsgBox("Some errors were found when compressing." + vbCrLf + "Do you want to get more details?", vbQuestion + vbYesNo) = vbYes Then
      Load frmErrors
      frmErrors.txtErrors = Errors
      frmErrors.Show vbModal
    End If
  End If
End Sub
Private Sub cmdDelete_Click()
  Dim TmpStr() As String, TmpStr2 As String
  Dim TmpInt As Integer, i As Integer, ip As Integer, NewVal As Integer, OldVal As Integer
  Dim TmpBol As Boolean

  If FileToDelete = 0 Then
    If MsgBox("Are you sure you want to delete selected files?" + vbCrLf + "Remember that if a file is repeated, all of them will be deleted", vbQuestion + vbYesNo) = vbNo Then Exit Sub
  Else
    lstFiles.ListIndex = FileToDelete - 1
  End If
  ReDim TmpStr(0)
  For i = 1 To lstFiles.ListCount
    If lstFiles.Selected(i - 1) = True Then
      TmpStr2 = ReduceDot(lstFiles.List(i - 1), "(")
      TmpBol = False
      For ip = 1 To UBound(TmpStr)
        If LCase(TmpStr(ip)) = LCase(TmpStr2) Then
          TmpBol = True
          Exit For
        End If
      Next ip
      If TmpBol = False Then
        ReDim Preserve TmpStr(UBound(TmpStr) + 1)
        TmpStr(UBound(TmpStr)) = TmpStr2
      End If
    End If
  Next i
  Open SelectedFile For Binary Access Read As #1
  Open Here + Temporal For Binary Access Write As #2
  Get #1, , CantiFiles
  ToFile 2, 1, CantiFiles
  If FileToDelete = 0 Then frmProgress.Show
  frmProgress.lblAction.Caption = "Searching file..."
  For i = 1 To CantiFiles
    NewVal = i / CantiFiles * 99
    If OldVal <> NewVal Then Progress NewVal
    OldVal = NewVal
    NameOfFile = Input$(Asc(Input$(1, 1)), 1)
    TmpBol = False
    For ip = 1 To UBound(TmpStr)
      If LCase(TmpStr(ip)) = LCase(NameOfFile) Then
        TmpBol = True
        Exit For
      End If
    Next ip
    If TmpBol = True Then
      If Len(NameOfFile) > 12 Then NameOfFile = Left$(NameOfFile, 10) + "..."
      frmProgress.lblAction.Caption = "Deleting " + NameOfFile
      Get #1, , LongOfFile
      Seek #1, Loc(1) + LongOfFile + 1
    Else
      TmpInt = TmpInt + 1
      ToFile 2, 2, Chr$(Len(NameOfFile)) + NameOfFile
      Get #1, , LongOfFile
      ToFile 2, 0, LongOfFile
      ToFile 2, 2, Input$(LongOfFile, 1)
    End If
  Next i
  Seek #2, 1
  ToFile 2, 1, TmpInt
  Close
  Kill SelectedFile
  Name Here + Temporal As SelectedFile
  If FileToDelete = 0 Then Unload frmProgress
  LoadFiles
  flsFiles_Click
  If FileToDelete = 0 And TmpInt = 0 Then
    If MsgBox("Archive does not contain any file." + vbCrLf + "Do you want to delete archive too?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
      Kill SelectedFile
      flsFiles.Refresh
      flsFiles_Click
    End If
  End If
  FileToDelete = 0
End Sub
Private Sub cmdExtract_Click()
  Dim File As String, TmpStr As String
  Dim i As Integer
  Dim Location As Long

  Errors = ""
  frmExtract.Show vbModal
  File = PathToExtract
  If File = "" Then Exit Sub
  If Right$(File, 1) <> "\" Then File = File + "\"
  Open SelectedFile For Binary Access Read As #1
  Get #1, , CantiFiles
  frmProgress.Show
  For i = 1 To CantiFiles
    TmpStr = Input$(Asc(Input$(1, 1)), 1)
    Get #1, , LongOfFile
    Location = Loc(1)
    If lstFiles.Selected(i - 1) = True Then
      If FileExist(File + TmpStr) = True Then
        Progress 0
        If MsgBox("File '" + TmpStr + "' already exists." + vbCrLf + "Overwrite it?", vbQuestion + vbYesNo) = vbNo Then
Again:
          Position = i - 1
          TmpStr = Dialog("g")
          If TmpStr <> "" Then
            If LongOfFile = 0 Then
              Open TmpStr For Output As #2
              Close #2
            Else
              Compresion.DecodeFile SelectedFile, TmpStr
            End If
          End If
        Else
          If CheckFile(File + TmpStr) = False Then
            MsgBox "Cannot overwrite file", vbCritical
            GoTo Again
          End If
          GoTo Continue
        End If
      Else
Continue:
        If LongOfFile = 0 Then
          Open File + TmpStr For Output As #2
          Close #2
        Else
          Compresion.DecodeFile SelectedFile, File + TmpStr
        End If
      End If
    End If
    Seek #1, Location + LongOfFile + 1
  Next i
  Close
  Unload frmProgress
  If Errors <> "" Then
    If MsgBox("Some errors were found when decompressing." + vbCrLf + "Do you want to get more details?", vbQuestion + vbYesNo) = vbYes Then
      Load frmErrors
      frmErrors.txtErrors = Errors
      frmErrors.Show vbModal
    End If
  End If
End Sub

Private Sub cmdNew_Click()
  frm1.Show
End Sub
  
Private Sub makedir()
On Error GoTo baf
Dim File As String
  
  File = flsFiles.Path & "\" & lblfn.Caption & ".tk2"
  If File = "" Then Exit Sub
  Open File For Output As #1
  Close
  Open File For Binary Access Write As #1
  ToFile 1, 1, 0
  Close
  flsFiles.Refresh
  
Exit Sub
baf:
End Sub

Private Sub newf()
On Error GoTo yak
Dim File As String
  
  File = Dialog("n")
  If File = "" Then Exit Sub
  Open File For Output As #1
  Close
  Open File For Binary Access Write As #1
  ToFile 1, 1, 0
  Close
  flsFiles.Refresh
 
  
Exit Sub
yak:
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Sub AddError(Chain As String)
  If Errors = "" Then
    Errors = Chain
  Else
    Errors = Errors + vbCrLf + Chain
  End If
End Sub

Sub Progress(Porcent As Integer)
  frmProgress.shaProgreso.Width = Porcent / 100 * frmProgress.picProgreso.Width
  frmProgress.lblPorc = Format$(Porcent) + "%"
  DoEvents
End Sub

Private Sub Command1_Click()
On Error GoTo bag
Dim path1 As String
path1 = dirFolders.Path
MkDir path1 & "\" & namafile
dirFolders.Refresh
dirFolders.Path = path1 & "\" & namafile

Exit Sub
bag:
End Sub

Private Sub dirFolders_Change()
  flsFiles.Path = dirFolders.Path
End Sub

Private Sub drvDrives_Change()
  On Error GoTo CheckError
  dirFolders.Path = drvDrives.Drive
  Exit Sub
CheckError:
  If Err.Number = 68 Then
    If MsgBox("This device is not ready", vbCritical + vbRetryCancel) = vbRetry Then
      drvDrives_Change
    Else
      drvDrives.Drive = OldDrive
    End If
  End If
End Sub

Private Sub flsFiles_Click()
  Dim TmpStr As String

  lblInfo.Caption = ""
  TmpStr = flsFiles.Filename
  If Right$(dirFolders.Path, 1) = "\" Then
    SelectedFile = dirFolders.Path + CheckFileCase(TmpStr, dirFolders.Path)
  Else
    SelectedFile = dirFolders.Path + "\" + CheckFileCase(TmpStr, dirFolders.Path + "\")
  End If
  If LCase(Right$(flsFiles.Filename, 4)) = ".tk2" Then
    ShowButtons True
    LoadFiles
  Else
    ShowButtons False
  End If
End Sub

Sub LoadFiles()
  Dim TmpInt As Integer, i As Integer

  On Error GoTo Check
  lstFiles.Clear
  Open SelectedFile For Binary Access Read As #1
  Get #1, , CantiFiles
  For i = 1 To CantiFiles
    TmpInt = Asc(Input$(1, 1))
    NameOfFile = Input$(TmpInt, 1)
    Get #1, , LongOfFile
    lstFiles.AddItem NameOfFile + " (" + ReduceSize(LongOfFile) + ")"
    Seek #1, Loc(1) + LongOfFile + 1
  Next i
  Close
  lblInfo.Caption = Format$(CantiFiles) + " file(s) in archive"
  Exit Sub
Check:
  flsFiles.Refresh
  lstFiles.Clear
  MsgBox "Unknown structure in archive '" + Simplificate(SelectedFile) + "'." + vbCrLf + "Cannot load it.", vbCritical
End Sub
Sub ShowButtons(Action As Boolean)
  cmdExtract.Enabled = False
  cmdDelete.Enabled = False
  If Action = True Then
    cmdAdd.Enabled = True
  Else
    cmdAdd.Enabled = False
    lstFiles.Clear
  End If
End Sub

Private Sub flsFiles_PathChange()
  ShowButtons False
End Sub

Private Sub Form_Load()
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
  Set Compresion = New clsCompresion
  If Right$(App.Path, 1) = "\" Then
    Here = App.Path
  Else
    Here = App.Path + "\"
  End If
  OldDrive = drvDrives.Drive
  dirFolders.Path = App.Path
End Sub

Private Sub lblfn_Change()
On Error GoTo ac:
Dim path1 As String
path1 = dirFolders.Path
MkDir path1 & "\" & lblfn.Caption
dirFolders.Refresh
dirFolders.Path = path1 & "\" & lblfn.Caption
Call makedir
flsFiles.ListIndex = 0
cmdAdd.SetFocus
Exit Sub
ac:
MsgBox ("Folder already exist")
End Sub


Private Sub lstFiles_Click()
  Dim TmpInt As Integer, i As Integer

  For i = 1 To lstFiles.ListCount
    If lstFiles.Selected(i - 1) = True Then TmpInt = TmpInt + 1
  Next i
  If TmpInt > 0 Then
    cmdExtract.Enabled = True
    cmdDelete.Enabled = True
  Else
    cmdExtract.Enabled = False
    cmdDelete.Enabled = False
  End If
End Sub

Private Sub lstFiles_DblClick()
  cmdExtract_Click
End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    mnuExtract.Enabled = cmdExtract.Enabled
    mnuDelete.Enabled = cmdDelete.Enabled
    mnuSelectAll.Enabled = CBool(lstFiles.ListCount)
    PopupMenu mnuPop
  End If
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show vbModal
End Sub

Function CheckFileCase(ByVal File As String, ByVal PathFile As String) As String
  Dim TmpStr As String

  If File <> "" And FileExist(PathFile + File) = False Then
    MsgBox "File has been deleted or renamed", vbCritical
    flsFiles.Refresh
    Exit Function
  End If
  TmpStr = Dir(PathFile + "*.*", vbNormal + vbHidden + vbSystem) ' Retrieve the first entry.
  Do While TmpStr <> ""
    If TmpStr <> "." And TmpStr <> ".." Then
      If Not (GetAttr(PathFile + File) And vbDirectory) = vbDirectory Then
        If LCase(TmpStr) = LCase(File) Then
          CheckFileCase = TmpStr
          Exit Function
        End If
      End If
    End If
    TmpStr = Dir
  Loop
  If File <> "" Then
    MsgBox "File not found", vbCritical
    End
  End If
End Function

Private Sub mnuDelete_Click()
  cmdDelete_Click
End Sub

Private Sub mnuExtract_Click()
  cmdExtract_Click
End Sub


Private Sub mnuSelectAll_Click()
  Dim i As Integer

  For i = 1 To lstFiles.ListCount
    lstFiles.Selected(i - 1) = True
  Next i
End Sub


Private Sub mnuusing_Click()
frmhelp.Show
End Sub
