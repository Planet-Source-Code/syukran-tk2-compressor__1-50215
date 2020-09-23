VERSION 5.00
Begin VB.Form frmExtract 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Select folder"
   ClientHeight    =   3465
   ClientLeft      =   1545
   ClientTop       =   1860
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3465
   ScaleWidth      =   5250
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   2415
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   1935
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Extract file"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Cancel extract file"
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New folder..."
         Height          =   375
         Left            =   240
         TabIndex        =   5
         ToolTipText     =   "Create new folder"
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
   Begin VB.DirListBox dirFolders 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.DriveListBox drvDrives 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblExtract 
      BackColor       =   &H00FFC0C0&
      Caption         =   "File(s) will be extracted to:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmExtract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldDrive As String

Private Sub cmdCancel_Click()
  PathToExtract = ""
  Unload Me
End Sub


Private Sub cmdNew_Click()
  Dim TmpStr As String

  TmpStr = InputBox$("Enter the new folder name")
  If TmpStr <> "" Then
    If Right$(txtPath.Text, 1) <> "\" Then
      TmpStr = txtPath.Text + "\" + TmpStr
    Else
      TmpStr = txtPath.Text + TmpStr
    End If
    MkDir TmpStr
    dirFolders.Refresh
  End If
End Sub

Private Sub cmdOk_Click()
  On Error GoTo Check
  If GetAttr(txtPath.Text) = vbDirectory Then:
  PathToExtract = txtPath.Text
  Unload Me
  Exit Sub
Check:
  If Err.Number = 53 Then
    MkDir txtPath.Text
    Resume Next
  End If
End Sub

Private Sub dirFolders_Change()
  txtPath.Text = dirFolders.Path
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


Private Sub Form_Load()
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
  OldDrive = drvDrives.Drive
  dirFolders.Path = App.Path
  dirFolders_Change
End Sub


