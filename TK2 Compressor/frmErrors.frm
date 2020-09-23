VERSION 5.00
Begin VB.Form frmErrors 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Errors"
   ClientHeight    =   3120
   ClientLeft      =   1530
   ClientTop       =   1845
   ClientWidth     =   6990
   Icon            =   "frmErrors.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3120
   ScaleWidth      =   6990
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Copy error report"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      ToolTipText     =   "Quit to main menu"
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtErrors 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frmErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText txtErrores.Text
End Sub

Private Sub Form_Load()
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
End Sub


