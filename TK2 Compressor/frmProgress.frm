VERSION 5.00
Begin VB.Form frmProgress 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please, wait...                                       -compressing file-"
   ClientHeight    =   1095
   ClientLeft      =   1530
   ClientTop       =   1755
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   73
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   Begin VB.PictureBox picProgreso 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   351
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      Begin VB.Label lblAction 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   30
         TabIndex        =   3
         Top             =   120
         Width           =   75
      End
      Begin VB.Label lblPorc 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   2280
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
      Begin VB.Shape shaProgreso 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.Label lblWait 
      BackColor       =   &H00FFC0C0&
      Caption         =   "The requested operation is being completed..."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
End Sub
