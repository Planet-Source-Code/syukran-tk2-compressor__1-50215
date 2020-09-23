VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About TK2 Compressor"
   ClientHeight    =   5250
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4365
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5250
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFC0C0&
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   4095
      Begin VB.Label lblWarning 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ecingtoto,Viperkid,Selferino,Dark Archos,Hanagata, Samwise,f_nola,Tasmanian,Maresmagians,PSC,VB lovers and anyone who supports me"
         ForeColor       =   &H00000000&
         Height          =   585
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Thanks to :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   4095
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Made in Malaysia [Nov,20,2003]"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   4095
      Begin VB.Label lblWarning 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Warning: This program is copyrighted. Its unauthorized reproduction may result in criminal charges."
         ForeColor       =   &H00000000&
         Height          =   465
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4095
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Syukran algorithm file compression. Copyright © 2003 by Syukran Hakim B Norazman. All rights reserved."
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3765
      End
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   360
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   510
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "TK2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   120
      Width           =   645
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   645
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPRESSOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   2925
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOk_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Me.Top = (Screen.Height / 2) - (Me.Height / 2)
  Me.Left = (Screen.Width / 2) - (Me.Width / 2)
End Sub
