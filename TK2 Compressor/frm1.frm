VERSION 5.00
Begin VB.Form frm1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1050
   ClientLeft      =   5430
   ClientTop       =   3495
   ClientWidth     =   2100
   Icon            =   "frm1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1050
   ScaleWidth      =   2100
   Begin VB.CommandButton Command1 
      Caption         =   "ok"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Please put your file name"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo al
If txt1.Text = "" Then txt1.Text = "File1"
namafile = "0-" & txt1.Text
frm1.Hide
frmMain.lblfn = namafile
If txt1.Text = "" Then
End If
txt1.Text = ""
Exit Sub
al:
End Sub

