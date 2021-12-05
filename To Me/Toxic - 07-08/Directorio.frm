VERSION 5.00
Begin VB.Form Direct 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fondo"
   ClientHeight    =   3285
   ClientLeft      =   3420
   ClientTop       =   900
   ClientWidth     =   8490
   Icon            =   "Directorio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Directorio.frx":000C
   ScaleHeight     =   3285
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox F2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   1065
      Left            =   120
      Pattern         =   "*.jpeg;*.jpg;*.bmp;*.gif;*.dib;*.jpe;*.jfif;*.tiff;*.tif"
      TabIndex        =   2
      Top             =   1920
      Width           =   3735
   End
   Begin VB.DirListBox F 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   990
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.DriveListBox D 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   3840
      Left            =   4440
      Picture         =   "Directorio.frx":CFB1
      Top             =   -720
      Width           =   3840
   End
End
Attribute VB_Name = "Direct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public File As String

Private Sub D_Change()
F.Path = D.Drive
End Sub

Private Sub F_Change()
F2.Path = F.Path
End Sub

Private Sub F2_Click()
File = F2.Path + "\" + F2.FileName
Fondo.Back.Picture = LoadPicture(File)
Fondo.Imoa
Fondo.FullScr
End Sub

Private Sub F2_DblClick()
Direct.Hide
End Sub

Private Sub F2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Or (KeyAscii = 27) Then Direct.Hide
End Sub

Private Sub Image2_Click()
Direct.Hide
End Sub
