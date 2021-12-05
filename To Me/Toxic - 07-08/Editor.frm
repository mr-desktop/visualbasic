VERSION 5.00
Begin VB.Form Editor 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mini - Terminal"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4140
   ClipControls    =   0   'False
   Icon            =   "Editor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Editor.frx":000C
   ScaleHeight     =   4095
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Cmd 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   345
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   120
      Picture         =   "Editor.frx":CFB1
      Top             =   360
      Width           =   3840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "->"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Editor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (UCase(Cmd.Text) = UCase("cmb")) Then ChPasw.Show
If (KeyAscii = 13) And (UCase(Cmd.Text) = UCase("clv")) Then Psw.Show
If (KeyAscii = 13) And (UCase(Cmd.Text) = UCase("image")) Then Direct.Show
If (KeyAscii = 13) And (UCase(Cmd.Text) = UCase("color")) Then Coloro.Show
If (KeyAscii = 13) And (UCase(Cmd.Text) = UCase("ver")) Then Jeus.Show
If (KeyAscii = 13) And (UCase(Cmd.Text) = UCase("?")) Then Help.Show
If (KeyAscii = 13) And (UCase(Cmd.Text) = UCase("full")) Then
If (Fondo.Full = True) And (Fondo.Tag = "0") Then
Fondo.Full = False
Fondo.Tag = "1"
End If
If (Fondo.Full = False) And (Fondo.Tag = "0") Then
Fondo.Full = True
Fondo.Tag = "1"
End If

Fondo.Tag = "0"
Fondo.FullScr
End If

If (KeyAscii = 13) Or (KeyAscii = 27) Then
Editor.Hide
End If
End Sub

Private Sub Image1_Click()
Editor.Hide
End Sub
