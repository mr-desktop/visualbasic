VERSION 5.00
Begin VB.Form Psw 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clave"
   ClientHeight    =   4320
   ClientLeft      =   12045
   ClientTop       =   10905
   ClientWidth     =   3750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Psw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Psw.frx":000C
   ScaleHeight     =   4320
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "x"
      TabIndex        =   0
      ToolTipText     =   "Inserte Su Contraseña"
      Top             =   240
      Width           =   2535
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "Psw.frx":CFB1
      Top             =   600
      Width           =   3840
   End
End
Attribute VB_Name = "Psw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then FinPsw
End Sub

Private Sub Image1_Click()
If (UCase(Pass.Text) = UCase(Fondo.Passito.Text)) Then
SavFl "Info.log", Fondo.Passito.Text
End
End If
Psw.Hide
End Sub

Private Sub Pass_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) And (UCase(Pass.Text) = UCase(Fondo.Passito.Text)) Then End
If (KeyAscii = 27) Then FinPsw
End Sub

Private Function FinPsw()
Psw.Hide
Pass.Text = ""
End Function
