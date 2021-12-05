VERSION 5.00
Begin VB.Form ChPasw 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar Clave"
   ClientHeight    =   4710
   ClientLeft      =   4695
   ClientTop       =   6180
   ClientWidth     =   3720
   ClipControls    =   0   'False
   Icon            =   "ChPasw.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   Picture         =   "ChPasw.frx":000C
   ScaleHeight     =   4710
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox pass2 
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
      ForeColor       =   &H000080FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MouseIcon       =   "ChPasw.frx":CFB1
      PasswordChar    =   "x"
      TabIndex        =   1
      ToolTipText     =   "Nueva Clave"
      Top             =   600
      Width           =   1695
   End
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
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MousePointer    =   99  'Custom
      PasswordChar    =   "x"
      TabIndex        =   0
      ToolTipText     =   "Clave Anterior"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   0
      Picture         =   "ChPasw.frx":D87B
      Top             =   1080
      Width           =   3840
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Aterior ->"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva ->"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "ChPasw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gamma As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then FinChPasw
End Sub

Private Sub Image1_Click()
If (Gamma = True) Then SvChPasw
FinChPasw
End Sub

Private Sub Label3_Click()
FinChPasw
End Sub

Private Sub Pass_Change()
If (UCase(Pass.Text) = Fondo.Passito.Text) Then
Gamma = True
End If
End Sub

Private Sub Pass_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then FinChPasw
End Sub

Private Sub Pass2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then FinChPasw
End Sub

Private Function FinChPasw()
ChPasw.Hide
Pass.Text = ""
pass2.Text = ""
Gamma = False
Pass.Enabled = True
End Function

Private Function SvChPasw()
Fondo.Passito.Text = UCase(pass2.Text)
SavFl "Info.log", Fondo.Passito.Text
End Function

