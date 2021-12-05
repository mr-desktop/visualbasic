VERSION 5.00
Begin VB.Form Calc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   1920
   ControlBox      =   0   'False
   Icon            =   "Calculo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Calculo.frx":6852
   ScaleHeight     =   1680
   ScaleWidth      =   1920
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   600
      TabIndex        =   0
      ToolTipText     =   "Promedio de la Nota"
      Top             =   120
      Width           =   735
   End
   Begin VB.Label O 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      ToolTipText     =   "Otra Total=0%; Examen=100%"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label E 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Extraordinario Total=30%; Examen=70%"
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label C 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Completivos Total=50%; Examen=50%"
      Top             =   600
      Width           =   735
   End
   Begin VB.Label G 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Generales Total=70%; Examen=30%"
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "Calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer

Private Sub Tt_Change()
If Not (Tt.Text = "") And (Tt.Text > "") Then A = Val(Tt.Text)
If (Tt.Text = "") Or (Val(Tt.Text) > 100) Then
Tt.Text = ""
A = 0
End If
End Sub

Private Sub Tt_KeyPress(KeyAscii As Integer)

If (KeyAscii = 13) Then
G.Caption = Int((-(A * 0.7) + 70) / 0.3)
C.Caption = Int((-(A * 0.5) + 70) / 0.5)
E.Caption = Int((-(A * 0.3) + 70) / 0.7)
O.Caption = Int((-(A * 0) + 70) / 1)
End If

If (KeyAscii = 27) Then End

If (KeyAscii = 8) Then
G.Caption = ""
C.Caption = ""
E.Caption = ""
O.Caption = ""
Tt.Text = ""
A = 0
End If

End Sub
