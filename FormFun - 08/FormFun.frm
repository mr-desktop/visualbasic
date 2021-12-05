VERSION 5.00
Begin VB.Form FormFun 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manipulador"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3975
   Icon            =   "FormFun.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Tex 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   6
      Text            =   "5"
      ToolTipText     =   "Segundos"
      Top             =   2520
      Width           =   975
   End
   Begin VB.Timer Tt 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label D2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Apagar Monitor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label D1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bloquear"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label C1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mostrar Botón de Windows"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label C2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ocultar Botón de Windows"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label A2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Capturar Ventana Seleccionada"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label A1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Capturar Pantalla Completa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label B2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cerrar Unidad Óptica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label B1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Abrir Unidad Óptica"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FormFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
  
'Api SendMessage
Private Declare Function SendMessage _
Lib "user32" _
Alias "SendMessageA" ( _
ByVal hwnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Any) As Long
  
'Constantes para SendMessage
Const WM_SYSCOMMAND = &H112&
Const SC_MONITORPOWER = &HF170&

Dim Ct As Integer

Private Sub A1_Click()
Pant (0)
End Sub

Private Sub A2_Click()
Pant (1)
End Sub

Private Sub B2_Click()
Call mciSendString("Set CDAudio Door Closed Wait", 0&, 0&, 0&)
End Sub

Private Sub B1_Click()
Call mciSendString("Set CDAudio Door Open Wait", 0&, 0&, 0&)
End Sub

Private Sub C2_Click()
WinB (0)
End Sub

Private Sub C1_Click()
WinB (1)
End Sub

Private Sub D1_Click()
Ct = Val(Tex.Text)
Tt.Enabled = True
BlockInput True
End Sub

Private Sub D2_Click()
SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal 2&
End Sub

Private Sub Tex_Change()
Tex.Text = Val(Tex.Text)
End Sub

Private Sub Tt_Timer()
Ct = Ct - 1

If (Ct <= 0) Then
BlockInput False
Tt.Enabled = False
Ct = 0
End If

Tex.Text = Str(Ct)

End Sub
