VERSION 5.00
Begin VB.Form Agnda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agenda"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2910
   Icon            =   "Agenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Agenda.frx":6852
   ScaleHeight     =   2535
   ScaleWidth      =   2910
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox A 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Nombre"
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox A 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Trabajo"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox A 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Celular"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox A 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Telefono"
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Ase 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Ase 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   6
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Ase 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      ForeColor       =   &H000000C0&
      Height          =   375
      Index           =   1
      Left            =   2400
      TabIndex        =   5
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Add 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agregar"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
End
Attribute VB_Name = "Agnda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Tex, Aa(1 To 12) As String
Dim C, CC(1 To 3) As Integer

Private Sub Ase_Click(Index As Integer)
CC(Index) = CC(Index) + 1
If (CC(Index) > 1) Then CC(Index) = 0

If (CC(Index) = 1) Then A(Index).Enabled = False
If (CC(Index) = 0) Then A(Index).Enabled = True
End Sub

Private Sub Form_Load()
Limp
End Sub

Private Sub Add_Click()
Dim fnum As Integer
If (A(0) = "") Then
MsgBox "Agrega Un Nombre"
GoTo Hell
End If
Tex = CurDir + "\Cont" + "\" + A(0).Text + ".txt"
Asig

On Error GoTo Ninguno
    fnum = FreeFile
    Open Tex For Output As fnum
        For C = 1 To 3
        Print #fnum, Aa(C)
        Next
    Close fnum
Ninguno:

Limp
MsgBox "El Contacto Ha Sido Agregado"
Hell:
End Sub

Function Limp()

For C = 0 To 3
A(C).Text = ""
Next

For C = 1 To 3
CC(C) = 0
Next

For C = 1 To 3
A(C).Enabled = False
Next

For C = 1 To 3
A(C).Enabled = True
Next

End Function

Function Asig()
If (CC(1) = 0) Then
Aa(1) = "Tel: " + A(1).Text
End If

If (CC(2) = 0) Then
Aa(2) = "Cel: " + A(2).Text
End If

If (CC(3) = 0) Then
Aa(3) = "E-Mail: " + A(3).Text
End If

End Function
