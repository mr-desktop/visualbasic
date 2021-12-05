VERSION 5.00
Begin VB.Form Note 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NoteEditor"
   ClientHeight    =   5640
   ClientLeft      =   540
   ClientTop       =   1170
   ClientWidth     =   6390
   Icon            =   "Note.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Note.frx":6852
   ScaleHeight     =   5640
   ScaleWidth      =   6390
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Pp2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      ToolTipText     =   "Palabra(s) Sustituta(s)"
      Top             =   600
      Width           =   6135
   End
   Begin VB.TextBox Pp1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Palabra a Encontrar"
      Top             =   120
      Width           =   6135
   End
   Begin VB.Frame Controles 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label J 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jumbo"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Lg 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Largo"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label A 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ancho"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label N 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image Dt 
         BorderStyle     =   1  'Fixed Single
         Height          =   660
         Left            =   1200
         Picture         =   "Note.frx":15B45
         ToolTipText     =   "Encontrar y Sustituir"
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   660
      End
      Begin VB.Label C 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1200
         TabIndex        =   5
         Top             =   120
         Width           =   660
      End
      Begin VB.Label R 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   660
      End
   End
   Begin VB.TextBox Txt 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4400
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Texto"
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Menu hlp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "Note"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Kk As Integer

'* * * Busqueda y Remplazo * * *
Private Sub Dt_Click()
Txt.Text = Replace(Txt.Text, Pp1.Text, Pp2.Text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hlp_Click()
Help.Show
End Sub

'* * * Posiciones * * *
Private Sub N_Click()
Pp1.Width = 6135
Pp2.Width = 6135
Txt.Height = 4400
Txt.Width = 6135
Note.Width = 6480
Note.Height = 6060
End Sub

Private Sub A_Click()
Pp1.Width = 9975
Pp2.Width = 9975
Txt.Height = 4400
Txt.Width = 9975
Note.Width = 10300
Note.Height = 6060
End Sub

Private Sub Lg_Click()
Pp1.Width = 6135
Pp2.Width = 6135
Txt.Height = 7695
Txt.Width = 6135
Note.Width = 6480
Note.Height = 9375
End Sub

Private Sub J_Click()
Pp1.Width = 9975
Pp2.Width = 9975
Txt.Height = 7695
Txt.Width = 9975
Note.Width = 10300
Note.Height = 9375
End Sub

'* * * Alineamiento * * *
Private Sub L_Click()
Txt.Alignment = 0
End Sub

Private Sub C_Click()
Txt.Alignment = 2
End Sub

Private Sub R_Click()
Txt.Alignment = 1
End Sub

'* * * Pulsaciones * * *
Private Sub Pp1_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then Dt_Click
End Sub

Private Sub Pp2_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then Dt_Click
End Sub

Private Sub Txt_Click()
Kk = Kk + 1
If (Kk > 1) Then Kk = 0
Controles.Visible = Kk
End Sub
