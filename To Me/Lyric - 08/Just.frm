VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form LyricS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscador de Liricas"
   ClientHeight    =   4740
   ClientLeft      =   4305
   ClientTop       =   1740
   ClientWidth     =   7005
   Icon            =   "Just.frx":0000
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   495
      Begin VB.Label Ll 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   495
      End
   End
   Begin RichTextLib.RichTextBox Ly 
      Height          =   6375
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11245
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"Just.frx":6852
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox Dd2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.ListBox D2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.DirListBox Dd 
      Height          =   315
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.FileListBox D 
      Height          =   480
      Left            =   3600
      Pattern         =   "*.txt"
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.PictureBox Pp 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   240
      Picture         =   "Just.frx":68CE
      ScaleHeight     =   4335
      ScaleWidth      =   6495
      TabIndex        =   7
      Top             =   240
      Width           =   6495
   End
   Begin VB.Menu SeFL 
      Caption         =   "Archivo"
      Begin VB.Menu Sv 
         Caption         =   "Guardar"
         Shortcut        =   ^G
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "LyricS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ct, Jj, Kk, Err As Integer
Dim Text, Cad, Rut As String

Private Sub Form_Load()
Kk = 0
Ly.SelAlignment = rtfCenter
ShowFl
ShowFl2
Dd.Path = CurDir
D.Path = CurDir
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If (Kk = 1) Then
  Jj = MsgBox("¿Desea Guardar Los Cambios?", vbYesNoCancel, App.Title)
  If (Jj = vbYes) Then Sv_Click
End If

If (Jj = vbCancel) Then
Cancel = 1
End If

End Sub

Private Sub D2_Click()
D.Path = Dd.List(D2.ListIndex)
Text = Dd.Path
ShowFl
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hlp_Click()
Help.Show
End Sub

Private Sub Ll_Click()
Fr_Click
End Sub

Private Sub Fr_Click()
If (Kk = 1) Then
  Jj = MsgBox("¿Desea Guardar Los Cambios?", vbYesNoCancel, App.Title)
  If (Jj = vbYes) Then Sv_Click
End If

If (Jj = vbCancel) Then
GoTo Fin
End If

Kk = 0
Me.Height = 5460
Pp.Height = 289
Fr.Visible = False
Ly.Visible = False
D2.Visible = True
Dd2.Visible = True

Fin:
End Sub

Private Sub Dd2_DblClick()
Cad = D.Path + "\" + D.List(Dd2.ListIndex)

If (Cad <> "") Then
  On Error GoTo Hell
  Ly.LoadFile Cad, 1
  Me.Height = 7875
  Pp.Height = 457
  Fr.Visible = True
  Ly.Visible = True
  D2.Visible = False
  Dd2.Visible = False
End If

If (2 = 3) Then
Hell:
Kk = 0
Fr_Click
End If

End Sub

'------------------- Funciones -------------------
Sub AbrirArchivo(Ruta As String, Texto As TextBox)
Dim fnum As Integer
On Error GoTo Ninguno
fnum = FreeFile
Open Ruta For Input As fnum
Do While Not EOF(fnum)
Line Input #fnum, txt
Texto.Text = Texto.Text & vbCrLf & txt
Loop
Close fnum
Ninguno:
End Sub

Private Sub ShowFl()
Dd2.Clear
If (D.ListCount <> 0) Then
  For Ct = 0 To D.ListCount
    Cad = Replace(D.List(Ct), ".txt", "")
    Dd2.AddItem Cad, Ct
  Next
End If
End Sub

Private Sub ShowFl2()
D2.Clear
If (Dd.ListCount <> 0) Then
  For Ct = 0 To Dd.ListCount
    Cad = Replace(Dd.List(Ct), CurDir, "")
    D2.AddItem Cad, Ct
  Next
End If
End Sub

Private Sub Ly_Change()
Kk = 1
End Sub

Private Sub Salir_Click()
Unload Me
End Sub

Private Sub Sv_Click()
Ly.SaveFile Cad, 1
Kk = 0
End Sub
