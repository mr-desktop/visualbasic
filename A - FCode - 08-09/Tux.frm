VERSION 5.00
Begin VB.Form Tox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Falling Code"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000080FF&
   Icon            =   "Tux.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Tux.frx":23D2
   ScaleHeight     =   7110
   ScaleWidth      =   10770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   9600
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      Begin VB.Label Coco1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Nivel"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Coco1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Puntuación"
         Top             =   6120
         Width           =   975
      End
      Begin VB.Label Coco1 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Vidas Perdidas"
         Top             =   6600
         Width           =   975
      End
      Begin VB.Label Coco1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Restantes para Pasar a otro Nivel"
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Coco1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Vidas"
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Timer Tt 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Top             =   4200
   End
   Begin VB.Label Rd 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Estás Listo?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   3000
      TabIndex        =   6
      Top             =   4440
      Width           =   4215
   End
   Begin VB.Image Im 
      BorderStyle     =   1  'Fixed Single
      Height          =   3660
      Left            =   2640
      Picture         =   "Tux.frx":71FA
      Top             =   1800
      Width           =   4860
   End
   Begin VB.Menu Sty 
      Caption         =   "Juego"
      Begin VB.Menu New 
         Caption         =   "Nuevo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Pausa 
         Caption         =   "Pausa"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Ended 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Est 
      Caption         =   "Estilo"
      Begin VB.Menu St 
         Caption         =   "Númerico"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu St 
         Caption         =   "Letras"
         Index           =   2
      End
      Begin VB.Menu St 
         Caption         =   "Alfanumérico"
         Index           =   3
      End
   End
   Begin VB.Menu Hlpm 
      Caption         =   "Ayuda"
      Begin VB.Menu Hlp 
         Caption         =   "Acerca de Fcode"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "Tox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Che, C, D(1 To 3), M2, Mm, Mm2, Nv, Stl, F, Xx, Yy, Ps As Integer
Dim M, Lt As Integer
Dim N, Ed As Boolean
Dim K, E(0 To 1) As String

'<-------------------- Variables -------------------->
'Lt = vidas
'Nv = nivel
'D(1) = vidas perdidas
'D(2) = restantes para pasar a otro nivel
'D(3) = puntuación

'Mm = indica cuando pasan 10 niveles
'Mm2 = indica si el modulo "Org" puede funcionar
'Ed = indica si el juego finalizo
'Ps = indica si el juego esta pausado

'M = velocidad
'Stl = estilo
'Che = indica la última modalidad de juego seleccionada
'K = tecla presionada
'C = contador

'F = número aleatorio que indica si caen números o letras
'E(0) = número aleatorio que cae
'E(1) = letra aleatoria que cae
'N = indica si el caracter tiene que volver arriba y cambiar

'Xx && Yy = posiciones por la cual se mueven los caracteres

'<-------------------- Modulo -------------------->
' <- Funciones y Subs->
'Inic: comienza el juego
'Inicio(Al): posiciona por donde comienza a caer el caracter
'CurDir: indica el directorio actual donde se ubica el juego
'Pas: pausa el juego
'Org: indica los colores y las imagenes de fondo de cada nivel
'Lost: modulo que se activa cuando pierdes el juego
'WinGame: modulo que se activa cuando ganas el juego

Private Sub Form_Activate()
Randomize
End Sub

Private Sub Form_GotFocus()
If (Ed = False) Then Pas
End Sub

Private Sub Form_Load()
Che = 1
St_Click (1)
End Sub

Private Sub Form_LostFocus()
If (Ed = False) Then Pas
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
    If (Im.Visible = True) Then
    Im_Click
    Else:
    Pas
    End If
End If

K = UCase(Chr(KeyAscii))
If (Ed = False) And (Ps = 0) Then

If (K = E(F)) Then
C = 0
D(2) = D(2) - 1
D(3) = D(3) + 1
Coco1(2).Caption = Str(D(2))
Coco1(4).Caption = Str(D(3))
If (D(2) <= 0) Then
Nv = Nv + 1
Coco1(5).Caption = Str(Nv)

M = M + 1
Mm = Mm + 1
If (Mm >= 10) Then
Mm2 = 1
M2 = M2 + 1
End If
Org

If (M >= 80) Then
    Ed = True
    Wingame
End If

Lt = Lt + 5
Coco1(3).Caption = Str(Lt)
D(2) = 10
End If
N = True
End If

If (K <> E(F)) Then
  If KeyAscii <> 13 Then
  If KeyAscii <> 27 Then
D(1) = D(1) + 1
Coco1(0).Caption = Str(D(1))
Lt = Lt - 1

If (Lt <= 0) Then
Ed = True
Lost
End If

Coco1(3).Caption = Str(Lt)
  End If
  End If
End If

End If
End Sub

'<----------------- R U Ready? -------------------->
Private Sub Rd_Click()
Im.Visible = False
Ps = 0
If (Ed = False) Then
Tt.Enabled = True
Me.Caption = "<< Falling Code >>"
End If
Rd.Visible = False
End Sub

'<----------------- Img Aguacate -------------------->
Private Sub Im_Click()
Rd_Click
End Sub

'<----------------- Menu -------------------->
Private Sub New_Click()
Inic
End Sub

Private Sub Pausa_Click()
Pas
End Sub

Private Sub Ended_Click()
End
End Sub

Private Sub Est_Click()
If (Ed = False) Then Pas
End Sub

Private Sub Sty_Click()
If (Ed = False) Then Pas
End Sub

Private Sub Hlpm_Click()
If (Ed = False) Then Pas
End Sub

Private Sub Hlp_Click()
If (Ed = False) Then Pas
Help.Show
End Sub

'<----------------- SubMenu Estilo -------------------->
Private Sub St_Click(Index As Integer)
St(Che).Checked = 0
St(Index).Checked = 1
Che = Index

If (Che = 3) Then
Stl = 2
Else:
Stl = 1
End If

If (Che = 2) Then F = 0
If (Che = 1) Then F = 1

Inic
End Sub

'<----------------- Contador -------------------->
Private Sub Tt_Timer()
If (Ed = False) And (Ps = 0) Then

Inicio (N)
C = C + Int(M) * 20

CurrentX = Xx
CurrentY = C
Print E(F)

If (C >= Me.Height - 700) Then
C = 0
D(1) = D(1) + 1
Coco1(0).Caption = Str(D(1))
Lt = Lt - 1
If (Lt <= 0) Then Lost
Coco1(3).Caption = Str(Lt)
N = True
End If

End If
End Sub

'<----------------- Modulo -------------------->
Sub Inic()
Mm2 = 1
M2 = 1
Org
Ed = False
Pas
Nv = 1
Lt = 10
M = 10
Mm = 0

Inicio (True)

C = 0

D(1) = 0
D(2) = 10
D(3) = 0
Lt = 10
Coco1(0).Caption = Str(D(1))
Coco1(2).Caption = Str(D(2))
Coco1(3).Caption = Str(Lt)
Coco1(4).Caption = Str(D(3))
Coco1(5).Caption = Str(Nv)
End Sub

Function Inicio(Al As Boolean)
If (Al = True) Then
C = 0
E(0) = Chr(Int(25 * Rnd) + 65)
E(1) = Chr(Int(9 * Rnd) + 48)
If (Stl = 2) Then F = Int(2 * Rnd)

Xx = Int((Me.Width - 1700) * Rnd)
N = False
Cls
End If
End Function

Function CurDir() As String
Dim Directorio As String
ChDir App.Path
ChDrive App.Path
Directorio = App.Path
If Len(Directorio) > 3 Then
Directorio = Directorio & "\"
End If
CurDir = Directorio
End Function

Sub Pas()
If (Ed = 0) Then
If (Ps = 1) Then GoTo Ef
Ps = 1
Me.Caption = "<< Pausa >>"
Tt.Enabled = False
Cls
Rd.Visible = True
Im.Visible = True
Else:
Im.Visible = True
End If

Ef:
End Sub

Sub Org()
On Error GoTo Nin
If (Mm2 = 1) Then
Me.Picture = LoadPicture(CurDir + "Org\" + Trim(Str(M2)) + ".jeus")
Im.Picture = LoadPicture(CurDir + "Org\Mod\" + Trim(Str(M2)) + ".coco")
Pas

Select Case M2
Case 1:
ForeColor = &H80FF&
Case 2:
ForeColor = &O0
Case 3:
ForeColor = &HFFFFC0
Case 4:
ForeColor = &O0
Case 5:
ForeColor = &HFFFFFF
Case 6:
ForeColor = &HFFFF&
Case 7:
ForeColor = &HFFFFFF
End Select

Mm = 0
Mm2 = 0
End If

If (2 = 3) Then
Nin:
Rd.ForeColor = vbWhite
Me.ForeColor = &H80FF&
End If

End Sub

Sub Lost()
Tt.Enabled = False
Ed = True
Cls

On Error GoTo Hell2
Me.Picture = LoadPicture(CurDir + "Org\Win\Lost.banana")
Im.Picture = LoadPicture(CurDir + "Org\Mod\Lost2.coco")
Hell2:
Im.Visible = True
Me.Caption = "<< Perdiste >>"
End Sub

Sub Wingame()
Tt.Enabled = False
Ed = True
Cls

On Error GoTo Hell
Me.Picture = LoadPicture(CurDir + "Org\Win\" + Trim(Str(Che)) + ".banana")
Im.Picture = LoadPicture(CurDir + "Org\Mod\Win.coco")
Hell:
Im.Visible = True
Me.Caption = "<< Ganaste >>"
End Sub

