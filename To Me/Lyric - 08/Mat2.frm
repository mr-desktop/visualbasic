VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form LyricS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Artistas"
   ClientHeight    =   8040
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   6735
   StartUpPosition =   1  'CenterOwner
   Begin VB.DirListBox Dd 
      Height          =   2115
      Left            =   3600
      TabIndex        =   21
      Top             =   4920
      Width           =   2895
   End
   Begin VB.FileListBox D 
      Height          =   2235
      Left            =   480
      TabIndex        =   20
      Top             =   4920
      Width           =   2895
   End
   Begin MSComDlg.CommonDialog CDial 
      Left            =   120
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Bk 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Back"
      Height          =   375
      Left            =   4080
      TabIndex        =   23
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Nt 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Next"
      Height          =   375
      Left            =   5280
      TabIndex        =   22
      Top             =   3840
      Width           =   975
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   5160
      TabIndex        =   18
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   19
      Left            =   5160
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   5160
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   18
      Left            =   5160
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   17
      Left            =   5160
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   5160
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   16
      Left            =   5160
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   5160
      TabIndex        =   15
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   15
      Left            =   5160
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   5160
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   3480
      TabIndex        =   13
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   14
      Left            =   3480
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   3480
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   13
      Left            =   3480
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   12
      Left            =   3480
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   3480
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   11
      Left            =   3480
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   10
      Left            =   3480
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1800
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   9
      Left            =   1800
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1800
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   8
      Left            =   1800
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   1800
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   7
      Left            =   1800
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   6
      Left            =   1800
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   5
      Left            =   1800
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   4
      Left            =   120
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   3
      Left            =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      Height          =   495
      Index           =   2
      Left            =   120
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Shape Brd 
      BorderColor     =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label FLesito 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu SeFL 
      Caption         =   "Archivo"
      Begin VB.Menu OpenFL 
         Caption         =   "Abrir ..."
         Shortcut        =   ^A
      End
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "LyricS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ct, Dt, Up, Stt, StF, Max As Integer

Private Sub Form_Load()
Dt = 1
Up = 0
Stt = 0
ShowFl
D.Path = CurDir
End Sub

Private Sub FLesito_Click(Index As Integer)
Brd(Index).BorderColor = &HFF0000
End Sub

Private Sub FLesito_DblClick(Index As Integer)
Brd(Index).BorderColor = &HFF0000

Dt = Dt + 1
If Dt > 1 Then Dt = 0

D.Path = Dd.List(Dd.ListCount - 1)
ShowFl
End Sub

Private Sub ShowFl()

Select Case Dt

Case 0
StF = 0
Max = D.ListCount
While (Max > 20)
Max = Max - 20
StF = StF + 1
Wend

If (D.ListCount <> 0) Then
  For Ct = 0 To (D.ListCount - 1)
    FLesito(Ct).Visible = True
    Brd(Ct).Visible = True
    FLesito(Ct).Caption = D.List(Ct + Up)
    Brd(Ct).BorderColor = &O0
  Next
End If

If (D.ListCount <> 20) Then
  For Ct = (D.ListCount) To 19
    FLesito(Ct).Visible = False
    Brd(Ct).Visible = False
  Next
End If

Case 1
StF = 0
Max = Dd.ListCount
While (Max > 20)
Max = Max - 20
StF = StF + 1
Wend

If (Dd.ListCount <> 0) Then
  For Ct = 0 To (Dd.ListCount - 1)
    FLesito(Ct).Visible = True
    Brd(Ct).Visible = True
    FLesito(Ct).Caption = Dd.List(Ct + Up)
    Brd(Ct).BorderColor = &O0
  Next
End If

If (Dd.ListCount <> 20) Then
  For Ct = (Dd.ListCount) To 19
    FLesito(Ct).Visible = False
    Brd(Ct).Visible = False
  Next
End If

End Select

End Sub

Private Sub Nt_Click()
If (Stt > 0) Then
Up = Up + 19
Stt = Stt + 1
End If
End Sub

Private Sub Bk_Click()
If (Stt > 1) Then
Up = Up + 19
Stt = Stt - 1
End If
End Sub

'------------------- Menu -------------------
Private Sub OpenFL_Click()

With CDial
.Filter = "Archivos de Texto|*.txt|Todos los Archivos|*.*"
.DialogTitle = "Seleccione un Archivo"
.InitDir = CurDir
.Flags = cdlOFNHideReadOnly
.ShowOpen
End With

If CDial.FileName = "" Then
  L.Caption = "No se ha seleccionado ningún archivo"
Else
  L.Caption = CDial.FileName
End If

L2.Caption = CDial.FileTitle
End Sub


Private Sub Salir_Click()
End
End Sub
