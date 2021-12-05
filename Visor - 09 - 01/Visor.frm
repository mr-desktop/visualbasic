VERSION 5.00
Begin VB.Form Visor 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visor"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9855
   Icon            =   "Visor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Ok 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK"
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox TxtCo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   5520
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.Label Bk 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label Nt 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   9120
      TabIndex        =   3
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label Comt 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   7560
      Width           =   6015
   End
   Begin VB.Image Imag 
      Height          =   7200
      Left            =   120
      Picture         =   "Visor.frx":23D2
      Stretch         =   -1  'True
      Top             =   240
      Width           =   9600
   End
End
Attribute VB_Name = "Visor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ct, Cmn, Cmx As Integer

'Ct: Contador
'Cmn: mínimo del contador (primera foto)
'Cmx: máximo del contador (última foto)

Private Sub Bk_Click()
Ct = Ct - 1
If (Ct < Cmn) Then Ct = Cmx
End Sub

Private Sub Comt_DblClick()
TxtCo.Text = ""
TxtCo.Visible = True
Ok.Visible = True
End Sub

Private Sub Form_Click()
Imag.Picture = LoadPicture(CurDir + "IMG (2).jpg")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Imag.Picture = LoadPicture(CurDir + "IMG (" + Len(Str(Ct)) + ").jpg")
End Sub

Private Sub Form_Load()
Ct = 2
Imag.Picture = LoadPicture(CurDir + "IMG (1).jpg")
End Sub

Private Sub Nt_Click()
Ct = Ct + 1
If (Ct > Cmx) Then Ct = Cmn
End Sub

Private Sub Ok_Click()
Comt.Caption = TxtCo.Text
TxtCo.Text = ""
TxtCo.Visible = False
Ok.Visible = False
End Sub

Private Sub TxtCo_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
TxtCo.Text = ""
TxtCo.Visible = False
Ok.Visible = False
End If
End Sub
