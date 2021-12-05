VERSION 5.00
Begin VB.Form Fondo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15360
   ClipControls    =   0   'False
   Icon            =   "Fondo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox Passito 
      DataSource      =   "Alpha"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Back 
      Height          =   9000
      Left            =   0
      Picture         =   "Fondo.frx":08CA
      Top             =   0
      Width           =   12000
   End
End
Attribute VB_Name = "Fondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Full As Boolean
Public Ww, Hh As Integer

Private Sub Back_DblClick()
Editor.Show
End Sub

Private Sub Form_DblClick()
Editor.Show
End Sub

Public Sub Form_Load()
Imoa
AbrTxt "Info.log", Fondo.Passito
Full = False
Fondo.Tag = "0"
FullScr
Passito.Text = UCase(Passito.Text)
Back.Width = Fondo.Width
Back.Height = Fondo.Height
End Sub

Private Sub Form_Resize()
FullScr
End Sub

Public Function FullScr()
If (Fondo.Full = False) Then
Fondo.Back.Stretch = False
Fondo.Back.Left = (Fondo.Width / 2) - (Ww / 2)
Fondo.Back.Top = (Fondo.Height / 2) - (Hh / 2)
Fondo.Back.Width = Ww
Fondo.Back.Height = Hh
End If

If (Fondo.Full = True) Then
Fondo.Back.Left = 0
Fondo.Back.Top = 0
Fondo.Back.Stretch = True
Fondo.Back.Width = Fondo.Width
Fondo.Back.Height = Fondo.Height
End If

End Function

Public Function Imoa()
Fondo.Ww = Fondo.Back.Width
Fondo.Hh = Fondo.Back.Height
End Function
