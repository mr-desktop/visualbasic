VERSION 5.00
Begin VB.Form Kuku 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Prew 
      Height          =   1560
      Left            =   120
      Picture         =   "Kuku.frx":0000
      Top             =   120
      Width           =   2085
   End
End
Attribute VB_Name = "Kuku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Ac = True
'Prew.Height = Me.Height
'Prew.Width = Me.Width
End Sub

Private Sub Form_LostFocus()
If (Acti = False) Then
Ac = 0
End
End If
End Sub

Private Sub Form_Resize()
'Prew.Height = Me.Height
'Prew.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
If (Acti = False) Then
Ac = 0
End
End If
End Sub
