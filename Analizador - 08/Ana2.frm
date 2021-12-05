VERSION 5.00
Begin VB.Form Ana 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Código ASCII"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   ControlBox      =   0   'False
   Icon            =   "Ana2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   3255
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox T 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.ListBox L 
      Enabled         =   0   'False
      Height          =   3180
      IntegralHeight  =   0   'False
      ItemData        =   "Ana2.frx":6852
      Left            =   120
      List            =   "Ana2.frx":6854
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "Ana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A(1 To 100), Aa(1 To 100), P As String
Dim B, C As Integer

Private Sub T_DblClick()
L.Clear
L.Enabled = True
For C = 1 To 255
L.AddItem (Chr(C) + " -> " + Str(C))
Next
End Sub

Private Sub T_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
End
End If

If (KeyAscii = 8) Then
A(B) = ""
Aa(B) = ""
B = B - 1
End If

If ((KeyAscii <> 13) And (KeyAscii <> 8)) Then
If (KeyAscii <> 27) Then
B = B + 1
A(B) = Chr(KeyAscii)
Aa(B) = Str(KeyAscii)
End If
End If

If (KeyAscii = 13) Then
L.Clear
L.Enabled = True
If (T.Text = "") Then
End
End If
For C = 1 To B
L.List(C - 1) = A(C) + " -> " + Aa(C)
L.ListIndex = C - 1
Next
L.ListIndex = 0
End If

End Sub
