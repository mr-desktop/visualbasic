VERSION 5.00
Begin VB.Form Rebound 
   BackColor       =   &H00000000&
   Caption         =   "Rebound - 1"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6315
   ForeColor       =   &H0000FF00&
   Icon            =   "Rebound.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tt 
      Interval        =   1
      Left            =   240
      Top             =   360
   End
   Begin VB.Menu Rbt 
      Caption         =   "Estilo"
      Begin VB.Menu Rbt1 
         Caption         =   "1"
      End
      Begin VB.Menu Rbt2 
         Caption         =   "2"
      End
      Begin VB.Menu Rbt3 
         Caption         =   "3"
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "Rebound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dd, Xf, Yf, Xx, Yy, Vl, Kk, Rt As Integer
Dim Xt, Yt, Tr, Jk As Boolean

Private Sub Form_DblClick()
Kk = 0

If (Tr = True) And (Kk = 0) Then
Me.WindowState = 0
Kk = 1
Tr = False
End If

If (Tr = False) And (Kk = 0) Then
Kk = 1
Me.WindowState = 2
Tr = True
End If

End Sub


Private Sub Form_Load()
Rt = 1
Yf = Me.Height - 400
Xf = Me.Width - 150
Xt = True
Yt = True

Vl = 40
End Sub

Private Sub Form_Resize()
Yf = Me.Height - 400
Xf = Me.Width - 150
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hlp_Click()
Help.Show
End Sub

Private Sub Rbt1_Click()
Me.Caption = "Rebound - " + "1"
ForeColor = &HFF00&
Cls
Rt = 1
End Sub

Private Sub Rbt2_Click()
Me.Caption = "Rebound - " + "2"
ForeColor = &H80FF&
Cls
Rt = 2
End Sub

Private Sub Rbt3_Click()
Me.Caption = "Rebound - " + "3"
ForeColor = &HFF0000
Cls
Rt = 3
End Sub

Private Sub Tt_Timer()
If (Rt = 2) Then Cls

If (Rt = 3) Then
If (Dd > 100) Then Jk = True
If (Dd < 1) Then Jk = False
If (Jk = False) Then Dd = Dd + 5
If (Jk = True) Then Dd = Dd - 5
End If

'Incremento en X'
If (Xt = True) Then
Xx = Xx + Vl
End If
'Final de Incremento en X'
If (Xt = True) And ((Xx + 1150) >= Xf) Then

If (Rt = 1) Then Cls
Xt = False
End If

'Decremento en X'
If (Xt = False) Then
Xx = Xx - Vl
End If
'Final Decremento en X'
If (Xt = False) And (Xx <= 0) Then
Cls
Xt = True
End If

'Incremento en Y'
If (Yt = True) Then
Yy = Yy + Vl
End If
'Final de Incremento en Y'
If (Yt = True) And ((Yy + 1550) >= Yf) Then
Cls
Yt = False
End If

'Decremento en Y'
If (Yt = False) Then
Yy = Yy - Vl
End If
'Final de Decremento en Y'
If (Yt = False) And (Yy <= 0) Then
Cls
Yt = True
End If

If (Rt = 3) Then Circle (Xx + 600, Yy + 600), (600), &HFF8080, 0, 0, (Dd / 100)
Circle (Xx + 600, Yy + 600), (600)
End Sub
