VERSION 5.00
Begin VB.Form Rebound 
   BackColor       =   &H00000000&
   Caption         =   "Rebote II - 1"
   ClientHeight    =   5205
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   6930
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFC0&
   Icon            =   "Rebote 2.5.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tt 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   600
      Top             =   4680
   End
   Begin VB.Timer Tt2 
      Interval        =   1
      Left            =   120
      Top             =   4680
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFF80&
      BorderWidth     =   2
      Index           =   11
      Visible         =   0   'False
      X1              =   2040
      X2              =   3600
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFF80&
      BorderWidth     =   2
      Index           =   10
      Visible         =   0   'False
      X1              =   2040
      X2              =   3600
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFF80&
      BorderWidth     =   2
      Index           =   9
      Visible         =   0   'False
      X1              =   2040
      X2              =   3600
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFF80&
      BorderWidth     =   2
      Index           =   8
      Visible         =   0   'False
      X1              =   2040
      X2              =   3600
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   7
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   6
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   5
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   4
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   2
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   1
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line L 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      Visible         =   0   'False
      X1              =   360
      X2              =   1920
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Menu menu 
      Caption         =   "Rebound"
      Begin VB.Menu Ended 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Sty 
      Caption         =   "Estilo"
      Begin VB.Menu Rbt 
         Caption         =   "1"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu Rbt 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu Rbt 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu Rbt 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu Rbt 
         Caption         =   "5"
         Index           =   5
      End
      Begin VB.Menu Rbt 
         Caption         =   "6"
         Index           =   6
      End
      Begin VB.Menu Rbt 
         Caption         =   "7"
         Index           =   7
      End
      Begin VB.Menu Rbt 
         Caption         =   "8"
         Index           =   8
      End
   End
   Begin VB.Menu Hlpmenu 
      Caption         =   "Ayuda"
      Begin VB.Menu Hlp 
         Caption         =   "Acerca de..."
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Rebound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Che, Ct, Cv, Xm, Ym, Xf, Yf As Integer
Dim Df, Ran, D, V, V2, Vv, Kk, S(1 To 4), Stl, M, G, Xt, Yt As Integer
Dim Xx, Yy, Cx, Cy, Xxa(1 To 8), Yya(1 To 8)  As Integer
Dim Col(0 To 6), Colo(0 To 3) As ColorConstants
Const Pix = 3.1415927 / 180
'Cv = Contador Personalizado
'Ct = Contador
'Xx, Cx = X
'Yy, Cy = Y
'M = Indica hacia donde gira el reloj
'Xt = Indica si X aumenta o Decrese
'Yt = Indica si Y aumenta o Decrese
'V = Tamaño
'D = Velocidad
'Stl = Estilo

Private Sub Ended_Click()
End
End Sub

Private Sub Form_Activate()
Che = 1
Rbt_Click (1)
Orga

Col(0) = vbYellow
Col(1) = &HFF8080
Col(2) = &H80C0FF
Col(3) = &H80FF80
Col(4) = &HC0FFFF
Col(5) = &H8080FF
Col(6) = vbWhite

Colo(0) = &HFF8080
Colo(1) = &HFFFFC0
Colo(2) = &HE0E0E0
Colo(3) = &HFFC0C0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii > 48) And ((KeyAscii < 57)) Then
Rbt_Click (Val(Chr(KeyAscii)))
End If

If (KeyAscii = 104) Then Hlp_Click
End Sub

Private Sub Form_Resize()
Orga
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Hlp_Click()
Help.Show
End Sub

Private Sub Rbt_Click(Index As Integer)
Rbt(Che).Checked = 0
Rbt(Index).Checked = 1
Che = Index

Select Case Che
Case 1:
'Gusano
L(0).BorderWidth = 3
L(0).Visible = 1
For Cv = 1 To 11
L(Cv).Visible = 0
Next

Me.Caption = "Rebound II - " + "1"
Tt.Enabled = 1
Tt2.Enabled = 0
Stl = 1
Ct = 1
M = 1
G = 3
D = 100
V = 400
Vv = 0

Case 2:
'Gaviota
For Cv = 0 To 1
L(Cv).BorderWidth = 3
L(Cv).Visible = 1
Next
For Cv = 2 To 11
L(Cv).Visible = 0
Next

Me.Caption = "Rebound II - " + "2"
Tt.Enabled = 0
Tt2.Enabled = 1
Stl = 2
Ct = 1
M = 1
G = 3
D = 100
V = 400
Vv = 0
Cx = 500
Cy = 500

Case 3:
'Triangulo
For Cv = 0 To 2
L(Cv).BorderWidth = 2
L(Cv).Visible = 1
Next
For Cv = 3 To 11
L(Cv).Visible = 0
Next

Me.Caption = "Rebound II - " + "3"
Tt.Enabled = 0
Tt2.Enabled = 1
Stl = 6
Ct = 1
M = 1
G = 3
D = 100
V = 800
Vv = 0
Cx = 1000
Cy = 1000
S(1) = 20
S(2) = 40

Case 4:
'Cuadrado
For Cv = 0 To 3
L(Cv).BorderWidth = 2
L(Cv).Visible = 1
Next
For Cv = 4 To 11
L(Cv).Visible = 0
Next

Me.Caption = "Rebound II - " + "4"
Tt.Enabled = 0
Tt2.Enabled = 1
Stl = 7
Ct = 1
M = 1
G = 4
D = 100
V = 800
Vv = 0
Cx = 1000
Cy = 1000
S(1) = 15
S(2) = 30
S(3) = 45

Case 5:
'Cubo
For Cv = 0 To 11
L(Cv).BorderWidth = 2
L(Cv).Visible = 1
Next

Me.Caption = "Rebound II - " + "5"
Tt.Enabled = 0
Tt2.Enabled = 1
Stl = 3
Ct = 1
M = 1
G = 4
Ran = 1
D = 100
V = 800
V2 = V
Vv = 500

Cx = 1000
Cy = 1000

S(1) = 15
S(2) = 30
S(3) = 45
Case 6:
'Estrella de 5 puntas
For Cv = 0 To 4
L(Cv).BorderWidth = 2
L(Cv).Visible = 1
Next
For Cv = 5 To 11
L(Cv).Visible = 0
Next

Me.Caption = "Rebound II - " + "6"
Tt.Enabled = 0
Tt2.Enabled = 1
Stl = 4
Ct = 1
M = 1
G = 5
Ran = 1
D = 100
V = 800
V2 = V
Vv = 0
Cx = 1000
Cy = 1000
Randomize
Aleat
S(1) = 12
S(2) = 24
S(3) = 36
S(4) = 48

Case 7:
'Pentagono
For Cv = 0 To 4
L(Cv).BorderWidth = 2
L(Cv).Visible = 1
Next
For Cv = 5 To 11
L(Cv).Visible = 0
Next

Me.Caption = "Rebound II - " + "7"
Tt.Enabled = 0
Tt2.Enabled = 1
Stl = 8
Ct = 1
M = 1
G = 5
D = 100
V = 800
V2 = V
Vv = 0
Cx = 1000
Cy = 1000
S(1) = 12
S(2) = 24
S(3) = 36
S(4) = 48

Case 8:
'Estrella de 4 puntas
For Cv = 0 To 7
L(Cv).BorderWidth = 2
L(Cv).Visible = 1
Next
For Cv = 8 To 11
L(Cv).Visible = 0
Next

Me.Caption = "Rebound II - " + "8"
Tt.Enabled = 0
Tt2.Enabled = 1
Stl = 5
Ct = 1
M = 1
G = 4
D = 100
V = 1000
V2 = V / 2
Vv = 0
Cx = 1000
Cy = 1000
S(1) = 15
S(2) = 30
S(3) = 45
S(4) = 53
End Select
End Sub

Sub Aleat()
If (Stl = 4) Or (Stl = 3) Then
Df = Ran

Ini:
If (Stl = 4) Then
Ran = (6 * Rnd)
Else:
Ran = (3 * Rnd)
End If
If (Df = Ran) Then GoTo Ini

End If
End Sub

Sub Orga()
Xf = Me.Width - 200
Yf = Me.Height - 750

If (Stl > 2) Then
Cx = 2000
Cy = 2000
Else:
Cx = 200
Cy = 200
End If

Xx = 0
Yy = 0
End Sub

Private Sub Tt2_Timer()
If (Stl = 3) Then

If (Kk = 1) Then
Vv = Vv + (20)
If (Vv > 500) Then Vv = 500
Else:
Vv = Vv - (20)
If (Vv < -500) Then Vv = -500
End If

End If

Ct = Ct + M
If (Ct > 60) Then Ct = 1
If (Ct < 0) Then Ct = 59
    
    Xx = Xm + (Sin(Ct * 6 * Pix)) * 1
    Yy = Ym - (Cos(Ct * 6 * Pix)) * 1
    
    '*12*
    Xxa(1) = Xx + (Sin(Ct * 6 * Pix)) * V
    Yya(1) = Yy - (Cos(Ct * 6 * Pix)) * V
    
    If (Stl = 2) Then
    Xxa(2) = Xx + (Sin(-Ct * 6 * Pix)) * V
    Yya(2) = Yy - (Cos(-Ct * 6 * Pix)) * V
    
    Xxa(3) = 0
    Yya(3) = 0
    End If

    If (Stl <> 2) Then
    '*3*
    Xxa(2) = Xx + (Sin((Ct + S(1)) * 6 * Pix)) * V
    Yya(2) = Yy - (Cos((Ct + S(1)) * 6 * Pix)) * V

    '*6*
    Xxa(3) = Xx + (Sin((Ct + S(2)) * 6 * Pix)) * V
    Yya(3) = Yy - (Cos((Ct + S(2)) * 6 * Pix)) * V
        
    '*9*
    Xxa(4) = Xx + (Sin((Ct + S(3)) * 6 * Pix)) * V
    Yya(4) = Yy - (Cos((Ct + S(3)) * 6 * Pix)) * V
        
    '*10* y *11*
    Xxa(5) = Xx + (Sin((Ct + S(4)) * 6 * Pix)) * (V2)
    Yya(5) = Yy - (Cos((Ct + S(4)) * 6 * Pix)) * (V2)

    '*7* y *8*
    Xxa(6) = Xx + (Sin((Ct + (38)) * 6 * Pix)) * (V2)
    Yya(6) = Yy - (Cos((Ct + (38)) * 6 * Pix)) * (V2)

    '*4* y *5*
    Xxa(7) = Xx + (Sin((Ct + (23)) * 6 * Pix)) * (V2)
    Yya(7) = Yy - (Cos((Ct + (23)) * 6 * Pix)) * (V2)
    
    '*1* y *2*
    Xxa(8) = Xx + (Sin((Ct + (8)) * 6 * Pix)) * (V2)
    Yya(8) = Yy - (Cos((Ct + (8)) * 6 * Pix)) * (V2)
    End If
    
'----------------------------------------------
If (Xt = 0) Then
Cx = Cx + D
End If

If (Yt = 0) Then
Cy = Cy + D
End If

If (Xt = 1) Then
Cx = Cx - D
End If

If (Yt = 1) Then
Cy = Cy - D
End If

For Cv = 1 To G

If ((Vv >= 0) And (Xxa(Cv) + Cx + Vv > Xf)) Or (Xxa(Cv) + Cx > Xf) Then
Xt = 1
M = M * -1
Kk = 1
Aleat
End If

If ((Vv >= 0) And (Yya(Cv) + Cy + Vv > Yf)) Or (Yya(Cv) + Cy > Yf) Then
Yt = 1
Kk = 1
Aleat
End If

If ((Vv <= 0) And (Xxa(Cv) + Cx + Vv < 0)) Or (Xxa(Cv) + Cx < 0) Then
Xt = 0
M = M * -1
Kk = 0
Aleat
End If

If ((Vv <= 0) And (Yya(Cv) + Cy + Vv < 0)) Or (Yya(Cv) + Cy < 0) Then
Yt = 0
Kk = 0
Aleat
End If

Next

'----------------------------------------------
If (Stl >= 6) Then
L(0).BorderColor = &HFF00&
L(1).BorderColor = &HFF8080
L(2).BorderColor = &H80FF&
L(3).BorderColor = &HFFFF&
L(4).BorderColor = &HFF80FF
L(5).BorderColor = &H80FF&
L(6).BorderColor = &HFF8080
L(7).BorderColor = &HFF00&
End If

If (Stl = 3) Then
For Cv = 0 To 11
L(Cv).BorderColor = Colo(Ran)
Next
End If

If (Stl = 4) Then
For Cv = 0 To 4
L(Cv).BorderColor = Col(Ran)
Next
End If

If (Stl = 6) Then
L(2).X1 = Xxa(1) + Cx
L(2).X2 = Xxa(3) + Cx
L(2).Y1 = Yya(1) + Cy
L(2).Y2 = Yya(3) + Cy
End If

If (Stl >= 6) Or (Stl = 3) Then
L(0).X1 = Xxa(1) + Cx
L(0).X2 = Xxa(2) + Cx
L(0).Y1 = Yya(1) + Cy
L(0).Y2 = Yya(2) + Cy

L(1).X1 = Xxa(2) + Cx
L(1).X2 = Xxa(3) + Cx
L(1).Y1 = Yya(2) + Cy
L(1).Y2 = Yya(3) + Cy
End If

If (Stl = 7) Or (Stl = 3) Then
L(3).X1 = Xxa(4) + Cx
L(3).X2 = Xxa(1) + Cx
L(3).Y1 = Yya(4) + Cy
L(3).Y2 = Yya(1) + Cy
End If

If (Stl >= 7) Or (Stl = 3) Then
L(2).X1 = Xxa(3) + Cx
L(2).X2 = Xxa(4) + Cx
L(2).Y1 = Yya(3) + Cy
L(2).Y2 = Yya(4) + Cy
End If

If (Stl = 3) Then
L(4).X1 = Xxa(1) + Cx + Vv
L(4).X2 = Xxa(2) + Cx + Vv
L(4).Y1 = Yya(1) + Cy + Vv
L(4).Y2 = Yya(2) + Cy + Vv

L(5).X1 = Xxa(2) + Cx + Vv
L(5).X2 = Xxa(3) + Cx + Vv
L(5).Y1 = Yya(2) + Cy + Vv
L(5).Y2 = Yya(3) + Cy + Vv

L(6).X1 = Xxa(3) + Cx + Vv
L(6).X2 = Xxa(4) + Cx + Vv
L(6).Y1 = Yya(3) + Cy + Vv
L(6).Y2 = Yya(4) + Cy + Vv

L(7).X1 = Xxa(4) + Cx + Vv
L(7).X2 = Xxa(1) + Cx + Vv
L(7).Y1 = Yya(4) + Cy + Vv
L(7).Y2 = Yya(1) + Cy + Vv

L(8).X1 = Xxa(1) + Cx
L(8).X2 = Xxa(1) + Cx + Vv
L(8).Y1 = Yya(1) + Cy
L(8).Y2 = Yya(1) + Cy + Vv

L(9).X1 = Xxa(2) + Cx
L(9).X2 = Xxa(2) + Cx + Vv
L(9).Y1 = Yya(2) + Cy
L(9).Y2 = Yya(2) + Cy + Vv

L(10).X1 = Xxa(3) + Cx
L(10).X2 = Xxa(3) + Cx + Vv
L(10).Y1 = Yya(3) + Cy
L(10).Y2 = Yya(3) + Cy + Vv

L(11).X1 = Xxa(4) + Cx
L(11).X2 = Xxa(4) + Cx + Vv
L(11).Y1 = Yya(4) + Cy
L(11).Y2 = Yya(4) + Cy + Vv
End If

If (Stl = 8) Then
L(4).X1 = Xxa(5) + Cx
L(4).X2 = Xxa(1) + Cx
L(4).Y1 = Yya(5) + Cy
L(4).Y2 = Yya(1) + Cy
End If

If (Stl >= 8) Then
L(3).X1 = Xxa(4) + Cx
L(3).X2 = Xxa(5) + Cx
L(3).Y1 = Yya(4) + Cy
L(3).Y2 = Yya(5) + Cy
End If

'--------------------------
If (Stl = 2) Then
L(0).X1 = Xxa(1) + Cx
L(0).X2 = Cx
L(0).Y1 = Yya(1) + Cy
L(0).Y2 = Cy
L(0).BorderColor = &HFF&

L(1).X1 = Xxa(2) + Cx
L(1).X2 = Cx
L(1).Y1 = Yya(2) + Cy
L(1).Y2 = Cy
L(1).BorderColor = &HFF0000
End If

If (Stl = 4) Then
L(0).X1 = Xxa(1) + Cx
L(0).X2 = Xxa(4) + Cx
L(0).Y1 = Yya(1) + Cy
L(0).Y2 = Yya(4) + Cy

L(1).X1 = Xxa(4) + Cx
L(1).X2 = Xxa(2) + Cx
L(1).Y1 = Yya(4) + Cy
L(1).Y2 = Yya(2) + Cy

L(2).X1 = Xxa(2) + Cx
L(2).X2 = Xxa(5) + Cx
L(2).Y1 = Yya(2) + Cy
L(2).Y2 = Yya(5) + Cy

L(3).X1 = Xxa(5) + Cx
L(3).X2 = Xxa(3) + Cx
L(3).Y1 = Yya(5) + Cy
L(3).Y2 = Yya(3) + Cy

L(4).X1 = Xxa(3) + Cx
L(4).X2 = Xxa(1) + Cx
L(4).Y1 = Yya(3) + Cy
L(4).Y2 = Yya(1) + Cy
End If

If (Stl = 5) Then
L(0).X1 = Xxa(1) + Cx
L(0).X2 = Xxa(5) + Cx
L(0).Y1 = Yya(1) + Cy
L(0).Y2 = Yya(5) + Cy
L(0).BorderColor = &HFF00&

L(1).X1 = Xxa(5) + Cx
L(1).X2 = Xxa(4) + Cx
L(1).Y1 = Yya(5) + Cy
L(1).Y2 = Yya(4) + Cy
L(1).BorderColor = &HFF8080

L(2).X1 = Xxa(4) + Cx
L(2).X2 = Xxa(6) + Cx
L(2).Y1 = Yya(4) + Cy
L(2).Y2 = Yya(6) + Cy
L(2).BorderColor = &HFF8080

L(3).X1 = Xxa(6) + Cx
L(3).X2 = Xxa(3) + Cx
L(3).Y1 = Yya(6) + Cy
L(3).Y2 = Yya(3) + Cy
L(3).BorderColor = &H80FF&

L(4).X1 = Xxa(3) + Cx
L(4).X2 = Xxa(7) + Cx
L(4).Y1 = Yya(3) + Cy
L(4).Y2 = Yya(7) + Cy
L(4).BorderColor = &H80FF&

L(5).X1 = Xxa(7) + Cx
L(5).X2 = Xxa(2) + Cx
L(5).Y1 = Yya(7) + Cy
L(5).Y2 = Yya(2) + Cy
L(5).BorderColor = &HFFFF&

L(6).X1 = Xxa(2) + Cx
L(6).X2 = Xxa(8) + Cx
L(6).Y1 = Yya(2) + Cy
L(6).Y2 = Yya(8) + Cy
L(6).BorderColor = &HFFFF&

L(7).X1 = Xxa(8) + Cx
L(7).X2 = Xxa(1) + Cx
L(7).Y1 = Yya(8) + Cy
L(7).Y2 = Yya(1) + Cy
L(7).BorderColor = &HFF00&
End If
End Sub

Private Sub Tt_Timer()
If (Xt = 0) Then
Xx = Xx + D
Xa = Xx - V
End If

If (Yt = 0) Then
Yy = Yy + D
Ya = Yy - V
End If

If (Xt = 1) Then
Xx = Xx - D
Xa = Xx + V
End If

If (Yt = 1) Then
Yy = Yy - D
Ya = Yy + V
End If

If (Xx > Xf) Then
Xt = 1
End If

If (Yy > Yf) Then
Yt = 1
End If

If (Xx < 100) Then
Xt = 0
End If

If (Yy < 100) Then
Yt = 0
End If

L(0).X1 = Xa
L(0).X2 = Xx
L(0).Y1 = Ya
L(0).Y2 = Yy
L(0).BorderColor = &HC000C0
End Sub
