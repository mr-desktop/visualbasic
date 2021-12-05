VERSION 5.00
Begin VB.Form Rebound 
   BackColor       =   &H00000000&
   Caption         =   "Rebote II - 1"
   ClientHeight    =   5235
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
   ForeColor       =   &H00C0FFFF&
   Icon            =   "Rebote 2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tt2 
      Interval        =   1
      Left            =   120
      Top             =   4680
   End
   Begin VB.Menu Reb 
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
Dim Df, Ran, D, V, V2, S(1 To 4), Stl, M, G, Xt, Yt As Integer
Dim Xx, Yy, Cx, Cy, Xxa(1 To 8), Yya(1 To 8) As Integer '>> subir o bajar
Dim Col(0 To 6) As ColorConstants
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

Private Sub Form_Load()
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
End Sub

Private Sub Form_Activate()
Cls
End Sub

Private Sub Ended_Click()
End
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii > 48) And ((KeyAscii < 56)) Then '>> subir o bajar
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

Select Case Index
Case 1: '>> subir o bajar >=1
'Simple 1
Me.Caption = "Rebound II - " + "1" '>> subir o bajar >=1
Me.DrawWidth = 3
Cls
Stl = 1 '>> subir o bajar >=1
Ct = 1
M = 1
G = 2
D = 100
V = 400
Cx = 100
Cy = 100

Case 2: '>> subir o bajar >=2
'Simple 2
Me.Caption = "Rebound II - " + "2" '>> subir o bajar >=2
Me.DrawWidth = 3
Cls
Stl = 2 '>> subir o bajar >=2
Ct = 1
M = 1
G = 3
D = 100
V = 400
Cx = 500
Cy = 500

Case 3: '>> subir o bajar >=3
'Triangulo
Me.Caption = "Rebound II - " + "3" '>> subir o bajar >=3
Me.DrawWidth = 1
Cls
Stl = 3 '>> subir o bajar >=3
Ct = 1
M = 1
G = 3
D = 100
V = 800
Cx = 1000
Cy = 1000
S(1) = 20
S(2) = 40

Case 4: '>> subir o bajar >=4
'Cuadrado
Me.Caption = "Rebound II - " + "4" '>> subir o bajar >=4
Me.DrawWidth = 1
Cls
Stl = 4 '>> subir o bajar >=4
Ct = 1
M = 1
G = 4
D = 100
V = 800
Cx = 1000
Cy = 1000
S(1) = 15
S(2) = 30
S(3) = 45

'Case 5: '>> subir o bajar >=5
'Cuadrado 2
'Me.Caption = "Rebound II - " + "5" '>> subir o bajar >=5
'Me.DrawWidth = 1
'Cls
'Stl = 5 '>> subir o bajar >=5
'Ct = 1
'M = 1
'G = 4
'D = 100
'V = 800
'Cx = 1000
'Cy = 1000
'S(1) = 15
'S(2) = 30
'S(3) = 45

Case 5: '>> subir o bajar >=5
'Estrella
Me.Caption = "Rebound II - " + "5" '>> subir o bajar >=5
Me.DrawWidth = 1
Cls
Stl = 5 '>> subir o bajar >=5
Ct = 1
M = 1
G = 5
D = 100
Cx = 1000
Cy = 1000
Randomize
Aleat
S(1) = 12
S(2) = 24
S(3) = 36
S(4) = 48

Case 6: '>> subir o bajar >=6
'Pentagono
Me.Caption = "Rebound II - " + "6" '>> subir o bajar >=6
Me.DrawWidth = 1
Cls
Stl = 6 '>> subir o bajar >=6
Ct = 1
M = 1
G = 5
D = 100
V = 800
V2 = V
Cx = 1000
Cy = 1000
S(1) = 12
S(2) = 24
S(3) = 36
S(4) = 48

Case 7: '>> subir o bajar >=7
'Estrella 4 picos
Me.Caption = "Rebound II - " + "7" '>> subir o bajar >=7
Me.DrawWidth = 1
Cls
Stl = 7 '>> subir o bajar >=7
Ct = 1
M = 1
G = 4
D = 100
V = 1000
V2 = V / 2
Cx = 1000
Cy = 1000
S(1) = 15
S(2) = 30
S(3) = 45
S(4) = 53
End Select

End Sub

Sub Aleat()
If (Stl = 5) Then '>> subir o bajar >=5
Df = Ran

Ini:
Ran = (6 * Rnd)
If (Df = Ran) Then GoTo Ini

V = 400
V2 = V
End If

Cls
End Sub

Sub Orga()
Xf = Me.Width - 200
Yf = Me.Height - 750

If (Stl > 2) Then '>> subir o bajar >=2
Cx = 2000
Cy = 2000
Else:
Cx = 200
Cy = 200
End If

Xx = 0
Yy = 0
Cls
End Sub

Private Sub Tt2_Timer()
If (Stl = 5) Then '>> subir o bajar >=5
V = V + (20)
V2 = V
End If

Ct = Ct + M
If (Ct > 60) Then Ct = 1
If (Ct < 0) Then Ct = 59
    
    Xx = Xm + (Sin(Ct * 6 * Pix)) * 1
    Yy = Ym - (Cos(Ct * 6 * Pix)) * 1
    
    '*12*
    Xxa(1) = Xx + (Sin(Ct * 6 * Pix)) * V
    Yya(1) = Yy - (Cos(Ct * 6 * Pix)) * V
    
    If (Stl = 1) Then '>> subir o bajar >=1
    Xxa(2) = 0
    Yya(2) = 0
    End If
    
    If (Stl = 2) Then '>> subir o bajar >=2
    Xxa(2) = Xx + (Sin(-Ct * 6 * Pix)) * V
    Yya(2) = Yy - (Cos(-Ct * 6 * Pix)) * V
    
    Xxa(3) = 0
    Yya(3) = 0
    End If

    If (Stl > 2) Then '>> subir o bajar >=2
    '*3*
    Xxa(2) = Xx + (Sin((Ct + S(1)) * 6 * Pix)) * V
    Yya(2) = Yy - (Cos((Ct + S(1)) * 6 * Pix)) * V

    '*6*
    Xxa(3) = Xx + (Sin((Ct + S(2)) * 6 * Pix)) * V
    Yya(3) = Yy - (Cos((Ct + S(2)) * 6 * Pix)) * V
    End If
    
    If (Stl > 3) Then '>> subir o bajar >=3
    '*9*
    Xxa(4) = Xx + (Sin((Ct + S(3)) * 6 * Pix)) * V
    Yya(4) = Yy - (Cos((Ct + S(3)) * 6 * Pix)) * V
    End If
    
    If (Stl >= 5) Then '>> subir o bajar >=5
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

If (Xxa(Cv) + Cx > Xf) Then
Xt = 1
M = M * -1
Aleat
End If

If (Yya(Cv) + Cy > Yf) Then
Yt = 1
Aleat
End If

If (Xxa(Cv) + Cx < 0) Then
Xt = 0
M = M * -1
Aleat
End If

If (Yya(Cv) + Cy < 40) Then
Yt = 0
Aleat
End If

Next

'----------------------------------------------
If (Stl = 7) Then '>> subir o bajar >=7
Line (Xxa(1) + Cx, Yya(1) + Cy)-(Xxa(5) + Cx, Yya(5) + Cy), &HFF00&
Line -(Xxa(4) + Cx, Yya(4) + Cy), &HFF8080
Line -(Xxa(6) + Cx, Yya(6) + Cy), &HFF8080
Line -(Xxa(3) + Cx, Yya(3) + Cy), &H80FF&
Line -(Xxa(7) + Cx, Yya(7) + Cy), &H80FF&
Line -(Xxa(2) + Cx, Yya(2) + Cy), &HFFFF&
Line -(Xxa(8) + Cx, Yya(8) + Cy), &HFFFF&
Line -(Xxa(1) + Cx, Yya(1) + Cy), &HFF00&

'Line (Xxa(1) + Cx, Yya(1) + Cy)-(Cx, Cy), &HFF00&
'Line -(Xxa(2) + Cx, Yya(2) + Cy), &HFF8080
'Line (Xxa(3) + Cx, Yya(3) + Cy)-(Cx, Cy), &HFF8080
'Line -(Xxa(4) + Cx, Yya(4) + Cy), &H80FF&
'Line (Xxa(5) + Cx, Yya(5) + Cy)-(Cx, Cy), &H80FF&
'Line -(Xxa(6) + Cx, Yya(6) + Cy), &HFFFF&
'Line (Xxa(7) + Cx, Yya(7) + Cy)-(Cx, Cy), &HFFFF&
'Line -(Xxa(8) + Cx, Yya(8) + Cy), &HFF00&

End If

If (Stl = 6) Then '>> subir o bajar >=6
Line (Xxa(1) + Cx, Yya(1) + Cy)-(Xxa(2) + Cx, Yya(2) + Cy), &HFF00&
Line -(Xxa(3) + Cx, Yya(3) + Cy), &HFF8080
Line -(Xxa(4) + Cx, Yya(4) + Cy), &H80FF&
Line -(Xxa(5) + Cx, Yya(5) + Cy), &HFFFF&
Line -(Xxa(1) + Cx, Yya(1) + Cy), &HFF80FF
End If

If (Stl = 5) Then '>> subir o bajar >=5
Line (Xxa(1) + Cx, Yya(1) + Cy)-(Xxa(4) + Cx, Yya(4) + Cy), Col(Ran)
Line -(Xxa(2) + Cx, Yya(2) + Cy), Col(Ran)
Line -(Xxa(5) + Cx, Yya(5) + Cy), Col(Ran)
Line -(Xxa(3) + Cx, Yya(3) + Cy), Col(Ran)
Line -(Xxa(1) + Cx, Yya(1) + Cy), Col(Ran)
End If

'If (Stl = 5) Then '>> subir o bajar >=5
'Line (Xxa(1) + Cx, Yya(1) + Cy)-(Cx, Cy), &HFF00&
'Line -(Xxa(2) + Cx, Yya(2) + Cy), &HFF8080
'Line (Xxa(3) + Cx, Yya(3) + Cy)-(Cx, Cy), &H80FF&
'Line -(Xxa(4) + Cx, Yya(4) + Cy), &HFFFF&
'End If

If (Stl = 4) Then '>> subir o bajar >=4
Line (Xxa(1) + Cx, Yya(1) + Cy)-(Xxa(2) + Cx, Yya(2) + Cy), &HFF00&
Line -(Xxa(3) + Cx, Yya(3) + Cy), &HFF8080
Line -(Xxa(4) + Cx, Yya(4) + Cy), &H80FF&
Line -(Xxa(1) + Cx, Yya(1) + Cy), &HFFFF&
End If

If (Stl = 3) Then '>> subir o bajar >=3
'Line (Xxa(1) + Cx, Yya(1) + Cy)-(Cx, Cy), &HFF00&
'Line -(Xxa(2) + Cx, Yya(2) + Cy), &HFF8080
'Line (Xxa(3) + Cx, Yya(3) + Cy)-(Cx, Cy), &H80FF&

Line (Xxa(1) + Cx, Yya(1) + Cy)-(Xxa(2) + Cx, Yya(2) + Cy), &HFF8080
Line -(Xxa(3) + Cx, Yya(3) + Cy), &H80FF&
Line -(Xxa(1) + Cx, Yya(1) + Cy), &HFFFF&
End If

If (Stl = 2) Then '>> subir o bajar >=2
Line (Xxa(1) + Cx, Yya(1) + Cy)-(Cx, Cy), &HFF&
Line -(Xxa(2) + Cx, Yya(2) + Cy), &HFF0000
End If

If (Stl = 1) Then '>> subir o bajar >=1
Line (Xxa(1) + Cx, Yya(1) + Cy)-(Cx, Cy), &H80FF&
End If

End Sub
