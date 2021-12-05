VERSION 5.00
Begin VB.Form Rebound 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Rebote II"
   ClientHeight    =   5235
   ClientLeft      =   150
   ClientTop       =   180
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
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Tit 
      Interval        =   1000
      Left            =   600
      Top             =   4680
   End
   Begin VB.Timer Tt2 
      Interval        =   1
      Left            =   120
      Top             =   4680
   End
End
Attribute VB_Name = "Rebound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ct, Cv, Xm, Ym, Xf, Yf, Tux As Integer
Dim Df, Ran, D, V, V2, S(1 To 4), Stl, M, G, Xt, Yt As Integer
Dim Xx, Yy, cx, cy, Xxa(1 To 8), Yya(1 To 8)  As Integer
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

Private Sub Form_Activate()
Orga
Me.WindowState = 2

Col(0) = vbYellow
Col(1) = &HFF8080
Col(2) = &H80C0FF
Col(3) = &H80FF80
Col(4) = &HC0FFFF
Col(5) = &H8080FF
Col(6) = vbWhite

Rebote (Sle)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (Tux = 1) Then Ended
End Sub

Private Sub Form_LostFocus()
If (Tux = 1) Then Ended
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If (Tux = 1) Then Ended
End Sub

Private Sub Form_Resize()
Orga
End Sub

Public Sub Ended()
ShowCursor True
End
End Sub

Public Sub Rebote(Index As Integer)

Select Case Index
Case 1:
Me.DrawWidth = 3
Cls
Stl = 1
Ct = 1
M = 1
G = 2
D = 100
V = 400
cx = 100
cy = 100

Case 2:
Me.DrawWidth = 3
Cls
Stl = 2
Ct = 1
M = 1
G = 3
D = 100
V = 400
cx = 500
cy = 500

Case 3:
Me.DrawWidth = 1
Cls
Stl = 3
Ct = 1
M = 1
G = 3
D = 100
V = 800
cx = 1000
cy = 1000
S(1) = 20
S(2) = 40

Case 4:
Me.DrawWidth = 1
Cls
Stl = 4
Ct = 1
M = 1
G = 4
D = 100
V = 800
cx = 1000
cy = 1000
S(1) = 15
S(2) = 30
S(3) = 45

Case 5:
Me.DrawWidth = 1
Cls
Stl = 5
Ct = 1
M = 1
G = 5
D = 100
cx = 1000
cy = 1000
Randomize
Aleat
S(1) = 12
S(2) = 24
S(3) = 36
S(4) = 48

Case 6:
Me.DrawWidth = 1
Cls
Stl = 6
Ct = 1
M = 1
G = 5
D = 100
V = 800
V2 = V
cx = 1000
cy = 1000
S(1) = 12
S(2) = 24
S(3) = 36
S(4) = 48

Case 7:
Me.DrawWidth = 1
Cls
Stl = 7
Ct = 1
M = 1
G = 4
D = 100
V = 1000
V2 = V / 2
cx = 1000
cy = 1000
S(1) = 15
S(2) = 30
S(3) = 45
S(4) = 53
End Select

End Sub

Sub Aleat()
If (Stl = 5) Then
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

If (Stl > 2) Then
cx = 2000
cy = 2000
Else:
cx = 200
cy = 200
End If

Xx = 0
Yy = 0
Cls
End Sub

Private Sub Form_Unload(Cancel As Integer)
Ended
End Sub

Private Sub Tim_Timer()
Ui = 1
End Sub

Private Sub Tit_Timer()
Tux = 1
Tit.Enabled = False
End Sub

Private Sub Tt2_Timer()
If (Stl = 5) Then
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
    
    If (Stl = 1) Then
    Xxa(2) = 0
    Yya(2) = 0
    End If
    
    If (Stl = 2) Then
    Xxa(2) = Xx + (Sin(-Ct * 6 * Pix)) * V
    Yya(2) = Yy - (Cos(-Ct * 6 * Pix)) * V
    
    Xxa(3) = 0
    Yya(3) = 0
    End If

    If (Stl > 2) Then
    '*3*
    Xxa(2) = Xx + (Sin((Ct + S(1)) * 6 * Pix)) * V
    Yya(2) = Yy - (Cos((Ct + S(1)) * 6 * Pix)) * V

    '*6*
    Xxa(3) = Xx + (Sin((Ct + S(2)) * 6 * Pix)) * V
    Yya(3) = Yy - (Cos((Ct + S(2)) * 6 * Pix)) * V
    End If
    
    If (Stl > 3) Then
    '*9*
    Xxa(4) = Xx + (Sin((Ct + S(3)) * 6 * Pix)) * V
    Yya(4) = Yy - (Cos((Ct + S(3)) * 6 * Pix)) * V
    End If
    
    If (Stl >= 5) Then
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
cx = cx + D
End If

If (Yt = 0) Then
cy = cy + D
End If

If (Xt = 1) Then
cx = cx - D
End If

If (Yt = 1) Then
cy = cy - D
End If

For Cv = 1 To G

If (Xxa(Cv) + cx > Xf) Then
Xt = 1
M = M * -1
Aleat
End If

If (Yya(Cv) + cy > Yf) Then
Yt = 1
Aleat
End If

If (Xxa(Cv) + cx < 0) Then
Xt = 0
M = M * -1
Aleat
End If

If (Yya(Cv) + cy < 40) Then
Yt = 0
Aleat
End If

Next

'----------------------------------------------
If (Stl = 7) Then
Line (Xxa(1) + cx, Yya(1) + cy)-(Xxa(5) + cx, Yya(5) + cy), &HFF00&
Line -(Xxa(4) + cx, Yya(4) + cy), &HFF8080
Line -(Xxa(6) + cx, Yya(6) + cy), &HFF8080
Line -(Xxa(3) + cx, Yya(3) + cy), &H80FF&
Line -(Xxa(7) + cx, Yya(7) + cy), &H80FF&
Line -(Xxa(2) + cx, Yya(2) + cy), &HFFFF&
Line -(Xxa(8) + cx, Yya(8) + cy), &HFFFF&
Line -(Xxa(1) + cx, Yya(1) + cy), &HFF00&
End If

If (Stl = 6) Then
Line (Xxa(1) + cx, Yya(1) + cy)-(Xxa(2) + cx, Yya(2) + cy), &HFF00&
Line -(Xxa(3) + cx, Yya(3) + cy), &HFF8080
Line -(Xxa(4) + cx, Yya(4) + cy), &H80FF&
Line -(Xxa(5) + cx, Yya(5) + cy), &HFFFF&
Line -(Xxa(1) + cx, Yya(1) + cy), &HFF80FF
End If

If (Stl = 5) Then
Line (Xxa(1) + cx, Yya(1) + cy)-(Xxa(4) + cx, Yya(4) + cy), Col(Ran)
Line -(Xxa(2) + cx, Yya(2) + cy), Col(Ran)
Line -(Xxa(5) + cx, Yya(5) + cy), Col(Ran)
Line -(Xxa(3) + cx, Yya(3) + cy), Col(Ran)
Line -(Xxa(1) + cx, Yya(1) + cy), Col(Ran)
End If

If (Stl = 4) Then
Line (Xxa(1) + cx, Yya(1) + cy)-(Xxa(2) + cx, Yya(2) + cy), &HFF00&
Line -(Xxa(3) + cx, Yya(3) + cy), &HFF8080
Line -(Xxa(4) + cx, Yya(4) + cy), &H80FF&
Line -(Xxa(1) + cx, Yya(1) + cy), &HFFFF&
End If

If (Stl = 3) Then
Line (Xxa(1) + cx, Yya(1) + cy)-(Xxa(2) + cx, Yya(2) + cy), &HFF8080
Line -(Xxa(3) + cx, Yya(3) + cy), &H80FF&
Line -(Xxa(1) + cx, Yya(1) + cy), &HFFFF&
End If

If (Stl = 2) Then
Line (Xxa(1) + cx, Yya(1) + cy)-(cx, cy), &HFF&
Line -(Xxa(2) + cx, Yya(2) + cy), &HFF0000
End If

If (Stl = 1) Then
Line (Xxa(1) + cx, Yya(1) + cy)-(cx, cy), &H80FF&
End If

End Sub
