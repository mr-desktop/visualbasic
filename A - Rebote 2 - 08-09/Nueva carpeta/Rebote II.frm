VERSION 5.00
Begin VB.Form Rebound 
   BackColor       =   &H00000000&
   Caption         =   "Rebote II - 1"
   ClientHeight    =   5235
   ClientLeft      =   165
   ClientTop       =   495
   ClientWidth     =   6930
   DrawWidth       =   3
   ForeColor       =   &H0000FF00&
   Icon            =   "Rebote II.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   6930
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tt2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1680
      Top             =   2400
   End
   Begin VB.Timer Tt3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   2400
   End
   Begin VB.Timer T2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   2400
   End
   Begin VB.Timer Tt 
      Interval        =   10
      Left            =   720
      Top             =   2400
   End
   Begin VB.Label L2 
      Caption         =   "Label1"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Ll 
      Caption         =   "Label1"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Menu Sty 
      Caption         =   "Estilo"
      Begin VB.Menu Rbt1 
         Caption         =   "1"
      End
      Begin VB.Menu Rbt2M 
         Caption         =   "2"
         Begin VB.Menu Rbt4 
            Caption         =   "1"
         End
         Begin VB.Menu Rbt2 
            Caption         =   "2"
         End
      End
      Begin VB.Menu Rbt3 
         Caption         =   "3"
      End
      Begin VB.Menu Rbt5M 
         Caption         =   "4"
         Begin VB.Menu Rbt8 
            Caption         =   "1"
         End
         Begin VB.Menu Rbt5 
            Caption         =   "2"
         End
         Begin VB.Menu Rbt6 
            Caption         =   "3"
         End
      End
      Begin VB.Menu Rbt9 
         Caption         =   "5"
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
Option Explicit
  
'   * Colocar un control Timer
'-------------------------------------------------
  
'Estructura de coordenadas para el api GetCursorPos
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Dim Mouse As POINTAPI

Dim M, Cv, Ct, Ct2, Cx, Cy, Xx, Xm, Yy, Ym, Xa, Xa2, Xxa(1 To 8), Ya, Ya2, Yya(1 To 8), Xf, Yf, V, D, Xt, Yt, Stl As Integer
Const Pix = 3.1415927 / 180

'Cv = Contador Personalizado
'Ct = Contador
'Ct2 = Contador Inverso
'Xx, Cx = X
'Yy, Cy = Y
'Xa = X Anterior / X2
'Xa2 = X2 Inversa
'Ya = Y Anterior / Y2
'Ya2 = Y2 Inversa
'Xt = Indica si X aumenta o Decrese
'Yt = Indica si Y aumenta o Decrese
'V = Tamaño
'D = Velocidad
'Stl = Estilo

Private Sub Form_Activate()
Rbt1_Click
End Sub

Private Sub Form_Resize()
Xf = Me.Width - 225
Yf = Me.Height - 550
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hlp_Click()
Help.Show
End Sub

Private Sub Rbt1_Click()
Me.Caption = "Rebound II - " + "1"
Tt.Enabled = True
T2.Enabled = 0
Tt3.Enabled = 0
Tt2.Enabled = False
Cls
D = 100
V = 400
Xf = Me.Width - 225
Yf = Me.Height - 800
ForeColor = &HFF8080
Stl = 1
Me.DrawWidth = 3
End Sub

Private Sub Rbt2_Click()
Me.Caption = "Rebound II - " + "2"
Tt.Enabled = True
T2.Enabled = 0
Tt3.Enabled = 0
Tt2.Enabled = False
Cls
D = 100
V = 400
Xf = Me.Width - 225
Yf = Me.Height - 800
ForeColor = &HFF&
Ct2 = 60
Ct = 1
Stl = 2
Me.DrawWidth = 3
End Sub

Private Sub Rbt3_Click()
Me.Caption = "Rebound II - " + "3"
Tt.Enabled = True
T2.Enabled = 0
Tt3.Enabled = 0
Tt2.Enabled = False
Cls
D = 100
V = 400
Xf = Me.Width - 225
Yf = Me.Height - 800
Stl = 3
ForeColor = &H80FF&
Me.DrawWidth = 3
End Sub

Private Sub Rbt4_Click()
Me.Caption = "Rebound II - " + "2"
Tt.Enabled = True
T2.Enabled = 0
Tt3.Enabled = 0
Tt2.Enabled = False
Cls
D = 100
V = 400
Xf = Me.Width - 225
Yf = Me.Height - 800
ForeColor = &HFF00&
Ct2 = 60
Ct = 1
Stl = 4
Me.DrawWidth = 3
End Sub

Private Sub Rbt5_Click()
Me.Caption = "Rebound II - " + "4"
Tt.Enabled = False
T2.Enabled = True
Tt3.Enabled = 0
Tt2.Enabled = False
Cls
D = 40
V = 800
Xf = Me.Width - 225
Yf = Me.Height - 800
Xm = (Me.Width / 2)
Ym = (Me.Height / 2) - 200
Cx = 10
Xx = Xm
Cy = 10
Yy = Ym
For Ct = 1 To 3
    Xxa(Ct) = Xm
    Yya(Ct) = Ym
Next
ForeColor = &HFF8080
Ct = 1
Stl = 5
M = 1
Me.DrawWidth = 3
End Sub

Private Sub Rbt6_Click()
Me.Caption = "Rebound II - " + "4"
Tt.Enabled = False
T2.Enabled = True
Tt3.Enabled = 0
Tt2.Enabled = False
Cls
D = 30
V = 800
Xf = Me.Width - 225
Yf = Me.Height - 800
Xm = (Me.Width / 2)
Ym = (Me.Height / 2) - 200
Cx = 10
Xx = Xm
Cy = 10
Yy = Ym
For Ct = 1 To 3
    Xxa(Ct) = Xm
    Yya(Ct) = Ym
Next
ForeColor = &HFF8080
Ct = 1
Stl = 6
M = 1
Me.DrawWidth = 3
End Sub

Private Sub Rbt8_Click()
Me.Caption = "Rebound II - " + "4"
Tt.Enabled = False
T2.Enabled = True
Tt3.Enabled = 0
Tt2.Enabled = False
Cls
D = 40
V = 800
Xf = Me.Width - 225
Yf = Me.Height - 800
Xm = (Me.Width / 2)
Ym = (Me.Height / 2) - 200
Cx = 10
Xx = Xm
Cy = 10
Yy = Ym
For Ct = 1 To 3
    Xxa(Ct) = Xm
    Yya(Ct) = Ym
Next
ForeColor = &HFF8080
Ct = 1
Stl = 8
M = 1
Me.DrawWidth = 3
End Sub

Private Sub Rbt9_Click()
Me.Caption = "Rebound II - " + "6"
Tt.Enabled = False
T2.Enabled = 0
Tt3.Enabled = 0
Tt2.Enabled = True
Cls
Stl = 9
Ct = 1
M = 1
D = 100
V = 1000
Cx = 1000
Cy = 1000
Me.DrawWidth = 1
End Sub

Private Sub T2_Timer()
    
    Ct = Ct + M
    If (Ct > 60) Then Ct = 1
    If (Ct < 0) Then Ct = 59
    Ll.Caption = Str(M) + " - " + Str(Ct)

    Xx = Xm + (Sin(Ct * 6 * Pix)) * 1
    Yy = Ym - (Cos(Ct * 6 * Pix)) * 1
    
    Xxa(1) = Xx + (Sin(Ct * 6 * Pix)) * V
    Yya(1) = Yy - (Cos(Ct * 6 * Pix)) * V

    Xxa(2) = Xx + (Sin((Ct + (20)) * 6 * Pix)) * V
    Yya(2) = Yy - (Cos((Ct + (20)) * 6 * Pix)) * V

    Xxa(3) = Xx + (Sin((Ct + (40)) * 6 * Pix)) * V
    Yya(3) = Yy - (Cos((Ct + (40)) * 6 * Pix)) * V

'----------------------------------------------
If (Xt = 0) Then
Cx = Cx + D * 3
End If

If (Yt = 0) Then
Cy = Cy + D * 3
End If

If (Xt = 1) Then
Cx = Cx - D * 3
End If

If (Yt = 1) Then
Cy = Cy - D * 3
End If
  
For Cv = 1 To 3

If (Xxa(Cv) + Cx > Xf) Then
Xt = 1
M = M * -1
If (Yt = 1) Or (Stl = 5) Then Cls
End If

If (Yya(Cv) + Cy > Yf) Then
Yt = 1
If (Xt = 1) Or (Stl = 5) Then Cls
End If

If (Xxa(Cv) + Cx < 100) Then
Xt = 0
M = M * -1
If (Yt = 0) Or (Stl = 5) Then Cls
End If

If (Yya(Cv) + Cy < 100) Then
Yt = 0
If (Xt = 0) Or (Stl = 5) Then Cls
End If

Next

'----------------------------------------------
If (Stl = 8) Then Cls
    Line (Xxa(1) + Cx, Yya(1) + Cy)-(Xxa(3) + Cx, Yya(3) + Cy)
    Line (Xxa(2) + Cx, Yya(2) + Cy)-(Xxa(3) + Cx, Yya(3) + Cy), &H80C0FF
    Line (Xxa(2) + Cx, Yya(2) + Cy)-(Xxa(1) + Cx, Yya(1) + Cy), &HFF&

End Sub

Private Sub Tt_Timer()
If (Stl = 3) Or (Stl = 4) Then Cls

If (Stl <> 3) Then
Ct = Ct + 1
If (Ct > 60) Then Ct = 1
Xa = Xx + (Sin(Ct * 6 * Pix)) * V
Ya = Yy - (Cos(Ct * 6 * Pix)) * V

If (Stl = 2) Or (Stl = 4) Then
Ct2 = Ct2 - 1
If (Ct2 < 1) Then Ct2 = 60
Xa2 = Xx + (Sin(Ct2 * 6 * Pix)) * V
Ya2 = Yy - (Cos(Ct2 * 6 * Pix)) * V
End If

End If

If (Xt = 0) Then
Xx = Xx + D
If (Stl = 3) Then Xa = Xx - V
End If

If (Yt = 0) Then
Yy = Yy + D
If (Stl = 3) Then Ya = Yy - V
End If

If (Xt = 1) Then
Xx = Xx - D
If (Stl = 3) Then Xa = Xx + V
End If

If (Yt = 1) Then
Yy = Yy - D
If (Stl = 3) Then Ya = Yy + V
End If

If (Xx > Xf) Then
Xt = 1
If (Stl <> 3) Then Cls
End If

If (Yy > Yf) Then
Yt = 1
If (Stl <> 3) Then Cls
End If

If (Xx < 100) Then
Xt = 0
If (Stl <> 3) Then Cls
End If

If (Yy < 100) Then
Yt = 0
If (Stl <> 3) Then Cls
End If

Line (Xa, Ya)-(Xx, Yy)
If (Stl = 2) Or (Stl = 4) Then Line (Xa2, Ya2)-(Xx, Yy), &HFF0000

End Sub

Private Sub Tt2_Timer()
Ct = Ct + M
    If (Ct > 60) Then Ct = 1
    If (Ct < 0) Then Ct = 59
    
    Xx = Xm + (Sin(Ct * 6 * Pix)) * 1
    Yy = Ym - (Cos(Ct * 6 * Pix)) * 1
    
    '*12*
    Xxa(1) = Xx + (Sin(Ct * 6 * Pix)) * V
    Yya(1) = Yy - (Cos(Ct * 6 * Pix)) * V

    '*3*
    Xxa(2) = Xx + (Sin((Ct + (15)) * 6 * Pix)) * V
    Yya(2) = Yy - (Cos((Ct + (15)) * 6 * Pix)) * V

    '*6*
    Xxa(3) = Xx + (Sin((Ct + (30)) * 6 * Pix)) * V
    Yya(3) = Yy - (Cos((Ct + (30)) * 6 * Pix)) * V
    
    '*9*
    Xxa(4) = Xx + (Sin((Ct + (45)) * 6 * Pix)) * V
    Yya(4) = Yy - (Cos((Ct + (45)) * 6 * Pix)) * V
    
    '*10* y *11*
    Xxa(5) = Xx + (Sin((Ct + (53)) * 6 * Pix)) * (V / 2)
    Yya(5) = Yy - (Cos((Ct + (53)) * 6 * Pix)) * (V / 2)

    '*7* y *8*
    Xxa(6) = Xx + (Sin((Ct + (38)) * 6 * Pix)) * (V / 2)
    Yya(6) = Yy - (Cos((Ct + (38)) * 6 * Pix)) * (V / 2)

    '*4* y *5*
    Xxa(7) = Xx + (Sin((Ct + (23)) * 6 * Pix)) * (V / 2)
    Yya(7) = Yy - (Cos((Ct + (23)) * 6 * Pix)) * (V / 2)
    
    '*1* y *2*
    Xxa(8) = Xx + (Sin((Ct + (8)) * 6 * Pix)) * (V / 2)
    Yya(8) = Yy - (Cos((Ct + (8)) * 6 * Pix)) * (V / 2)
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

For Cv = 1 To 4

If (Xxa(Cv) + Cx > Xf) Then
Xt = 1
M = M * -1
Cls
End If

If (Yya(Cv) + Cy > Yf) Then
Yt = 1
Cls
End If

If (Xxa(Cv) + Cx < 0) Then
Xt = 0
M = M * -1
Cls
End If

If (Yya(Cv) + Cy < 40) Then
Yt = 0
Cls
End If

Next

'----------------------------------------------

Line (Xxa(1) + Cx, Yya(1) + Cy)-(Xxa(5) + Cx, Yya(5) + Cy), &HFF00&
Line (Xxa(5) + Cx, Yya(5) + Cy)-(Xxa(4) + Cx, Yya(4) + Cy), &HFF8080
Line (Xxa(4) + Cx, Yya(4) + Cy)-(Xxa(6) + Cx, Yya(6) + Cy), &HFF8080
Line (Xxa(6) + Cx, Yya(6) + Cy)-(Xxa(3) + Cx, Yya(3) + Cy), &H80FF&
Line (Xxa(3) + Cx, Yya(3) + Cy)-(Xxa(7) + Cx, Yya(7) + Cy), &H80FF&
Line (Xxa(7) + Cx, Yya(7) + Cy)-(Xxa(2) + Cx, Yya(2) + Cy), &HFFFF&
Line (Xxa(2) + Cx, Yya(2) + Cy)-(Xxa(8) + Cx, Yya(8) + Cy), &HFFFF&
Line (Xxa(8) + Cx, Yya(8) + Cy)-(Xxa(1) + Cx, Yya(1) + Cy), &HFF00&

End Sub
