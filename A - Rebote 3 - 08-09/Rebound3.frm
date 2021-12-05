VERSION 5.00
Begin VB.Form Rebound 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rebound III"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8310
   ForeColor       =   &H00FF8080&
   Icon            =   "Rebound3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tt 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu Re 
      Caption         =   "Rebound"
      Begin VB.Menu Ended 
         Caption         =   "Salir"
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
Dim Xxa(1 To 8), Yya(1 To 8), Ct, Cy, Stl, Cx, D, S1, S2, Xt, Yt, Xf, Yf, Xx, Yy, M, Xm, Ym, V As Integer
Const Pix = 3.1415927 / 180

Private Sub Ended_Click()
End
End Sub

Private Sub Form_Load()
Cls
S1 = -7595
S2 = -6460
D = 5
V = 80
Xf = Me.Width - 225
Yf = Me.Height - 800
Xm = 200
Ym = 200
Cx = 10
Xx = Xm
Cy = 10
Yy = Ym
For Ct = 1 To 4
    Xxa(Ct) = Xm
    Yya(Ct) = Ym
Next
Ct = 1
M = 1
Stl = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub hlp_Click()
Help.Show
End Sub

Private Sub stl1_Click()

Stl = 1
S1 = -7595
S2 = -6460
D = 5
V = 80
Cx = 0
Cy = 0
End Sub


Private Sub Tt_Timer()
    Ct = Ct + M
    If (Ct > 60) Then Ct = 1
    If (Ct < 0) Then Ct = 59
    'Cls
    'Print Str(Ct) + " " + Str(M) + " " + Str(Cx) + " " + Str(Cy)
    
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

If (Xxa(Cv) + Cx > Xf + S1) Then
Xt = 1
M = M * -1

End If

If (Yya(Cv) + Cy > Yf + S2) Then
Yt = 1

End If

If (Xxa(Cv) + Cx < 0) Then
Xt = 0
M = M * -1

End If

If (Yya(Cv) + Cy < 40) Then
Yt = 0

End If

Next

'----------------------------------------------


'Estrella de 4 Puntas
Dim lpPts(15)  As POINTAPI
Dim hRustedRgn As Long

    lpPts(0).x = 0:             lpPts(0).y = 0
    lpPts(1).x = 0:             lpPts(1).y = Me.Height
    lpPts(2).x = Me.Width:      lpPts(2).y = Me.Height
    lpPts(3).x = Me.Width:      lpPts(3).y = 0
    
    '*12*
    lpPts(4).x = Xxa(1) + Cx:      lpPts(4).y = 0
    lpPts(5).x = Xxa(1) + Cx:      lpPts(5).y = Yya(1) + Cy
    '*10* y *11*
    lpPts(6).x = Xxa(5) + Cx:      lpPts(6).y = Yya(5) + Cy
    
    '*9*
    lpPts(7).x = Xxa(4) + Cx:       lpPts(7).y = Yya(4) + Cy
    
    '*7* y *8*
    lpPts(8).x = Xxa(6) + Cx:      lpPts(8).y = Yya(6) + Cy
    
    '*6*
    lpPts(9).x = Xxa(3) + Cx:      lpPts(9).y = Yya(3) + Cy
    
    '*4* y *5*
    lpPts(10).x = Xxa(7) + Cx:    lpPts(10).y = Yya(7) + Cy
    
    '*3*
    lpPts(11).x = Xxa(2) + Cx:    lpPts(11).y = Yya(2) + Cy
    
    '*2*  y *1*
    lpPts(12).x = Xxa(8) + Cx:    lpPts(12).y = Yya(8) + Cy
    
    '*12*
    lpPts(13).x = Xxa(1) + Cx:     lpPts(13).y = Yya(1) + Cy
    
    lpPts(14).x = Xxa(1) + Cx:     lpPts(14).y = 0
    
hRustedRgn = CreatePolygonRgn(lpPts(0), UBound(lpPts), ALTERNATE)
Call SetWindowRgn(Me.hWnd, hRustedRgn, True)

End Sub
