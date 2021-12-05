VERSION 5.00
Begin VB.Form Rebound 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rebound III"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   ForeColor       =   &H0000FF00&
   Icon            =   "ReboundIII.frx":0000
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
   Begin VB.Label Hlp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ayuda"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Rebound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dd, Xf, Yf, Xx, Yy, Vl, Rt As Integer
Dim Xt, Yt, Jk As Boolean

Private Sub Form_Load()
Yf = Me.Height ' - 3300
Xf = Me.Width ' - 3570

Xt = True
Yt = True

Vl = 5
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Hlp_Click()
Help.Show
End Sub

Private Sub Rbt1_Click()
Me.Caption = "Rebound III - " + "1"
Rt = 1
End Sub

Private Sub Tt_Timer()

'Incremento en X'
If (Xt = True) Then
Xx = Xx + Vl
End If
'Final de Incremento en X'
If (Xt = True) And ((Xx + 7980) >= Xf) Then

Xt = False
End If

'Decremento en X'
If (Xt = False) Then
Xx = Xx - Vl
End If
'Final Decremento en X'
If (Xt = False) And (Xx <= 0) Then
Xt = True
End If

'Incremento en Y'
If (Yt = True) Then
Yy = Yy + Vl
End If
'Final de Incremento en Y'
If (Yt = True) And ((Yy + 7140) >= Yf) Then
Yt = False
End If

'Decremento en Y'
If (Yt = False) Then
Yy = Yy - Vl
End If
'Final de Decremento en Y'
If (Yt = False) And (Yy <= 20) Then
Yt = True
End If

'Estrella de 4 Puntas
Dim lpPts(15)  As POINTAPI
Dim hRustedRgn As Long

    lpPts(0).x = 0:             lpPts(0).y = 0
    lpPts(1).x = 0:             lpPts(1).y = Me.Height
    lpPts(2).x = Me.Width:      lpPts(2).y = Me.Height
    lpPts(3).x = Me.Width:      lpPts(3).y = 0
    
    lpPts(4).x = 75 + Xx:        lpPts(4).y = 0
    lpPts(5).x = 75 + Xx:        lpPts(5).y = 0 + Yy
    lpPts(6).x = 50 + Xx:        lpPts(6).y = 50 + Yy
    lpPts(7).x = 0 + Xx:         lpPts(7).y = 75 + Yy
    lpPts(8).x = 50 + Xx:        lpPts(8).y = 100 + Yy
    lpPts(9).x = 75 + Xx:        lpPts(9).y = 150 + Yy
    lpPts(10).x = 100 + Xx:      lpPts(10).y = 100 + Yy
    lpPts(11).x = 150 + Xx:      lpPts(11).y = 75 + Yy
    lpPts(12).x = 100 + Xx:      lpPts(12).y = 50 + Yy
    lpPts(13).x = 75 + Xx:       lpPts(13).y = 0 + Yy
    lpPts(14).x = 75 + Xx:       lpPts(14).y = 0
    
hRustedRgn = CreatePolygonRgn(lpPts(0), UBound(lpPts), ALTERNATE)
Call SetWindowRgn(Me.hWnd, hRustedRgn, True)

End Sub
