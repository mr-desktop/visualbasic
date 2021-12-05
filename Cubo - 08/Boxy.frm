VERSION 5.00
Begin VB.Form Box 
   BackColor       =   &H80000008&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inferno-Box - 1"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5100
   ForeColor       =   &H0000FF00&
   Icon            =   "Boxy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   120
   End
   Begin VB.Menu St 
      Caption         =   "Estilo"
      Begin VB.Menu St1 
         Caption         =   "1"
      End
      Begin VB.Menu St2 
         Caption         =   "2"
      End
      Begin VB.Menu St3 
         Caption         =   "3"
      End
      Begin VB.Menu St4 
         Caption         =   "4"
      End
      Begin VB.Menu St5 
         Caption         =   "5"
      End
      Begin VB.Menu St6 
         Caption         =   "6"
      End
      Begin VB.Menu St7 
         Caption         =   "7"
      End
   End
   Begin VB.Menu Hlp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x(1 To 8), y(1 To 8), i, xx, yy, aA As Integer
Const pix = 3.14159265358979

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Hlp_Click()
Help.Show
End Sub

Private Sub St1_Click()
Me.Caption = "Inferno-Box - " + "1"
aA = 0
End Sub

Private Sub St2_Click()
Me.Caption = "Inferno-Box - " + "2"
aA = 1
End Sub

Private Sub St3_Click()
Me.Caption = "Inferno-Box - " + "3"
aA = 2
End Sub

Private Sub St4_Click()
Me.Caption = "Inferno-Box - " + "4"
aA = 3
End Sub

Private Sub St5_Click()
Me.Caption = "Inferno-Box - " + "5"
aA = 4
End Sub

Private Sub St6_Click()
Me.Caption = "Inferno-Box - " + "6"
aA = 5
End Sub

Private Sub St7_Click()
Me.Caption = "Inferno-Box - " + "7"
aA = 6
End Sub

Private Sub Timer1_Timer()
yy = Sqr(2 * (1560 / 2) ^ 2)
xx = Sqr(2 * (2160 / 2) ^ 2)

i = i + 10
Box (aA)

x(5) = x(1) + 1000
x(6) = x(2) + 1000
x(7) = x(3) + 1000
x(8) = x(4) + 1000

y(5) = y(1) + 1000
y(6) = y(2) + 1000
y(7) = y(3) + 1000
y(8) = y(4) + 1000

Cls

'1
Line (x(5), y(5))-(x(8), y(8))
'2
Line (x(6), y(6))-(x(5), y(5))
'3
Line (x(7), y(7))-(x(8), y(8))
'4
Line (x(6), y(6))-(x(7), y(7))
'5
Line (x(1), y(1))-(x(4), y(4))
'6
Line (x(2), y(2))-(x(1), y(1))
'7
Line (x(3), y(3))-(x(4), y(4))
'8
Line (x(2), y(2))-(x(3), y(3))
'9
Line (x(5), y(5))-(x(1), y(1))
'10
Line (x(6), y(6))-(x(2), y(2))
'11
Line (x(8), y(8))-(x(4), y(4))
'12
Line (x(7), y(7))-(x(3), y(3))

End Sub

Public Function Box(N As Integer)

If (N = 6) Then
x(4) = xx * Cos(i * pix / 180) + 2000
y(2) = yy * Sin(i * pix / 180) + 2000

x(2) = xx * Cos((i + 2 * 45) * pix / 180) + 2000
y(1) = yy * Sin((i + 2 * 45) * pix / 180) + 2000

x(1) = xx * Cos((i + 6 * 45) * pix / 180) + 2000
y(3) = yy * Sin((i + 6 * 45) * pix / 180) + 2000

x(3) = xx * Cos((i + 4 * 45) * pix / 180) + 2000
y(4) = yy * Sin((i + 4 * 45) * pix / 180) + 2000
End If

If (N = 5) Then
x(3) = xx * Cos(i * pix / 180) + 2000
y(2) = yy * Sin(i * pix / 180) + 2000

x(1) = xx * Cos((i + 2 * 45) * pix / 180) + 2000
y(1) = yy * Sin((i + 2 * 45) * pix / 180) + 2000

x(4) = xx * Cos((i + 6 * 45) * pix / 180) + 2000
y(4) = yy * Sin((i + 6 * 45) * pix / 180) + 2000

x(2) = xx * Cos((i + 4 * 45) * pix / 180) + 2000
y(3) = yy * Sin((i + 4 * 45) * pix / 180) + 2000
End If

If (N = 2) Then
x(1) = xx * Cos(i * pix / 180) + 2000
y(1) = yy * Sin(i * pix / 180) + 2000

x(2) = xx * Cos((i + 2 * 45) * pix / 180) + 2000
y(2) = yy * Sin((i + 2 * 45) * pix / 180) + 2000

x(3) = xx * Cos((i + 6 * 45) * pix / 180) + 2000
y(3) = yy * Sin((i + 6 * 45) * pix / 180) + 2000

x(4) = xx * Cos((i + 4 * 45) * pix / 180) + 2000
y(4) = yy * Sin((i + 4 * 45) * pix / 180) + 2000
End If

If (N = 3) Then
x(1) = xx * Cos(i * pix / 180) + 2000
y(4) = yy * Sin(i * pix / 180) + 2000

x(2) = xx * Cos((i + 2 * 45) * pix / 180) + 2000
y(3) = yy * Sin((i + 2 * 45) * pix / 180) + 2000

x(3) = xx * Cos((i + 6 * 45) * pix / 180) + 2000
y(2) = yy * Sin((i + 6 * 45) * pix / 180) + 2000

x(4) = xx * Cos((i + 4 * 45) * pix / 180) + 2000
y(1) = yy * Sin((i + 4 * 45) * pix / 180) + 2000
End If

If (N = 4) Then
x(2) = xx * Cos(i * pix / 180) + 2000
y(1) = yy * Sin(i * pix / 180) + 2000

x(1) = xx * Cos((i + 2 * 45) * pix / 180) + 2000
y(2) = yy * Sin((i + 2 * 45) * pix / 180) + 2000

x(3) = xx * Cos((i + 6 * 45) * pix / 180) + 2000
y(4) = yy * Sin((i + 6 * 45) * pix / 180) + 2000

x(4) = xx * Cos((i + 4 * 45) * pix / 180) + 2000
y(3) = yy * Sin((i + 4 * 45) * pix / 180) + 2000
End If

If (N = 1) Then
x(2) = xx * Cos(i * pix / 180) + 2000
y(2) = yy * Sin(i * pix / 180) + 2000

x(1) = xx * Cos((i + 2 * 45) * pix / 180) + 2000
y(1) = yy * Sin((i + 2 * 45) * pix / 180) + 2000

x(3) = xx * Cos((i + 6 * 45) * pix / 180) + 2000
y(3) = yy * Sin((i + 6 * 45) * pix / 180) + 2000

x(4) = xx * Cos((i + 4 * 45) * pix / 180) + 2000
y(4) = yy * Sin((i + 4 * 45) * pix / 180) + 2000
End If

If (N = 0) Then
x(4) = xx * Cos(i * pix / 180) + 2000
y(2) = yy * Sin(i * pix / 180) + 2000

x(3) = xx * Cos((i + 2 * 45) * pix / 180) + 2000
y(1) = yy * Sin((i + 2 * 45) * pix / 180) + 2000

x(1) = xx * Cos((i + 6 * 45) * pix / 180) + 2000
y(3) = yy * Sin((i + 6 * 45) * pix / 180) + 2000

x(2) = xx * Cos((i + 4 * 45) * pix / 180) + 2000
y(4) = yy * Sin((i + 4 * 45) * pix / 180) + 2000
End If

End Function
