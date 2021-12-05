VERSION 5.00
Begin VB.Form bluecal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4605
   ClientLeft      =   4170
   ClientTop       =   1110
   ClientWidth     =   3990
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "bluecal.frx":0000
   ScaleHeight     =   4605
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   7920
      Top             =   2400
   End
   Begin VB.TextBox txt1 
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   3495
   End
   Begin VB.Shape Shapedem 
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lbldem 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Jesús Rivas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   3495
   End
   Begin VB.Shape Shape100 
      Height          =   4335
      Left            =   120
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lbld 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "1/x"
      Height          =   375
      Left            =   1800
      TabIndex        =   20
      Top             =   3840
      Width           =   375
   End
   Begin VB.Shape Shaped 
      Height          =   615
      Left            =   1680
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblsqr 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "sqr"
      Height          =   375
      Left            =   2520
      TabIndex        =   19
      Top             =   3840
      Width           =   375
   End
   Begin VB.Shape Shapesqr 
      Height          =   615
      Left            =   2400
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lbloff 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Off"
      Height          =   375
      Left            =   3240
      TabIndex        =   18
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape Shapeoff 
      Height          =   615
      Left            =   3120
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbleql 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "="
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   3840
      Width           =   375
   End
   Begin VB.Shape Shapeeql 
      Height          =   615
      Left            =   3120
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lblce 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "C"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape Shapece 
      Height          =   615
      Left            =   2400
      Top             =   1560
      Width           =   615
   End
   Begin VB.Shape Shapemul 
      Height          =   615
      Left            =   2400
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lblmul 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shapedot 
      Height          =   615
      Left            =   960
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lbldot 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   3840
      Width           =   375
   End
   Begin VB.Shape Shapediv 
      Height          =   615
      Left            =   3120
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lbldiv 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "/"
      Height          =   375
      Left            =   3240
      TabIndex        =   12
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shape8 
      Height          =   615
      Left            =   960
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lbl8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "8"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shape7 
      Height          =   615
      Left            =   240
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lbl7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "7"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shapeminus 
      Height          =   615
      Left            =   3120
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblminus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "-"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2400
      Width           =   375
   End
   Begin VB.Shape Shapeplus 
      Height          =   615
      Left            =   2400
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblplus 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "+"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2400
      Width           =   375
   End
   Begin VB.Shape Shape0 
      Height          =   615
      Left            =   240
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label lbl0 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   375
   End
   Begin VB.Shape Shape9 
      Height          =   615
      Left            =   1680
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label lbl9 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "9"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Width           =   375
   End
   Begin VB.Shape Shape6 
      Height          =   615
      Left            =   1680
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lbl6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "6"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2400
      Width           =   375
   End
   Begin VB.Shape Shape5 
      Height          =   615
      Left            =   960
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lbl5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "5"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2400
      Width           =   375
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   240
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "4"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Width           =   375
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   1680
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "3"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   960
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "2"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   375
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   240
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "1"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   375
   End
End
Attribute VB_Name = "bluecal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Double
Dim pflag As Boolean
Dim mflag As Boolean
Dim mulflag As Boolean
Dim divflag As Boolean
Dim i As Integer
Private Sub Label7_Click()
txt1.Text = Val(txt1.Text) + "."
End Sub

Private Sub Form_Load()
i = 7
pflag = False
mflag = False
divflag = False
mulflag = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &H80000012
lbl2.ForeColor = &H80000012
lbl3.ForeColor = &H80000012
lbl4.ForeColor = &H80000012
lbl5.ForeColor = &H80000012
lbl6.ForeColor = &H80000012
lbl7.ForeColor = &H80000012
lbl8.ForeColor = &H80000012
lbl9.ForeColor = &H80000012
lbl0.ForeColor = &H80000012
lblplus.ForeColor = &H80000012
lblminus.ForeColor = &H80000012
lbldiv.ForeColor = &H80000012
lbldot.ForeColor = &H80000012
lblmul.ForeColor = &H80000012
lbleql.ForeColor = &H80000012
lbld.ForeColor = &H80000012
lbloff.ForeColor = &H80000012
lblce.ForeColor = &H80000012
lblsqr.ForeColor = &H80000012


Shape1.BorderColor = &H80000012
Shape2.BorderColor = &H80000012
Shape3.BorderColor = &H80000012
Shape4.BorderColor = &H80000012
Shape5.BorderColor = &H80000012
Shape6.BorderColor = &H80000012
Shape7.BorderColor = &H80000012
Shape8.BorderColor = &H80000012
Shape9.BorderColor = &H80000012
Shape0.BorderColor = &H80000012
Shapeplus.BorderColor = &H80000012
Shapeminus.BorderColor = &H80000012
Shapediv.BorderColor = &H80000012
Shapemul.BorderColor = &H80000012
Shapedot.BorderColor = &H80000012
Shapeeql.BorderColor = &H80000012
Shapeoff.BorderColor = &H80000012
Shapece.BorderColor = &H80000012
Shapesqr.BorderColor = &H80000012
Shaped.BorderColor = &H80000012




End Sub
Private Sub lbl0_Click()
txt1.Text = txt1.Text + "0"
End Sub

Private Sub lbl0_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl0.ForeColor = &HFF0000
Shape0.BorderColor = &HFF0000
End Sub

Private Sub lbl1_Click()
txt1.Text = txt1.Text + "1"
End Sub

Private Sub lbl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl1.ForeColor = &HFF0000
Shape1.BorderColor = &HFF0000
End Sub

Private Sub lbl2_Click()
txt1.Text = txt1.Text + "2"

End Sub

Private Sub lbl2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl2.ForeColor = &HFF0000
Shape2.BorderColor = &HFF0000
End Sub

Private Sub lbl3_Click()
txt1.Text = txt1.Text + "3"
End Sub

Private Sub lbl3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl3.ForeColor = &HFF0000
Shape3.BorderColor = &HFF0000
End Sub

Private Sub lbl4_Click()
txt1.Text = txt1.Text + "4"
End Sub

Private Sub lbl4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl4.ForeColor = &HFF0000
Shape4.BorderColor = &HFF0000
End Sub

Private Sub lbl5_Click()
txt1.Text = txt1.Text + "5"
End Sub

Private Sub lbl5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl5.ForeColor = &HFF0000
Shape5.BorderColor = &HFF0000
End Sub

Private Sub lbl6_Click()
txt1.Text = txt1.Text + "6"
End Sub

Private Sub lbl6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl6.ForeColor = &HFF0000
Shape6.BorderColor = &HFF0000
End Sub

Private Sub lbl7_Click()
txt1.Text = txt1.Text + "7"
End Sub

Private Sub lbl7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl7.ForeColor = &HFF0000
Shape7.BorderColor = &HFF0000
End Sub

Private Sub lbl8_Click()
txt1.Text = txt1.Text + "8"
End Sub

Private Sub lbl8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl8.ForeColor = &HFF0000
Shape8.BorderColor = &HFF0000
End Sub

Private Sub lbl9_Click()
txt1.Text = txt1.Text + "9"
End Sub

Private Sub lbl9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbl9.ForeColor = &HFF0000
Shape9.BorderColor = &HFF0000
End Sub

Private Sub lblce_Click()
txt1.Text = " "

End Sub

Private Sub lblce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblce.ForeColor = &HFF0000
Shapece.BorderColor = &HFF0000

End Sub

Private Sub lbld_Click()
temp = Val(txt1.Text)
txt1.Text = 1 / temp
End Sub

Private Sub lbld_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbld.ForeColor = &HFF0000
Shaped.BorderColor = &HFF0000
End Sub

Private Sub lbldiv_Click()
divflag = True
temp = Val(txt1.Text)
txt1.Text = " "


End Sub

Private Sub lbldiv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldiv.ForeColor = &HFF0000
Shapediv.BorderColor = &HFF0000
End Sub

Private Sub lbldot_Click()
txt1.Text = txt1.Text + "."
End Sub

Private Sub lbldot_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbldot.ForeColor = &HFF0000
Shapedot.BorderColor = &HFF0000
End Sub

Private Sub lbleql_Click()

If divflag = True Then txt1.Text = temp / Val(txt1.Text)
divflag = False
If pflag = True Then txt1.Text = temp + Val(txt1.Text)
pflag = False
If mflag = True Then txt1.Text = temp - Val(txt1.Text)
mflag = False
If mulflag = True Then txt1.Text = temp * Val(txt1.Text)
mulflag = False
End Sub

Private Sub lbleql_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbleql.ForeColor = &HFF0000
Shapeeql.BorderColor = &HFF0000

End Sub

Private Sub lblminus_Click()
mflag = True
temp = Val(txt1.Text)
txt1.Text = " "
End Sub

Private Sub lblminus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblminus.ForeColor = &HFF0000
Shapeminus.BorderColor = &HFF0000
End Sub

Private Sub lblmul_Click()
mulflag = True
temp = Val(txt1.Text)
txt1.Text = " "
End Sub

Private Sub lblmul_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblmul.ForeColor = &HFF0000
Shapemul.BorderColor = &HFF0000
End Sub

Private Sub lbloff_Click()
End
End Sub

Private Sub lbloff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lbloff.ForeColor = &HFF0000
Shapeoff.BorderColor = &HFF0000

End Sub

Private Sub lblplus_Click()
pflag = True
temp = Val(txt1.Text)
txt1.Text = " "
End Sub

Private Sub lblplus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblplus.ForeColor = &HFF0000
Shapeplus.BorderColor = &HFF0000
End Sub

Private Sub lblsqr_Click()
txt1.Text = Val(txt1.Text) * 2

End Sub

Private Sub lblsqr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblsqr.ForeColor = &HFF0000
Shapesqr.BorderColor = &HFF0000

End Sub


Private Sub Timer1_Timer()

If (i > 7) Then i = 6
i = i + 1
lbldem.ForeColor = QBColor(i)


End Sub
