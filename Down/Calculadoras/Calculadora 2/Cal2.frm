VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CALCUCOLOR"
   ClientHeight    =   2880
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Cal2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   3885
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Check1 
      BackColor       =   &H008080FF&
      Caption         =   "Col"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtdisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   840
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmdEqual 
      BackColor       =   &H008080FF&
      Caption         =   "="
      Height          =   375
      Left            =   2040
      MouseIcon       =   "Cal2.frx":030A
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Result"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "x²"
      Height          =   375
      Left            =   240
      MouseIcon       =   "Cal2.frx":074C
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Square"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "%"
      Height          =   375
      Left            =   240
      MouseIcon       =   "Cal2.frx":0B8E
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Percentage"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "+/-"
      Height          =   375
      Left            =   240
      MouseIcon       =   "Cal2.frx":0FD0
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Change Sign"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FF8080&
      Caption         =   "Hide"
      Height          =   375
      Left            =   1440
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3600
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   16
      Left            =   2640
      Max             =   256
      SmallChange     =   8
      TabIndex        =   27
      Top             =   3120
      Value           =   192
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   16
      Left            =   1440
      Max             =   256
      SmallChange     =   8
      TabIndex        =   26
      Top             =   3120
      Value           =   192
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   16
      Left            =   240
      Max             =   256
      SmallChange     =   8
      TabIndex        =   25
      Top             =   3120
      Value           =   192
      Width           =   1095
   End
   Begin VB.CommandButton cmdPoint 
      BackColor       =   &H00C0E0FF&
      Caption         =   "."
      Height          =   375
      Left            =   840
      MouseIcon       =   "Cal2.frx":1412
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Decimal"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MC"
      Height          =   375
      Left            =   3240
      MouseIcon       =   "Cal2.frx":1854
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Memory Clear"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MR"
      Height          =   375
      Left            =   3240
      MouseIcon       =   "Cal2.frx":1C96
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Memory Recall"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "C"
      Height          =   375
      Left            =   240
      MouseIcon       =   "Cal2.frx":20D8
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Clear"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd0 
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   375
      Left            =   1440
      MouseIcon       =   "Cal2.frx":251A
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "0"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdPlus 
      BackColor       =   &H00C0C0FF&
      Caption         =   "+"
      Height          =   375
      Left            =   2640
      MouseIcon       =   "Cal2.frx":295C
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Add"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton cmdMinus 
      BackColor       =   &H00C0C0FF&
      Caption         =   "-"
      Height          =   375
      Left            =   2640
      MouseIcon       =   "Cal2.frx":2D9E
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Subtract"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmdMultiply 
      BackColor       =   &H00C0C0FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MouseIcon       =   "Cal2.frx":31E0
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Multiply"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmdDivide 
      BackColor       =   &H00C0C0FF&
      Caption         =   "/"
      Height          =   375
      Left            =   2640
      MouseIcon       =   "Cal2.frx":3622
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Divide"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "3"
      Height          =   375
      Left            =   2040
      MouseIcon       =   "Cal2.frx":3A64
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "3"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmd2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "2"
      Height          =   375
      Left            =   1440
      MouseIcon       =   "Cal2.frx":3EA6
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "2"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmd1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "1"
      Height          =   375
      Left            =   840
      MouseIcon       =   "Cal2.frx":42E8
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "1"
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton cmd6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "6"
      Height          =   375
      Left            =   2040
      MouseIcon       =   "Cal2.frx":472A
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "6"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmd5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "5"
      Height          =   375
      Left            =   1440
      MouseIcon       =   "Cal2.frx":4B6C
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "5"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmd4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "4"
      Height          =   375
      Left            =   840
      MouseIcon       =   "Cal2.frx":4FAE
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "4"
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton cmd9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "9"
      Height          =   375
      Left            =   2040
      MouseIcon       =   "Cal2.frx":53F0
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "9"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "8"
      Height          =   375
      Left            =   1440
      MouseIcon       =   "Cal2.frx":5832
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "8"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton cmd7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "7"
      Height          =   375
      Left            =   840
      MouseIcon       =   "Cal2.frx":5C74
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "7"
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MS"
      Height          =   375
      Left            =   3240
      MouseIcon       =   "Cal2.frx":60B6
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Memory Store"
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   34
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BLUE"
      Height          =   210
      Left            =   3000
      TabIndex        =   30
      Top             =   2880
      Width           =   390
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GREEN"
      Height          =   210
      Left            =   1800
      TabIndex        =   29
      Top             =   2880
      Width           =   510
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RED"
      Height          =   210
      Left            =   600
      TabIndex        =   28
      Top             =   2880
      Width           =   300
   End
   Begin VB.Label lblMemory 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   3465
      MousePointer    =   1  'Arrow
      TabIndex        =   24
      Top             =   960
      Width           =   75
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      MousePointer    =   1  'Arrow
      TabIndex        =   23
      Top             =   840
      Width           =   495
   End
   Begin VB.Menu Ed 
      Caption         =   "&Editar"
      Begin VB.Menu Cpy 
         Caption         =   "&Copiar"
         Shortcut        =   ^C
      End
      Begin VB.Menu Pste 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Hlp 
      Caption         =   "&Ayuda"
      Begin VB.Menu Key 
         Caption         =   "&Teclas rápidas"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Enrico, Lorenzo, MEM, dot, pressequaltrue
Dim equalpress, PERCENT, valpercent

Private Sub About_Click()
Form1.Enabled = False
Form2.Show
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
   Form1.Height = 3555
   Form1.BackColor = &H8000000F
   Label2.ForeColor = vbBlack
Else
   If HScroll1.Value + HScroll2.Value + HScroll3.Value <= 150 Then
      Label2.ForeColor = vbWhite
   End If
   Form1.Height = 4755
   Form1.BackColor = RGB(HScroll1, HScroll2, HScroll3)
End If

txtdisplay.SetFocus
End Sub

Private Sub cmd0_Click()
If Lorenzo <= 4 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If

If Val(txtdisplay) > 0 Or Val(txtdisplay) < 0 Or dot = 1 Then
   txtdisplay = txtdisplay + "0"
Else
   txtdisplay = "0"
   equalpress = 1
End If

txtdisplay.SetFocus
End Sub

Private Sub cmd1_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "1"


txtdisplay.SetFocus
End Sub

Private Sub cmd2_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "2"

txtdisplay.SetFocus
End Sub

Private Sub cmd3_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "3"

txtdisplay.SetFocus
End Sub

Private Sub cmd4_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "4"

txtdisplay.SetFocus
End Sub

Private Sub cmd5_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "5"

txtdisplay.SetFocus
End Sub

Private Sub cmd6_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "6"

txtdisplay.SetFocus
End Sub

Private Sub cmd7_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If

txtdisplay = txtdisplay + "7"

txtdisplay.SetFocus
End Sub

Private Sub cmd8_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "8"

txtdisplay.SetFocus
End Sub

Private Sub cmd9_Click()
If Lorenzo <= 4 And Lorenzo > 0 Then
   pressequaltrue = pressequaltrue + 1
End If

If equalpress = 1 Then
   txtdisplay = ""
   equalpress = 0
End If
txtdisplay = txtdisplay + "9"

txtdisplay.SetFocus
End Sub

Private Sub cmdClear_Click()
txtdisplay = ""
equalpress = 1
Lorenzo = 0
dot = 0
PERCENT = 0
txtdisplay.SetFocus
End Sub

Private Sub cmdDivide_Click()
Enrico = Val(txtdisplay)
txtdisplay = ""
Lorenzo = 4
dot = 0
pressequaltrue = 1

txtdisplay.SetFocus
End Sub

Private Sub cmdEqual_Click()
If pressequaltrue > 1 Then
   If Lorenzo = 1 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico + valpercent
        
      Else
         txtdisplay = Enrico + Val(txtdisplay)
      End If
   ElseIf Lorenzo = 2 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico - valpercent
               
      Else
         txtdisplay = Enrico - Val(txtdisplay)
      End If
   ElseIf Lorenzo = 3 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico * valpercent
      
      Else
         txtdisplay = Enrico * Val(txtdisplay)
      End If
   ElseIf Lorenzo = 4 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico / valpercent
         
      Else
         txtdisplay = Enrico / Val(txtdisplay)
      End If
   End If
   Lorenzo = 0
   pressequaltrue = 0
   dot = 0
   PERCENT = 0
End If

txtdisplay.SetFocus

equalpress = 1
End Sub

Private Sub cmdMinus_Click()
Enrico = Val(txtdisplay)
txtdisplay = ""
Lorenzo = 2
dot = 0
pressequaltrue = 1

txtdisplay.SetFocus
End Sub

Private Sub cmdMultiply_Click()
Enrico = Val(txtdisplay)
txtdisplay = ""
Lorenzo = 3
dot = 0
pressequaltrue = 1

txtdisplay.SetFocus
End Sub


Private Sub cmdPlus_Click()
Enrico = Val(txtdisplay)
txtdisplay = ""
Lorenzo = 1
dot = 0
pressequaltrue = 1

txtdisplay.SetFocus
End Sub

Private Sub cmdPoint_Click()
If txtdisplay = "" Or equalpress = 1 Then
   txtdisplay = "0"
End If

If dot = 0 Then
   txtdisplay = txtdisplay + "."
   dot = 1
End If

equalpress = 0

txtdisplay.SetFocus
End Sub

Private Sub Command1_Click()
If Not Val(txtdisplay) = 0 Then
    MEM = txtdisplay
    lblMemory.Caption = "M"
End If


txtdisplay.SetFocus
End Sub

Private Sub Command2_Click()
txtdisplay = MEM
equalpress = 1
dot = 0

txtdisplay.SetFocus
End Sub

Private Sub Command3_Click()
lblMemory = ""
MEM = ""
MC = 1
txtdisplay.SetFocus
End Sub


Private Sub Command4_Click()
Form1.Height = 3555

txtdisplay.SetFocus
End Sub



Private Sub Command5_Click()
If equalpress = 0 Then
   txtdisplay = Val(txtdisplay) - 2 * Val(txtdisplay)
   dot = 0
   If Val(txtdisplay) = 0 Then
      equalpress = 1
   End If
End If

txtdisplay.SetFocus
End Sub

Private Sub Command6_Click()
If pressequaltrue > 1 Then
   txtdisplay = (Val(txtdisplay) / 100) * Enrico
   valpercent = Val(txtdisplay)
   PERCENT = 1
End If

txtdisplay.SetFocus
End Sub

Private Sub Command7_Click()
If equalpress = 0 Then
    txtdisplay = Val(txtdisplay) * Val(txtdisplay)
    equalpress = 1
    dot = 0
End If

txtdisplay.SetFocus
End Sub

Private Sub Cpy_Click()
Clipboard.Clear
Clipboard.SetText txtdisplay
If Not Clipboard.GetText = Empty Then
    Pste.Enabled = True
End If
End Sub
Private Sub Form_Load()
If Clipboard.GetText = Empty Then
    Pste.Enabled = False
End If
End Sub

Private Sub HScroll1_Change()
If HScroll1.Value + HScroll2.Value + HScroll3.Value <= 150 Then
   Label2.ForeColor = vbWhite
   Label3.ForeColor = vbWhite
   Label4.ForeColor = vbWhite
   Label5.ForeColor = vbWhite
Else
   Label2.ForeColor = vbBlack
   Label3.ForeColor = vbBlack
   Label4.ForeColor = vbBlack
   Label5.ForeColor = vbBlack
End If

Form1.BackColor = RGB(HScroll1, HScroll2, HScroll3)

txtdisplay.SetFocus
End Sub

Private Sub HScroll2_Change()
If HScroll1.Value + HScroll2.Value + HScroll3.Value <= 150 Then
   Label2.ForeColor = vbWhite
   Label3.ForeColor = vbWhite
   Label4.ForeColor = vbWhite
   Label5.ForeColor = vbWhite
Else
   Label2.ForeColor = vbBlack
   Label3.ForeColor = vbBlack
   Label4.ForeColor = vbBlack
   Label5.ForeColor = vbBlack
End If

Form1.BackColor = RGB(HScroll1, HScroll2, HScroll3)

txtdisplay.SetFocus
End Sub

Private Sub HScroll3_Change()
If HScroll1.Value + HScroll2.Value + HScroll3.Value <= 150 Then
   Label2.ForeColor = vbWhite
   Label3.ForeColor = vbWhite
   Label4.ForeColor = vbWhite
   Label5.ForeColor = vbWhite
Else
   Label2.ForeColor = vbBlack
   Label3.ForeColor = vbBlack
   Label4.ForeColor = vbBlack
   Label5.ForeColor = vbBlack
End If

Form1.BackColor = RGB(HScroll1, HScroll2, HScroll3)

txtdisplay.SetFocus
End Sub

Private Sub Key_Click()
Form1.Enabled = False
Form3.Show
End Sub

Private Sub Pste_Click()
txtdisplay = Clipboard.GetText
End Sub

Private Sub txtDisplay_KeyPress(KeyAscii As Integer)
If KeyAscii = 48 Then
   If Lorenzo <= 4 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If

   If Val(txtdisplay) > 0 Or Val(txtdisplay) < 0 Or dot = 1 Then
      txtdisplay = txtdisplay + "0"
   Else
      txtdisplay = "0"
      equalpress = 1
   End If

   txtdisplay.SetFocus
   
ElseIf KeyAscii = 49 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "1"

ElseIf KeyAscii = 50 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "2"

ElseIf KeyAscii = 51 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "3"

ElseIf KeyAscii = 52 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "4"

ElseIf KeyAscii = 53 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "5"

ElseIf KeyAscii = 54 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "6"

ElseIf KeyAscii = 55 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "7"

ElseIf KeyAscii = 56 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "8"

ElseIf KeyAscii = 57 Then
   If Lorenzo <= 4 And Lorenzo > 0 Then
      pressequaltrue = pressequaltrue + 1
   End If

   If equalpress = 1 Then
      txtdisplay = ""
      equalpress = 0
   End If
   txtdisplay = txtdisplay + "9"

ElseIf KeyAscii = 42 Then
   Enrico = Val(txtdisplay)
   txtdisplay = ""
   Lorenzo = 3
   dot = 0
   pressequaltrue = 1

ElseIf KeyAscii = 43 Then
   Enrico = Val(txtdisplay)
   txtdisplay = ""
   Lorenzo = 1
   dot = 0
   pressequaltrue = 1

ElseIf KeyAscii = 45 Then
   Enrico = Val(txtdisplay)
   txtdisplay = ""
   Lorenzo = 2
   dot = 0
   pressequaltrue = 1

ElseIf KeyAscii = 47 Then
   Enrico = Val(txtdisplay)
   txtdisplay = ""
   Lorenzo = 4
   dot = 0
   pressequaltrue = 1

ElseIf KeyAscii = 13 Then
  If pressequaltrue > 1 Then
   If Lorenzo = 1 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico + valpercent
        
      Else
         txtdisplay = Enrico + Val(txtdisplay)
      End If
   ElseIf Lorenzo = 2 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico - valpercent
               
      Else
         txtdisplay = Enrico - Val(txtdisplay)
      End If
   ElseIf Lorenzo = 3 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico * valpercent
      
      Else
         txtdisplay = Enrico * Val(txtdisplay)
      End If
   ElseIf Lorenzo = 4 Then
      If PERCENT = 1 Then
         txtdisplay = Enrico / valpercent
         
      Else
         txtdisplay = Enrico / Val(txtdisplay)
      End If
   End If
   Lorenzo = 0
   pressequaltrue = 0
   dot = 0
   PERCENT = 0
   End If
   equalpress = 1

ElseIf KeyAscii = 46 Then
   If txtdisplay = "" Or equalpress = 1 Then
      txtdisplay = "0"
   End If
   
   If dot = 0 Then
      txtdisplay = txtdisplay + "."
      dot = 1
   End If
   
   equalpress = 0

   txtdisplay.SetFocus

ElseIf KeyAscii = 83 Then
    If Not Val(txtdisplay) = 0 Then
        MEM = txtdisplay
        lblMemory.Caption = "M"
    End If
   

ElseIf KeyAscii = 82 Then
   txtdisplay = MEM
   equalpress = 1
   dot = 0

ElseIf KeyAscii = 67 Then
   lblMemory = ""
   MEM = ""

ElseIf KeyAscii = 64 Then
   If equalpress = 0 Then
        txtdisplay = Val(txtdisplay) * Val(txtdisplay)
        equalpress = 1
        dot = 0
   End If

ElseIf KeyAscii = 37 Then
   If pressequaltrue > 1 Then
      txtdisplay = (Val(txtdisplay) / 100) * Enrico
      valpercent = Val(txtdisplay)
      PERCENT = 1
   End If

ElseIf KeyAscii = 63 Then
   If equalpress = 0 Then
      txtdisplay = Val(txtdisplay) - 2 * Val(txtdisplay)
      dot = 0
      If Val(txtdisplay) = 0 Then
         equalpress = 1
      End If
   End If

ElseIf KeyAscii = 27 Then
   txtdisplay = ""
   equalpress = 1
   Lorenzo = 0
   dot = 0
   PERCENT = 0
End If

End Sub
