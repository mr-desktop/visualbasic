VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   3675
   ClientLeft      =   5625
   ClientTop       =   4365
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command16 
      BackColor       =   &H008080FF&
      Caption         =   "Off"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFC0C0&
      Height          =   3375
      ItemData        =   "Form1.frx":26379
      Left            =   3480
      List            =   "Form1.frx":2637B
      Sorted          =   -1  'True
      TabIndex        =   17
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdEquals 
      BackColor       =   &H00C0C0FF&
      Caption         =   "="
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0080C0FF&
      Caption         =   "÷"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0080C0FF&
      Caption         =   "x"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0080C0FF&
      Caption         =   "-"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0080C0FF&
      Caption         =   "+"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Clear"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "9"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "8"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "7"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "6"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "5"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "4"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "3"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "2"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "1"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'www.slurpcandy.co.uk/welcome.htm
'Slurp's Calculator v1.0
'21st October 00
'Author: Mark A. Hill

Private Sub cmdEquals_Click()
Call Equals
End Sub

Private Sub Command1_Click()
Call One
End Sub

Private Sub Command10_Click()
Call zero
End Sub

Private Sub Command11_Click()
Call clear
End Sub

Private Sub Command12_Click()
Call Add
End Sub

Private Sub Command13_Click()
Call Subtract
End Sub

Private Sub Command14_Click()
Call multiply
End Sub

Private Sub Command15_Click()
Call Divide
End Sub

Private Sub Command16_Click()
Call EndIt
End Sub

Private Sub Command2_Click()
Call two
End Sub

Private Sub Command3_Click()
Call three
End Sub

Private Sub Command4_Click()
Call four
End Sub

Private Sub Command5_Click()
Call five
End Sub

Private Sub Command6_Click()
Call six
End Sub

Private Sub Command7_Click()
Call seven
End Sub

Private Sub Command8_Click()
Call eight
End Sub

Private Sub Command9_Click()
Call nine
End Sub

Private Sub Form_Load()
Form1.Label1.RightToLeft = True
End Sub
