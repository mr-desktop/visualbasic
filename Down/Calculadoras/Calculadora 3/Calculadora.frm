VERSION 5.00
Begin VB.Form Calculadora 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   -60
   ClientWidth     =   2070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Calculadora.frx":0000
   ScaleHeight     =   3945
   ScaleWidth      =   2070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command23 
      BackColor       =   &H00000080&
      Caption         =   "Off"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H0080C0FF&
      Caption         =   "!x"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H0080C0FF&
      Caption         =   "In"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H0080C0FF&
      Caption         =   "-x"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H008080FF&
      Caption         =   "Pi"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2520
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00FF8080&
      Caption         =   "x^2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00FF8080&
      Caption         =   "Sqr"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0FFC0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0C0FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0C0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2040
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
End
Attribute VB_Name = "Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nn, bb As Double

Private Sub Command1_Click()
On Error Resume Next
If Label1.Caption = "+" Then
Text1.Text = Val(Label2.Caption) + Text1.Text
End If

If Label1.Caption = "-" Then
Text1.Text = Val(Label2.Caption) - Text1.Text
End If

If Label1.Caption = "*" Then
Text1.Text = Val(Label2.Caption) * Text1.Text
End If

If Label1.Caption = "/" Then
Text1.Text = Val(Label2.Caption) / Text1.Text
End If

End Sub

Private Sub Command10_Click()
Text1.Text = Text1.Text + "9"
End Sub

Private Sub Command11_Click()
Text1.Text = Text1.Text + "8"
End Sub

Private Sub Command12_Click()
Text1.Text = Text1.Text + "7"
End Sub

Private Sub Command13_Click()
Text1.Text = Text1.Text + "6"
End Sub

Private Sub Command14_Click()
Text1.Text = Text1.Text + "5"
End Sub

Private Sub Command15_Click()
Text1.Text = Text1.Text + "0"
End Sub

Private Sub Command16_Click()
Text1.Text = ""
Label2.Caption = ""
End Sub

Private Sub Command17_Click()
Text1.Text = Sqr(Val(Text1.Text))
End Sub

Private Sub Command18_Click()
Text1.Text = Text1.Text * Text1.Text
End Sub

Private Sub Command19_Click()
Text1.Text = Text1.Text + "3.1415926535897932384626433832795"
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text + "4"
End Sub

Private Sub Command20_Click()
Text1.Text = Val(Text1.Text * -1)
End Sub

Private Sub Command21_Click()
Text1.Text = 1 / Val(Text1.Text)
End Sub

Private Sub Command22_Click()
bb = 1
For nn = 1 To Text1.Text
bb = bb * nn
Next
Text1.Text = Val(bb)
End Sub

Private Sub Command23_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next
Label1.Caption = ""
Label1.Caption = "+"
Label2.Caption = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command4_Click()
On Error Resume Next
Label1.Caption = ""
Label1.Caption = "-"
Label2.Caption = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command5_Click()
On Error Resume Next
Label1.Caption = ""
Label1.Caption = "*"
Label2.Caption = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command6_Click()
On Error Resume Next
Label1.Caption = ""
Label1.Caption = "/"
Label2.Caption = Text1.Text
Text1.Text = ""
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text + "3"
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text + "2"
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text + "1"
End Sub
