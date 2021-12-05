VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keyboard Shortcut"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Cal23.frx":0000
   ScaleHeight     =   2415
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&OK"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   900
      ItemData        =   "Cal23.frx":40324
      Left            =   240
      List            =   "Cal23.frx":4036D
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   300
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Enabled = True
Unload Me
End Sub

Private Sub Label5_Click()

End Sub

