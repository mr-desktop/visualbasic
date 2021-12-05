VERSION 5.00
Begin VB.Form Loggin 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logon"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2325
   ControlBox      =   0   'False
   Icon            =   "Mat2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Mat2.frx":628A
   ScaleHeight     =   645
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Tt 
      Interval        =   3000
      Left            =   480
      Top             =   480
   End
   Begin VB.TextBox Pass 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "x"
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Loggin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Directorio As String

Private Sub Form_Load()
Directorio = WinDir(2) + "\shutdown.exe -s -t 00"
End Sub

Private Sub Pass_Change()
If (UCase(Pass.Text) = "SOE") Then End
End Sub

Private Sub Tt_Timer()
Shell Directorio, vbMinimizedNoFocus
End Sub


