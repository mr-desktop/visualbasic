VERSION 5.00
Begin VB.Form Help2 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "J.e.u.s"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7875
   Icon            =   "Help2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   3480
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "x"
      TabIndex        =   1
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Help2.frx":0E42
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   840
      Picture         =   "Help2.frx":0E73
      Top             =   360
      Width           =   2400
   End
   Begin VB.Label Cnc 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      Top             =   3480
      Width           =   735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2760
      Left            =   3960
      Picture         =   "Help2.frx":4637
      Top             =   360
      Width           =   3660
   End
End
Attribute VB_Name = "Help2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cnc_Click()
'bloquea teclado
BlockInput True
End Sub

Private Sub Form_Activate()
'ocultar en task manager
App.TaskVisible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
BlockInput True
Help.Show
End Sub

Private Sub Text2_Change()
If (UCase(Text2.Text) = "SOE") Then End
End Sub

Private Sub Timer1_Timer()
    Static ct As Integer
    ct = ct + 1
    If ct = 30 Then
      BlockInput False
      ct = 0
    End If
End Sub
