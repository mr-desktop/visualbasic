VERSION 5.00
Begin VB.Form Help 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "J.e.u.s"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4230
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   4230
   StartUpPosition =   1  'CenterOwner
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
      Height          =   1215
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Help.frx":0E42
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   2700
      Left            =   240
      Picture         =   "Help.frx":0E79
      Top             =   0
      Width           =   3600
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
Me.Hide
End Sub
