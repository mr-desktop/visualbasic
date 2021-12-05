VERSION 5.00
Begin VB.Form Jeus 
   BackColor       =   &H00800080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hc"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10635
   Icon            =   "Hc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Hc.frx":000C
   ScaleHeight     =   9495
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Inferno ® Bloqueo Versión 5.2 Copyright © 2006-2007"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   2295
      Left            =   5640
      TabIndex        =   0
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   3840
      Left            =   6720
      Picture         =   "Hc.frx":CFB1
      Top             =   -360
      Width           =   3840
   End
   Begin VB.Image Image2 
      Height          =   8250
      Left            =   600
      Picture         =   "Hc.frx":12BDE
      Top             =   600
      Width           =   9300
   End
End
Attribute VB_Name = "Jeus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
Jeus.Hide
End Sub

Private Sub Image3_Click()
Jeus.Hide
End Sub
