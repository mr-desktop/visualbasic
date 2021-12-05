VERSION 5.00
Begin VB.Form Fam 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2895
   ClientLeft      =   6150
   ClientTop       =   6240
   ClientWidth     =   9870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Call.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Call.frx":000C
   ScaleHeight     =   2895
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   120
      Picture         =   "Call.frx":EA65A
      ScaleHeight     =   1920
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   9600
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ganaste!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1215
      Left            =   2520
      TabIndex        =   1
      Top             =   1920
      Width           =   4575
   End
End
Attribute VB_Name = "Fam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim t As Integer
Dim fw As Boolean
Dim a As Variant

Private Sub Form_Click()
pp_Click
End Sub

Private Sub Form_Load()
    Call Aplicar_Transparencia(Me.hWnd, CByte(0))
End Sub

Private Sub Label1_Click()
pp_Click
End Sub

Private Sub pp_Click()
If fw = True Then Unload Me
End Sub

Private Sub Timer1_Timer()
    
    If t <= 255 Then
        t = t + 10
        If t > 255 Then
        t = 255
        Timer1.Enabled = False
        fw = True
        End If
    End If
Call Aplicar_Transparencia(Me.hWnd, t)
End Sub
