VERSION 5.00
Begin VB.Form Fam 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1365
   ClientLeft      =   6150
   ClientTop       =   6240
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Call.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox pp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   0
      Picture         =   "Call.frx":57E2
      ScaleHeight     =   1500
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   -120
      Width           =   7500
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

Private Sub Form_Load()
    Call Aplicar_Transparencia(Me.hWnd, CByte(0))
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
