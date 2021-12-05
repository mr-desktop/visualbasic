VERSION 5.00
Begin VB.Form frmAlarm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2265
   ClientLeft      =   675
   ClientTop       =   3825
   ClientWidth     =   5250
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "frmAlarm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5250
   Begin VB.Timer tmrAlarm 
      Left            =   4440
      Top             =   1680
   End
   Begin VB.PictureBox Picture2 
      Height          =   1815
      Left            =   240
      Picture         =   "frmAlarm.frx":030A
      ScaleHeight     =   1755
      ScaleWidth      =   1875
      TabIndex        =   2
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Yes"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblNalle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Time to let go !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Dim SYNC As Long
Dim wav As String

Private Sub cmdOK_Click()
    tmrAlarm.Interval = 0
    Unload Me
End Sub

Private Sub Form_Activate()
tmrAlarm.Interval = 1000
End Sub

Private Sub Picture2_Click()
    Unload Me
End Sub

Private Sub tmrAlarm_Timer()
Dim SYNC As Long
Dim R As Integer
Dim wav As String
If frmMain.tmrTime.Tag = 1 Then
SYNC = SND_ASYNC
wav = App.Path & "/" & "sound.wav"
R = sndPlaySound(ByVal wav, SYNC)
Else
    frmAlarm.Caption = "Time to wake up !!"
End If
End Sub
