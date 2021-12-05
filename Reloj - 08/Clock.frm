VERSION 5.00
Begin VB.Form reloj 
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reloj Analogo"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7845
   Icon            =   "Clock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7845
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tt2 
      Interval        =   75
      Left            =   5640
      Top             =   4440
   End
   Begin VB.Timer tt1 
      Interval        =   1000
      Left            =   1560
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      Caption         =   "Comenzar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Line sec 
      BorderColor     =   &H00404000&
      X1              =   1800
      X2              =   1815
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Line min 
      BorderColor     =   &H00000040&
      BorderWidth     =   2
      X1              =   1800
      X2              =   1815
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Label cro 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label td 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "88:88:88"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   4680
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Line l2 
      BorderWidth     =   5
      X1              =   5895
      X2              =   5880
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line l1 
      BorderColor     =   &H00400040&
      BorderWidth     =   3
      X1              =   1800
      X2              =   1815
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Image Image4 
      Height          =   1995
      Left            =   4080
      Picture         =   "Clock.frx":6852
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3480
   End
   Begin VB.Image Image3 
      Height          =   1995
      Left            =   240
      Picture         =   "Clock.frx":7081
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   3480
   End
   Begin VB.Image Cra 
      Height          =   3000
      Left            =   4320
      Picture         =   "Clock.frx":B6BC
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   270
      Picture         =   "Clock.frx":17768
      Stretch         =   -1  'True
      Top             =   195
      Width           =   3000
   End
   Begin VB.Menu Hlp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "reloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim h, m, s, x, y, x2, y2, x3, y3 As Integer
Dim crm, crs, crc, beta As Integer
Dim sg As String

Const pix = 3.1415927 / 180

Private Sub Command1_Click()
beta = beta + 1
Command1.Caption = "Parar"
If (beta = 2) Then
beta = 0
Command1.Caption = "Seguir"
End If
End Sub


Private Sub cro_Click()
beta = 0
crc = 0
crs = 0
crm = 0
End Sub

Private Sub Form_Load()
td.Caption = Time
h = Hour(Time)
m = Minute(Time)
s = Second(Time)
crc = 0
crs = 0
crm = 0

l1.X1 = l1.x2 + Sin(h * 28 * pix) * 1000
l1.Y1 = l1.y2 - Cos(h * 28 * pix) * 1000
Min.X1 = Min.x2 + Sin(m * 6 * pix) * 1050
Min.Y1 = Min.y2 - Cos(m * 6 * pix) * 1050
sec.X1 = sec.x2 + Sin(s * 6 * pix) * 1200
sec.Y1 = sec.y2 - Cos(s * 6 * pix) * 1200
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Hlp_Click()
Help.Show
End Sub

Private Sub tt1_Timer()
td.Caption = Time
h = Hour(Time)
m = Minute(Time)
s = Second(Time)

l1.X1 = l1.x2 + Sin(h * 28 * pix) * 1000
l1.Y1 = l1.y2 - Cos(h * 28 * pix) * 1000
Min.X1 = Min.x2 + Sin(m * 6 * pix) * 1050
Min.Y1 = Min.y2 - Cos(m * 6 * pix) * 1050
sec.X1 = sec.x2 + Sin(s * 6 * pix) * 1200
sec.Y1 = sec.y2 - Cos(s * 6 * pix) * 1200

End Sub

Private Sub tt2_Timer()
crc = crc + beta

If (crc > 9) Then
crs = crs + beta
crc = 0
End If

If (crs > 59) Then
crm = crm + beta
crs = 0
End If

If (crm > 59) Then
crm = 0
End If

sg = Trim(Str(crm)) + ":" + Trim(Str(crs)) + ":" + Trim(Str(crc))
cro.Caption = sg
l2.X1 = l2.x2 + Sin(crc * 37 * pix) * 1000
l2.Y1 = l2.y2 - Cos(crc * 37 * pix) * 1000
End Sub
