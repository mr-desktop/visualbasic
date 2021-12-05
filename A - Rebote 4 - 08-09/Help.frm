VERSION 5.00
Begin VB.Form Config 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "J.e.u.s"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4590
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   3855
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4095
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
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Help.frx":0E42
         Top             =   2520
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   2700
         Left            =   240
         Picture         =   "Help.frx":0E86
         Top             =   240
         Width           =   3600
      End
   End
   Begin VB.ListBox Lis 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   1500
      ItemData        =   "Help.frx":20A0
      Left            =   240
      List            =   "Help.frx":20B9
      TabIndex        =   0
      Top             =   4200
      Width           =   2895
   End
   Begin VB.Label Selec 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label Cas 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Ace 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
End
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objShell As Object

Private Sub Ace_Click()
Set objShell = CreateObject("Wscript.Shell")
On Error Resume Next
objShell.RegWrite Ruta, Str(Sle)
Sle = Val(objShell.RegRead(Ruta))
Set objShell = Nothing

Final
End Sub

Private Sub Cas_Click()
Final
End Sub

Private Sub Form_Activate()
Selec.Caption = Sle
Acti = True
ShowCursor True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Or (KeyAscii = 13) Then Unload Me
End Sub

Private Sub Form_Load()
Acti = True
ShowCursor True
End Sub

Private Sub Form_LostFocus()
Final
End Sub

Private Sub Form_Unload(Cancel As Integer)
Final
End Sub

Private Sub Lis_Click()
Sle = Lis.ListIndex + 1
Selec = Sle
End Sub

Sub Final()
If (Ac = False) Then ShowCursor False
Acti = False
Unload Me
End Sub
