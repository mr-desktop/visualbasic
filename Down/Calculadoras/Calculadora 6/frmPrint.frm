VERSION 5.00
Begin VB.Form frmPrint 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   15000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   15000
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.ListBox lstKvitto 
      Height          =   14880
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   2700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RECEIPT"
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
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Wrightadjustment listbox
Private Const LB_SETTABSTOPS = &H192
Private Declare Function SendMessage Lib "user32" _
Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Must be a long or integer array
Dim mTabs(0) As Long

Private Sub Form_Load()
'Alignment wright
'Set the a tab stop to a negative number
    mTabs(0) = -100
End Sub
