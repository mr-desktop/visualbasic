VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Apagar monitor"
   ClientHeight    =   1260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3930
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   3930
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Apagar"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Api SendMessage
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As _
        Long, ByVal wParam As Long, lParam As Any) As Long

'Constantes para SendMessage
Const WM_SYSCOMMAND = &H112&
Const SC_MONITORPOWER = &HF170&

Private Sub Command1_Click()
    
    
    If MsgBox("Pagar el monitor por 15 segundos sin mover el raton ?", vbQuestion) = vbNo Then Exit Sub
        
        Timer1.Enabled = True
        'Apaga el monitor
        SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal 2&
    
End Sub

Private Sub Form_Load()
    '15 segundos de lapso
    Timer1.Interval = 15000
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    
    Timer1.Enabled = False
    'Prende el monitor
    SendMessage Me.hwnd, WM_SYSCOMMAND, SC_MONITORPOWER, ByVal -1&
End Sub

