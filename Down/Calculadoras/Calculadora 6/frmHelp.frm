VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   0  'None
   ClientHeight    =   5790
   ClientLeft      =   1260
   ClientTop       =   150
   ClientWidth     =   4425
   Icon            =   "frmHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   295
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print help"
      Height          =   225
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   1500
   End
   Begin VB.CommandButton cmdTillbaka 
      Caption         =   "Back to Calculator"
      Height          =   225
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Print Receipt"
      Height          =   255
      Index           =   15
      Left            =   1320
      TabIndex        =   34
      Top             =   4440
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Alarm"
      Height          =   255
      Index           =   9
      Left            =   1320
      TabIndex        =   33
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "P or R"
      Height          =   255
      Index           =   15
      Left            =   360
      TabIndex        =   32
      Top             =   4440
      Width           =   525
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "A or S"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   31
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F1"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   30
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F2 or Del"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   29
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F3"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   28
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F4 or Home"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   27
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F5 or M"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   26
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F6"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   25
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   24
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F7"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   23
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      Height          =   255
      Index           =   0
      Left            =   1320
      TabIndex        =   22
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "C -  button"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   20
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "CE - button"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   19
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "% - button"
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   18
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Monitor"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   17
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "No monitor"
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   16
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Readout and receipt"
      Height          =   255
      Index           =   7
      Left            =   1320
      TabIndex        =   15
      Top             =   2520
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F8"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   14
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt only"
      Height          =   255
      Index           =   8
      Left            =   1320
      TabIndex        =   13
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "Esc or End"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   12
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Height          =   255
      Index           =   10
      Left            =   1320
      TabIndex        =   11
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "To Clipboard if you want to use it in another application"
      Height          =   495
      Index           =   11
      Left            =   1320
      TabIndex        =   10
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F9"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F10"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   8
      Top             =   3480
      Width           =   405
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F11"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   405
   End
   Begin VB.Label lblFKnapp 
      BackStyle       =   0  'Transparent
      Caption         =   "F12"
      Height          =   255
      Index           =   14
      Left            =   360
      TabIndex        =   6
      Top             =   3960
      Width           =   405
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      Height          =   255
      Index           =   12
      Left            =   1320
      TabIndex        =   5
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Swedish VAT ??"
      Height          =   255
      Index           =   13
      Left            =   1320
      TabIndex        =   4
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label lblFText 
      BackStyle       =   0  'Transparent
      Caption         =   "Reduce Swedish VAT ??"
      Height          =   255
      Index           =   14
      Left            =   1320
      TabIndex        =   3
      Top             =   3960
      Width           =   3015
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   7320
      Picture         =   "frmHelp.frx":030A
      Top             =   0
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   7680
      Picture         =   "frmHelp.frx":0554
      Top             =   360
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   7680
      Picture         =   "frmHelp.frx":0913
      Top             =   720
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   7320
      Top             =   720
      Width           =   195
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   7
      Left            =   2400
      MousePointer    =   8  'Size NW SE
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   6
      Left            =   2160
      MousePointer    =   6  'Size NE SW
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   5
      Left            =   1920
      MousePointer    =   6  'Size NE SW
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   4
      Left            =   1680
      MousePointer    =   8  'Size NW SE
      Top             =   120
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   3
      Left            =   720
      MousePointer    =   7  'Size N S
      Top             =   360
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   2
      Left            =   720
      MousePointer    =   7  'Size N S
      Top             =   120
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   1
      Left            =   480
      MousePointer    =   9  'Size W E
      Top             =   120
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   0
      Left            =   360
      MousePointer    =   9  'Size W E
      Top             =   120
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator - help"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   7320
      Picture         =   "frmHelp.frx":0B5D
      Top             =   480
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   7320
      Picture         =   "frmHelp.frx":0DA7
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   5880
      Picture         =   "frmHelp.frx":0FF1
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   6240
      Picture         =   "frmHelp.frx":173B
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "frmHelp.frx":1E85
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   6960
      Picture         =   "frmHelp.frx":25CF
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5880
      Picture         =   "frmHelp.frx":2D19
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   6240
      Picture         =   "frmHelp.frx":3463
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "frmHelp.frx":3BAD
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   6960
      Picture         =   "frmHelp.frx":42F7
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI

Private Sub cmdPrint_Click()
Printer.FontName = "Arial"
Printer.FontSize = 14: Printer.FontBold = True
Printer.CurrentX = 1: Printer.CurrentY = 1
Printer.Print "Help for person with Alzheimer Light"
Printer.FontSize = 8
Printer.CurrentX = 1
Printer.Print "F1                Help"
Printer.Print "F2 or Del     C-Button"
Printer.Print "F3                CE-Button"
Printer.Print "F4 or Home  %-Button"
Printer.Print "F5                Monitor"
Printer.Print "F6                No monitor"
Printer.Print "F7                Readout and receipt"
Printer.Print "F8                Receipt only"
Printer.Print "F9                Copy to clipboard"
Printer.Print "F10               Show abouform"
Printer.Print "F11               Add swedish VAT 25 %"
Printer.Print "F12               Reduce swedish VAT 20 %"
Printer.Print "A or S          Set Alarm"
Printer.Print "P or R          Print Receipt"
Printer.Print "Esc or End  Exit application"
End Sub

Private Sub cmdTillbaka_Click()
    imgTitleClose_Click
End Sub

Private Sub Form_Load()
'Make "the Form"
    MakeWindow Me, True
    AlwaysOnTop Me, True
End Sub

Private Sub imgTitleClose_Click()
    Unload Me
End Sub

Private Sub imgTitleHelp_Click()
    Unload Me
End Sub

Private Sub imgTitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMaxRestore_Click()
    ChangeState Me
End Sub

Private Sub imgTitleMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub imgTitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub lblTitle_DblClick()
    ChangeState Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub Resizer_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Temp = GetCursorPos(OldCursorPos)
End Sub

Private Sub Resizer_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Temp = GetCursorPos(NewCursorPos)
    ResizeForm Me, OldCursorPos, NewCursorPos, Index
End Sub

Private Sub imgTitleMain_DblClick()
    ChangeState Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyDelete
    imgTitleClose_Click
Case vbKeyEscape
    imgTitleClose_Click
Case vbKeyEnd
    imgTitleClose_Click
End Select
End Sub



