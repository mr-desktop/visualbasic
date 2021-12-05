VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3855
   ClientLeft      =   6360
   ClientTop       =   390
   ClientWidth     =   5595
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "Show this form at Startup"
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSplash.frx":0000
      Height          =   1815
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.Image imgLogo 
      Height          =   2385
      Left            =   360
      Picture         =   "frmSplash.frx":014F
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1815
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   7320
      Top             =   0
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   7680
      Picture         =   "frmSplash.frx":DA49
      Top             =   360
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   7680
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
      Caption         =   "Calculator - Splash"
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
      Width           =   1800
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   7320
      Top             =   480
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   7320
      Picture         =   "frmSplash.frx":DE08
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   5880
      Picture         =   "frmSplash.frx":E04F
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   6240
      Picture         =   "frmSplash.frx":E4EA
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "frmSplash.frx":E948
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   6960
      Picture         =   "frmSplash.frx":EBC4
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5880
      Picture         =   "frmSplash.frx":EE51
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   6240
      Picture         =   "frmSplash.frx":EF0A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "frmSplash.frx":EFA2
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   6960
      Picture         =   "frmSplash.frx":F030
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Temp
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI
Private Sub chkLoadTipsAtStartup_Click()
    ' save whether or not this form should be displayed at startup
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub cmdOK_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Dim ShowAtStartup As Long
    'Make "the Form"
    MakeWindow Me, True
    AlwaysOnTop Me, True

    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 1)
    If ShowAtStartup = 0 Then
        frmMain.Show
        Me.Hide
        Exit Sub
    End If
        
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    chkLoadTipsAtStartup.BackColor = RGB(207, 207, 207)
    cmdOK.BackColor = RGB(207, 207, 207)

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



