VERSION 5.00
Begin VB.Form frmSetAlarm 
   BorderStyle     =   0  'None
   ClientHeight    =   2700
   ClientLeft      =   3060
   ClientTop       =   705
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   Begin VB.CheckBox chkSound 
      Caption         =   "Sound"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.CheckBox chkMessage 
      Caption         =   "Message"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Set"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   675
   End
   Begin VB.Frame frmSetTime 
      Height          =   1455
      Left            =   360
      TabIndex        =   6
      Top             =   600
      Width           =   1455
      Begin VB.TextBox txtSecond 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   3
         Top             =   960
         Width           =   400
      End
      Begin VB.TextBox txtHour 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   1
         Top             =   240
         Width           =   400
      End
      Begin VB.TextBox txtMinute 
         Height          =   285
         Left            =   840
         MaxLength       =   2
         TabIndex        =   2
         Top             =   600
         Width           =   400
      End
      Begin VB.Label lblSec 
         BackStyle       =   0  'Transparent
         Caption         =   "Second"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblHour 
         BackStyle       =   0  'Transparent
         Caption         =   "Hour"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblMinute 
         BackStyle       =   0  'Transparent
         Caption         =   "Minute"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2160
      Width           =   675
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   5520
      Picture         =   "frmSetAlarm.frx":0000
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleMaxRestore 
      Height          =   195
      Left            =   5880
      Top             =   480
      Width           =   195
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   0
      Left            =   0
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   375
      Index           =   1
      Left            =   240
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   2
      Left            =   480
      MousePointer    =   7  'Size N S
      Top             =   0
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   75
      Index           =   3
      Left            =   480
      MousePointer    =   7  'Size N S
      Top             =   240
      Width           =   855
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   4
      Left            =   1440
      MousePointer    =   8  'Size NW SE
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   5
      Left            =   1680
      MousePointer    =   6  'Size NE SW
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   6
      Left            =   1920
      MousePointer    =   6  'Size NE SW
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image Resizer 
      Height          =   165
      Index           =   7
      Left            =   2160
      MousePointer    =   8  'Size NW SE
      Top             =   0
      Width           =   165
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   5880
      Picture         =   "frmSetAlarm.frx":024A
      Top             =   120
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alarmsettings"
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
      Left            =   3960
      TabIndex        =   0
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   5520
      Picture         =   "frmSetAlarm.frx":0609
      Top             =   840
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   5520
      Picture         =   "frmSetAlarm.frx":0841
      Top             =   480
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   4080
      Picture         =   "frmSetAlarm.frx":0A88
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   4440
      Picture         =   "frmSetAlarm.frx":0F23
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   4800
      Picture         =   "frmSetAlarm.frx":1381
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   5160
      Picture         =   "frmSetAlarm.frx":15FD
      Top             =   120
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   4080
      Picture         =   "frmSetAlarm.frx":188A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   4440
      Picture         =   "frmSetAlarm.frx":1943
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   4800
      Picture         =   "frmSetAlarm.frx":19DB
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   5160
      Picture         =   "frmSetAlarm.frx":1A69
      Stretch         =   -1  'True
      Top             =   600
      Width           =   285
   End
End
Attribute VB_Name = "frmSetAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkMessage_Click()
' Tagging it for the timer to catch
    frmMain.tmrTime.Tag = 0
End Sub

Private Sub chkSound_Click()
' Tagging it for the timer to catch
    frmMain.tmrTime.Tag = 1
End Sub
Private Sub Form_Activate()
    txtHour.SetFocus
End Sub

Private Sub Form_Load()
    MakeWindow Me, True
    AlwaysOnTop Me, True
    txtHour.Text = "00"
    txtMinute.Text = "00"
    txtSecond.Text = "00"
    chkMessage.Value = 1
    frmSetTime.BackColor = RGB(207, 207, 207)
    chkMessage.BackColor = RGB(207, 207, 207)
    chkSound.BackColor = RGB(207, 207, 207)
    OKButton.BackColor = RGB(207, 207, 207)
    cmdClear.BackColor = RGB(207, 207, 207)
End Sub

Private Sub imgTitleClose_Click()
    Unload Me
End Sub

Private Sub imgTitleHelp_Click()
    MsgBox "Do you realy need help with this ??"
End Sub

Private Sub imgTitleLeft_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub imgTitleRight_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub
Private Sub cmdClear_Click()
' Clears all the text boxes
    txtHour.Text = "00"
    txtMinute.Text = "00"
    txtSecond.Text = "00"
    chkMessage.Value = 0
    chkSound.Value = 0
    frmMain.lblAlarm.Caption = "Off"
    End Sub

Private Sub OKButton_Click()
' Sets the alarm time
 If txtHour.Text = "" Then
    txtHour.SetFocus
    Exit Sub
    Else
    frmMain.Alarmtime = txtHour.Text + ":" + txtMinute.Text + ":" + txtSecond.Text
    frmMain.lblAlarm.Caption = frmMain.Alarmtime
    Unload Me
  End If
End Sub

Private Sub txtHour_Change()
' Checking the data range of the text box
If txtHour.Text = "" Then Exit Sub
    If txtHour.Text > 24 Or txtHour.Text < 0 Then ' The data range
       ' If txtHour.Text = 0 Then txtHour.Text = 12 ' 0 hours is the same as 12 hours
        MsgBox "Digits between 1 and 24 allowed", 16, "A Calc"
        txtHour.Text = "00"
        txtHour.SelStart = 0
        txtHour.SelLength = Len(txtHour)
        txtHour.SetFocus
    End If
End Sub
Private Sub txtHour_GotFocus()
    txtHour.SelLength = Len(txtHour)
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
' More data validation, this time whenever a key is pressed
If KeyAscii = vbKeyReturn Then
txtMinute.SetFocus
  Exit Sub
End If
If KeyAscii = vbKeyBack Then ' Eliminating the backspace key from the list
      Exit Sub
End If
   If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then ' Allowing only numbers into the field
        KeyAscii = 0 ' Re-assigning the ascii value of the key pressed
    End If
End Sub

Private Sub txtMinute_Change()
If txtMinute.Text = "" Then Exit Sub
' Checking the data range of the text box
    If txtMinute.Text > 60 Or txtMinute.Text < 0 Then
        MsgBox "Use digits between 0 och 60", 16, "A Calc"
        txtMinute.Text = "00"
        txtMinute.SelStart = 0
        txtMinute.SelLength = Len(txtMinute)
        txtMinute.SetFocus
    End If
End Sub

Private Sub txtMinute_GotFocus()
    txtMinute.SelLength = Len(txtMinute)
End Sub

Private Sub txtMinute_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtSecond.SetFocus
    Exit Sub
    End If
' More data validation
    If KeyAscii = vbKeyBack Then
        Exit Sub
    End If
    If KeyAscii < vbKey0 Or KeyAscii > vbKey9 Then
            KeyAscii = 0
    End If
End Sub

Private Sub txtSecond_Change()
    If txtSecond.Text = "" Then Exit Sub
    If txtSecond.Text > 60 Or txtMinute.Text < 0 Then
        MsgBox "Use digits between 0 och 60", 16, "A Calc"
        txtSecond.Text = "00"
        txtSecond.SelStart = 0
        txtSecond.SelLength = Len(txtSecond)
        txtSecond.SetFocus
    End If
End Sub
Private Sub txtSecond_GotFocus()
    txtSecond.SelLength = Len(txtSecond)
End Sub

Private Sub txtSecond_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
OKButton.SetFocus
  Exit Sub
End If

End Sub
