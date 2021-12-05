VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   ClientHeight    =   4020
   ClientLeft      =   1020
   ClientTop       =   615
   ClientWidth     =   4860
   DrawStyle       =   5  'Transparent
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   268
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   Begin VB.PictureBox picScroll 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   2535
      Left            =   360
      ScaleHeight     =   165
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   1
      Top             =   720
      Width           =   4095
   End
   Begin VB.Label lblMail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail Me !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      Caption         =   "Stop"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   3360
      Width           =   330
   End
   Begin VB.Image imgTitleHelp 
      Height          =   195
      Left            =   7320
      Picture         =   "frmAbout.frx":000C
      Top             =   0
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleMaximize 
      Height          =   195
      Left            =   7680
      Picture         =   "frmAbout.frx":0256
      Top             =   360
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleRestore 
      Height          =   195
      Left            =   7680
      Picture         =   "frmAbout.frx":0615
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
      Left            =   240
      MousePointer    =   9  'Size W E
      Top             =   120
      Width           =   75
      Visible         =   0   'False
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator - about"
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
      Width           =   1695
   End
   Begin VB.Image imgTitleMinimize 
      Height          =   195
      Left            =   7320
      Picture         =   "frmAbout.frx":085F
      Top             =   480
      Width           =   195
      Visible         =   0   'False
   End
   Begin VB.Image imgTitleClose 
      Height          =   195
      Left            =   7320
      Picture         =   "frmAbout.frx":0AA9
      Top             =   240
      Width           =   195
   End
   Begin VB.Image imgTitleLeft 
      Height          =   450
      Left            =   5880
      Picture         =   "frmAbout.frx":0CF3
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleRight 
      Height          =   450
      Left            =   6240
      Picture         =   "frmAbout.frx":143D
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "frmAbout.frx":1B87
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgWindowBottomRight 
      Height          =   450
      Left            =   6960
      Picture         =   "frmAbout.frx":22D1
      Top             =   0
      Width           =   285
   End
   Begin VB.Image imgTitleMain 
      Height          =   450
      Left            =   5880
      Picture         =   "frmAbout.frx":2A1B
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowBottom 
      Height          =   450
      Left            =   6240
      Picture         =   "frmAbout.frx":3165
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowLeft 
      Height          =   450
      Left            =   6600
      Picture         =   "frmAbout.frx":38AF
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
   Begin VB.Image imgWindowRight 
      Height          =   450
      Left            =   6960
      Picture         =   "frmAbout.frx":3FF9
      Stretch         =   -1  'True
      Top             =   480
      Width           =   285
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Temp
Dim OldCursorPos As PointAPI
Dim NewCursorPos As PointAPI
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Const DT_BOTTOM As Long = &H8
Const DT_CALCRECT As Long = &H400
Const DT_CENTER As Long = &H1
Const DT_EXPANDTABS As Long = &H40
Const DT_EXTERNALLEADING As Long = &H200
Const DT_LEFT As Long = &H0
Const DT_NOCLIP As Long = &H100
Const DT_NOPREFIX As Long = &H800
Const DT_RIGHT As Long = &H2
Const DT_SINGLELINE As Long = &H20
Const DT_TABSTOP As Long = &H80
Const DT_TOP As Long = &H0
Const DT_VCENTER As Long = &H4
Const DT_WORDBREAK As Long = &H10
Const conSwNormal = 1

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Const ScrollText As String = "Calculator" & vbCrLf & _
                                vbCrLf & vbCrLf & _
                                vbCrLf & vbCrLf & _
                                "Do you miss any functionality ???" & _
                                 vbCrLf & vbCrLf & _
                               vbCrLf & "Send me a mail !!!" & _
                                vbCrLf & "d1001.johansson@swipnet.se" & _
                                vbCrLf & vbCrLf & _
                                vbCrLf & "Thanks to: " & vbCrLf & _
                                vbCrLf & "Herman Lui, Basic math functions" & vbCrLf & _
                                vbCrLf & "Robert Wright, Skin form" & vbCrLf & _
                                vbCrLf & "Production and idea:" & vbCrLf & _
                                "Björn Johansson" & vbCrLf & _
                                vbCrLf & vbCrLf & _
                                vbCrLf & "Muuuuuuuuuuu "
                             
Dim EndingFlag As Boolean
Private Type PointAPI
    X As Long
    Y As Long
End Type

Private Sub Form_Load()
'Make "the Form"
    MakeWindow Me, True
    AlwaysOnTop Me, True
    
    picScroll.ForeColor = vbYellow
    picScroll.FontSize = 10
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

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblExit.ForeColor = vbWhite
End Sub

Private Sub lblMail_Click()
    ShellExecute hwnd, "open", "mailto:d1001.johansson@swipnet.se", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub lblMail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMail.ForeColor = vbWhite
End Sub

Private Sub lblMail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblMail.ForeColor = vbBlue
End Sub

Private Sub lblTitle_DblClick()
    ChangeState Me
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoDrag Me
End Sub

Private Sub picScroll_Click()
    ShellExecute hwnd, "open", "mailto:d1001.johansson@swipnet.se", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub imgTitleMain_DblClick()
    ChangeState Me
End Sub

Private Sub Form_Activate()
    RunMain
End Sub

Private Sub RunMain()
Dim LastFrameTime As Long
Const IntervalTime As Long = 40
Dim rt As Long
Dim DrawingRect As RECT
Dim UpperX As Long, UpperY As Long 'Upper left point of drawing rect
Dim RectHeight As Long

'show the form
frmAbout.Refresh

'Get the size of the drawing rectangle by suppying the DT_CALCRECT constant
rt = DrawText(picScroll.hdc, ScrollText, -1, DrawingRect, DT_CALCRECT)

If rt = 0 Then 'err
    MsgBox "Error scrolling text", vbExclamation
    EndingFlag = True
Else
    DrawingRect.Top = picScroll.ScaleHeight
    DrawingRect.Left = 0
    DrawingRect.Right = picScroll.ScaleWidth
    'Store the height of The rect
    RectHeight = DrawingRect.Bottom
    DrawingRect.Bottom = DrawingRect.Bottom + picScroll.ScaleHeight
End If


Do While Not EndingFlag
    
    If GetTickCount() - LastFrameTime > IntervalTime Then
                    
        picScroll.Cls
        
        DrawText picScroll.hdc, ScrollText, -1, DrawingRect, DT_CENTER Or DT_WORDBREAK
        
        'update the coordinates of the rectangle
        DrawingRect.Top = DrawingRect.Top - 1
        DrawingRect.Bottom = DrawingRect.Bottom - 1
        
        'control the scolling and reset if it goes out of bounds
        If DrawingRect.Top < -(RectHeight) Then 'time to reset
            DrawingRect.Top = picScroll.ScaleHeight
            DrawingRect.Bottom = RectHeight + picScroll.ScaleHeight
        End If
        
        picScroll.Refresh
        
        LastFrameTime = GetTickCount()
        
    End If
    
    DoEvents
Loop

Unload Me
Set frmAbout = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    EndingFlag = True
End Sub

Private Sub lblExit_Click()
    EndingFlag = True
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblExit.ForeColor = vbRed
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

