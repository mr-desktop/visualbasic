Attribute VB_Name = "SkinForm"
Option Explicit

' *********************************************
' * Code by Robert Wright - <rob@xenonic.com> *
' *********************************************

' *********************************************
' *        -> This code is FREEWARE <-        *
' * You are free to use this any of this VB   *
' * Project (including the images) in your    *
' * own programs.  A "thanks" in your About   *
' * box would be nice, but you don't have to. *
' *                                           *
' * All I ask is that you vote for me on PSC! *
' *********************************************


' Used to set the shape of the form
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
' Used to create the rounded rectangle region
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' Used to make the form draggable
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
' Also used to make the form draggable
Public Declare Function ReleaseCapture Lib "user32" () As Long
' Used to make the window always on top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, y, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
' Used to get the cursor position
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
' Various bits and pieces used by the above functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type PointAPI
    x As Long
    y As Long
End Type

Dim Resizable As Integer

Public Sub AlwaysOnTop(TheForm As Form, Toggle As Boolean)
' TheForm:  The form you want to make always on top or not
' Toggle:   (True/False) - True for always on top, False for normal
    
    If Toggle = True Then
        SetWindowPos TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    Else
        SetWindowPos TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    End If
End Sub

Public Sub DoDrag(TheForm As Form)
' TheForm:  The form you want to start dragging
    
    If TheForm.WindowState <> vbMaximized Then
        ReleaseCapture
        SendMessage TheForm.hwnd, &HA1, 2, 0&
    End If
End Sub

Public Sub MakeWindow(TheForm As Form, IsResizable As Boolean)
' TheForm:           The form you want to make graphical
' IsResizable:       (True/False) - True for resizable at runtime

' Declare some variables
    Dim FormWidth As Long
    Dim FormHeight As Long
    Dim Temp As Integer

' Set the Resizable variable
    Resizable = IIf(IsResizable = True, 1, 0)
    
' Store the form's width and height in pixels in a variable
    FormWidth = (TheForm.Width / Screen.TwipsPerPixelX)
    FormHeight = (TheForm.Height / Screen.TwipsPerPixelY)
    
' Set various parameters of the form
    TheForm.BackColor = RGB(207, 207, 207)
    TheForm.Caption = TheForm!lblTitle.Caption
    
' Set the position of the title label
    TheForm!lblTitle.Left = 16
    TheForm!lblTitle.Top = 7
    
' Make the form "rounded rectangle" shaped (call to the sub below)
    DoTransparency TheForm
    
' Move the image blocks into place and stretch them accordingly
    With TheForm!imgTitleLeft
        .Top = 0
        .Left = 0
    End With
    
    With TheForm!imgTitleRight
        .Top = 0
        .Left = FormWidth - 19
    End With
    
    With TheForm!imgTitleMain
        .Top = 0
        .Left = 19
        .Width = FormWidth - 19
    End With
    
    With TheForm!imgWindowLeft
        .Top = 30
        .Left = 0
        .Height = FormHeight - 60
    End With
    
    With TheForm!imgWindowBottomLeft
        .Top = FormHeight - 30
        .Left = 0
    End With
    
    With TheForm!imgWindowBottom
        .Top = FormHeight - 30
        .Left = 19
        .Width = FormWidth - 38
    End With
    
    With TheForm!imgWindowBottomRight
        .Top = FormHeight - 30
        .Left = FormWidth - 19
    End With
    
    With TheForm!imgWindowRight
        .Top = 30
        .Left = FormWidth - 19
        .Height = FormHeight - 38
    End With
    
' Position the title buttons (close, minimize, help)
    With TheForm!imgTitleClose
        .Top = 8
        .Left = FormWidth - 22
    End With
    
    With TheForm!imgTitleMaxRestore
        .Top = 8
        .Left = FormWidth - 39
    End With
    
    With TheForm!imgTitleMinimize
        .Top = 8
        .Left = FormWidth - 56
    End With
    
    With TheForm!imgTitleHelp
        .Top = 8
        .Left = FormWidth - 73
    End With
    
' Position the resizing invisible images
    If IsResizable = True Then
        For Temp = 0 To 7
            TheForm!Resizer(Temp).Visible = True
        Next Temp
        
        With TheForm!Resizer(0)
            .Top = 30
            .Left = 0
            .Height = FormHeight - 60
        End With
        
        With TheForm!Resizer(1)
            .Top = 30
            .Left = FormWidth - 5
            .Height = FormHeight - 60
        End With
        
        With TheForm!Resizer(2)
            .Top = 0
            .Left = 19
            .Width = FormWidth - 39
        End With
        
        With TheForm!Resizer(3)
            .Top = FormHeight - 5
            .Left = 19
            .Width = FormWidth - 39
        End With
        
        With TheForm!Resizer(4)
            .Top = FormHeight - 11
            .Left = FormWidth - 11
        End With
        
        With TheForm!Resizer(5)
            .Top = FormHeight - 11
            .Left = 0
        End With
        
        With TheForm!Resizer(6)
            .Top = 0
            .Left = FormWidth - 11
        End With
        
        With TheForm!Resizer(7)
            .Top = 0
            .Left = 0
        End With
    End If

End Sub

Public Sub DoTransparency(TheForm As Form)
' TheForm:  The form you want to be rounded rectangle shape
    
    Dim TempRegions(6) As Long
    Dim FormWidthInPixels As Long
    Dim FormHeightInPixels As Long
    Dim a
    
' Convert the form's height and width from twips to pixels
    FormWidthInPixels = TheForm.Width / Screen.TwipsPerPixelX
    FormHeightInPixels = TheForm.Height / Screen.TwipsPerPixelY
    
' Make a rounded rectangle shaped region with the dimensions of the form
    a = CreateRoundRectRgn(0, 0, FormWidthInPixels, FormHeightInPixels, 24, 24)
    
' Set this region as the shape for "TheForm"
    a = SetWindowRgn(TheForm.hwnd, a, True)
End Sub

Public Sub ResizeForm(TheForm As Form, OldCursorPos As PointAPI, NewCursorPos As PointAPI, ResizeMode As Integer)
On Error Resume Next
    
' TheForm:      The form you want to resize
' OldCursorPos: The old cursor position (MouseDown)
' NewCursorPos: The new cursor position (MouseUp)
' ResizeMode:   0 - Left side
'               1 - Right side
'               2 - Top side
'               3 - Bottom side
'               4 - Bottom right corner
'               5 - Bottom left corner
'               6 - Top right corner
'               7 - Top left corner
    
' Declare some variables
    Dim DifferenceX
    Dim DifferenceY
    
' Put the difference between the first cursor pos and the second into variables
    DifferenceX = (NewCursorPos.x - OldCursorPos.x) * Screen.TwipsPerPixelX
    DifferenceY = (NewCursorPos.y - OldCursorPos.y) * Screen.TwipsPerPixelY
    
' Determine which resizing mode (above) has been called and resize accordingly
    Select Case ResizeMode
    Case 0
        TheForm.Move TheForm.Left + DifferenceX, TheForm.Top, TheForm.Width - DifferenceX, TheForm.Height
    Case 1
        TheForm.Move TheForm.Left, TheForm.Top, TheForm.Width + DifferenceX, TheForm.Height
    Case 2
        TheForm.Move TheForm.Left, TheForm.Top + DifferenceY, TheForm.Width, TheForm.Height - DifferenceY
    Case 3
        TheForm.Move TheForm.Left, TheForm.Top, TheForm.Width, TheForm.Height + DifferenceY
    Case 4
        TheForm.Move TheForm.Left, TheForm.Top, TheForm.Width + DifferenceX, TheForm.Height + DifferenceY
    Case 5
        TheForm.Move TheForm.Left + DifferenceX, TheForm.Top, TheForm.Width - DifferenceX, TheForm.Height + DifferenceY
    Case 6
        TheForm.Move TheForm.Left, TheForm.Top + DifferenceY, TheForm.Width + DifferenceX, TheForm.Height - DifferenceY
    Case 7
        TheForm.Move TheForm.Left + DifferenceX, TheForm.Top + DifferenceY, TheForm.Width - DifferenceX, TheForm.Height - DifferenceY
    End Select
    
' Check to see if the form has been resized below the minimum size
    If TheForm.Width < 57 * Screen.TwipsPerPixelX Then TheForm.Width = 57 * Screen.TwipsPerPixelX
    If TheForm.Height < 90 * Screen.TwipsPerPixelY Then TheForm.Height = 90 * Screen.TwipsPerPixelY
    
' After resizing the form, make the form "rounded rectangle" shaped
    MakeWindow TheForm, True
End Sub

Public Sub ChangeState(TheForm As Form)
' TheForm:  The form you want to change state (maximized, normal)
    
    If TheForm.WindowState = vbNormal Then
        TheForm.WindowState = vbMaximized
        TheForm!imgTitleMaxRestore.Picture = TheForm!imgTitleRestore.Picture
        MakeWindow TheForm, False
    Else
        TheForm.WindowState = vbNormal
        TheForm!imgTitleMaxRestore.Picture = TheForm!imgTitleMaximize.Picture
        MakeWindow TheForm, IIf(Resizable = 1, True, False)
    End If
End Sub
