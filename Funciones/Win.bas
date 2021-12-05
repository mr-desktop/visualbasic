Attribute VB_Name = "Win"
'* * * * * * * * * * * * * * 1 * * * * * * * * * * * * * * * *
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public tWnd As Long, bWnd As Long, sSave As String * 250

'* * * * * * * * * * * * * * 2 * * * * * * * * * * * * * * * *
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'* * * * * * * * * * * * * * 3 * * * * * * * * * * * * * * * *
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long


Public Function WinB(N As Integer)
If (N = 1) Then
'Show Win Button
    SetWindowPos bWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW
End If

If (N = 0) Then
'Hide Win Button
    tWnd = FindWindow("Shell_traywnd", vbNullString)
    bWnd = GetWindow(tWnd, GW_CHILD)
    Do
    GetClassName bWnd, sSave, 250
    If LCase(Left$(sSave, 6)) = "button" Then Exit Do
    bWnd = GetWindow(bWnd, GW_HWNDNEXT)
    Loop
    SetWindowPos bWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW
End If
End Function


Public Function Pant(N As Integer)
If (N = 0) Then
'Captura toda la pantalla
    keybd_event 44, 0, 0&, 0&
End If

If (N = 1) Then
'Captura la ventana activa
    keybd_event 44, 1, 0&, 0&
End If
End Function

'Fondo de Windows
Public Function Fondo(Ruta As String)
Dim fallo As Integer
fallo = SystemParametersInfo(20, 0, Ruta, 0)
End Function
