Attribute VB_Name = "Funciones"
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Const GW_CHILD = 5
Const GW_HWNDNEXT = 2

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Dim tWnd As Long, bWnd As Long, sSave As String * 250

'OcultarWB()
'MostrarWB()
'AbrTxt(String[Path],TextBox[Texto])
'BorrCapt(String[Path])
'CopyFl(String[PathI],String[PathF])
'CrearCpt(String[Path])
'ExeFl(String[Path])
'BorrFl(String[Path])
'SavFl(String[PathI],String)
'MoveFl(String[PathI],String[PathF])

Public Function OcultarWB()
    tWnd = FindWindow("Shell_traywnd", vbNullString)
    bWnd = GetWindow(tWnd, GW_CHILD)
    Do
    GetClassName bWnd, sSave, 250
    If LCase(Left$(sSave, 6)) = "button" Then Exit Do
    bWnd = GetWindow(bWnd, GW_HWNDNEXT)
    Loop
    SetWindowPos bWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW
End Function

Public Function MostrarWB()
    SetWindowPos bWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW
End Function

Sub AbrTxt(Path As String, Texto As TextBox)
Dim fnum As Integer
On Error GoTo Ninguno
fnum = FreeFile
    Open Path For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, txt
        Texto.Text = Texto.Text & vbCrLf & txt
    Loop
    Close fnum
Ninguno:
SavFl Path, Texto.Text, True
End Sub

Sub BorrCpt(Path As String)
    Dim sName As String
    Dim sFullName As String

    Dim Dirs() As String
    Dim DirsNo As Integer
    Dim i As Integer


    If Not Right(Path, 1) = "\" Then
        Path = Path & "\"
    End If

    sName = Dir(Path & "*.*")

    While Len(sName) > 0
        sFullName = Path & sName
        SetAttr sFullName, vbNormal
        Kill sFullName
        sName = Dir
    Wend
    
    sName = Dir(Path & "*.*", vbHidden)

    While Len(sName) > 0
        sFullName = Path & sName
        SetAttr sFullName, vbNormal
        Kill sFullName
        sName = Dir
    Wend
    
    DirsNo = 0
    sName = Dir(Path, vbDirectory)

    While Len(sName) > 0

        If sName <> "." And sName <> ".." Then
            DirsNo = DirsNo + 1
            ReDim Preserve Dirs(DirsNo) As String
            Dirs(DirsNo - 1) = sName
        End If
        sName = Dir
    Wend

    For i = 0 To DirsNo - 1
        RmDir (Path & Dirs(i) & "\")
        RmDir Path & Dirs(i)
    Next
End Sub

Public Sub CopyFl(PathI As String, PathF As String)
On Error GoTo error
FileCopy PathI$, PathF$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub CrearCpt(Path As String)
On Error GoTo error
MkDir Path$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ExeFl(Path As String)
On Error GoTo error
ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & (Path))
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub BorrFl(Path As String)
On Error GoTo error
Kill Path$ & "*.*"
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Sub SavFl(Path As String, Texto As String, Optional Nuevo As Boolean)
Dim fnum As Integer
On Error GoTo Ninguno
    If (Nuevo <> True) Then
    BorrFl Path
    End If
    fnum = FreeFile
    Open Path For Output As fnum
        Print #fnum, Texto
    Close fnum
Ninguno:
End Sub

Public Sub MoveFl(PathI As String, PathF As String)
On Error GoTo error
FileCopy PathI$, PathF$
Kill PathI$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub
