Attribute VB_Name = "Funciones"
'3* AbTxt(String[Path],[Boolean]Crear)
'4* SvFl(String[PathF],String[Texto])
'5* BrFl(String[Path])
'6* MvFl(String[PathI],String[PathF])
'7* CopFl(String[PathI],String[PathF])
'8* BrCp(String[Path])
'9* CrCp(String[Path])
'10* Execute(String[Path])
'11* Ran(Int[Min], Int[Max])
'12* Cifrar(String[Texto])
'13* Descifrar(String[Texto])

Public Function AbTxt(Path As String)
Dim fnum As Integer
On Error GoTo Ninguno
fnum = FreeFile
    Open Path For Input As fnum
    Do While Not EOF(fnum)
        Line Input #fnum, txt
        AbrTxt = AbrTxt & vbCrLf & txt
    Loop
    Close fnum
Ninguno:
End Function

Public Sub SvFl(PathF As String, Texto As String)
Dim fnum As Integer
On Error GoTo Ninguno
    fnum = FreeFile
    Open PathF For Output As fnum
        Print #fnum, Texto
    Close fnum
Ninguno:
End Sub

Public Sub BrFl(Path As String)
On Error GoTo error
Kill Path$ & "*.*"
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub MvFl(PathI As String, PathF As String)
On Error GoTo error
FileCopy PathI$, PathF$
Kill PathI$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub CopFl(PathI As String, PathF As String)
On Error GoTo error
FileCopy PathI$, PathF$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub BrCp(Path As String)
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
        BorrarCarpeta (Path & Dirs(i) & "\")
        RmDir Path & Dirs(i)
    Next
End Sub

Public Sub CrCp(Path As String)
On Error GoTo error
MkDir Path$
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub Execute(Path As String)
On Error GoTo error
ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & (Path))
Exit Sub
error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Function Ran(Min, Max As Integer) As Integer
Randomize
Min = Min - 1
Max = Max + 1
Ran = Int((Min - Max + 1) * Rnd + Max)
End Function

Function Cifrar(Texto As String)
        Dim C As Integer
        Dim K As String
        For C = 1 To Len(Texto)
                If (Asc(Mid(Texto, C, 1)) > 100) Then
                        K = K + Trim(Str(Asc(Mid(Texto, C, 1)))) + ","
                End If
                If (Asc(Mid(Texto, C, 1)) < 100) Then
                        K = K + "0" + Trim(Str(Asc(Mid(Texto, C, 1)))) + ","
                End If
        Next
        K = StrReverse(Mid(K, 1, (Len(K) - 1)))
        Encript = K
End Function

Function Descifrar(Texto As String)
        Dim C As Integer
        Dim K, K2 As String
        K2 = StrReverse(Texto)
        C = 1
        While (C < Len(K2))
                K = K + Chr(Val(Mid(K2, C, 3)))
                C = C + 4
        Wend
        DesEncript = K
End Function
