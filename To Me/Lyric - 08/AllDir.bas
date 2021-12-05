Attribute VB_Name = "AllDir"
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Obtiene Los Directorios De Windows
Public Function WinDir(N As Integer) As String
Dim AaA As String * 128
Dim Aa As Integer

If (N = 2) Then
Aa = GetSystemDirectory(AaA, 128)
End If

If (N = 1) Then
Aa = GetWindowsDirectory(AaA, 128)
End If

WinDir = RTrim$(LCase$(Left$(AaA, Aa)))
End Function

'Obtiene El Directorio Del Programa Ejecutado
Public Function CurDir() As String
Dim Directorio As String
ChDir App.Path
ChDrive App.Path
Directorio = App.Path
If Len(Directorio) > 3 Then
Directorio = Directorio & "\"
End If
CurDir = Directorio
End Function
