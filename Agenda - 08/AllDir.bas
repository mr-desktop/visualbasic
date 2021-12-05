Attribute VB_Name = "AllDir"
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

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


