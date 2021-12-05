Attribute VB_Name = "AllDir"
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
