VERSION 5.00
Begin VB.Form Der 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Text"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   ControlBox      =   0   'False
   ForeColor       =   &H00FF0000&
   Icon            =   "Derivado4.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Derivado4.frx":6852
   ScaleHeight     =   5355
   ScaleWidth      =   3510
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox AST 
      BackColor       =   &H00FFC0C0&
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Text            =   "105"
      Top             =   120
      Width           =   375
   End
   Begin VB.ListBox L 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   4380
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox S 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Der"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim U(1 To 22), F(1 To 6), CO(1 To 23) As String
Dim P, P2 As String
Dim C, D, E, Mc, M2, N, Z As Integer

'Mc = Numero Maximo de Caracteres'

Private Sub Form_Activate()
'* u = letras *'
U(1) = "b"
U(2) = "c"
U(3) = "d"
U(4) = "f"
U(5) = "g"
U(6) = "h"
U(7) = "j"
U(8) = "k"
U(9) = "l"
U(10) = "m"
U(11) = "n"
U(12) = "p"
U(13) = "q"
U(14) = "r"
U(15) = "s"
U(16) = "t"
U(17) = "v"
U(18) = "w"
U(19) = "x"
U(20) = "y"
U(21) = "z"
U(22) = "b"

'* vocales *'
F(1) = "a"
F(2) = "e"
F(3) = "i"
F(4) = "o"
F(5) = "u"
F(6) = "a"

'* correpciones *'
CO(1) = "á"
CO(2) = "à"
CO(3) = "â"
CO(4) = "ä"
CO(5) = "é"
CO(6) = "è"
CO(7) = "ê"
CO(8) = "ë"
CO(9) = "í"
CO(10) = "ì"
CO(11) = "î"
CO(12) = "ï"
CO(13) = "ó"
CO(14) = "ò"
CO(15) = "ô"
CO(16) = "ö"
CO(17) = "ú"
CO(18) = "ù"
CO(19) = "û"
CO(20) = "ü"
CO(21) = "ý"
CO(22) = "ÿ"
CO(23) = "ñ"

End Sub

Private Sub S_Change()
If (Mc < 26) Then
    P = LCase(S.Text)
    Mc = Len(P)
End If
End Sub

Private Sub S_KeyPress(KeyAscii As Integer)
If (KeyAscii = 13) Then
AST.Text = Val(AST.Text)
If (AST.Text < 1) Then AST.Text = Val(AST.Text) * -1
L.Clear
N = 0
L.AddItem (P + " -> L + " + Str(N))
For E = 1 To Val(AST.Text)
    Ppro
    Spro
Next
End If

If (KeyAscii = 27) Then
End
End If

End Sub

Sub Ppro()
Dim k As Boolean
k = False

'corrigiendo la palabra
For C = 1 To Mc
For D = 1 To 23
If (Mid(P, C, 1) = CO(D)) Then
    If (D > 0) And (D < 5) Then Z = 1
    If (D > 4) And (D < 9) Then Z = 2
    If (D > 8) And (D < 13) Then Z = 3
    If (D > 12) And (D < 17) Then Z = 4
    If (D > 16) And (D < 21) Then Z = 5
    If (D > 20) And (D < 23) Then Z = 6
    If (Z <> 0) Then P2 = P2 + F(Z)
    If (D = 23) Then
    P2 = P2 + "nh"
    M2 = M2 + 1
    End If
    Z = 0
    k = True
End If
Next
If (k = False) Then
    P2 = P2 + Mid(P, C, 1)
End If
k = False

Next
P = P2
P2 = ""
Mc = Mc + M2
M2 = 0
k = False
'hasta aquí

For C = 1 To Mc
For D = 1 To 21
    If (Mid(P, C, 1) = U(D)) Then
    P2 = P2 + U(D + 1)
    k = True
    End If
Next
If (k = False) Then
P2 = P2 + Mid(P, C, 1)
End If
k = False

Next
P = P2
P2 = ""

For C = 1 To Mc
For D = 1 To 5
    If (Mid(P, C, 1) = F(D)) Then
    P2 = P2 + F(D + 1)
    k = True
    End If
Next

If (k = False) Then
P2 = P2 + Mid(P, C, 1)
End If
k = False

Next
P = P2
k = False
P2 = ""

End Sub

Sub Spro()
N = N + 1
L.AddItem (P + " -> L + " + Str(N))
End Sub
