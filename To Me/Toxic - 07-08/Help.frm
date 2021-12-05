VERSION 5.00
Begin VB.Form Help 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   ClipControls    =   0   'False
   Icon            =   "Help.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Help.frx":000C
   ScaleHeight     =   3705
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1740
      ItemData        =   "Help.frx":CFB1
      Left            =   120
      List            =   "Help.frx":CFCA
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label LS 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sentencias"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   3840
      Left            =   1080
      Picture         =   "Help.frx":CFF4
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Ju, Ca As String

Private Sub Image1_Click()
Help.Hide
End Sub

Private Sub List1_Click()
Ju = List1.ListIndex
If (Ju = 0) Then
Ca = "Formulario Que Cambia La Imagen De Fondo"
End If
If (Ju = 1) Then
Ca = "Formulario Que Cambia El Color De Fondo"
End If
If (Ju = 2) Then
Ca = "Comando Que Expande o Centraliza La Imagen De Fondo"
End If
If (Ju = 3) Then
Ca = "Muestra La Versión Del Producto Y El Nombre Del Autor"
End If
If (Ju = 4) Then
Ca = "Formulario Que Cambia La Clave"
End If
If (Ju = 5) Then
Ca = "Formulario Que Confirma La Clave Para Salir"
End If
If (Ju = 6) Then
Ca = "Formulario Que Ofrece Ayuda"
End If
LS.Caption = Ca
End Sub
