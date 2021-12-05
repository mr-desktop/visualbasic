VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculadora científica"
   ClientHeight    =   3525
   ClientLeft      =   2280
   ClientTop       =   2355
   ClientWidth     =   5220
   Icon            =   "Calculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Calculator.frx":0442
   ScaleHeight     =   3525
   ScaleWidth      =   5220
   Begin VB.CommandButton cmdMod 
      BackColor       =   &H00800080&
      Caption         =   "Mod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Remainder "
      Top             =   2880
      Width           =   540
   End
   Begin VB.CommandButton cmdRandom 
      BackColor       =   &H00800080&
      Caption         =   "Rnd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Random numbers between 0 and 1"
      Top             =   2880
      Width           =   540
   End
   Begin VB.CommandButton cmdInverse 
      BackColor       =   &H00C0C0FF&
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Inverse"
      Top             =   1920
      Width           =   540
   End
   Begin VB.CommandButton cmdPCR 
      BackColor       =   &H00C0FFC0&
      Caption         =   "nCr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Combination"
      Top             =   960
      Width           =   1020
   End
   Begin VB.CommandButton cmdxy 
      BackColor       =   &H00FFC0C0&
      Caption         =   "x^y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "y^x"
      Top             =   2400
      Width           =   540
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3045
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   945
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      BackColor       =   &H00C0FFC0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   945
      Width           =   855
   End
   Begin VB.CommandButton cmdnPr 
      BackColor       =   &H00800080&
      Caption         =   "nPr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Permuntations"
      Top             =   2880
      Width           =   540
   End
   Begin VB.CommandButton cmdPercent 
      BackColor       =   &H00C0C0FF&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2880
      Width           =   540
   End
   Begin VB.CommandButton cmdCos 
      BackColor       =   &H00FF8080&
      Caption         =   "cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4425
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1950
      Width           =   540
   End
   Begin VB.CommandButton cmdSign 
      BackColor       =   &H00FF8080&
      Caption         =   "sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3855
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1950
      Width           =   540
   End
   Begin VB.CommandButton cmdTan 
      BackColor       =   &H00FF8080&
      Caption         =   "tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1920
      Width           =   540
   End
   Begin VB.CommandButton cmdMem 
      BackColor       =   &H00FFC0FF&
      Caption         =   "MC"
      Height          =   435
      Index           =   2
      Left            =   4425
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Clears Memory"
      Top             =   1470
      Width           =   540
   End
   Begin VB.CommandButton cmdMem 
      BackColor       =   &H00FFC0FF&
      Caption         =   "RCL"
      Height          =   435
      Index           =   1
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Recall memory"
      Top             =   1470
      Width           =   540
   End
   Begin VB.CommandButton cmdMem 
      BackColor       =   &H00FFC0FF&
      Caption         =   "STO"
      Height          =   435
      Index           =   0
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Store number into memory"
      Top             =   1470
      Width           =   540
   End
   Begin VB.CommandButton cmdPi 
      BackColor       =   &H00C0C0FF&
      Caption         =   " pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "3.141592654"
      Top             =   2400
      Width           =   540
   End
   Begin VB.CommandButton cmdAc 
      BackColor       =   &H00C0FFC0&
      Caption         =   "AC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   945
      Width           =   855
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   945
      Width           =   855
   End
   Begin VB.CommandButton cmdSciFunctions 
      BackColor       =   &H00C0C0FF&
      Caption         =   "x!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Factorial"
      Top             =   1440
      Width           =   540
   End
   Begin VB.CommandButton cmdSciFunctions 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sqr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Square Route"
      Top             =   2400
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
      BackColor       =   &H008080FF&
      Caption         =   "="
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
      BackColor       =   &H0080C0FF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Division"
      Top             =   2880
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
      BackColor       =   &H0080C0FF&
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Multiplication"
      Top             =   2400
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
      BackColor       =   &H0080C0FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Subtract"
      Top             =   1920
      Width           =   540
   End
   Begin VB.CommandButton cmdBasicFunctions 
      BackColor       =   &H0080C0FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add"
      Top             =   1440
      Width           =   540
   End
   Begin VB.CommandButton cmdDecimal 
      BackColor       =   &H00C0E0FF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1380
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2910
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdChangeSign 
      BackColor       =   &H00C0E0FF&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   795
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Change Sign"
      Top             =   2910
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   210
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2910
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   9
      Left            =   1380
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1470
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   8
      Left            =   795
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1470
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   210
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1470
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   1380
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1950
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   795
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1950
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   210
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1950
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   3
      Left            =   1380
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2430
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   795
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2430
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton cmdButton 
      BackColor       =   &H00C0E0FF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   210
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2430
      UseMaskColor    =   -1  'True
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Height          =   135
      Left            =   2415
      MaskColor       =   &H80000004&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   1050
      UseMaskColor    =   -1  'True
      Width           =   135
   End
   Begin VB.Label lblMem 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4515
      TabIndex        =   26
      Top             =   420
      Width           =   2790
   End
   Begin VB.Label lblScreen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   4800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim EraseNext As Boolean, Divide As Double, Times As Double, yX As Double
Const pi = 3.141592654
Dim Add As Double, codeMod As Double, nPr As Double
Dim Subtract As Double, nCr As Double

Private Sub cmdAc_Click()
On Error Resume Next
    lblScreen.Caption = ""
    Form_Load
    Add = 0
    Subtract = 0
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdBasicFunctions_Click(Index As Integer)
    On Error Resume Next
    Call SubBasicFuntions(Index)
End Sub

Private Sub cmdButton_Click(Index As Integer)
On Error Resume Next
    Call PressKey(Index + 48)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdChangeSign_Click()
    On Error Resume Next
        If Len(lblScreen.Caption) = 0 Or lblScreen.Caption = "-" Then
        lblScreen.Caption = "-"
        Exit Sub
        End If
    lblScreen.Caption = lblScreen.Caption * -1
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdClear_Click(Index As Integer)
    On Error Resume Next

    Select Case Index
    
    Case 0
    lblScreen.Caption = ""
    Times = 1
    Divide = 1
    
    Case 1
    lblScreen.Caption = ""
    End Select
    
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdCos_Click()
    On Error Resume Next
    lblScreen.Caption = Round(Cos(Radians(Val(lblScreen.Caption))), 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdDecimal_Click()
    On Error Resume Next
    Dim i As Integer
        If EraseNext = True Then
        lblScreen.Caption = ""
        EraseNext = False
        End If
            For i = 1 To Len(lblScreen.Caption)
                If Mid(lblScreen.Caption, i, 1) = "." Then
                MsgBox ("ILLEGAL"), , "NYI"
                Exit Sub
                End If
            Next i
    lblScreen.Caption = lblScreen.Caption & "."
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdInverse_Click()
    On Error Resume Next
    lblScreen.Caption = 1 / Val(lblScreen.Caption)
End Sub

Private Sub cmdMem_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0 'sto
    lblMem.Caption = lblScreen.Caption
    
    Case 1 'rcl
    lblScreen.Caption = lblMem.Caption
    lblMem.Caption = ""
    
    Case 2
    lblMem.Caption = "" 'memclear
    End Select
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdMod_Click()
    EraseNext = 1
    codeMod = Val(lblScreen.Caption)
End Sub

Private Sub cmdnPr_Click()
    On Error Resume Next
    nPr = Val(lblScreen.Caption)
    EraseNext = True
End Sub

Private Sub cmdPCR_Click()
    On Error Resume Next
        If Val(lblScreen.Caption) = 0 Then
        lblScreen.Caption = 1
        Exit Sub
        End If
    nCr = Val(lblScreen.Caption)
    EraseNext = True
End Sub

Private Sub cmdPercent_Click()
On Error Resume Next
    lblScreen.Caption = (Val(lblScreen.Caption) * 0.01) + Val(lblScreen.Caption)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdPi_Click()
    On Error Resume Next
    lblScreen.Caption = FormatNumber(pi, 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdRandom_Click()
    On Error Resume Next
    lblScreen.Caption = Rnd
End Sub

Private Sub cmdSciFunctions_Click(Index As Integer)
    On Error Resume Next
    Dim Screen As Double, TheScreen As Double
    On Error Resume Next
        Select Case Index
        Case 0 'squared
            If Len(lblScreen.Caption) = 0 Then
            MsgBox ("ILLEGAL, 0 ²!!"), , "NYI"
            Exit Sub
            End If
        lblScreen.Caption = Val(lblScreen.Caption) ^ 2
        Case 1 'sqroute
            If Val(lblScreen.Caption) < 0 Then
            MsgBox ("ILLEGAL, sqr ( )!!"), , "NYI"
            Exit Sub
            End If
        lblScreen.Caption = Sqr(Val(lblScreen.Caption))
        Case 2 'factorial
            If Val(lblScreen.Caption) < 0 Then
            MsgBox ("ILLEGAL,FACT ( )!!"), , "NYI"
            Exit Sub
            Else
                If Val(lblScreen.Caption) = 0 Then
                lblScreen.Caption = 1
                Exit Sub
                End If
            End If
        Screen = Val(lblScreen.Caption)
        TheScreen = Val(lblScreen.Caption)
        lblScreen.Caption = Factorial(TheScreen, Screen)
        End Select
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub CmdBack_Click()
    On Error Resume Next
        If (lblScreen.Caption <> "") Then
        lblScreen.Caption = Mid(lblScreen.Caption, 1, Len(lblScreen.Caption) - 1)
        End If
cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdSign_Click()
    On Error Resume Next
    lblScreen.Caption = Round(Sin(Radians(Val(lblScreen.Caption))), 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdTan_Click()
    On Error Resume Next
    lblScreen.Caption = Round(Tan(Radians(Val(lblScreen.Caption))), 9)
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdX3_Click()
    On Error Resume Next
    lblScreen.Caption = Val(lblScreen.Caption) ^ 3
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub cmdxy_Click()
    Dim y As Integer
        If Len(lblScreen.Caption) = 0 Then
        MsgBox ("ILLEGAL, 0^x!!"), vbCritical, "NYI"
        Exit Sub
        End If
    yX = Val(lblScreen.Caption)
    cmdBasicFunctions(0).SetFocus
    EraseNext = True
End Sub

Private Sub CmdRnd_Click()
    On Error Resume Next
    lblScreen.Caption = Rnd
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Call PressKey(KeyCode)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Randomize
    Times = 1
    Divide = 1
    Add = 0
    Subtract = 0
End Sub

Private Function Radians(ByRef Degrees As Double)
    On Error Resume Next
'converts a number to radians for tan functions
    Radians = Degrees * pi / 180
End Function

Public Sub PressKey(ByVal Ind As Integer)
    On Error Resume Next
        If Ind >= 48 And Ind <= 57 Then
        Ind = Ind - 48
            If EraseNext = True Then
            EraseNext = False
            lblScreen.Caption = Ind
            Exit Sub
            Else
                lblScreen.Caption = lblScreen.Caption & Ind
                Exit Sub
            End If
        End If
    
             If Ind >= 96 And Ind < 106 Then
             Ind = Ind - 96
                 If EraseNext = True Then
                 EraseNext = False
                 lblScreen.Caption = Ind
                 Exit Sub
                 Else
                     lblScreen.Caption = lblScreen.Caption & Ind
                     Exit Sub
             End If
                 End If
    
    Select Case Ind
    Case 107
    Ind = 1
    Call SubBasicFuntions(Ind)
    
    Case 109
    Ind = 2
    Call SubBasicFuntions(Ind)
    
    Case 106
    Ind = 3
    Call SubBasicFuntions(Ind)
    
    Case 111
    Ind = 4
    Call SubBasicFuntions(Ind)
    
    Case 110
    cmdDecimal_Click
    Case 8
    CmdBack_Click
    End Select
End Sub

Private Function Factorial(ByVal TheLabel As Double, Label As Double) As Double
'does the factorial of a number, used in combination and permutation
    On Error Resume Next
    Do Until Label = 1
    Label = Label - 1
    TheLabel = (Label) * TheLabel
    Loop
Factorial = TheLabel
End Function

Private Function Perm(ByRef x As Integer, N As Double, Temprary As Double, r As Double) As Double
'This does all the permutations, returns perm of 2 numbers
    On Error Resume Next
    Call PermComErrorCheck(N, r)
    x = 1
        Do Until x = r
        N = N * (Temprary - x)
        x = x + 1
        Loop
    EraseNext = True
    Perm = N
End Function

Private Sub lblScreen_Change()
Dim a As Integer
a = InStr(lblScreen, ",")
If a > 0 Then
lblScreen = Left(lblScreen, a - 1) & "." & Mid(lblScreen, a + 1, Len(lblScreen) - a - 1)
End If

    On Error Resume Next
        If Len(lblScreen.Caption) >= 20 Then
        CmdBack_Click
        Beep
        End If
End Sub

Private Sub Equal()
    On Error Resume Next
    Static NFirst As Double 'N in perm and comb
    Static Rsec As Double  'R in perm and comb
    Static temp As Double 'r in perm and comb
    Dim Counter As Integer
        If Divide <> 1 Then
            If Divide = 0 Or Val(lblScreen.Caption) = 0 Then
            MsgBox ("¡División por cero!"), vbCritical, "NYI"
            EraseNext = True
            Exit Sub
            End If
        lblScreen.Caption = Divide / Val(lblScreen.Caption)
        Divide = 1
            Else
            If Times <> 1 Then
            lblScreen.Caption = Val(lblScreen.Caption) * Times
            Times = 1
            Else
                If Add <> 0 Then
                lblScreen.Caption = Val(lblScreen.Caption) + Add
                Add = 0
                Else
                    If Subtract <> 0 Then
                    lblScreen.Caption = Subtract - Val(lblScreen.Caption)
                    Subtract = 0
                    Else
                        If codeMod <> 0 Then
                        lblScreen.Caption = codeMod Mod Val(lblScreen.Caption)
                        codeMod = 0
                        Else
                            If nPr <> 0 Then
                                If Val(lblScreen.Caption) = 0 Then
                                lblScreen.Caption = 1
                                Exit Sub
                                End If
                                    NFirst = nPr 'WORKING HERE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                    Counter = 1
                                    Rsec = Val(lblScreen.Caption)
                                    EraseNext = False
                                    lblScreen.Caption = Perm(Counter, nPr, NFirst, Rsec) 'perm is a function
                                    nPr = 0
                                    Else
                                        If yX <> 0 Then
                                        Dim y As Double, Rsec1 As Double
                                        y = Val(lblScreen.Caption)
                                        lblScreen.Caption = yX ^ y
                                        yX = 0
                                        Else
                                            If nCr <> 0 Then 'gets perm
                                            NFirst = nCr
                                            Rsec = Val(lblScreen.Caption)
                                            Rsec1 = Rsec
                                            Counter = 1
                                            lblScreen.Caption = Perm(Counter, nCr, NFirst, Rsec) / Factorial(Rsec, Rsec1)
                                            nCr = 0
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
    EraseNext = True
    cmdBasicFunctions(0).SetFocus
End Sub

Private Sub SubBasicFuntions(ByVal Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 1
        Add = Val(lblScreen.Caption) + Add
        EraseNext = True
        lblScreen.Caption = Add
    
        Case 2
        Subtract = Val(lblScreen.Caption) - Subtract
        EraseNext = True
        lblScreen.Caption = Subtract
        
        Case 3
        Times = Val(lblScreen.Caption) * Times
        EraseNext = True
        lblScreen.Caption = Times
        
        Case 4
        Divide = Val(lblScreen.Caption) / Divide
        EraseNext = True
        lblScreen.Caption = Divide
        
        Case 0
        Call Equal
        End Select
End Sub

Private Sub PermComErrorCheck(ByVal First As Double, Second As Double)
On Error Resume Next
'number must be greater than or = to 0
    If (First < 0) Or (Second < 0) Or (Second > First) Then
    MsgBox ("ILLEGAL PERM ()!!"), , "NYI"
    End If
End Sub
