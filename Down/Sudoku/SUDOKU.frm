VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "B I E N V E N I D O      A L   ""  S U D O K U """
   ClientHeight    =   8790
   ClientLeft      =   3405
   ClientTop       =   645
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   7215
   Begin VB.TextBox aux 
      Height          =   555
      Left            =   600
      TabIndex        =   94
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Vsolución 
      Caption         =   "&Ver Solución"
      Height          =   495
      Left            =   3840
      TabIndex        =   93
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   480
      TabIndex        =   11
      Top             =   240
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   8
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   7
         Left            =   720
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   6
         Left            =   120
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   17
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   720
         MaxLength       =   1
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   3
         Left            =   120
         MaxLength       =   1
         TabIndex        =   15
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   2
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   720
         MaxLength       =   1
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   120
         MaxLength       =   1
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Salir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   5520
      TabIndex        =   10
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton Asolucion 
      Caption         =   "&Analizar Solución "
      Height          =   495
      Left            =   2160
      TabIndex        =   9
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton Njuego 
      Caption         =   "&Nuevo Juego"
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   4800
      TabIndex        =   7
      Top             =   5280
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   80
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   92
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   79
         Left            =   720
         MaxLength       =   1
         TabIndex        =   91
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   78
         Left            =   120
         MaxLength       =   1
         TabIndex        =   90
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   77
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   89
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   76
         Left            =   720
         MaxLength       =   1
         TabIndex        =   88
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   75
         Left            =   120
         MaxLength       =   1
         TabIndex        =   87
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   74
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   86
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   73
         Left            =   720
         MaxLength       =   1
         TabIndex        =   85
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   72
         Left            =   120
         MaxLength       =   1
         TabIndex        =   84
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2640
      TabIndex        =   6
      Top             =   5280
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   71
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   83
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   70
         Left            =   720
         MaxLength       =   1
         TabIndex        =   82
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   69
         Left            =   120
         MaxLength       =   1
         TabIndex        =   81
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   68
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   80
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   67
         Left            =   720
         MaxLength       =   1
         TabIndex        =   79
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   66
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   65
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   77
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   64
         Left            =   720
         MaxLength       =   1
         TabIndex        =   76
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   63
         Left            =   120
         MaxLength       =   1
         TabIndex        =   75
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   480
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   62
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   74
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   61
         Left            =   720
         MaxLength       =   1
         TabIndex        =   73
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   60
         Left            =   120
         MaxLength       =   1
         TabIndex        =   72
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   59
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   71
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   58
         Left            =   720
         MaxLength       =   1
         TabIndex        =   70
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   57
         Left            =   120
         MaxLength       =   1
         TabIndex        =   69
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   56
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   68
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   55
         Left            =   720
         MaxLength       =   1
         TabIndex        =   67
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   54
         Left            =   120
         MaxLength       =   1
         TabIndex        =   66
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   4800
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   53
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   65
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   52
         Left            =   720
         MaxLength       =   1
         TabIndex        =   64
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   51
         Left            =   120
         MaxLength       =   1
         TabIndex        =   63
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   50
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   62
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   49
         Left            =   720
         MaxLength       =   1
         TabIndex        =   61
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   48
         Left            =   120
         MaxLength       =   1
         TabIndex        =   60
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   47
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   59
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   46
         Left            =   720
         MaxLength       =   1
         TabIndex        =   58
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   45
         Left            =   120
         MaxLength       =   1
         TabIndex        =   57
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   44
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   56
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   43
         Left            =   720
         MaxLength       =   1
         TabIndex        =   55
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   42
         Left            =   120
         MaxLength       =   1
         TabIndex        =   54
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   41
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   53
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   40
         Left            =   720
         MaxLength       =   1
         TabIndex        =   52
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   39
         Left            =   120
         MaxLength       =   1
         TabIndex        =   51
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   38
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   50
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   37
         Left            =   720
         MaxLength       =   1
         TabIndex        =   49
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   36
         Left            =   120
         MaxLength       =   1
         TabIndex        =   48
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   35
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   47
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   34
         Left            =   720
         MaxLength       =   1
         TabIndex        =   46
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   33
         Left            =   120
         MaxLength       =   1
         TabIndex        =   45
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   32
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   44
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   31
         Left            =   720
         MaxLength       =   1
         TabIndex        =   43
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   30
         Left            =   120
         MaxLength       =   1
         TabIndex        =   42
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   29
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   28
         Left            =   720
         MaxLength       =   1
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   27
         Left            =   120
         MaxLength       =   1
         TabIndex        =   39
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   4800
      TabIndex        =   1
      Top             =   240
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   26
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   38
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   25
         Left            =   720
         MaxLength       =   1
         TabIndex        =   37
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   24
         Left            =   120
         MaxLength       =   1
         TabIndex        =   36
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   23
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   35
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   22
         Left            =   720
         MaxLength       =   1
         TabIndex        =   34
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   21
         Left            =   120
         MaxLength       =   1
         TabIndex        =   33
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   20
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   19
         Left            =   720
         MaxLength       =   1
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   18
         Left            =   120
         MaxLength       =   1
         TabIndex        =   30
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   17
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   29
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   16
         Left            =   720
         MaxLength       =   1
         TabIndex        =   28
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   15
         Left            =   120
         MaxLength       =   1
         TabIndex        =   27
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   14
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   26
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   13
         Left            =   720
         MaxLength       =   1
         TabIndex        =   25
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   12
         Left            =   120
         MaxLength       =   1
         TabIndex        =   24
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   11
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   10
         Left            =   720
         MaxLength       =   1
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   9
         Left            =   120
         MaxLength       =   1
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Autor: Valderrama Freddy
'Correo: favs2855@cantv.net
'
'



Private Function GenerarRndNumber(Upper As Integer, Lower As Integer) As Integer
'Se declara una funcion para buscar numeros aleatorios antre el intervalo 1-9
Randomize
GenerarRndNumber = Int((Upper - Lower + 1) * Rnd + Lower)
End Function
Private Function Llenar(x As Integer, y As Integer) As Boolean
'Borra casillas aleatorias del nuevo juego
a = GenerarRndNumber((x), (y))
b = GenerarRndNumber((x), (y))
c = GenerarRndNumber((x), (y))
d = GenerarRndNumber((x), (y))
e = GenerarRndNumber((x), (y))
Do While a = b
b = GenerarRndNumber((x), (y))
Loop
Do While (a = c Or c = b)
c = GenerarRndNumber((x), (y))
Loop
Do While (a = d Or d = b Or d = c)
d = GenerarRndNumber((x), (y))
Loop
Do While (e = a Or e = b Or e = c Or e = d)
e = GenerarRndNumber((x), (y))
Loop
Text(a).Text = ""
Text(b).Text = ""
Text(c).Text = ""
Text(d).Text = ""
Text(e).Text = ""

End Function


Private Sub Asolucion_Click()
' Verifica lo solución creada por el usuario con el Arrays correspondiente
a0 = Array(7, 8, 6, 1, 5, 4, 3, 2, 9, 5, 9, 2, 8, 3, 7, 6, 4, 1, 4, 1, 3, 9, 2, 6, 8, 5, 7, 2, 7, 8, 9, 1, 5, 6, 4, 3, 1, 6, 5, 3, 8, 4, 2, 7, 9, 3, 4, 9, 6, 7, 2, 1, 8, 5, 5, 3, 2, 4, 6, 1, 8, 9, 7, 9, 1, 8, 7, 2, 3, 4, 5, 6, 7, 6, 4, 5, 9, 8, 2, 3, 1)
a1 = Array(2, 6, 1, 5, 9, 3, 4, 7, 8, 4, 3, 5, 8, 2, 7, 6, 1, 9, 7, 8, 9, 1, 6, 4, 5, 3, 2, 3, 8, 5, 9, 4, 7, 6, 1, 2, 2, 9, 1, 5, 8, 6, 3, 7, 4, 4, 7, 6, 3, 2, 1, 8, 9, 5, 8, 3, 4, 7, 5, 9, 1, 2, 6, 9, 5, 2, 1, 6, 3, 7, 4, 8, 6, 1, 7, 2, 4, 8, 9, 5, 3)
a2 = Array(8, 1, 2, 6, 7, 5, 4, 3, 9, 7, 5, 3, 4, 2, 9, 1, 8, 6, 6, 9, 4, 8, 3, 1, 7, 5, 2, 9, 4, 6, 1, 5, 3, 7, 2, 8, 8, 1, 5, 6, 7, 2, 9, 3, 4, 2, 7, 3, 4, 8, 9, 5, 1, 6, 5, 6, 4, 2, 9, 1, 3, 8, 7, 3, 9, 7, 5, 6, 8, 2, 4, 1, 1, 2, 8, 3, 4, 7, 9, 6, 5)
a = aux.Text
Select Case a
       Case 1: a = a0
       Case 2: a = a1
       Case 3: a = a2
End Select


For i = 0 To 80
If Text(i).Text = "" Then
x = x + 1
Else
If Text(i).Text = a(i) Then
Else
x = x + 1
End If
End If
Next i

If x > 0 Then
MsgBox "Numero(s) Duplicado(s), o celdas Vacias Favor Verifique"
Else
MsgBox "Felicitaciones, Juego Terminado con Exito"
End If
End Sub

Private Sub Form_Load()
' limpia toda la matriz
For i = 0 To 80
Text(i).Text = ""
Next i

End Sub

Sub Njuego_Click()
'genera un nuevo juevo
a0 = Array(7, 8, 6, 1, 5, 4, 3, 2, 9, 5, 9, 2, 8, 3, 7, 6, 4, 1, 4, 1, 3, 9, 2, 6, 8, 5, 7, 2, 7, 8, 9, 1, 5, 6, 4, 3, 1, 6, 5, 3, 8, 4, 2, 7, 9, 3, 4, 9, 6, 7, 2, 1, 8, 5, 5, 3, 2, 4, 6, 1, 8, 9, 7, 9, 1, 8, 7, 2, 3, 4, 5, 6, 7, 6, 4, 5, 9, 8, 2, 3, 1)
a1 = Array(2, 6, 1, 5, 9, 3, 4, 7, 8, 4, 3, 5, 8, 2, 7, 6, 1, 9, 7, 8, 9, 1, 6, 4, 5, 3, 2, 3, 8, 5, 9, 4, 7, 6, 1, 2, 2, 9, 1, 5, 8, 6, 3, 7, 4, 4, 7, 6, 3, 2, 1, 8, 9, 5, 8, 3, 4, 7, 5, 9, 1, 2, 6, 9, 5, 2, 1, 6, 3, 7, 4, 8, 6, 1, 7, 2, 4, 8, 9, 5, 3)
a2 = Array(8, 1, 2, 6, 7, 5, 4, 3, 9, 7, 5, 3, 4, 2, 9, 1, 8, 6, 6, 9, 4, 8, 3, 1, 7, 5, 2, 9, 4, 6, 1, 5, 3, 7, 2, 8, 8, 1, 5, 6, 7, 2, 9, 3, 4, 2, 7, 3, 4, 8, 9, 5, 1, 6, 5, 6, 4, 2, 9, 1, 3, 8, 7, 3, 9, 7, 5, 6, 8, 2, 4, 1, 1, 2, 8, 3, 4, 7, 9, 6, 5)

a = GenerarRndNumber(0, 4)
aux.Text = a
Select Case a
       Case 1: a = a0
       Case 2: a = a1
       Case 3: a = a2
End Select


b = Array(-1, 8, 17, 26, 35, 44, 53, 62, 71)
c = Array(9, 18, 27, 36, 45, 54, 63, 72, 81)

For i = 0 To 80
Text(i).Text = ""
Text(i).Text = a(i)
Next i


Dim x As Integer
Dim y As Integer

For i = 0 To 8
x = b(i)
y = c(i)
prueba = Llenar(x, y)
Next i

For i = 0 To 80
If Text(i).Text = "" Then
Text(i).Enabled = True
Else
Text(i).Enabled = False
End If
Next i

End Sub



Private Sub Salir_Click()
' Solo termina el programa o juego
End
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case KeyAscii
            Case 48 To 57:          '  NUMEROS
            Case Else:                '  CUALQUIER OTRA LETRA ANULAR
   KeyAscii = 0
  End Select
End Sub

 Sub Vsolución_Click()
 'Muestra la Matriz completa que el usuario esta ejecutando
a0 = Array(7, 8, 6, 1, 5, 4, 3, 2, 9, 5, 9, 2, 8, 3, 7, 6, 4, 1, 4, 1, 3, 9, 2, 6, 8, 5, 7, 2, 7, 8, 9, 1, 5, 6, 4, 3, 1, 6, 5, 3, 8, 4, 2, 7, 9, 3, 4, 9, 6, 7, 2, 1, 8, 5, 5, 3, 2, 4, 6, 1, 8, 9, 7, 9, 1, 8, 7, 2, 3, 4, 5, 6, 7, 6, 4, 5, 9, 8, 2, 3, 1)
a1 = Array(2, 6, 1, 5, 9, 3, 4, 7, 8, 4, 3, 5, 8, 2, 7, 6, 1, 9, 7, 8, 9, 1, 6, 4, 5, 3, 2, 3, 8, 5, 9, 4, 7, 6, 1, 2, 2, 9, 1, 5, 8, 6, 3, 7, 4, 4, 7, 6, 3, 2, 1, 8, 9, 5, 8, 3, 4, 7, 5, 9, 1, 2, 6, 9, 5, 2, 1, 6, 3, 7, 4, 8, 6, 1, 7, 2, 4, 8, 9, 5, 3)
a2 = Array(8, 1, 2, 6, 7, 5, 4, 3, 9, 7, 5, 3, 4, 2, 9, 1, 8, 6, 6, 9, 4, 8, 3, 1, 7, 5, 2, 9, 4, 6, 1, 5, 3, 7, 2, 8, 8, 1, 5, 6, 7, 2, 9, 3, 4, 2, 7, 3, 4, 8, 9, 5, 1, 6, 5, 6, 4, 2, 9, 1, 3, 8, 7, 3, 9, 7, 5, 6, 8, 2, 4, 1, 1, 2, 8, 3, 4, 7, 9, 6, 5)
a = aux.Text
Select Case a
       Case 1: a = a0
       Case 2: a = a1
       Case 3: a = a2
End Select

For i = 0 To 80
Text(i).Text = ""
Text(i).Text = a(i)
Next i
End Sub
