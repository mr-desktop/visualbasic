VERSION 5.00
Begin VB.Form Trost 
   BackColor       =   &H8000000D&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tetravex"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   11175
   ClipControls    =   0   'False
   Icon            =   "Tetra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Tetra.frx":1CFA
   ScaleHeight     =   5640
   ScaleWidth      =   11175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   24
      Left            =   9960
      TabIndex        =   147
      Top             =   4440
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   41
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   40
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   49
         Left            =   360
         TabIndex        =   151
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   48
         Left            =   360
         TabIndex        =   150
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   49
         Left            =   720
         TabIndex        =   149
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   48
         Left            =   0
         TabIndex        =   148
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   23
      Left            =   8880
      TabIndex        =   142
      Top             =   4440
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   39
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   38
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   39
         Left            =   360
         TabIndex        =   146
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   38
         Left            =   360
         TabIndex        =   145
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   47
         Left            =   720
         TabIndex        =   144
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   46
         Left            =   0
         TabIndex        =   143
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   22
      Left            =   7800
      TabIndex        =   137
      Top             =   4440
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   21
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   20
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   29
         Left            =   360
         TabIndex        =   141
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   28
         Left            =   360
         TabIndex        =   140
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   45
         Left            =   720
         TabIndex        =   139
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   44
         Left            =   0
         TabIndex        =   138
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   21
      Left            =   6720
      TabIndex        =   132
      Top             =   4440
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   19
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   18
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   19
         Left            =   360
         TabIndex        =   136
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   18
         Left            =   360
         TabIndex        =   135
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   43
         Left            =   720
         TabIndex        =   134
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   42
         Left            =   0
         TabIndex        =   133
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   20
      Left            =   5640
      TabIndex        =   127
      Top             =   4440
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   1
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   0
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   131
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   130
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   41
         Left            =   720
         TabIndex        =   129
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   40
         Left            =   0
         TabIndex        =   128
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   19
      Left            =   9960
      TabIndex        =   122
      Top             =   3360
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   43
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   42
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   47
         Left            =   360
         TabIndex        =   126
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   46
         Left            =   360
         TabIndex        =   125
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   39
         Left            =   720
         TabIndex        =   124
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   38
         Left            =   0
         TabIndex        =   123
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   14
      Left            =   9960
      TabIndex        =   117
      Top             =   2280
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   45
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   44
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   45
         Left            =   360
         TabIndex        =   121
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   44
         Left            =   360
         TabIndex        =   120
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   29
         Left            =   720
         TabIndex        =   119
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   28
         Left            =   0
         TabIndex        =   118
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   9
      Left            =   9960
      TabIndex        =   112
      Top             =   1200
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   47
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   46
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   43
         Left            =   360
         TabIndex        =   116
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   42
         Left            =   360
         TabIndex        =   115
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   19
         Left            =   720
         TabIndex        =   114
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   18
         Left            =   0
         TabIndex        =   113
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   4
      Left            =   9960
      TabIndex        =   107
      Top             =   120
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   49
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   48
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   41
         Left            =   360
         TabIndex        =   111
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   40
         Left            =   360
         TabIndex        =   110
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   109
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   0
         TabIndex        =   108
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   18
      Left            =   8880
      TabIndex        =   102
      Top             =   3360
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   37
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   36
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   37
         Left            =   360
         TabIndex        =   106
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   36
         Left            =   360
         TabIndex        =   105
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   37
         Left            =   720
         TabIndex        =   104
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   36
         Left            =   0
         TabIndex        =   103
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   17
      Left            =   7800
      TabIndex        =   97
      Top             =   3360
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   23
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   22
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   27
         Left            =   360
         TabIndex        =   101
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   26
         Left            =   360
         TabIndex        =   100
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   35
         Left            =   720
         TabIndex        =   99
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   34
         Left            =   0
         TabIndex        =   98
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   16
      Left            =   6720
      TabIndex        =   92
      Top             =   3360
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   17
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   16
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   17
         Left            =   360
         TabIndex        =   96
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   16
         Left            =   360
         TabIndex        =   95
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   33
         Left            =   720
         TabIndex        =   94
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   32
         Left            =   0
         TabIndex        =   93
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   15
      Left            =   5640
      TabIndex        =   87
      Top             =   3360
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   3
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   2
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   91
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   90
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   31
         Left            =   720
         TabIndex        =   89
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   30
         Left            =   0
         TabIndex        =   88
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   13
      Left            =   8880
      TabIndex        =   82
      Top             =   2280
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   35
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   34
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   35
         Left            =   360
         TabIndex        =   86
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   34
         Left            =   360
         TabIndex        =   85
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   27
         Left            =   720
         TabIndex        =   84
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   26
         Left            =   0
         TabIndex        =   83
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   8
      Left            =   8880
      TabIndex        =   77
      Top             =   1200
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   33
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   32
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   33
         Left            =   360
         TabIndex        =   81
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   32
         Left            =   360
         TabIndex        =   80
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   17
         Left            =   720
         TabIndex        =   79
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   16
         Left            =   0
         TabIndex        =   78
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   3
      Left            =   8880
      TabIndex        =   72
      Top             =   120
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   31
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   30
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   31
         Left            =   360
         TabIndex        =   76
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   30
         Left            =   360
         TabIndex        =   75
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   720
         TabIndex        =   74
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   0
         TabIndex        =   73
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   12
      Left            =   7800
      TabIndex        =   67
      Top             =   2280
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   25
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   24
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   25
         Left            =   360
         TabIndex        =   71
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   24
         Left            =   360
         TabIndex        =   70
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   25
         Left            =   720
         TabIndex        =   69
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   24
         Left            =   0
         TabIndex        =   68
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   11
      Left            =   6720
      TabIndex        =   62
      Top             =   2280
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   15
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   14
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   360
         TabIndex        =   66
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   360
         TabIndex        =   65
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   23
         Left            =   720
         TabIndex        =   64
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   22
         Left            =   0
         TabIndex        =   63
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   10
      Left            =   5640
      TabIndex        =   57
      Top             =   2280
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   5
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   4
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   61
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   60
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   21
         Left            =   720
         TabIndex        =   59
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   20
         Left            =   0
         TabIndex        =   58
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   7
      Left            =   7800
      TabIndex        =   52
      Top             =   1200
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   27
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   26
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   23
         Left            =   360
         TabIndex        =   56
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   22
         Left            =   360
         TabIndex        =   55
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   15
         Left            =   720
         TabIndex        =   54
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   14
         Left            =   0
         TabIndex        =   53
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   2
      Left            =   7800
      TabIndex        =   47
      Top             =   120
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   29
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   28
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   21
         Left            =   360
         TabIndex        =   51
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   20
         Left            =   360
         TabIndex        =   50
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   49
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   0
         TabIndex        =   48
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   6
      Left            =   6720
      TabIndex        =   42
      Top             =   1200
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   13
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   12
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   360
         TabIndex        =   46
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   360
         TabIndex        =   45
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   13
         Left            =   720
         TabIndex        =   44
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   12
         Left            =   0
         TabIndex        =   43
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   5
      Left            =   5640
      TabIndex        =   37
      Top             =   1200
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   7
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   6
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   41
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   40
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   720
         TabIndex        =   39
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   0
         TabIndex        =   38
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   1
      Left            =   6720
      TabIndex        =   32
      Top             =   120
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   11
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   10
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   360
         TabIndex        =   36
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   35
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   34
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   33
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Bl 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   0
      Left            =   5640
      TabIndex        =   27
      Top             =   120
      Width           =   1095
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   9
         X1              =   1080
         X2              =   0
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Line Qw 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   7
         Index           =   8
         X1              =   0
         X2              =   1080
         Y1              =   0
         Y2              =   1080
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   31
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Yy 
         Alignment       =   2  'Center
         BackColor       =   &H00004000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   30
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   29
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Xx 
         Alignment       =   2  'Center
         BackColor       =   &H00800080&
         BackStyle       =   0  'Transparent
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   24
      Left            =   9960
      TabIndex        =   176
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   23
      Left            =   8880
      TabIndex        =   175
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   22
      Left            =   7800
      TabIndex        =   174
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   21
      Left            =   6720
      TabIndex        =   173
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   20
      Left            =   5640
      TabIndex        =   172
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   19
      Left            =   9960
      TabIndex        =   171
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   14
      Left            =   9960
      TabIndex        =   170
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   9960
      TabIndex        =   169
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   9960
      TabIndex        =   168
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   18
      Left            =   8880
      TabIndex        =   167
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   17
      Left            =   7800
      TabIndex        =   166
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   16
      Left            =   6720
      TabIndex        =   165
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   15
      Left            =   5640
      TabIndex        =   164
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   13
      Left            =   8880
      TabIndex        =   163
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   8880
      TabIndex        =   162
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   8880
      TabIndex        =   161
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   12
      Left            =   7800
      TabIndex        =   160
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   11
      Left            =   6720
      TabIndex        =   159
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   10
      Left            =   5640
      TabIndex        =   158
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   7
      Left            =   7800
      TabIndex        =   157
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   7800
      TabIndex        =   156
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   6720
      TabIndex        =   155
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   5640
      TabIndex        =   154
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   6720
      TabIndex        =   153
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Ind 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   5640
      TabIndex        =   152
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   2280
      TabIndex        =   26
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   7
      Left            =   2280
      TabIndex        =   25
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   10
      Left            =   120
      TabIndex        =   24
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   11
      Left            =   1200
      TabIndex        =   23
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   12
      Left            =   2280
      TabIndex        =   22
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   3
      Left            =   3360
      TabIndex        =   21
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   8
      Left            =   3360
      TabIndex        =   20
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   13
      Left            =   3360
      TabIndex        =   19
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   15
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   16
      Left            =   1200
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   17
      Left            =   2280
      TabIndex        =   16
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   18
      Left            =   3360
      TabIndex        =   15
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   4
      Left            =   4440
      TabIndex        =   14
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   9
      Left            =   4440
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   14
      Left            =   4440
      TabIndex        =   12
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   19
      Left            =   4440
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   20
      Left            =   120
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   21
      Left            =   1200
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   22
      Left            =   2280
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   23
      Left            =   3360
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   24
      Left            =   4440
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   6
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1095
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Indi 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label g 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Gen 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Menu Tra 
      Caption         =   "Tetravex"
      Begin VB.Menu Nw 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu Nna 
         Caption         =   "Nivel"
         Begin VB.Menu Xx2 
            Caption         =   "2x2"
         End
         Begin VB.Menu Xx3 
            Caption         =   "3x3"
         End
         Begin VB.Menu Xx4 
            Caption         =   "4x4"
         End
         Begin VB.Menu Xx5 
            Caption         =   "5x5"
         End
      End
      Begin VB.Menu Ex 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Hlp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "Trost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lev, Ju, H, Lmn, Kk, K, Kp(0 To 24) As Integer
Dim pp(0 To 24), S, Hti, Hto As Integer

Private Sub Bl_Click(Index As Integer)
S = Index
End Sub

Private Sub Ex_Click()
End
End Sub

Private Sub Form_Load()
H = 24
S = 100
Lev = 2
Gen_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Hlp_Click()
Help.Show
End Sub

Private Sub Ind_Click(Index As Integer)
If (Indi(Index).Top = Bl(Index).Top) Then
    If (Indi(Index).Left = Bl(Index).Left) Then
        If (Index = S) Then
        Ju = Ju - 1
        g.Caption = Str(Ju)
        End If
    End If
End If

If (S <> 100) Then
Bl(S).Top = Ind(Index).Top
Bl(S).Left = Ind(Index).Left
S = 100
End If
End Sub

Private Sub Indi_Click(Index As Integer)
If (S <> 100) Then

If (Indi(S).Top = Bl(S).Top) Then
    If (Indi(S).Left = Bl(S).Left) Then
        Ju = Ju - 1
        g.Caption = Str(Ju)
    End If
End If

Bl(S).Top = Indi(Index).Top
Bl(S).Left = Indi(Index).Left
S = 100
End If

If (Indi(Index).Top = Bl(Index).Top) Then
    If (Indi(Index).Left = Bl(Index).Left) Then
        Ju = Ju + 1
        g.Caption = Str(Ju)
            If (Lev = 5) And (Ju = 25) Then
                Winner
            End If
            If (Lev = 4) And (Ju = 16) Then
                Winner
            End If
            If (Lev = 3) And (Ju = 9) Then
                Winner
            End If
            If (Lev = 2) And (Ju = 4) Then
                Winner
            End If
    End If
End If

End Sub


Private Sub Gen_Click()
Ju = 0
g.Caption = 0
    For K = 0 To 49
        
        Hti = Ran(0, 9)
        Hto = Ran(0, 9)
        
        If (K < 1) Then
        Xx(49).Caption = Hti
        Yy(49).Caption = Hto
        End If
        
        If (K >= 1) Then
        Xx(K - 1).Caption = Hti
        Yy(K - 1).Caption = Hto
        End If
        
        Xx(K).Caption = Hti
        Yy(K).Caption = Hto
       
        K = K + 1
    Next
    
    For K = 0 To 24
    Bl(K).Visible = True
    Ind(K).Visible = True
    Indi(K).Visible = True
    Next
           
    If (Lev = 4) Then
    Lev4
    End If

    If (Lev = 3) Then
    Lev3
    Lev4
    End If
    
    If (Lev = 2) Then
    Lev2
    Lev3
    Lev4
    End If
Ordenar
End Sub

Private Function Lev2()
        Bl(2).Visible = False
        Bl(7).Visible = False
        Bl(12).Visible = False
        Bl(11).Visible = False
        Bl(10).Visible = False

        Ind(2).Visible = False
        Ind(7).Visible = False
        Ind(11).Visible = False
        Ind(12).Visible = False
        Ind(10).Visible = False
        
        Indi(2).Visible = False
        Indi(7).Visible = False
        Indi(11).Visible = False
        Indi(12).Visible = False
        Indi(10).Visible = False
End Function

Private Function Lev3()
        Bl(3).Visible = False
        Bl(8).Visible = False
        Bl(13).Visible = False
        Bl(18).Visible = False
        Bl(23).Visible = False
        Bl(16).Visible = False
        Bl(15).Visible = False
        Bl(17).Visible = False

        Ind(3).Visible = False
        Ind(8).Visible = False
        Ind(13).Visible = False
        Ind(18).Visible = False
        Ind(23).Visible = False
        Ind(16).Visible = False
        Ind(15).Visible = False
        Ind(17).Visible = False
        
        Indi(3).Visible = False
        Indi(8).Visible = False
        Indi(13).Visible = False
        Indi(18).Visible = False
        Indi(23).Visible = False
        Indi(16).Visible = False
        Indi(15).Visible = False
        Indi(17).Visible = False
End Function

Private Function Lev4()
        Bl(4).Visible = False
        Bl(9).Visible = False
        Bl(14).Visible = False
        Bl(19).Visible = False
        Bl(24).Visible = False
        Bl(20).Visible = False
        Bl(21).Visible = False
        Bl(22).Visible = False
        Bl(23).Visible = False

        Ind(4).Visible = False
        Ind(9).Visible = False
        Ind(14).Visible = False
        Ind(19).Visible = False
        Ind(24).Visible = False
        Ind(20).Visible = False
        Ind(21).Visible = False
        Ind(22).Visible = False
        Ind(23).Visible = False
        
        Indi(4).Visible = False
        Indi(9).Visible = False
        Indi(14).Visible = False
        Indi(19).Visible = False
        Indi(24).Visible = False
        Indi(20).Visible = False
        Indi(21).Visible = False
        Indi(22).Visible = False
        Indi(23).Visible = False
End Function

Public Function Winner()
Fam.Show
End Function

Public Function Ran(N1, N2 As Integer) As Integer
Randomize
N1 = N1 - 1
N2 = N2 + 1
Ran = Int((N1 - N2 + 1) * Rnd + N2)
End Function

Public Function Ordenar()
If Lev = 2 Then
Trost.Width = 8040
Trost.Height = 3135
pp(0) = 0
pp(1) = 1

pp(2) = 5
pp(3) = 6
H = 3
End If

If Lev = 3 Then
Trost.Width = 9090
Trost.Height = 4245
pp(0) = 0
pp(1) = 1
pp(2) = 2

pp(3) = 5
pp(4) = 6
pp(5) = 7

pp(6) = 10
pp(7) = 11
pp(8) = 12
H = 8
End If

If Lev = 4 Then
Trost.Width = 10170
Trost.Height = 5280

pp(0) = 0
pp(1) = 1
pp(2) = 2
pp(3) = 3

pp(4) = 5
pp(5) = 6
pp(6) = 7
pp(7) = 8

pp(8) = 10
pp(9) = 11
pp(10) = 12
pp(11) = 13

pp(12) = 15
pp(13) = 16
pp(14) = 17
pp(15) = 18

H = 15
End If

If Lev = 5 Then
Trost.Width = 11280
Trost.Height = 6435
pp(0) = 0
pp(1) = 1
pp(2) = 2
pp(3) = 3
pp(4) = 4

pp(5) = 5
pp(6) = 6
pp(7) = 7
pp(8) = 8
pp(9) = 9

pp(10) = 10
pp(11) = 11
pp(12) = 12
pp(13) = 13
pp(14) = 14

pp(15) = 15
pp(16) = 16
pp(17) = 17
pp(18) = 18
pp(19) = 19

pp(20) = 20
pp(21) = 21
pp(22) = 22
pp(23) = 23
pp(24) = 24
H = 24
End If

For Kk = 0 To 24
    Kp(Kk) = 100
Next

For Kk = 0 To Val(H)
Inic:
    Lmn = Ran(0, Val(H))
    For K = 0 To 24
        If (Lmn = Kp(K)) Then
            GoTo Inic
        End If
    Next
Kp(Kk) = Lmn

Bl(pp(Kk)).Top = Ind(pp(Lmn)).Top
Bl(pp(Kk)).Left = Ind(pp(Lmn)).Left

Next
End Function

Private Sub Nw_Click()
Gen_Click
End Sub

Private Sub Xx2_Click()
Lev = 2
Gen_Click
End Sub

Private Sub Xx3_Click()
Lev = 3
Gen_Click
End Sub

Private Sub Xx4_Click()
Lev = 4
Gen_Click
End Sub

Private Sub Xx5_Click()
Lev = 5
Gen_Click
End Sub
