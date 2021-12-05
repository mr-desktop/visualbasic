VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00400000&
   Caption         =   "Busca Minas"
   ClientHeight    =   8145
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   11040
   Icon            =   "Busca Mina.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11040
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox ooo 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1950
      Left            =   240
      TabIndex        =   483
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Tt2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   0
   End
   Begin VB.Timer Tt1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   0
   End
   Begin VB.Label Counter 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Index           =   1
      Left            =   2400
      TabIndex        =   482
      ToolTipText     =   "Número de Minas"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Counter 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Index           =   0
      Left            =   7200
      TabIndex        =   481
      ToolTipText     =   "Tiempo"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Mensaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Loser"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   3720
      TabIndex        =   480
      Top             =   5040
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Image Lost 
      BorderStyle     =   1  'Fixed Single
      Height          =   2925
      Left            =   2400
      Picture         =   "Busca Mina.frx":1CFA
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   479
      Left            =   10560
      TabIndex        =   479
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   478
      Left            =   10200
      TabIndex        =   478
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   477
      Left            =   9840
      TabIndex        =   477
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   476
      Left            =   9480
      TabIndex        =   476
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   475
      Left            =   9120
      TabIndex        =   475
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   474
      Left            =   8760
      TabIndex        =   474
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   473
      Left            =   8400
      TabIndex        =   473
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   472
      Left            =   8040
      TabIndex        =   472
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   471
      Left            =   7680
      TabIndex        =   471
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   470
      Left            =   7320
      TabIndex        =   470
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   469
      Left            =   6960
      TabIndex        =   469
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   468
      Left            =   6600
      TabIndex        =   468
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   467
      Left            =   6240
      TabIndex        =   467
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   466
      Left            =   5880
      TabIndex        =   466
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   465
      Left            =   5520
      TabIndex        =   465
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   464
      Left            =   5160
      TabIndex        =   464
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   463
      Left            =   4800
      TabIndex        =   463
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   462
      Left            =   4440
      TabIndex        =   462
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   461
      Left            =   4080
      TabIndex        =   461
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   460
      Left            =   3720
      TabIndex        =   460
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   459
      Left            =   3360
      TabIndex        =   459
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   458
      Left            =   120
      TabIndex        =   458
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   457
      Left            =   480
      TabIndex        =   457
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   456
      Left            =   840
      TabIndex        =   456
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   455
      Left            =   1200
      TabIndex        =   455
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   454
      Left            =   1560
      TabIndex        =   454
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   453
      Left            =   1920
      TabIndex        =   453
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   452
      Left            =   2280
      TabIndex        =   452
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   451
      Left            =   2640
      TabIndex        =   451
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   450
      Left            =   3000
      TabIndex        =   450
      Tag             =   "0"
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   449
      Left            =   10560
      TabIndex        =   449
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   448
      Left            =   10200
      TabIndex        =   448
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   447
      Left            =   9840
      TabIndex        =   447
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   446
      Left            =   9480
      TabIndex        =   446
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   445
      Left            =   9120
      TabIndex        =   445
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   444
      Left            =   8760
      TabIndex        =   444
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   443
      Left            =   8400
      TabIndex        =   443
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   442
      Left            =   8040
      TabIndex        =   442
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   441
      Left            =   7680
      TabIndex        =   441
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   440
      Left            =   7320
      TabIndex        =   440
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   439
      Left            =   6960
      TabIndex        =   439
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   438
      Left            =   6600
      TabIndex        =   438
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   437
      Left            =   6240
      TabIndex        =   437
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   436
      Left            =   5880
      TabIndex        =   436
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   435
      Left            =   5520
      TabIndex        =   435
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   434
      Left            =   5160
      TabIndex        =   434
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   433
      Left            =   4800
      TabIndex        =   433
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   432
      Left            =   4440
      TabIndex        =   432
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   431
      Left            =   4080
      TabIndex        =   431
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   430
      Left            =   3720
      TabIndex        =   430
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   429
      Left            =   3360
      TabIndex        =   429
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   428
      Left            =   120
      TabIndex        =   428
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   427
      Left            =   480
      TabIndex        =   427
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   426
      Left            =   840
      TabIndex        =   426
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   425
      Left            =   1200
      TabIndex        =   425
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   424
      Left            =   1560
      TabIndex        =   424
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   423
      Left            =   1920
      TabIndex        =   423
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   422
      Left            =   2280
      TabIndex        =   422
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   421
      Left            =   2640
      TabIndex        =   421
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   420
      Left            =   3000
      TabIndex        =   420
      Tag             =   "0"
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   419
      Left            =   10560
      TabIndex        =   419
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   418
      Left            =   10200
      TabIndex        =   418
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   417
      Left            =   9840
      TabIndex        =   417
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   416
      Left            =   9480
      TabIndex        =   416
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   415
      Left            =   9120
      TabIndex        =   415
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   414
      Left            =   8760
      TabIndex        =   414
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   413
      Left            =   8400
      TabIndex        =   413
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   412
      Left            =   8040
      TabIndex        =   412
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   411
      Left            =   7680
      TabIndex        =   411
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   410
      Left            =   7320
      TabIndex        =   410
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   409
      Left            =   6960
      TabIndex        =   409
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   408
      Left            =   6600
      TabIndex        =   408
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   407
      Left            =   6240
      TabIndex        =   407
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   406
      Left            =   5880
      TabIndex        =   406
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   405
      Left            =   5520
      TabIndex        =   405
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   404
      Left            =   5160
      TabIndex        =   404
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   403
      Left            =   4800
      TabIndex        =   403
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   402
      Left            =   4440
      TabIndex        =   402
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   401
      Left            =   4080
      TabIndex        =   401
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   400
      Left            =   3720
      TabIndex        =   400
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   399
      Left            =   3360
      TabIndex        =   399
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   398
      Left            =   120
      TabIndex        =   398
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   397
      Left            =   480
      TabIndex        =   397
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   396
      Left            =   840
      TabIndex        =   396
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   395
      Left            =   1200
      TabIndex        =   395
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   394
      Left            =   1560
      TabIndex        =   394
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   393
      Left            =   1920
      TabIndex        =   393
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   392
      Left            =   2280
      TabIndex        =   392
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   391
      Left            =   2640
      TabIndex        =   391
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   390
      Left            =   3000
      TabIndex        =   390
      Tag             =   "0"
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   389
      Left            =   10560
      TabIndex        =   389
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   388
      Left            =   10200
      TabIndex        =   388
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   387
      Left            =   9840
      TabIndex        =   387
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   386
      Left            =   9480
      TabIndex        =   386
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   385
      Left            =   9120
      TabIndex        =   385
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   384
      Left            =   8760
      TabIndex        =   384
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   383
      Left            =   8400
      TabIndex        =   383
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   382
      Left            =   8040
      TabIndex        =   382
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   381
      Left            =   7680
      TabIndex        =   381
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   380
      Left            =   7320
      TabIndex        =   380
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   379
      Left            =   6960
      TabIndex        =   379
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   378
      Left            =   6600
      TabIndex        =   378
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   377
      Left            =   6240
      TabIndex        =   377
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   376
      Left            =   5880
      TabIndex        =   376
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   375
      Left            =   5520
      TabIndex        =   375
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   374
      Left            =   5160
      TabIndex        =   374
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   373
      Left            =   4800
      TabIndex        =   373
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   372
      Left            =   4440
      TabIndex        =   372
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   371
      Left            =   4080
      TabIndex        =   371
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   370
      Left            =   3720
      TabIndex        =   370
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   369
      Left            =   3360
      TabIndex        =   369
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   368
      Left            =   120
      TabIndex        =   368
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   367
      Left            =   480
      TabIndex        =   367
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   366
      Left            =   840
      TabIndex        =   366
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   365
      Left            =   1200
      TabIndex        =   365
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   364
      Left            =   1560
      TabIndex        =   364
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   363
      Left            =   1920
      TabIndex        =   363
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   362
      Left            =   2280
      TabIndex        =   362
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   361
      Left            =   2640
      TabIndex        =   361
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   360
      Left            =   3000
      TabIndex        =   360
      Tag             =   "0"
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   359
      Left            =   10560
      TabIndex        =   359
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   358
      Left            =   10200
      TabIndex        =   358
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   357
      Left            =   9840
      TabIndex        =   357
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   356
      Left            =   9480
      TabIndex        =   356
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   355
      Left            =   9120
      TabIndex        =   355
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   354
      Left            =   8760
      TabIndex        =   354
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   353
      Left            =   8400
      TabIndex        =   353
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   352
      Left            =   8040
      TabIndex        =   352
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   351
      Left            =   7680
      TabIndex        =   351
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   350
      Left            =   7320
      TabIndex        =   350
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   349
      Left            =   6960
      TabIndex        =   349
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   348
      Left            =   6600
      TabIndex        =   348
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   347
      Left            =   6240
      TabIndex        =   347
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   346
      Left            =   5880
      TabIndex        =   346
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   345
      Left            =   5520
      TabIndex        =   345
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   344
      Left            =   5160
      TabIndex        =   344
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   343
      Left            =   4800
      TabIndex        =   343
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   342
      Left            =   4440
      TabIndex        =   342
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   341
      Left            =   4080
      TabIndex        =   341
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   340
      Left            =   3720
      TabIndex        =   340
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   339
      Left            =   3360
      TabIndex        =   339
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   338
      Left            =   120
      TabIndex        =   338
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   337
      Left            =   480
      TabIndex        =   337
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   336
      Left            =   840
      TabIndex        =   336
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   335
      Left            =   1200
      TabIndex        =   335
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   334
      Left            =   1560
      TabIndex        =   334
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   333
      Left            =   1920
      TabIndex        =   333
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   332
      Left            =   2280
      TabIndex        =   332
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   331
      Left            =   2640
      TabIndex        =   331
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   330
      Left            =   3000
      TabIndex        =   330
      Tag             =   "0"
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   329
      Left            =   10560
      TabIndex        =   329
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   328
      Left            =   10200
      TabIndex        =   328
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   327
      Left            =   9840
      TabIndex        =   327
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   326
      Left            =   9480
      TabIndex        =   326
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   325
      Left            =   9120
      TabIndex        =   325
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   324
      Left            =   8760
      TabIndex        =   324
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   323
      Left            =   8400
      TabIndex        =   323
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   322
      Left            =   8040
      TabIndex        =   322
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   321
      Left            =   7680
      TabIndex        =   321
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   320
      Left            =   7320
      TabIndex        =   320
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   319
      Left            =   6960
      TabIndex        =   319
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   318
      Left            =   6600
      TabIndex        =   318
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   317
      Left            =   6240
      TabIndex        =   317
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   316
      Left            =   5880
      TabIndex        =   316
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   315
      Left            =   5520
      TabIndex        =   315
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   314
      Left            =   5160
      TabIndex        =   314
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   313
      Left            =   4800
      TabIndex        =   313
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   312
      Left            =   4440
      TabIndex        =   312
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   311
      Left            =   4080
      TabIndex        =   311
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   310
      Left            =   3720
      TabIndex        =   310
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   309
      Left            =   3360
      TabIndex        =   309
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   308
      Left            =   120
      TabIndex        =   308
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   307
      Left            =   480
      TabIndex        =   307
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   306
      Left            =   840
      TabIndex        =   306
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   305
      Left            =   1200
      TabIndex        =   305
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   304
      Left            =   1560
      TabIndex        =   304
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   303
      Left            =   1920
      TabIndex        =   303
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   302
      Left            =   2280
      TabIndex        =   302
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   301
      Left            =   2640
      TabIndex        =   301
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   300
      Left            =   3000
      TabIndex        =   300
      Tag             =   "0"
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   299
      Left            =   10560
      TabIndex        =   299
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   298
      Left            =   10200
      TabIndex        =   298
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   297
      Left            =   9840
      TabIndex        =   297
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   296
      Left            =   9480
      TabIndex        =   296
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   295
      Left            =   9120
      TabIndex        =   295
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   294
      Left            =   8760
      TabIndex        =   294
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   293
      Left            =   8400
      TabIndex        =   293
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   292
      Left            =   8040
      TabIndex        =   292
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   291
      Left            =   7680
      TabIndex        =   291
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   290
      Left            =   7320
      TabIndex        =   290
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   289
      Left            =   6960
      TabIndex        =   289
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   288
      Left            =   6600
      TabIndex        =   288
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   287
      Left            =   6240
      TabIndex        =   287
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   286
      Left            =   5880
      TabIndex        =   286
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   285
      Left            =   5520
      TabIndex        =   285
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   284
      Left            =   5160
      TabIndex        =   284
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   283
      Left            =   4800
      TabIndex        =   283
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   282
      Left            =   4440
      TabIndex        =   282
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   281
      Left            =   4080
      TabIndex        =   281
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   280
      Left            =   3720
      TabIndex        =   280
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   279
      Left            =   3360
      TabIndex        =   279
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   278
      Left            =   120
      TabIndex        =   278
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   277
      Left            =   480
      TabIndex        =   277
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   276
      Left            =   840
      TabIndex        =   276
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   275
      Left            =   1200
      TabIndex        =   275
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   274
      Left            =   1560
      TabIndex        =   274
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   273
      Left            =   1920
      TabIndex        =   273
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   272
      Left            =   2280
      TabIndex        =   272
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   271
      Left            =   2640
      TabIndex        =   271
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   270
      Left            =   3000
      TabIndex        =   270
      Tag             =   "0"
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   269
      Left            =   10560
      TabIndex        =   269
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   268
      Left            =   10200
      TabIndex        =   268
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   267
      Left            =   9840
      TabIndex        =   267
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   266
      Left            =   9480
      TabIndex        =   266
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   265
      Left            =   9120
      TabIndex        =   265
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   264
      Left            =   8760
      TabIndex        =   264
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   263
      Left            =   8400
      TabIndex        =   263
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   262
      Left            =   8040
      TabIndex        =   262
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   261
      Left            =   7680
      TabIndex        =   261
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   260
      Left            =   7320
      TabIndex        =   260
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   259
      Left            =   6960
      TabIndex        =   259
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   258
      Left            =   6600
      TabIndex        =   258
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   257
      Left            =   6240
      TabIndex        =   257
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   256
      Left            =   5880
      TabIndex        =   256
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   255
      Left            =   5520
      TabIndex        =   255
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   254
      Left            =   5160
      TabIndex        =   254
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   253
      Left            =   4800
      TabIndex        =   253
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   252
      Left            =   4440
      TabIndex        =   252
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   251
      Left            =   4080
      TabIndex        =   251
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   250
      Left            =   3720
      TabIndex        =   250
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   249
      Left            =   3360
      TabIndex        =   249
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   248
      Left            =   120
      TabIndex        =   248
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   247
      Left            =   480
      TabIndex        =   247
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   246
      Left            =   840
      TabIndex        =   246
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   245
      Left            =   1200
      TabIndex        =   245
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   244
      Left            =   1560
      TabIndex        =   244
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   243
      Left            =   1920
      TabIndex        =   243
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   242
      Left            =   2280
      TabIndex        =   242
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   241
      Left            =   2640
      TabIndex        =   241
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   240
      Left            =   3000
      TabIndex        =   240
      Tag             =   "0"
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   239
      Left            =   10560
      TabIndex        =   239
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   238
      Left            =   10200
      TabIndex        =   238
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   237
      Left            =   9840
      TabIndex        =   237
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   236
      Left            =   9480
      TabIndex        =   236
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   235
      Left            =   9120
      TabIndex        =   235
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   234
      Left            =   8760
      TabIndex        =   234
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   233
      Left            =   8400
      TabIndex        =   233
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   232
      Left            =   8040
      TabIndex        =   232
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   231
      Left            =   7680
      TabIndex        =   231
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   230
      Left            =   7320
      TabIndex        =   230
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   229
      Left            =   6960
      TabIndex        =   229
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   228
      Left            =   6600
      TabIndex        =   228
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   227
      Left            =   6240
      TabIndex        =   227
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   226
      Left            =   5880
      TabIndex        =   226
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   225
      Left            =   5520
      TabIndex        =   225
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   224
      Left            =   5160
      TabIndex        =   224
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   223
      Left            =   4800
      TabIndex        =   223
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   222
      Left            =   4440
      TabIndex        =   222
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   221
      Left            =   4080
      TabIndex        =   221
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   220
      Left            =   3720
      TabIndex        =   220
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   219
      Left            =   3360
      TabIndex        =   219
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   218
      Left            =   120
      TabIndex        =   218
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   217
      Left            =   480
      TabIndex        =   217
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   216
      Left            =   840
      TabIndex        =   216
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   215
      Left            =   1200
      TabIndex        =   215
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   214
      Left            =   1560
      TabIndex        =   214
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   213
      Left            =   1920
      TabIndex        =   213
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   212
      Left            =   2280
      TabIndex        =   212
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   211
      Left            =   2640
      TabIndex        =   211
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   210
      Left            =   3000
      TabIndex        =   210
      Tag             =   "0"
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   209
      Left            =   10560
      TabIndex        =   209
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   208
      Left            =   10200
      TabIndex        =   208
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   207
      Left            =   9840
      TabIndex        =   207
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   206
      Left            =   9480
      TabIndex        =   206
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   205
      Left            =   9120
      TabIndex        =   205
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   204
      Left            =   8760
      TabIndex        =   204
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   203
      Left            =   8400
      TabIndex        =   203
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   202
      Left            =   8040
      TabIndex        =   202
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   201
      Left            =   7680
      TabIndex        =   201
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   200
      Left            =   7320
      TabIndex        =   200
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   199
      Left            =   6960
      TabIndex        =   199
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   198
      Left            =   6600
      TabIndex        =   198
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   197
      Left            =   6240
      TabIndex        =   197
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   196
      Left            =   5880
      TabIndex        =   196
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   195
      Left            =   5520
      TabIndex        =   195
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   194
      Left            =   5160
      TabIndex        =   194
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   193
      Left            =   4800
      TabIndex        =   193
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   192
      Left            =   4440
      TabIndex        =   192
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   191
      Left            =   4080
      TabIndex        =   191
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   190
      Left            =   3720
      TabIndex        =   190
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   189
      Left            =   3360
      TabIndex        =   189
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   188
      Left            =   120
      TabIndex        =   188
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   187
      Left            =   480
      TabIndex        =   187
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   186
      Left            =   840
      TabIndex        =   186
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   185
      Left            =   1200
      TabIndex        =   185
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   184
      Left            =   1560
      TabIndex        =   184
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   183
      Left            =   1920
      TabIndex        =   183
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   182
      Left            =   2280
      TabIndex        =   182
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   181
      Left            =   2640
      TabIndex        =   181
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   180
      Left            =   3000
      TabIndex        =   180
      Tag             =   "0"
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   179
      Left            =   10560
      TabIndex        =   179
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   178
      Left            =   10200
      TabIndex        =   178
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   177
      Left            =   9840
      TabIndex        =   177
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   176
      Left            =   9480
      TabIndex        =   176
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   175
      Left            =   9120
      TabIndex        =   175
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   174
      Left            =   8760
      TabIndex        =   174
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   173
      Left            =   8400
      TabIndex        =   173
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   172
      Left            =   8040
      TabIndex        =   172
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   171
      Left            =   7680
      TabIndex        =   171
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   170
      Left            =   7320
      TabIndex        =   170
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   169
      Left            =   6960
      TabIndex        =   169
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   168
      Left            =   6600
      TabIndex        =   168
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   167
      Left            =   6240
      TabIndex        =   167
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   166
      Left            =   5880
      TabIndex        =   166
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   165
      Left            =   5520
      TabIndex        =   165
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   164
      Left            =   5160
      TabIndex        =   164
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   163
      Left            =   4800
      TabIndex        =   163
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   162
      Left            =   4440
      TabIndex        =   162
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   161
      Left            =   4080
      TabIndex        =   161
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   160
      Left            =   3720
      TabIndex        =   160
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   159
      Left            =   3360
      TabIndex        =   159
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   158
      Left            =   120
      TabIndex        =   158
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   157
      Left            =   480
      TabIndex        =   157
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   156
      Left            =   840
      TabIndex        =   156
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   155
      Left            =   1200
      TabIndex        =   155
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   154
      Left            =   1560
      TabIndex        =   154
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   153
      Left            =   1920
      TabIndex        =   153
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   152
      Left            =   2280
      TabIndex        =   152
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   151
      Left            =   2640
      TabIndex        =   151
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   150
      Left            =   3000
      TabIndex        =   150
      Tag             =   "0"
      Top             =   4080
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   149
      Left            =   10560
      TabIndex        =   149
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   148
      Left            =   10200
      TabIndex        =   148
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   147
      Left            =   9840
      TabIndex        =   147
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   146
      Left            =   9480
      TabIndex        =   146
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   145
      Left            =   9120
      TabIndex        =   145
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   144
      Left            =   8760
      TabIndex        =   144
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   143
      Left            =   8400
      TabIndex        =   143
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   142
      Left            =   8040
      TabIndex        =   142
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   141
      Left            =   7680
      TabIndex        =   141
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   140
      Left            =   7320
      TabIndex        =   140
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   139
      Left            =   6960
      TabIndex        =   139
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   138
      Left            =   6600
      TabIndex        =   138
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   137
      Left            =   6240
      TabIndex        =   137
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   136
      Left            =   5880
      TabIndex        =   136
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   135
      Left            =   5520
      TabIndex        =   135
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   134
      Left            =   5160
      TabIndex        =   134
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   133
      Left            =   4800
      TabIndex        =   133
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   132
      Left            =   4440
      TabIndex        =   132
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   131
      Left            =   4080
      TabIndex        =   131
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   130
      Left            =   3720
      TabIndex        =   130
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   129
      Left            =   3360
      TabIndex        =   129
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   128
      Left            =   120
      TabIndex        =   128
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   127
      Left            =   480
      TabIndex        =   127
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   126
      Left            =   840
      TabIndex        =   126
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   125
      Left            =   1200
      TabIndex        =   125
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   124
      Left            =   1560
      TabIndex        =   124
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   123
      Left            =   1920
      TabIndex        =   123
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   122
      Left            =   2280
      TabIndex        =   122
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   121
      Left            =   2640
      TabIndex        =   121
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   120
      Left            =   3000
      TabIndex        =   120
      Tag             =   "0"
      Top             =   3720
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   119
      Left            =   10560
      TabIndex        =   119
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   118
      Left            =   10200
      TabIndex        =   118
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   117
      Left            =   9840
      TabIndex        =   117
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   116
      Left            =   9480
      TabIndex        =   116
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   115
      Left            =   9120
      TabIndex        =   115
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   114
      Left            =   8760
      TabIndex        =   114
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   113
      Left            =   8400
      TabIndex        =   113
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   112
      Left            =   8040
      TabIndex        =   112
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   111
      Left            =   7680
      TabIndex        =   111
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   110
      Left            =   7320
      TabIndex        =   110
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   109
      Left            =   6960
      TabIndex        =   109
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   108
      Left            =   6600
      TabIndex        =   108
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   107
      Left            =   6240
      TabIndex        =   107
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   106
      Left            =   5880
      TabIndex        =   106
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   105
      Left            =   5520
      TabIndex        =   105
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   104
      Left            =   5160
      TabIndex        =   104
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   103
      Left            =   4800
      TabIndex        =   103
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   102
      Left            =   4440
      TabIndex        =   102
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   101
      Left            =   4080
      TabIndex        =   101
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   100
      Left            =   3720
      TabIndex        =   100
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   99
      Left            =   3360
      TabIndex        =   99
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   98
      Left            =   120
      TabIndex        =   98
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   97
      Left            =   480
      TabIndex        =   97
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   96
      Left            =   840
      TabIndex        =   96
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   95
      Left            =   1200
      TabIndex        =   95
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   94
      Left            =   1560
      TabIndex        =   94
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   93
      Left            =   1920
      TabIndex        =   93
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   92
      Left            =   2280
      TabIndex        =   92
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   91
      Left            =   2640
      TabIndex        =   91
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   90
      Left            =   3000
      TabIndex        =   90
      Tag             =   "0"
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   89
      Left            =   10560
      TabIndex        =   89
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   88
      Left            =   10200
      TabIndex        =   88
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   87
      Left            =   9840
      TabIndex        =   87
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   86
      Left            =   9480
      TabIndex        =   86
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   85
      Left            =   9120
      TabIndex        =   85
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   84
      Left            =   8760
      TabIndex        =   84
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   83
      Left            =   8400
      TabIndex        =   83
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   82
      Left            =   8040
      TabIndex        =   82
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   81
      Left            =   7680
      TabIndex        =   81
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   80
      Left            =   7320
      TabIndex        =   80
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   79
      Left            =   6960
      TabIndex        =   79
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   78
      Left            =   6600
      TabIndex        =   78
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   77
      Left            =   6240
      TabIndex        =   77
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   76
      Left            =   5880
      TabIndex        =   76
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   75
      Left            =   5520
      TabIndex        =   75
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   74
      Left            =   5160
      TabIndex        =   74
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   73
      Left            =   4800
      TabIndex        =   73
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   72
      Left            =   4440
      TabIndex        =   72
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   71
      Left            =   4080
      TabIndex        =   71
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   70
      Left            =   3720
      TabIndex        =   70
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   69
      Left            =   3360
      TabIndex        =   69
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   68
      Left            =   120
      TabIndex        =   68
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   67
      Left            =   480
      TabIndex        =   67
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   66
      Left            =   840
      TabIndex        =   66
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   65
      Left            =   1200
      TabIndex        =   65
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   64
      Left            =   1560
      TabIndex        =   64
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   63
      Left            =   1920
      TabIndex        =   63
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   62
      Left            =   2280
      TabIndex        =   62
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   61
      Left            =   2640
      TabIndex        =   61
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   60
      Left            =   3000
      TabIndex        =   60
      Tag             =   "0"
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   59
      Left            =   10560
      TabIndex        =   59
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   58
      Left            =   10200
      TabIndex        =   58
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   57
      Left            =   9840
      TabIndex        =   57
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   56
      Left            =   9480
      TabIndex        =   56
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   55
      Left            =   9120
      TabIndex        =   55
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   54
      Left            =   8760
      TabIndex        =   54
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   53
      Left            =   8400
      TabIndex        =   53
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   52
      Left            =   8040
      TabIndex        =   52
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   51
      Left            =   7680
      TabIndex        =   51
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   50
      Left            =   7320
      TabIndex        =   50
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   49
      Left            =   6960
      TabIndex        =   49
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   48
      Left            =   6600
      TabIndex        =   48
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   47
      Left            =   6240
      TabIndex        =   47
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   46
      Left            =   5880
      TabIndex        =   46
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   45
      Left            =   5520
      TabIndex        =   45
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   44
      Left            =   5160
      TabIndex        =   44
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   43
      Left            =   4800
      TabIndex        =   43
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   42
      Left            =   4440
      TabIndex        =   42
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   41
      Left            =   4080
      TabIndex        =   41
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   40
      Left            =   3720
      TabIndex        =   40
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   39
      Left            =   3360
      TabIndex        =   39
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   38
      Left            =   120
      TabIndex        =   38
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   37
      Left            =   480
      TabIndex        =   37
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   36
      Left            =   840
      TabIndex        =   36
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   35
      Left            =   1200
      TabIndex        =   35
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   34
      Left            =   1560
      TabIndex        =   34
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   33
      Left            =   1920
      TabIndex        =   33
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   32
      Left            =   2280
      TabIndex        =   32
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   31
      Left            =   2640
      TabIndex        =   31
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   30
      Left            =   3000
      TabIndex        =   30
      Tag             =   "0"
      Top             =   2640
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   29
      Left            =   10560
      TabIndex        =   29
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   28
      Left            =   10200
      TabIndex        =   28
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   27
      Left            =   9840
      TabIndex        =   27
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   26
      Left            =   9480
      TabIndex        =   26
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   25
      Left            =   9120
      TabIndex        =   25
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   24
      Left            =   8760
      TabIndex        =   24
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   23
      Left            =   8400
      TabIndex        =   23
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   22
      Left            =   8040
      TabIndex        =   22
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   21
      Left            =   7680
      TabIndex        =   21
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   20
      Left            =   7320
      TabIndex        =   20
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   19
      Left            =   6960
      TabIndex        =   19
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   18
      Left            =   6600
      TabIndex        =   18
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   17
      Left            =   6240
      TabIndex        =   17
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   5880
      TabIndex        =   16
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   5520
      TabIndex        =   15
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   5160
      TabIndex        =   14
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   4800
      TabIndex        =   13
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   4440
      TabIndex        =   12
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   4080
      TabIndex        =   11
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   3720
      TabIndex        =   10
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   3360
      TabIndex        =   9
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   3000
      TabIndex        =   8
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   2640
      TabIndex        =   7
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   1920
      TabIndex        =   5
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   1560
      TabIndex        =   4
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   840
      TabIndex        =   2
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Label Terr 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "0"
      Top             =   2280
      Width           =   345
   End
   Begin VB.Image T1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1995
      Left            =   4680
      Picture         =   "Busca Mina.frx":F208
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1995
   End
   Begin VB.Image T2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1995
      Left            =   4680
      Picture         =   "Busca Mina.frx":12045
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1995
   End
   Begin VB.Menu KKl 
      Caption         =   "Nivel"
      Begin VB.Menu PpY 
         Caption         =   "Principiante"
      End
      Begin VB.Menu Inter 
         Caption         =   "Intermedio"
      End
      Begin VB.Menu Juio 
         Caption         =   "Experto"
      End
   End
   Begin VB.Menu Hlp 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Experto 30x16 90m
'Medio 16x16 40m
'Principiante 9x9 10m
Dim C, Cc, C2, Ct, L, W, S(1 To 90), Nmin, Ll As Integer

Private Sub Form_Load()
L = 3
Nmin = 90
T1_Click
End Sub

Public Function Ran(N1, N2 As Integer) As Integer
Randomize
N1 = N1 - 1
N2 = N2 + 1
Ran = Int((N1 - N2 + 1) * Rnd + N2)
End Function

Public Function GMinas()
For C = 1 To 90
S(C) = 600
Next
    For Cc = 1 To Val(Nmin)
Inicio:
        Ll = Ran(0, 479)
        For C = 1 To Nmin
            If (Ll = Val(S(C))) Then
                GoTo Inicio
            End If
        Next
        S(Cc) = Ll
        ooo.AddItem (Str(Ll))
        Terr(Ll).Tag = "1"
Next

End Function

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Hlp_Click()
Help.Show
End Sub

Private Sub Inter_Click()
Nmin = 40
T1_Click
End Sub

Private Sub Juio_Click()
Nmin = 90
T1_Click
End Sub

Private Sub PpY_Click()
Nmin = 10
T1_Click
End Sub

Private Sub T1_Click()
ooo.Clear
GMinas
W = 0
Lost.Visible = 0
C2 = 0
Mensaje.Visible = 0
T1.Visible = 1
Ct = 0
Tt2.Enabled = 0
Counter(0).Caption = 0
Counter(1).Caption = 0
For C = 0 To 479
Terr(C).BackColor = &H4000&
Terr(C).Caption = ""
Terr(C).Enabled = True
Terr(C).BorderStyle = 1
Next
End Sub

Private Sub T2_Click()
T1_Click
End Sub

Private Sub Terr_Click(Index As Integer)
If (Terr(Index).Tag <> "1") Then C2 = C2 + 1

Counter(1).Caption = Str(C2)

T1.Visible = 0
Tt1.Enabled = 1
Tt2.Enabled = 1

Terr(Index).Enabled = False
Terr(Index).BorderStyle = 0

If (Terr(Index).Tag = "1") Then
Loser
Terr(Index).BackColor = &H80FF&
Terr(Index).Caption = "x"
End If

If (Terr(Index).Tag <> "1") Then
Terr(Index).BackColor = &HFF8080
End If

If (C2 = Nmin) Then
Ganas
End If

End Sub

Private Sub tt1_Timer()
T1.Visible = 1
Tt1.Enabled = 0
End Sub


Function Loser()
W = 1

For C = 1 To Nmin
Terr(S(C)).BackColor = &H80FF&
Terr(S(C)).Caption = "x"
Next

Tt1.Enabled = False
Lost.Visible = True
Mensaje.Visible = True
T1.Visible = 0
Ct = 0
Tt2.Enabled = 0

For C = 0 To 479
Terr(C).Enabled = False
Next

End Function

Function Ganas()

End Function

Private Sub Tt2_Timer()
Ct = Ct + 1
Counter(0).Caption = Ct
End Sub
