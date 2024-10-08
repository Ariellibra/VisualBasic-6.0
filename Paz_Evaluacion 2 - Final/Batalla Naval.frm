VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF0000&
   Caption         =   "Form1"
   ClientHeight    =   11325
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   26205
   LinkTopic       =   "Form1"
   ScaleHeight     =   11325
   ScaleWidth      =   26205
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H0080C0FF&
      Height          =   2415
      Left            =   8280
      TabIndex        =   113
      Top             =   4080
      Width           =   2415
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   1320
         TabIndex        =   116
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   115
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H008080FF&
         Caption         =   "Atacar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   114
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   1440
         TabIndex        =   118
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   117
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   7260
      Left            =   19080
      TabIndex        =   111
      Top             =   2640
      Width           =   6855
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17400
      Style           =   1  'Graphical
      TabIndex        =   110
      Top             =   10440
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Height          =   2415
      Left            =   8280
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton Command1 
         BackColor       =   &H008080FF&
         Caption         =   "Atacar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Height          =   7455
      Left            =   11520
      TabIndex        =   1
      Top             =   2640
      Width           =   7215
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   5
         Left            =   6360
         TabIndex        =   106
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   4
         Left            =   5280
         TabIndex        =   105
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   3
         Left            =   4200
         TabIndex        =   104
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   2
         Left            =   3120
         TabIndex        =   103
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   1
         Left            =   2040
         TabIndex        =   102
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   0
         Left            =   960
         TabIndex        =   101
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   100
         Top             =   6480
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   99
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   98
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   97
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   96
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Y\X"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   35
         Left            =   6240
         TabIndex        =   79
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   34
         Left            =   5160
         TabIndex        =   78
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   33
         Left            =   4080
         TabIndex        =   77
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   32
         Left            =   3000
         TabIndex        =   76
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   31
         Left            =   1920
         TabIndex        =   75
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   30
         Left            =   840
         TabIndex        =   74
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   29
         Left            =   6240
         TabIndex        =   73
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   28
         Left            =   5160
         TabIndex        =   72
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   27
         Left            =   4080
         TabIndex        =   71
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   26
         Left            =   3000
         TabIndex        =   70
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   25
         Left            =   1920
         TabIndex        =   69
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   24
         Left            =   840
         TabIndex        =   68
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   23
         Left            =   6240
         TabIndex        =   67
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   22
         Left            =   5160
         TabIndex        =   66
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   21
         Left            =   4080
         TabIndex        =   65
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   20
         Left            =   3000
         TabIndex        =   64
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   19
         Left            =   1920
         TabIndex        =   63
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   18
         Left            =   840
         TabIndex        =   62
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   17
         Left            =   6240
         TabIndex        =   61
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   16
         Left            =   5160
         TabIndex        =   60
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   15
         Left            =   4080
         TabIndex        =   59
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   14
         Left            =   3000
         TabIndex        =   58
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   13
         Left            =   1920
         TabIndex        =   57
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   12
         Left            =   840
         TabIndex        =   56
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   11
         Left            =   6240
         TabIndex        =   55
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   10
         Left            =   5160
         TabIndex        =   54
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   9
         Left            =   4080
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   8
         Left            =   3000
         TabIndex        =   52
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   7
         Left            =   1920
         TabIndex        =   51
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   6
         Left            =   840
         TabIndex        =   50
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   5
         Left            =   6240
         TabIndex        =   49
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   4
         Left            =   5160
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   3
         Left            =   4080
         TabIndex        =   47
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   2
         Left            =   3000
         TabIndex        =   46
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   1
         Left            =   1920
         TabIndex        =   45
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Cpu 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   0
         Left            =   840
         TabIndex        =   44
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   2640
      Width           =   7215
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   5
         Left            =   6360
         TabIndex        =   93
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   4
         Left            =   5280
         TabIndex        =   92
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   3
         Left            =   4200
         TabIndex        =   91
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   2
         Left            =   3120
         TabIndex        =   90
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   1
         Left            =   2040
         TabIndex        =   89
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   5
         Left            =   120
         TabIndex        =   88
         Top             =   6480
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   87
         Top             =   5400
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   86
         Top             =   4320
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   85
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   84
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Y\X"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   83
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   0
         Left            =   960
         TabIndex        =   82
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   81
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Index           =   35
         Left            =   6240
         TabIndex        =   43
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   34
         Left            =   5160
         TabIndex        =   42
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   33
         Left            =   4080
         TabIndex        =   41
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   32
         Left            =   3000
         TabIndex        =   40
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   31
         Left            =   1920
         TabIndex        =   39
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   30
         Left            =   840
         TabIndex        =   38
         Top             =   6360
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   29
         Left            =   6240
         TabIndex        =   37
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   28
         Left            =   5160
         TabIndex        =   36
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   27
         Left            =   4080
         TabIndex        =   35
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   26
         Left            =   3000
         TabIndex        =   34
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   25
         Left            =   1920
         TabIndex        =   33
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   24
         Left            =   840
         TabIndex        =   32
         Top             =   5280
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   23
         Left            =   6240
         TabIndex        =   31
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   22
         Left            =   5160
         TabIndex        =   30
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   21
         Left            =   4080
         TabIndex        =   29
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   20
         Left            =   3000
         TabIndex        =   28
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   19
         Left            =   1920
         TabIndex        =   27
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   18
         Left            =   840
         TabIndex        =   26
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   17
         Left            =   6240
         TabIndex        =   25
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   16
         Left            =   5160
         TabIndex        =   24
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   15
         Left            =   4080
         TabIndex        =   23
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   14
         Left            =   3000
         TabIndex        =   22
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   13
         Left            =   1920
         TabIndex        =   21
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   12
         Left            =   840
         TabIndex        =   20
         Top             =   3120
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   11
         Left            =   6240
         TabIndex        =   19
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   10
         Left            =   5160
         TabIndex        =   18
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   9
         Left            =   4080
         TabIndex        =   17
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   8
         Left            =   3000
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   7
         Left            =   1920
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   6
         Left            =   840
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   5
         Left            =   6240
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   4
         Left            =   5160
         TabIndex        =   12
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   3
         Left            =   4080
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   0
         Left            =   840
         TabIndex        =   9
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Player 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cordenadas Barcos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7800
      TabIndex        =   119
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Movimientos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   19080
      TabIndex        =   112
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "BATALLA NAVAL"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   5880
      TabIndex        =   109
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "Barcos Enemigos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11520
      TabIndex        =   108
      Top             =   1440
      Width           =   7215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Mis Barcos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   107
      Top             =   1440
      Width           =   7215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7800
      TabIndex        =   80
      Top             =   1440
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cordX, cordY, cord, cordCpu, rndCpuX, rndCpuY As Integer
Dim BarcoPlayer(2) As New Barco
Dim BarcoCpu(2) As Barco
Dim n, h, turnos As Integer
Dim turnoActual As String
Dim barcosPlayer, barcosCpu As Integer

Private Sub Command1_Click()
    
    cordX = Text1(0).Text
    cordY = Text1(1).Text
    
    If cordX > 6 Or cordX < 1 Or cordY > 6 Or cordY < 1 Then
        
        MsgBox "Elija una cordenada correcta", vbCritical, "Error"
    
    Else
    
        Ataque
        turnos = turnos + 1
        Score
        
        If barcosCpu = 3 Then
        
            MsgBox "Gano el jugador", vbExclamation, "GANADOR"
            Command1.Enabled = False
            Text1(0).Enabled = False
            Text1(1).Enabled = False
    
        ElseIf barcosPlayer = 3 Then
        
            MsgBox "Gano la Maquina", vbExclamation, "PERDEDOR"
            Command1.Enabled = False
            Text1(0).Enabled = False
            Text1(1).Enabled = False
        
        Else
        
            CpuAtaque
            Score
    
        End If
        
    End If
    
End Sub
Private Sub CpuCrearBarco()
    Dim posicionesOcupadas(2) As Integer
    Dim ocupado As Boolean
    Dim cordCpu As Integer
    Dim i As Integer
    Dim j As Integer
    
    For n = 0 To 2
        Do
            ' Calcular la posición de CPU
            cordCpu = CordenadaRnd
            ocupado = False
            
            ' Verificar si la coordenada ya está ocupada
            For j = 0 To n - 1
                If cordCpu = posicionesOcupadas(j) Then
                    ocupado = True
                    Exit For
                End If
            Next j
        Loop While ocupado
        
        ' Guardar la posición válida
        posicionesOcupadas(n) = cordCpu
        
        ' Inicializar barco de CPU
        Set BarcoCpu(n) = New Barco
        BarcoCpu(n).Constructor "Cpu", cordCpu, False
        
        ' Cambiar el color de fondo del cuadro correspondiente
        Cpu(BarcoCpu(n).GetPosci).BackColor = &H4080&
    Next n
End Sub

Private Function CordenadaRnd() As Integer
        Dim cordFunction As Integer

        Randomize
        
        ' Generar coordenadas aleatorias
        rndCpuX = CInt((Rnd * 5) + 1)
        rndCpuY = CInt((Rnd * 5) + 1)
        
        ' Calcular la posición de CPU
        cordFunction = CInt((6 * (rndCpuY - 1)) + (rndCpuX - 1))
        
        CordenadaRnd = cordFunction

End Function

Private Sub CpuAtaque()
    Dim fallo, repite As Boolean
    repite = True
    fallo = False
    
    MsgBox "Se viene el ataque de la Maquina", vbInformation, "Ataque: Maquina"
    
    Do
        cord = CordenadaRnd
        
        If Player(cord).Caption = "Agua" Or Player(cord).Caption = "Hundido" Then
                
            repite = True
            
        Else
        
            For n = 0 To 2
        
                If BarcoPlayer(n).GetPosci = cord Then
        
                    MsgBox "Hundido", vbCritical, "En el Blanco"
                    BarcoPlayer(n).SetHundido (True)
                    Player(cord).BackColor = &HFF&
                    Player(cord).Caption = "Hundido"
                    fallo = False
                    repite = False
                    n = 2
            
                Else
            
                    fallo = True
            
                End If
        
            Next n
    
            If fallo = True Then
    
                MsgBox "Fallo el tiro", vbInformation, "En el Agua"
                Player(cord).BackColor = &HC00000
                Player(cord).Caption = "Agua"
                repite = False
            
            End If
        
        End If
        
    Loop Until (repite = False)
    
    List1.AddItem ("La Maquina Ataco en la posicion: Y: " & rndCpuY & " X: " & rndCpuX)
    List1.AddItem ("El resultado fue: " & Player(cord).Caption)
    
    turnoActual = "Player"
    
    MsgBox "Se viene el ataque del Jugador", vbInformation, "Ataque: Jugador"
    
    turnos = turnos + 1
    
    
End Sub

Private Sub Score()
    barcosPlayer = 0
    barcosCpu = 0
    
    For n = 0 To 2
        
        If BarcoPlayer(n).GetHundido = True Then
            
            barcosPlayer = barcosPlayer + 1
            
        End If
        
        If BarcoCpu(n).GetHundido = True Then
            
            barcosCpu = barcosCpu + 1
            
        End If
    
    Next n

    Label5.Caption = _
    "Turnos: " & turnos & vbCrLf & _
    "Turno Actual: " & turnoActual & vbCrLf & _
    "Tus Barcos: " & barcosPlayer & "/3 " & vbCrLf & _
    "Barcos Enemigos: " & barcosCpu & "/3 " & vbCrLf
    
    
End Sub

Private Function Ataque()
    Dim fallo As Boolean
    fallo = False
    
    cord = (6 * (cordY - 1)) + (cordX - 1)
    
    For n = 0 To 2
        
        If BarcoCpu(n).GetPosci = cord Then
        
            MsgBox "Hundido", vbCritical, "En el Blanco"
            BarcoCpu(n).SetHundido (True)
            Cpu(cord).BackColor = &HFF&
            Cpu(cord).Caption = "Hundido"
            fallo = False
            n = 2
            
        Else
            
            fallo = True
            
        End If
        
    Next n
    
    If fallo = True Then
    
        MsgBox "Fallo el tiro", vbInformation, "En el Agua"
        Cpu(cord).BackColor = &HC00000
        Cpu(cord).Caption = "Agua"
            
    End If
    
    turnoActual = "Cpu"
    
    List1.AddItem ("El jugador Ataco en la posicion: Y: " & cordY & " X: " & cordX)
    List1.AddItem ("El resultado fue: " & Cpu(cord).Caption)
    
End Function

Public Sub CrearBarcos()
    Dim termino As Boolean
    termino = False
    h = 0
    
    Do
        MsgBox "Elija las cordenadas de sus barcos", vbInformation, "Elejir posiciones"
        
        cordX = Text2(0).Text
        cordY = Text2(1).Text
        cord = CInt((6 * (cordY - 1)) + (cordX - 1))
    
    
        BarcoPlayer(h).Constructor "Player", cord, False
        Player(BarcoPlayer(h).GetPosci).BackColor = &H4080&
        Player(cord).Caption = "Barco"
    
        h = h + 1
        
    
    Loop While (termino = True)
    

End Sub

Private Sub Command2_Click()
    
    cordX = Text2(0).Text
    cordY = Text2(1).Text
    cord = CInt((6 * (cordY - 1)) + (cordX - 1))
    
    
    BarcoPlayer(h).Constructor "Player", cord, False
    Player(BarcoPlayer(h).GetPosci).BackColor = &H4080&
    Player(cord).Caption = "Barco"
    
    h = h + 1

End Sub

Private Sub Command3_Click()

    End

End Sub

Private Sub Command6_Click()

    CpuAtaque
    Score

End Sub

Private Sub Command4_Click()

    cordX = Text1(2).Text
    cordY = Text1(3).Text
    cord = CInt((6 * (cordY - 1)) + (cordX - 1))
    
    If cordX > 6 Or cordX < 1 Or cordY > 6 Or cordY < 1 Then
        
        MsgBox "Elija una cordenada correcta", vbCritical, "Error"
        
    Else
        
        BarcoPlayer(h).Constructor "Player", cord, False
        Player(BarcoPlayer(h).GetPosci).BackColor = &H4080&
        Player(cord).Caption = "Barco"
    
        h = h + 1
        
        If h > 2 Then
        
            Frame5.Visible = False
            Frame3.Visible = True
            Label1.Caption = "Cordenadas Ataque"
            
        End If
    
    End If
    
End Sub

Private Sub Form_Activate()
    
    n = 0
    h = 0
    barcosPlayer = 0
    barcosCpu = 0
    
    turnoActual = "Jugador"
    
    CpuCrearBarco
    
    Label5.Caption = _
    "Turnos: 0" & vbCrLf & _
    "Turno Actual: " & turnoActual & vbCrLf & _
    "Tus Barcos: " & "0/3 " & vbCrLf & _
    "Barcos Enemigos: " & "0/3 " & vbCrLf

End Sub

