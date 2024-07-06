VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   12240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21390
   LinkTopic       =   "Form1"
   ScaleHeight     =   12240
   ScaleWidth      =   21390
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command21 
      BackColor       =   &H0080C0FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   25080
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   13200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command20 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   12000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   12000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   23760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   20520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   23760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   20520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   23760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   20520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Falso"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Verdadero"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H0080FF80&
      Caption         =   "Si"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   11040
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton Command23 
      BackColor       =   &H008080FF&
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6480
      Width           =   2415
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Falicidades, Aprobaste!!"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   19920
      TabIndex        =   34
      Top             =   9960
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hay cinco grupos sanguíneos diferentes"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10440
      TabIndex        =   7
      Top             =   6720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "La ""A"" es la letra más utilizada en el idioma inglés"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10440
      TabIndex        =   9
      Top             =   9960
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "El rugido de un león puede oírse hasta a ocho kilómetros de distancia"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   19920
      TabIndex        =   8
      Top             =   6720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Alexander Fleming descubrió la penicilina"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Francia es el segundo país más grande de Europa"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   19920
      TabIndex        =   5
      Top             =   3480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "El Nilo es el río más largo del mundo"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10440
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "El bíceps es el músculo más fuerte del hombre"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Vicente Aleixandre fue el primer Nobel español"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   19920
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Los caracoles pueden dormir un mes"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10440
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "La batalla de Hastings tuvo lugar en 1066"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Empezar el examen?"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   10440
      TabIndex        =   31
      Top             =   4320
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_Click()

End Sub

Private Sub Command1_Click()

    Label2.Visible = True
    Command1.Visible = False
    Command2.Visible = False
    Command3.Visible = True
    Command4.Visible = True
    Label1.BackColor = &H80FF80
    Label1.Caption = "La batalla de Hastings tuvo lugar en 1066. (Verdadero)"
    Label2.Caption = "Los caracoles pueden dormir un mes"
    
    
End Sub

Private Sub Command2_Click()
    
    Label1.BackColor = &H8080FF
    Label2.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label2.Visible = True
    
End Sub

Private Sub Command4_Click()

    Label3.Visible = True
    Command3.Visible = False
    Command4.Visible = False
    Command5.Visible = True
    Command6.Visible = True
    Label2.BackColor = &H80FF80
    Label2.Caption = "Los caracoles pueden dormir un mes. (Falso) Pueden dormir hasta tres meses"
    Label3.Caption = "Vicente Aleixandre fue el primer Nobel español"
    
End Sub

Private Sub Command3_Click()
    
    Label2.BackColor = &H8080FF
    Label3.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label3.Visible = True
    
End Sub

Private Sub Command6_Click()

    Label4.Visible = True
    Command5.Visible = False
    Command6.Visible = False
    Command7.Visible = True
    Command8.Visible = True
    Label3.BackColor = &H80FF80
    Label3.Caption = "Vicente Aleixandre fue el primer Nobel español. (Falso) Fue Camilo José Cela"
    Label4.Caption = "El bíceps es el músculo más fuerte del hombre"
    
End Sub

Private Sub Command5_Click()
    
    Label3.BackColor = &H8080FF
    Label4.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label4.Visible = True
    
End Sub

Private Sub Command8_Click()

    Label5.Visible = True
    Command7.Visible = False
    Command8.Visible = False
    Command9.Visible = True
    Command10.Visible = True
    Label4.BackColor = &H80FF80
    Label4.Caption = "El bíceps es el músculo más fuerte del hombre. (Falso) Es la lengua."
    Label5.Caption = "El Nilo es el río más largo del mundo"
    
End Sub

Private Sub Command7_Click()
    
    Label4.BackColor = &H8080FF
    Label5.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label5.Visible = True
    
End Sub

Private Sub Command9_Click()

    Label6.Visible = True
    Command9.Visible = False
    Command10.Visible = False
    Command11.Visible = True
    Command12.Visible = True
    Label5.BackColor = &H80FF80
    Label5.Caption = "El Nilo es el río más largo del mundo. (Verdadero)"
    Label6.Caption = "Francia es el segundo país más grande de Europa"
    
End Sub

Private Sub Command10_Click()
    
    Label5.BackColor = &H8080FF
    Label6.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label6.Visible = True
    
End Sub

Private Sub Command11_Click()

    Label7.Visible = True
    Command11.Visible = False
    Command12.Visible = False
    Command13.Visible = True
    Command14.Visible = True
    Label6.BackColor = &H80FF80
    Label6.Caption = "Francia es el segundo país más grande de Europa. (Verdadero)"
    Label7.Caption = "Alexander Fleming descubrió la penicilina"
    
End Sub

Private Sub Command12_Click()
    
    Label6.BackColor = &H8080FF
    Label7.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label7.Visible = True
    
End Sub

Private Sub Command13_Click()

    Label8.Visible = True
    Command14.Visible = False
    Command13.Visible = False
    Command15.Visible = True
    Command16.Visible = True
    Label7.BackColor = &H80FF80
    Label7.Caption = "Alexander Fleming descubrió la penicilina. (Verdadero)"
    Label8.Caption = "Hay cinco grupos sanguíneos diferentes"
    
End Sub

Private Sub Command14_Click()
    
    Label7.BackColor = &H8080FF
    Label8.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label8.Visible = True
    
End Sub

Private Sub Command16_Click()

    Label9.Visible = True
    Command15.Visible = False
    Command16.Visible = False
    Command17.Visible = True
    Command18.Visible = True
    Label8.BackColor = &H80FF80
    Label8.Caption = "Hay cinco grupos sanguíneos diferentes. (Falso) Hay cuatro: A, B, AB y O"
    Label9.Caption = "El rugido de un león puede oírse hasta a ocho kilómetros de distancia"

End Sub

Private Sub Command15_Click()

    Label8.BackColor = &H8080FF
    Label9.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label9.Visible = True

End Sub

Private Sub Command17_Click()

    Label10.Visible = True
    Command17.Visible = False
    Command18.Visible = False
    Command19.Visible = True
    Command20.Visible = True
    Label9.BackColor = &H80FF80
    Label9.Caption = "El rugido de un león puede oírse hasta a ocho kilómetros de distancia. (Verdadero)"
    Label10.Caption = "La 'A' es la letra más utilizada en el idioma inglés"

End Sub

Private Sub Command18_Click()

    Label9.BackColor = &H8080FF
    Label10.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label10.Visible = True

End Sub

Private Sub Command20_Click()

    Label10.Visible = True
    Command19.Visible = False
    Command20.Visible = False
    Label10.BackColor = &H80FF80
    Label10.Height = 2895
    Label10.Caption = "La 'A' es la letra más utilizada en el idioma inglés. (Falso) La 'E' es la letra más común y aparece en el 11% de las palabras inglesas, según los diccionarios Oxford."
    Label12.Caption = "Feliciades completaste el examen!!"
    Label12.Visible = True
    Label12.BackColor = &HC0FFFF

End Sub

Private Sub Command19_Click()

    Label10.BackColor = &H8080FF
    Label12.Caption = "Respueta incorrecta, intentelo de nuevo"
    Label12.Visible = True

End Sub

Private Sub Command21_Click()
    End
End Sub

Private Sub Command22_Click()

    Label1.Visible = True
    Command1.Visible = True
    Command2.Visible = True
    Label11.Visible = False
    Command22.Visible = False
    Command23.Visible = False
    Command21.Visible = True
    
End Sub

Private Sub Command23_Click()
    End
End Sub


