VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00008000&
   Caption         =   "Form2"
   ClientHeight    =   11310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   24750
   LinkTopic       =   "Form2"
   ScaleHeight     =   11310
   ScaleWidth      =   24750
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command18 
      BackColor       =   &H0080C0FF&
      Caption         =   "Limpiar Carrito"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   20400
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ver Carrito"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   18600
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mostrar Resumen"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0080C0FF&
      Caption         =   "Finalizar Compra"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   22200
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text13 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   25
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   23
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   21
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   19
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   17
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16440
      TabIndex        =   13
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13920
      TabIndex        =   11
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11400
      TabIndex        =   9
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16200
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   22800
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10320
      Width           =   1575
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   22560
      TabIndex        =   53
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label22 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18600
      TabIndex        =   52
      Top             =   7440
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label21 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   22560
      TabIndex        =   51
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label20 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   21720
      TabIndex        =   50
      Top             =   2400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label19 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   18600
      TabIndex        =   49
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label18 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   9360
      TabIndex        =   48
      Top             =   7920
      Width           =   6375
   End
   Begin VB.Label Label17 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   47
      Top             =   7920
      Width           =   6495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Collar para perro $ 25.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   720
      TabIndex        =   32
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      Caption         =   "Pollera: $ 35.00"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   13320
      TabIndex        =   44
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      Caption         =   "Pantalón largo:  $ 88.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   10800
      TabIndex        =   43
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      Caption         =   "Pantalón corto:  $ 58.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   8280
      TabIndex        =   42
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      Caption         =   "Par de medias:  $ 20.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   5760
      TabIndex        =   41
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      Caption         =   "Buzo: $ 125.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   3240
      TabIndex        =   40
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00004040&
      Caption         =   "    Zapatillas:     $ 485.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   720
      TabIndex        =   39
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Manta para perro grande:  $ 75.00"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   15840
      TabIndex        =   38
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Manta para perro chico:  $ 35.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   13320
      TabIndex        =   37
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Bolsa de piedras para gatos por 5 Paquetes : $ 400.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10800
      TabIndex        =   36
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Adorno para pez: $ 10.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8280
      TabIndex        =   35
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Hueso para perro: $ 15.00"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   5760
      TabIndex        =   34
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      Caption         =   "Caja para gato:  $ 120.00 "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3240
      TabIndex        =   33
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Tienda Web"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sector Mascotas"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   45
      Top             =   1200
      Width           =   17895
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sector Ropa"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   46
      Top             =   4560
      Width           =   15375
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim balance As Double 'toda la plata del dia
Dim promedioVenta As Double 'promedio de venta
Dim contVentas As Double 'ventas individuales
Dim totalPro As Integer 'total de cantidad de productos del dia

Dim cantPro As Long 'cantidad de productos

Dim ventaPro As Double 'total de esa venta unitaria
Dim ventaProCarrito As String ' total de esa venta carrito

Dim totalVentaIndi As Double 'acumulador de ventas individual
Dim totalCantPro As Integer 'acumuador de productos individual

Dim nombProMascotas As String 'strings con los nombres de los pro
Dim nombProRopa As String

Dim hayMascotas As Boolean
Dim hayRopa As Boolean

'oraciones a imprimir
Dim ropa As String
Dim mascota As String

Dim carritoNom As String
Dim carritoCant As String
Dim carritoPrecio As String

Private Sub Command1_Click()

    End

End Sub

Private Sub Command15_Click() 'finalizar compra
    
    'parte guardado global
    totalPro = totalPro + totalCantPro
    balance = balance + totalVentaIndi
    
    'carrito falso
    
    If Not totalVentaIndi = 0 Then
        contVentas = contVentas + 1
        
        mascota = "Se realizo la compra de " & nombProMascotas & " y son productos del sector mascotas"
        ropa = "Se realizo la compra de " & nombProRopa & " y son productos del sector de ropa"
        
        If hayMascotas = True And hayRopa = True Then
        
            Label18.Caption = mascota & vbCrLf & ropa & vbCrLf & " y el total es $ " & totalVentaIndi
            
        ElseIf hayMascotas = True And hayRopa = False Then
            
            Label18.Caption = mascota & vbCrLf & " y el total es $ " & totalVentaIndi
        
        ElseIf hayMascotas = False And hayRopa = True Then
            
            Label18.Caption = ropa & vbCrLf & " y el total es $ " & totalVentaIndi
        
        End If
        
        Command18_Click
        
    End If
    
    nombProMascotas = ""
    nombProRopa = ""
    totalVentaIndi = 0
    totalCantPro = 0
    
End Sub

Private Sub Command16_Click()
    
    If Not contVentas = 0 Then
        promedioVenta = balance / contVentas
    
        Label17.Caption = _
        "Resumen del Dia:" & vbCrLf & _
        "Cantidad de Productos Vendidos: " & totalPro & vbCrLf & _
        "Total de Dinero recaudado: $ " & balance & vbCrLf & _
        "Promedio de Venta: $ " & promedioVenta
    End If

End Sub

Private Sub Command17_Click()

    If Not totalVentaIndi = 0 Then
    
        If totalVentaIndi < 19 Then
            
            Label18.Caption = "La compra es inferior a $19, si supera los $40 tiene el envio gratis!" + _
            ", si no desea el descuento, haga click en 'Finalizar Compra', sino agregue otros items" + _
            " o haga click en 'Limpiar Carrito' para empezar de nuevo"
        
        ElseIf totalVentaIndi >= 19 And totalVentaIndi <= 40 Then
            
            Label18.Caption = "La compra es inferior a $40, si supera los $40 tiene el envio gratis!" + _
            ", si no desea el descuento, haga click en 'Finalizar Compra', sino agregue otros items" + _
            " o haga click en 'Limpiar Carrito' para empezar de nuevo"
        
        ElseIf totalVentaIndi > 200 Then
            
            Label18.Caption = "Tiene un descuento del 10% y el ENVIO GRATIS"
        
        ElseIf totalVentaIndi > 40 Then
            
            Label18.Caption = "Tiene el ENVIO GRATIS"
        
        End If
        
        Label19.Visible = True
        Label20.Visible = True
        Label21.Visible = True
        Label22.Visible = True
        Label23.Visible = True
        
        Command18.Visible = True
        Command15.Visible = True
    
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
    End If
    
End Sub

Private Sub Command18_Click()
    
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    Text4.Enabled = True
    Text5.Enabled = True
    Text6.Enabled = True
    Text7.Enabled = True
    Text8.Enabled = True
    Text9.Enabled = True
    Text10.Enabled = True
    Text11.Enabled = True
    Text12.Enabled = True
    Text13.Enabled = True
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Text11.Text = ""
    Text12.Text = ""
    Text13.Text = ""
    
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True
    Command10.Enabled = True
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    Command14.Enabled = True
    
    Command18.Visible = False
    Command15.Visible = False
    
    Label19.Visible = False
    Label20.Visible = False
    Label21.Visible = False
    Label22.Visible = False
    Label23.Visible = False
    
    'borra el carrito
    
    hayMascotas = False
    hayRopa = False
    
    nombProMascotas = ""
    totalVentaIndi = 0
    totalCantPro = 0
    
    carritoNom = ""
    carritoCant = ""
    carritoPrecio = ""
    
    Label19.Caption = ""
    Label20.Caption = ""
    Label21.Caption = ""

End Sub

Private Sub Command2_Click()

    If IsNumeric(Text1.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text1.Text) 'cantidad de productos
        ventaPro = 25 * cantPro 'precio del producto
        ventaProCarrito = 25 * cantPro
        
        hayMascotas = True
        nombProMascotas = nombProMascotas + "Collar para Perros, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Collar para Perros" + vbCrLf
        carritoCant = carritoCant + "x" + Text1.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text1.Enabled = False
        Command2.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
    
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
    
End Sub

Private Sub Command3_Click()

    If IsNumeric(Text2.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text2.Text) 'cantidad de productos
        ventaPro = 120 * cantPro 'precio del producto
        ventaProCarrito = 120 * cantPro
        
        hayMascotas = True
        nombProMascotas = nombProMascotas + "Collar para Gatos, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Collar para Gatos" & vbCrLf
        carritoCant = carritoCant + "x" + Text2.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text2.Enabled = False
        Command3.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
    
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
    
    
End Sub

Private Sub Command4_Click()

    If IsNumeric(Text3.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text3.Text) 'cantidad de productos
        ventaPro = 15 * cantPro 'precio del producto
        ventaProCarrito = 15 * cantPro
        
        hayMascotas = True
        nombProMascotas = nombProMascotas + "Hueso para perro, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Hueso para perro" & vbCrLf
        carritoCant = carritoCant + "x" + Text3.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text3.Enabled = False
        Command4.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command5_Click()

    If IsNumeric(Text4.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text4.Text) 'cantidad de productos
        ventaPro = 10 * cantPro 'precio del producto
        ventaProCarrito = 10 * cantPro
    
        hayMascotas = True
        nombProMascotas = nombProMascotas + "Adorno para pez, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Adorno para pez" & vbCrLf
        carritoCant = carritoCant + "x" + Text4.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text4.Enabled = False
        Command5.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
    
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
    
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command6_Click()

    If IsNumeric(Text5.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text5.Text) 'cantidad de productos
        ventaPro = 400 * cantPro 'precio del producto
        ventaProCarrito = 400 * cantPro
    
        hayMascotas = True
        nombProMascotas = nombProMascotas + "Piedras para Gatos x5u, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Piedras p/ Gatos x5u" & vbCrLf
        carritoCant = carritoCant + "x" + Text5.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text5.Enabled = False
        Command6.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command7_Click()

    If IsNumeric(Text6.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text6.Text) 'cantidad de productos
        ventaPro = 35 * cantPro 'precio del producto
        ventaProCarrito = 35 * cantPro
    
        hayMascotas = True
        nombProMascotas = nombProMascotas + "Manta p/ perro chica, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Manta p/ perro chica" & vbCrLf
        carritoCant = carritoCant + "x" + Text6.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text6.Enabled = False
        Command7.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command8_Click()

    If IsNumeric(Text7.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text7.Text) 'cantidad de productos
        ventaPro = 75 * cantPro 'precio del producto
        ventaProCarrito = 75 * cantPro
    
        hayMascotas = True
        nombProMascotas = nombProMascotas + "Manta p/ perro grande, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Manta p/ perro grande" & vbCrLf
        carritoCant = carritoCant + "x" + Text7.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text7.Enabled = False
        Command8.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command9_Click()

    If IsNumeric(Text8.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text8.Text) 'cantidad de productos
        ventaPro = 485 * cantPro 'precio del producto
        ventaProCarrito = 485 * cantPro
    
        hayRopa = True
        nombProRopa = nombProRopa + "Zapatillas, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Zapatillas" & vbCrLf
        carritoCant = carritoCant + "x" + Text8.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text8.Enabled = False
        Command9.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command10_Click()

    If IsNumeric(Text9.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text9.Text) 'cantidad de productos
        ventaPro = 125 * cantPro 'precio del producto
        ventaProCarrito = 125 * cantPro
    
        hayRopa = True
        nombProRopa = nombProRopa + "Buzo, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Buzo" & vbCrLf
        carritoCant = carritoCant + "x" + Text9.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text9.Enabled = False
        Command10.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command11_Click()

    If IsNumeric(Text10.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text10.Text) 'cantidad de productos
        ventaPro = 20 * cantPro 'precio del producto
        ventaProCarrito = 20 * cantPro
    
        hayRopa = True
        nombProRopa = nombProRopa + "Par de Medias, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Par de Medias" & vbCrLf
        carritoCant = carritoCant + "x" + Text10.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text10.Enabled = False
        Command11.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command12_Click()

    If IsNumeric(Text11.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text11.Text) 'cantidad de productos
        ventaPro = 58 * cantPro 'precio del producto
        ventaProCarrito = 58 * cantPro
    
        hayRopa = True
        nombProRopa = nombProRopa + "Pantalon corto, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Pantalon corto" & vbCrLf
        carritoCant = carritoCant + "x" + Text11.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text11.Enabled = False
        Command12.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command13_Click()

    If IsNumeric(Text12.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text12.Text) 'cantidad de productos
        ventaPro = 88 * cantPro 'precio del producto
        ventaProCarrito = 88 * cantPro
    
        hayRopa = True
        nombProRopa = nombProRopa + "Pantalon largo, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Pantalon largo" & vbCrLf
        carritoCant = carritoCant + "x" + Text12.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text12.Enabled = False
        Command13.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Command14_Click()

    If IsNumeric(Text13.Text) = True Then
    
    'parte venta individual
        cantPro = CInt(Text13.Text) 'cantidad de productos
        ventaPro = 35 * cantPro 'precio del producto
        ventaProCarrito = 35 * cantPro
    
        hayRopa = True
        nombProRopa = nombProRopa + "Pollera, "
    
    'acumulador para ventas indivual
        totalVentaIndi = totalVentaIndi + ventaPro 'compre 4 collares de parros y 4 gato
        totalCantPro = totalCantPro + cantPro
    
    'carrito
        carritoNom = carritoNom + "Pollera" & vbCrLf
        carritoCant = carritoCant + "x" + Text13.Text + vbCrLf
        carritoPrecio = carritoPrecio + "$ " + ventaProCarrito + vbCrLf
    
        Text13.Enabled = False
        Command14.Enabled = False
    
    'actualiza en tiempo real el carrito
        Label19.Caption = carritoNom
        Label20.Caption = carritoCant
        Label21.Caption = carritoPrecio
        Label23.Caption = totalVentaIndi
        
        If totalVentaIndi > 200 Then
        
            totalVentaIndi = totalVentaIndi * 0.9
            Label22.Caption = "Total:                                        $" & vbCrLf & "Descuento: $ " & (totalVentaIndi * 0.1)
            
        End If
        
    Else
        
        MsgBox "Ingrese un Numero. Vuelva a intentarlo", , "Error"
        
    End If
    
End Sub

Private Sub Form_Activate()

    Label22.Caption = "Total:                                        $"

End Sub

Private Sub Text1_GotFocus()

    Text1.BackColor = &H80FF80

End Sub
Private Sub Text1_LostFocus()
    
    Text1.BackColor = &HFFFFC0

End Sub

Private Sub Text2_GotFocus()

    Text2.BackColor = &H80FF80

End Sub
Private Sub Text2_LostFocus()
    
    Text2.BackColor = &HFFFFC0

End Sub

Private Sub Text3_GotFocus()

    Text3.BackColor = &H80FF80

End Sub
Private Sub Text3_LostFocus()
    
    Text3.BackColor = &HFFFFC0

End Sub

Private Sub Text4_GotFocus()

    Text4.BackColor = &H80FF80

End Sub
Private Sub Text4_LostFocus()
    
    Text4.BackColor = &HFFFFC0

End Sub

Private Sub Text5_GotFocus()

    Text5.BackColor = &H80FF80

End Sub
Private Sub Text5_LostFocus()
    
    Text5.BackColor = &HFFFFC0

End Sub

Private Sub Text6_GotFocus()

    Text6.BackColor = &H80FF80

End Sub
Private Sub Text6_LostFocus()
    
    Text6.BackColor = &HFFFFC0

End Sub

Private Sub Text7_GotFocus()

    Text7.BackColor = &H80FF80

End Sub
Private Sub Text7_LostFocus()
    
    Text7.BackColor = &HFFFFC0

End Sub
Private Sub Text8_GotFocus()

    Text8.BackColor = &H80FF80

End Sub
Private Sub Text8_LostFocus()
    
    Text8.BackColor = &HFFFFC0

End Sub
Private Sub Text9_GotFocus()

    Text9.BackColor = &H80FF80

End Sub
Private Sub Text9_LostFocus()
    
    Text9.BackColor = &HFFFFC0

End Sub
Private Sub Text10_GotFocus()

    Text10.BackColor = &H80FF80

End Sub
Private Sub Text10_LostFocus()
    
    Text10.BackColor = &HFFFFC0

End Sub

Private Sub Text11_GotFocus()

    Text11.BackColor = &H80FF80

End Sub
Private Sub Text11_LostFocus()
    
    Text11.BackColor = &HFFFFC0

End Sub

Private Sub Text12_GotFocus()

    Text12.BackColor = &H80FF80

End Sub
Private Sub Text12_LostFocus()
    
    Text12.BackColor = &HFFFFC0

End Sub

Private Sub Text13_GotFocus()

    Text13.BackColor = &H80FF80

End Sub
Private Sub Text13_LostFocus()
    
    Text13.BackColor = &HFFFFC0

End Sub
