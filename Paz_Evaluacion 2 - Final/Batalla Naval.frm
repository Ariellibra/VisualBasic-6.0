VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22980
   LinkTopic       =   "Form1"
   ScaleHeight     =   11460
   ScaleWidth      =   22980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Cambiar Turno"
      Height          =   855
      Left            =   6240
      TabIndex        =   89
      Top             =   8160
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command4"
      Height          =   615
      Left            =   17400
      TabIndex        =   88
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   15960
      TabIndex        =   87
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   735
      Left            =   20760
      TabIndex        =   50
      Top             =   6480
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame3"
      Height          =   2775
      Left            =   20160
      TabIndex        =   44
      Top             =   3360
      Width           =   2655
      Begin VB.TextBox Text2 
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   49
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   48
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   480
         TabIndex        =   45
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "X"
         Height          =   495
         Left            =   840
         TabIndex        =   47
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Y"
         Height          =   495
         Left            =   840
         TabIndex        =   46
         Top             =   1320
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   2775
      Left            =   20160
      TabIndex        =   2
      Top             =   360
      Width           =   2655
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Y"
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "X"
         Height          =   495
         Left            =   840
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Barcos Enemigos"
      Height          =   6975
      Left            =   7440
      TabIndex        =   1
      Top             =   720
      Width           =   8175
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   35
         Left            =   5640
         TabIndex        =   86
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   34
         Left            =   4560
         TabIndex        =   85
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   33
         Left            =   3480
         TabIndex        =   84
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   32
         Left            =   2400
         TabIndex        =   83
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   31
         Left            =   1320
         TabIndex        =   82
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   30
         Left            =   240
         TabIndex        =   81
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   29
         Left            =   5640
         TabIndex        =   80
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   28
         Left            =   4560
         TabIndex        =   79
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   27
         Left            =   3480
         TabIndex        =   78
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   26
         Left            =   2400
         TabIndex        =   77
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   25
         Left            =   1320
         TabIndex        =   76
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   24
         Left            =   240
         TabIndex        =   75
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   23
         Left            =   5640
         TabIndex        =   74
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   22
         Left            =   4560
         TabIndex        =   73
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   21
         Left            =   3480
         TabIndex        =   72
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   20
         Left            =   2400
         TabIndex        =   71
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   19
         Left            =   1320
         TabIndex        =   70
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   18
         Left            =   240
         TabIndex        =   69
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   17
         Left            =   5640
         TabIndex        =   68
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   16
         Left            =   4560
         TabIndex        =   67
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   15
         Left            =   3480
         TabIndex        =   66
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   14
         Left            =   2400
         TabIndex        =   65
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   13
         Left            =   1320
         TabIndex        =   64
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   12
         Left            =   240
         TabIndex        =   63
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   11
         Left            =   5640
         TabIndex        =   62
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   10
         Left            =   4560
         TabIndex        =   61
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   9
         Left            =   3480
         TabIndex        =   60
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   8
         Left            =   2400
         TabIndex        =   59
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   7
         Left            =   1320
         TabIndex        =   58
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   6
         Left            =   240
         TabIndex        =   57
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   5
         Left            =   5640
         TabIndex        =   56
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   4
         Left            =   4560
         TabIndex        =   55
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   3
         Left            =   3480
         TabIndex        =   54
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   2
         Left            =   2400
         TabIndex        =   53
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   1
         Left            =   1320
         TabIndex        =   52
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Cpu 
         BackColor       =   &H00FFFF80&
         Caption         =   "Label5"
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   51
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Mis Barcos"
      Height          =   6975
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   35
         Left            =   5640
         TabIndex        =   43
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   34
         Left            =   4560
         TabIndex        =   42
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   33
         Left            =   3480
         TabIndex        =   41
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   32
         Left            =   2400
         TabIndex        =   40
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   31
         Left            =   1320
         TabIndex        =   39
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   30
         Left            =   240
         TabIndex        =   38
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   29
         Left            =   5640
         TabIndex        =   37
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   28
         Left            =   4560
         TabIndex        =   36
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   27
         Left            =   3480
         TabIndex        =   35
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   26
         Left            =   2400
         TabIndex        =   34
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   25
         Left            =   1320
         TabIndex        =   33
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   24
         Left            =   240
         TabIndex        =   32
         Top             =   4680
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   23
         Left            =   5640
         TabIndex        =   31
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   22
         Left            =   4560
         TabIndex        =   30
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   21
         Left            =   3480
         TabIndex        =   29
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   20
         Left            =   2400
         TabIndex        =   28
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   19
         Left            =   1320
         TabIndex        =   27
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   18
         Left            =   240
         TabIndex        =   26
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   17
         Left            =   5640
         TabIndex        =   25
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   16
         Left            =   4560
         TabIndex        =   24
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   15
         Left            =   3480
         TabIndex        =   23
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   14
         Left            =   2400
         TabIndex        =   22
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   13
         Left            =   1320
         TabIndex        =   21
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   12
         Left            =   240
         TabIndex        =   20
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   11
         Left            =   5640
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   10
         Left            =   4560
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   9
         Left            =   3480
         TabIndex        =   17
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   8
         Left            =   2400
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   7
         Left            =   1320
         TabIndex        =   15
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   6
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   5
         Left            =   5640
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   4
         Left            =   4560
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   3
         Left            =   3480
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   2
         Left            =   2400
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label1"
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Player 
         BackColor       =   &H00FFFFC0&
         Caption         =   " Label1"
         Height          =   855
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
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
Dim n As Integer

Private Sub Command1_Click()
    
    cordX = Text1(0).Text
    cordY = Text1(1).Text
    
    cord = (6 * (cordY - 1)) + (cordX - 1)
    
    Ataque (cord)

End Sub
Private Sub CpuCrearBarco()
    Dim posicionesOcupadas(2) As Integer
    Dim ocupado As Boolean
    Dim j As Integer
    
    For n = 0 To 2
        Randomize
    
        ' Generar coordenadas aleatorias
        rndCpuX = CInt((Rnd * 5) + 1)
        rndCpuY = CInt((Rnd * 5) + 1)
        
        ' Calcular la posición de CPU
        cordCpu = CInt((6 * (rndCpuY - 1)) + (rndCpuX - 1))
        
        For j = 0 To 2
        
            If cordCpu = posicionesOcupadas(n) Then
                
                rndCpuX = CInt((Rnd * 5) + 1)
                rndCpuY = CInt((Rnd * 5) + 1)
        
                ' Calcular la posición de CPU
                cordCpu = CInt((6 * (rndCpuY - 1)) + (rndCpuX - 1))
            
            Else
            
                posicionesOcupadas(n) = cordCpu
                j = 2
            
            End If
        
        Next j
        
        ' Inicializar barco de CPU
        Set BarcoCpu(n) = New Barco
        BarcoCpu(n).Constructor "Cpu", cordCpu, False
        
        ' Cambiar el color de fondo del cuadro correspondiente
        Cpu(BarcoCpu(n).GetPosci).BackColor = &H4080&
    
    Next n
    
End Sub
Private Sub CpuAtaque()
    Dim fallo As Boolean
    fallo = False
    
    Randomize
    
    ' Generar coordenadas aleatorias
    rndCpuX = CInt((Rnd * 5) + 1)
    rndCpuY = CInt((Rnd * 5) + 1)
        
    ' Calcular la posición de CPU
    cord = CInt((6 * (rndCpuY - 1)) + (rndCpuX - 1))
    
    For n = 0 To 2
        
        If BarcoPlayer(n).GetPosci = cord Then
        
            MsgBox "Hundido", vbCritical, "En el Blanco"
            BarcoPlayer(n).SetHundido (True)
            Player(cord).BackColor = &HFF&
            fallo = False
            n = 2
            
        Else
            
            fallo = True
            
        End If
        
    Next n
    
    If fallo = True Then
    
        MsgBox "Fallo el tiro", vbInformation, "En el Agua"
        Player(cord).BackColor = &HC00000
            
    End If

End Sub

Private Function Ataque(cord As Integer)
    Dim fallo As Boolean
    fallo = False
    
    For n = 0 To 2
        
        If BarcoCpu(n).GetPosci = cord Then
        
            MsgBox "Hundido", vbCritical, "En el Blanco"
            BarcoCpu(n).SetHundido (True)
            Cpu(cord).BackColor = &HFF&
            fallo = False
            n = 2
            
        Else
            
            fallo = True
            
        End If
        
    Next n
    
    If fallo = True Then
    
        MsgBox "Fallo el tiro", vbInformation, "En el Agua"
        Cpu(cord).BackColor = &HC00000
            
    End If
    
End Function

Private Sub Command2_Click()
    
    cordX = Text2(0).Text
    cordY = Text2(1).Text
    
    Print ((6 * (cordY - 1)) + (cordX - 1))
    
    BarcoPlayer(n).Constructor "Player", (6 * (cordY - 1)) + (cordX - 1), False
    Player(BarcoPlayer(n).GetPosci).BackColor = &H4080&
    
    n = n + 1

End Sub

Private Sub Command3_Click()

    CpuCrearBarco

End Sub

Private Sub Command4_Click()
    
    Unload Form1
    
End Sub

Private Sub Command5_Click()
    
    Load Form1
    
End Sub

Private Sub Command6_Click()

    CpuAtaque

End Sub

Private Sub Form_Activate()
    
    n = 0

End Sub

