VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
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
      Height          =   735
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FF80&
      Caption         =   "Comprar Moto"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cargar Nafta"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Detener"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Arrancar"
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
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
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moto1 As New Motocicleta

Private Sub actualizarLabel()

    Label1.Caption = "Datos de la moto:" & vbCrLf & _
                     "Color: " & moto1.GetColor & vbCrLf & _
                     "Matrícula: " & moto1.GetMatricula & vbCrLf & _
                     "Cilindrada: " & moto1.GetCilindrada & vbCrLf & _
                     "Combustible actual: " & moto1.GetCombustibleActual & " litros" & vbCrLf & _
                     "Capacidad del tanque: " & moto1.GetCapacidadTanque & " litros" & vbCrLf & _
                     "Número de ruedas: " & moto1.GetNumeroRuedas & vbCrLf & _
                     "Marca: " & moto1.GetMarca & vbCrLf & _
                     "Modelo: " & moto1.GetModelo & vbCrLf & _
                     "Fecha de fabricación: " & moto1.GetFechaFabricacion & vbCrLf & _
                     "Velocidad punta: " & moto1.GetVelocidadPunta & " km/h" & vbCrLf & _
                     "Peso: " & moto1.GetPeso & " kg"
                     
End Sub

Private Sub Command1_Click()

    moto1.arrancar
    actualizarLabel
    
End Sub

Private Sub Command2_Click()

    moto1.detener
    actualizarLabel
    
End Sub

Private Sub Command3_Click()

    moto1.cargarNafta
    actualizarLabel
    
End Sub
Private Sub Command4_Click()

    moto1.moto InputBox("Ingrese el color de la moto:"), _
                InputBox("Ingrese la matrícula de la moto:"), _
                CDbl(InputBox("Ingrese la cilindrada de la moto:")), _
                CDbl(InputBox("Ingrese la capacidad del tanque de la moto:")), _
                CInt(InputBox("Ingrese el número de ruedas de la moto:")), _
                InputBox("Ingrese la marca de la moto:"), _
                InputBox("Ingrese el modelo de la moto:"), _
                InputBox("Ingrese la fecha de fabricación de la moto:"), _
                CInt(InputBox("Ingrese la velocidad punta de la moto:")), _
                CDbl(InputBox("Ingrese el peso de la moto:"))
                
    actualizarLabel
    
End Sub


Private Sub Command5_Click()
    
    End

End Sub

Private Sub Form_Activate()

    Label1.Caption = "Esperando datos de la moto..."
    
End Sub

