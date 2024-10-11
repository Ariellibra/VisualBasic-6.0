VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Crear Mascotas con Constructor"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Crear Mascotas con Set"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Mostrar Datos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   5040
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objMascotas1 As New Mascotas
Dim objMascotas2 As New Mascotas
Dim grande As Boolean

Private Sub Command1_Click()

    Label1.Caption = "Mascota 1:" & vbCrLf & _
                     "Nombre: " & objMascotas1.GetName & vbCrLf & _
                     "Edad: " & objMascotas1.GetEdad & vbCrLf & _
                     "Peso: " & objMascotas1.GetPeso & vbCrLf & _
                     "Raza: " & objMascotas1.GetRaza

    Label2.Caption = "Mascota 2:" & vbCrLf & _
                     "Nombre: " & objMascotas2.GetName & vbCrLf & _
                     "Edad: " & objMascotas2.GetEdad & vbCrLf & _
                     "Peso: " & objMascotas2.GetPeso & vbCrLf & _
                     "Raza: " & objMascotas2.GetRaza

    grande = objMascotas1.esGrande
    
    If grande = True Then
        Label1.Caption = Label1.Caption & vbCrLf & "Es Mayor de edad"
    Else
        Label1.Caption = Label1.Caption & vbCrLf & "Es Menor de edad"
    End If

End Sub

Private Sub Command2_Click()

    objMascotas1.SetName InputBox("Ingrese el nombre de la mascota:")
    objMascotas1.SetEdad CInt(InputBox("Ingrese la edad de la mascota:"))
    objMascotas1.SetPeso CDbl(InputBox("Ingrese el peso de la mascota:"))
    objMascotas1.SetRaza InputBox("Ingrese la raza de la mascota:")

End Sub

Private Sub Command3_Click()

    objMascotas2.Constructor "Moona", 3, 7.5, "Callejera"
    
End Sub

Private Sub Command4_Click()
    
    End
    
End Sub

