VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   1815
   End
   Begin VB.ListBox List1 
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
      Height          =   4860
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   8895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cargar Datos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Datos Alumnos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nombre As String
Dim apellido As String
Dim edad As String
Dim fechaNacimiento As String
Dim dni As String
Dim telefono As String
Dim linea As String


Private Sub guardarDatos()

    Open App.Path & "\Alumnos.txt" For Append As #1
    
    nombre = InputBox("Ingrese el Nombre:", "Datos Personales")
    apellido = InputBox("Ingrese el Apellido:", "Datos Personales")
    
    Do
        edad = InputBox("Ingrese la Edad (solo números):", "Datos Personales")
        
    Loop Until IsNumeric(edad)
    
    fechaNacimiento = InputBox("Ingrese la Fecha de Nacimiento (dd/mm/yyyy):", "Datos Personales")
    
    Do
        dni = InputBox("Ingrese el DNI (solo números):", "Datos Personales")
        
    Loop Until IsNumeric(dni)
    
    Do
        telefono = InputBox("Ingrese el Teléfono (solo números):", "Datos Personales")
    Loop Until IsNumeric(telefono)
    
    Write #1, nombre, apellido, edad, fechaNacimiento, dni, telefono
    
    Close #1
    
    MsgBox "Datos guardados en Alumnos.txt"
    
End Sub

Private Sub cargarDatos()

    Open App.Path & "\Alumnos.txt" For Input As #1
    
    List1.Clear
    
    Do While Not EOF(1)
    
        Line Input #1, linea
        List1.AddItem (linea)
                      
    Loop
    
    Close #1
    
End Sub

Private Sub Command1_Click()

    guardarDatos
    cargarDatos
    
End Sub


Private Sub Command2_Click()
    
    End

End Sub

