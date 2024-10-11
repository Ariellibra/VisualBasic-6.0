VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19335
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   19335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
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
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox Text1 
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
      Index           =   4
      Left            =   15720
      TabIndex        =   19
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Modificar Nota"
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
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   360
      Width           =   2175
   End
   Begin VB.TextBox Text1 
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
      Index           =   3
      Left            =   16920
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   18120
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   16920
      TabIndex        =   13
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   15720
      TabIndex        =   10
      Top             =   5160
      Width           =   975
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   12000
      TabIndex        =   9
      Top             =   3720
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Juntar Notas y Nombres"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Juntar Notas y Nombres"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   2400
      TabIndex        =   6
      Top             =   3720
      Width           =   7215
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   12000
      TabIndex        =   5
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Cargar Nombres"
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Modificar Nota"
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
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   2175
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   4920
      TabIndex        =   2
      Top             =   360
      Width           =   4695
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Cargar Notas"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15720
      TabIndex        =   20
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Posicion Nota"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16920
      TabIndex        =   18
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   18120
      TabIndex        =   17
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16920
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Posicion Nota"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   15720
      TabIndex        =   11
      Top             =   4560
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n, l, j, k, mayor, menor As Integer
Dim promedio As Double
Dim nota(5, 5) As Integer
Dim flag As Boolean
Dim nombres(5) As String
Private Sub cargarNotas()
    For n = 0 To 4
        Dim notaValida As Boolean
        notaValida = False
        
        Do While Not notaValida
            nota(j, n) = CInt(InputBox("Ingrese la nota: " & n + 1 & " (0 a 10)", "Cargar Notas"))
            
            If nota(j, n) >= 0 And nota(j, n) <= 10 Then
                notaValida = True
            Else
                MsgBox "Nota fuera de rango. Por favor, ingrese un número entre 0 y 10."
            End If
        Loop
        
        If n = 0 Then
            List1.AddItem ("")
            List1.List(j) = List1.List(j) & "[ " & nota(j, n) & ", "
        ElseIf n = 4 Then
            List1.List(j) = List1.List(j) & nota(j, n) & " ]"
        Else
            List1.List(j) = List1.List(j) & nota(j, n) & ", "
        End If
    Next n
    
    j = j + 1
    
    mayorMenorPromedio j - 1
End Sub
Private Sub CambiarListaOriginal(ByVal index As Integer)
    List1.List(index) = ""
    
    For n = 0 To 4
        If n = 0 Then
            List1.List(index) = List1.List(index) & "[ " & nota(index, n) & ", "
        ElseIf n = 4 Then
            List1.List(index) = List1.List(index) & nota(index, n) & " ]"
        Else
            List1.List(index) = List1.List(index) & nota(index, n) & ", "
        End If
    Next n

    mayorMenorPromedio index
    
End Sub


Private Sub mayorMenorPromedio(ByVal index As Integer)
    Dim promedio As Double
    Dim mayor, menor As Integer
    promedio = 0
    
    For n = 0 To 4
        If n = 0 Then
            mayor = nota(index, n)
            menor = nota(index, n)
        ElseIf nota(index, n) <= menor Then
            menor = nota(index, n)
        ElseIf nota(index, n) >= mayor Then
            mayor = nota(index, n)
        End If
        promedio = promedio + nota(index, n)
    Next n

    ' Asegúrate de que List2 tenga el mismo número de elementos que List1
    If List2.ListCount <= index Then
        List2.AddItem ("")
    End If
    
    List2.List(index) = "Nota Mayor: " & mayor & " Menor: " & menor & " Promedio: " & promedio / 5
End Sub



Private Sub cargarNombres()
    Dim i As Integer

    For n = 0 To 4
        Dim nombreValido As Boolean
        nombreValido = False
        
        Do While Not nombreValido
            nombres(n) = InputBox("Ingrese un nombre sin repetir", "Cargar Nombres")
            nombreValido = True
            
            For i = 0 To n - 1
                If nombres(n) = nombres(i) Then
                    MsgBox "El nombre ya existe. Por favor, ingrese un nombre diferente."
                    nombreValido = False
                    Exit For
                End If
            Next i
        Loop
        
        List3.AddItem (nombres(n))
    Next n
    
End Sub

Private Sub juntarNombresyNotas()

    List4.Clear
    For n = 0 To 4
        
        List4.AddItem ("Nombre del Alumno: " & List3.List(n) & " obtuvo: " & List1.List(n))
        
    Next n
    
End Sub

Private Sub juntarNombresyNotas2()

    List5.Clear
    For n = 0 To 4
        
        List5.AddItem ("[" & List3.List(n) & "] - " & List1.List(n))
        
    Next n
    
End Sub
Private Sub cambiarNota()

    Dim ind As Integer
    ind = List5.ListIndex
    
    If ind <> -1 Then
        nota(ind, (Text1(0) - 1)) = Text1(1)
        CambiarListaOriginal ind
    Else
        MsgBox "Seleccione un elemento de la lista primero", , "Error"
    End If
    
    juntarNombresyNotas
    juntarNombresyNotas2
    
End Sub
Private Sub cambiarNotaPorNombre()

    Dim nombreBuscado As String
    Dim nombreEncontrado As Boolean
    Dim t As Integer
    
    nombreBuscado = Text1(4)
    nombreEncontrado = False
    
    For t = 0 To List3.ListCount - 1
    
        If List3.List(t) = nombreBuscado Then
            nombreEncontrado = True
            Exit For
        End If
        
    Next t
    
    If nombreEncontrado Then
    
        nota(t, (Text1(3) - 1)) = Text1(2)
        CambiarListaOriginal t
    Else
        MsgBox "Nombre no encontrado. Por favor, ingrese un nombre válido."
    End If
    
    juntarNombresyNotas
    juntarNombresyNotas2
    
End Sub

Private Sub Command1_Click()
    
    cargarNotas
    
End Sub

Private Sub Command2_Click()
    
    cambiarNota
    
    
End Sub

Private Sub Command3_Click()
    
    cargarNombres
    
End Sub

Private Sub Command4_Click()
    
    juntarNombresyNotas

End Sub

Private Sub Command5_Click()
    
    juntarNombresyNotas2
    
End Sub

Private Sub Command6_Click()

    cambiarNotaPorNombre

End Sub

Private Sub Command7_Click()

    End

End Sub

Private Sub Form_Activate()
    
    j = 0
    
End Sub

