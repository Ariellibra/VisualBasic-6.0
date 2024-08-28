VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23265
   LinkTopic       =   "Form1"
   ScaleHeight     =   12300
   ScaleWidth      =   23265
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List5 
      Height          =   2985
      Left            =   12000
      TabIndex        =   9
      Top             =   3720
      Width           =   3615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Juntar Notas y Nombres"
      Height          =   615
      Left            =   9960
      TabIndex        =   8
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Juntar Notas y Nombres"
      Height          =   615
      Left            =   360
      TabIndex        =   7
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ListBox List4 
      Height          =   2985
      Left            =   2400
      TabIndex        =   6
      Top             =   3720
      Width           =   7215
   End
   Begin VB.ListBox List3 
      Height          =   2985
      Left            =   12000
      TabIndex        =   5
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cargar Nombres"
      Height          =   615
      Left            =   9960
      TabIndex        =   4
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calcular Promedio"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   2985
      Left            =   6000
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Notas"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
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
        
        nota(j, n) = CInt(InputBox("Ingrese la nota: " & n + 1, "Cargar Notas"))
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
    
End Sub

Private Sub mayorMenorPromedio()
    
    promedio = 0
    
    For n = 0 To 4
        If n = 0 Then
            mayor = nota(k, n)
            menor = nota(k, n)
        ElseIf nota(k, n) <= menor Then
            menor = nota(k, n)
        ElseIf nota(k, n) >= mayor Then
            mayor = nota(k, n)
        End If
        promedio = promedio + nota(k, n)
        
    Next n
    
    List2.AddItem ("Nota Mayor: " & mayor & " Menor: " & menor & " Promedio: " & promedio / 5)
    
    k = k + 1

End Sub

Private Sub cargarNombres()

    For n = 0 To 4
        
        nombres(n) = InputBox("Ingrese un nombre sin repetir", "Cargar Nombres")
        List3.AddItem (nombres(n))
    Next n
    
End Sub

Private Sub juntarNombresyNotas()
    
    For n = 0 To 4
        
        List4.AddItem ("Nombre del Alumno: " & List3.List(n) & " obtuvo: " & List1.List(n))
        
    Next n
    
End Sub

Private Sub juntarNombresyNotas2()
    
    For n = 0 To 4
        
        List5.AddItem ("[" & List3.List(n) & "] - " & List1.List(n))
        
    Next n
    
End Sub

Private Sub Command1_Click()
    
    cargarNotas
    
End Sub

Private Sub Command2_Click()

    mayorMenorPromedio

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

Private Sub Form_Activate()
    
    j = 0
    
End Sub

