VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim num1 As String
Dim num2 As String
Dim suma As Integer
Dim preg As Boolean
Dim preg_str As String
Dim contador As Integer

Private Sub Form_Activate()
    
    Form2.Hide
    
    contador = 0
    
    preg = True
    
    Do While preg = True
        
        num1 = InputBox("Ingrese los numeros a sumar", "Formulario 2")
        
        contador = contador + 1
            
        suma = suma + CInt(num1)
            
        If contador = 2 Then
                
            MsgBox "La suma es: " & suma, vbInformation, "Suma"
                
            num1 = InputBox("Desea seguir sumando numeros?" & vbCrLf & "Aprete S o N", "Formulario 2")
                
            If LCase(num1) = "s" Then
                preg = True
                suma = 0
                contador = 0
            Else
                preg = False
            End If
                
        End If
    
    Loop
    
    Unload Me
    Form3.Show
    
End Sub

