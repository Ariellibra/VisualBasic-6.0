VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim corre As Boolean
Dim cont, cont15, cont50, cont25 As Integer
Dim num1 As Integer
Dim num15, num50, num25 As String



Private Sub Form_Activate()
    
    cont = 1
    
    Do
        num1 = CInt(InputBox("Ingrese 1 numero positivo", "Formulario 1"))
        
        If num1 >= 0 Then
            
            If num1 < 15 Then
                
                cont15 = cont15 + 1
                num15 = num15 & num1 & ", "
            
            ElseIf num1 > 50 Then
                
                cont50 = cont50 + 1
                num50 = num50 & num1 & ", "
            
            ElseIf num1 >= 25 Then
                
                If num1 <= 45 Then
                
                    cont25 = cont25 + 1
                    num25 = num25 & num1 & ", "
                
                End If
            
            End If
            
        End If
        
        cont = cont + 1
        
    Loop Until cont = 10
    
    MsgBox _
    "La cantidad de numeros Menores a 15 son: " & cont15 & _
    "Los numeros son: " & num15 & _
    "La cantidad de numeros Mayores a 50 son: " & cont50 & _
    "Los numeros son: " & num50 & _
    "La cantidad de numeros entre 25 y 45 son: " & cont45 & _
    "Los numeros son: " & num45, vbInformation, "Formulario 1"

    
End Sub

Private Sub Form_Load()

End Sub
