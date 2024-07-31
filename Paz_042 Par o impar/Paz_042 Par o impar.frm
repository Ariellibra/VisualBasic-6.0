VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   14235
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim num As String
Dim res As Integer


Private Sub ParImpar()
    
    

    Do
     
        num = InputBox("Ingrese un numero", "ParImpar")
        
        Calculo (num)
        
        Pregunta
    
    Loop Until LCase(num) = "fin"
    
End Sub

Private Sub Calculo(num1 As String)
    
    res = CInt(CDbl(num1) / 2) * 2
    
End Sub

Private Sub Pregunta()
    
    If LCase(num) = "fin" Then
        
    ElseIf LCase(num) = "" Then
                
        MsgBox "Ingrese algo", vbCritical, "Error"
            
    ElseIf res = num Then
        
        MsgBox "Es Par", vbCritical, "Par"
            
    Else
        MsgBox "Es Impar", vbCritical, "Impar"
        
    End If
    

End Sub

Private Sub Form_Activate()
    
    
    ParImpar
    End
    

End Sub

