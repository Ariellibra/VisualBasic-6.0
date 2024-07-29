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

Dim num As Double

Private Sub ParImpar()
     


    If CInt(num / 2) * 2 = num Then
        
        MsgBox "Es Par", vbCritical, "Par"
    Else
        MsgBox "Es Impar", vbCritical, "Impar"
        
    End If
    
End Sub

Private Sub Calculo()
    
    ParImpar
    
End Sub

Private Sub Form_Activate()
    
    num = InputBox("Ingrese un numero", "ParImpar")
    
    
    Calculo
    

End Sub

