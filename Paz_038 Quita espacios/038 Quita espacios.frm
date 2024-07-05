VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frase, frase2 As String
Dim n As Integer

Private Sub Form_Activate()

    Form1.Hide
    
    frase = InputBox("Escriba su frase", "Formulario 1: Quitar espacios")
    
    For n = 1 To Len(frase)
        
        If Mid(frase, n, 1) <> " " Then
        
            frase2 = frase2 + Mid(frase, n, 1)
        
        End If
        
    Next n
    
    MsgBox "La frase sin espacios es: " & frase2, vbInformation, "Formulario 1: Quitar espacios"
    
    End
    
End Sub
