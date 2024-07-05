VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frase As String
Dim chara As String
Dim n As Integer
Dim cont As Integer

Private Sub Form_Activate()

    Form2.Hide
    
    frase = InputBox("Escriba una frase", "Formulario 2: Frase")
    chara = InputBox("Escriba un caracter", "Formulario 2: Frase")
    
    For n = 1 To Len(frase)
        
        If Mid(frase, n, 1) = chara Then
        cont = cont + 1
        End If
        
    Next n
    
    MsgBox "La frase era: " & frase & " Y el caracter : " & chara & " se repite: " & cont, vbInformation, "Formulario 2: Frase"
    
    Unload Me
    Form3.Show
    
End Sub

