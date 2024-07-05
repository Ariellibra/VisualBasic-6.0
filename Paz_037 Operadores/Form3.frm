VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim frase, frase2 As String
Dim chara, chara2 As String
Dim n As Integer

Private Sub Form_Activate()

    Form3.Hide
    
    frase = InputBox("Escriba una frase", "Formulario 3: Cambiar caracteres")
    chara = InputBox("Escriba un caracter a remplazar", "Formulario 3: Cambiar caracteres")
    chara2 = InputBox("Escriba el nuevo caracter", "Formulario 3: Cambiar caracteres")
    
    
    For n = 1 To Len(frase)
        
        If Mid(frase, n, 1) = chara Then
            
            frase2 = frase2 + chara2
            
        Else
        
        frase2 = frase2 + Mid(frase, n, 1)
        
        End If
        
    Next n
    
    MsgBox "La frase era: " & frase & " Y el caracter a cambiar era el : " & chara & vbCrLf & " La nueva frase es: " & frase2, vbInformation, "Formulario 3: Cambiar caracteres"
    
    Unload Me
    Form4.Show
    
End Sub

