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

Dim txt1 As String
Dim txt2 As Integer
Dim frase As String
Dim n As Integer

Private Sub Form_Activate()
    
    Unload Me
    
    txt1 = InputBox("Ingrese una palabra o frase", "Formulario 2")
    
    For n = 0 To Len(txt1) - 1
        
        frase = frase + Mid(txt1, Len(txt1) - n, 1)
        
    Next n
    
    MsgBox frase, vbInformation, "Formulario 2"
    
    Form3.Show
    Exit Sub
    
End Sub

