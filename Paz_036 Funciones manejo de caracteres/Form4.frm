VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form4"
   ScaleHeight     =   6255
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim txt1 As String
Dim txt2 As Integer
Dim frase As String
Dim fraseComp As String
Dim n As Integer

Private Sub Form_Activate()
    
    Unload Me
    
    txt1 = InputBox("Ingrese una palabra", "Formulario 4")
    
    For n = 0 To Len(txt1) - 1
        
        frase = frase + Mid(txt1, Len(txt1) - n, 1)
        
    Next n
    
    If txt1 = frase Then
        
        MsgBox txt1 & " y " & frase & " Son palindromos", vbInformation, "Formulario 4"
    
    Else
        
        MsgBox txt1 & " y " & frase & " No son palindromos", vbInformation, "Formulario 4"
        
    End If
    
    End
    
End Sub

