VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   ScaleHeight     =   12180
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim numArray(9), n, i, mayor, menor As Integer

Private Sub LlenoArray()

    Randomize
    
    For n = 0 To 9
        
        numArray(n) = CInt((Rnd * 100) + 1)
         
    Next n
    
    For n = 0 To 9
        
        Print (numArray(n) & " " & n)
        
    Next n
    
End Sub

Private Sub MayorMenor()
    
    menor = 1002

    For n = 0 To 9
        
        If numArray(n) = "" Then
        
        
        ElseIf numArray(n) < menor Then
            
            menor = numArray(n)
            
        End If
        
    Next n
    
    Print (menor)
    
    For n = 0 To 9
        
        If numArray(n) = menor Then
        
            numArray(n) = ""
            Print (menor & " " & n)
            
        End If
        
    Next n
    
'    For n = 0 To 9
'
'        Print (numArray(n) & " " & n)
'
'    Next n
    
    
End Sub

Private Sub FuncionPorDiez()
    i = 1
    Do
        MayorMenor
        i = i + 1
    Loop Until i = 10

End Sub

Private Sub OrdenarMayorMenor()
    
    LlenoArray
    FuncionPorDiez

End Sub

Private Sub Form_Activate()
    
    OrdenarMayorMenor
    
End Sub

