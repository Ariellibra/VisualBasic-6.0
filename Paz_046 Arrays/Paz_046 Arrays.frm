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
Dim arrayNum(9) As Long
Dim res As Double
Dim azar, n, j As Integer


Private Sub Form_Activate()
    
    LlenaArray
    
End Sub

Private Sub LlenaArray()
    Randomize
    
    For n = 0 To 9
    
        arrayNum(n) = CInt((Rnd() * 10) + 1)
        res = 1
        
        For j = 1 To arrayNum(n)
            
            res = res * j
            
        Next j
        
        Print (arrayNum(n) & " = " & res)
    
    Next n

End Sub

Private Sub Factorial(dato As Integer)

    
End Sub
