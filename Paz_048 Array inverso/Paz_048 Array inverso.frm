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
Dim n As Integer
Dim askiArray(4) As String



Private Sub Codigo()
    
    For n = 0 To 4
        
        askiArray(4 - n) = InputBox("", "Traductor Ascii")
        Print (Asc(askiArray(4 - n)))
        
    Next n
    
End Sub


Private Sub Form_Activate()
    
    Codigo
    
End Sub

