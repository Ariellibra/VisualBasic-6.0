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

Dim txt1 As String
Dim txt2 As Integer
Dim n As Integer


Private Sub Form_Activate()
    
    txt1 = InputBox("Ingrese una palabra", "Formulario 1")
    
    For n = 1 To Len(txt1)
        
        Print (Mid(txt1, n, 1))
        
    Next n
    
    Form2.Show
    Exit Sub
    
End Sub

