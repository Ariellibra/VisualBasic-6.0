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

Dim cantTotal As Long
Dim edad As Integer
Dim n As Integer

Private Sub Form_Activate()
    
    
    
    Exit Sub
    
    edad = 12
    cantTotal = 15
    
    Do
        cantTotal = cantTotal * 2
        edad = edad + 1
        
    Loop Until cantTotal >= 4560
    
    MsgBox _
    "La Edad a la que llegue a $ 4560 es: " & edad & vbCrLf & _
    "Y la cantidad total seria de: $ " & cantTotal, vbInformation, "Formulario 2"
    
    
End Sub

