VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim email, email2 As String
Dim n As Integer

Private Sub Form_Activate()

    Form4.Hide
    
    email = InputBox("Escriba su email", "Formulario 4: Email")
    
    For n = 1 To Len(email)
        
        If Mid(email, n, 1) = "@" Then
            
            email2 = email2 + "@ceu.es"
            Exit For
            
        Else
        
        email2 = email2 + Mid(email, n, 1)
        
        End If
        
    Next n
    
    MsgBox "Su nuevo email es: " & email2, vbInformation, "Formulario 4: Email"
    
    End
    
End Sub
