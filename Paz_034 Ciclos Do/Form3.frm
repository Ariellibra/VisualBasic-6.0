VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form3"
   ScaleHeight     =   9690
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oracion As String
Dim num1 As String
Dim letra1 As String
Dim n As Integer
Dim bande As Boolean


Private Sub Form_Load()

    Form3.Hide
    
    n = 1
    bande = True
    
    
    Do While bande = True
        
        letra1 = InputBox("Ingrese una letra" & vbCrLf & "La letra puede ser X o Z", "Formulario 3")
        
        If LCase(letra1) = "x" Or LCase(letra1) = "z" Then
        
            num1 = InputBox("Ingrese un numero entre 1 y 15", "Formulario 3")
            
            If CInt(num1) > 0 And CInt(num1) <= 15 Then
            
                For n = 1 To CInt(num1)
                    
                    oracion = oracion + letra1
                Next n
                
                MsgBox oracion, vbInformation, "Formulario 3"
                
                bande = False
                
            End If
        
        ElseIf letra1 = "" Then
        
            bande = False
            
        End If
    
    
    Loop
    
    End
    
End Sub
