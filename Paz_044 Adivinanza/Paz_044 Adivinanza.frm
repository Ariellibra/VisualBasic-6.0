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
Dim contador, num1, azar, rango As Integer
Dim adivino As Boolean


Private Sub Form_Activate()
    
    Juego
    
End Sub

Private Sub Juego()
    
    Randomize
    
    rango = CInt(InputBox("Ingrese un numero para el rango de numeros a divinar, este tiene que ser mayor a 1", "Juego: Rango"))
    
    Do
        
        num1 = CInt(InputBox("Ingrese el numero para adivinar", "Juego: Adivinanza"))
        
        azar = (Rnd() + rango) + 1
        
        If num1 = azar Then
            
            contador = contador + 1
            MsgBox "Excelente, el numero era " & azar & " y la cantidad de intentos fue: " & contador, vbInformation, "Juego: Ganaste"
            adivino = True
            
        ElseIf num1 > azar Then
            
            contador = contador + 1
            MsgBox "Tu numero es muy grande, Intentos: " & contador, vbCritical, "Juego: Perdiste"
            
        ElseIf num1 < azar Then
        
            contador = contador + 1
            MsgBox "Tu numero es muy chico, Intentos: " & contador, vbCritical, "Juego: Perdiste"
            
        End If
    
    Loop Until adivino = True
    

End Sub

Private Sub Form_Load()

End Sub


