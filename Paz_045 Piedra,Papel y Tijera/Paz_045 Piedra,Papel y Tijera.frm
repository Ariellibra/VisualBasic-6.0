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
Dim ganadas, contador, num1, azar As Integer
Dim elec As String
Dim adivino As Boolean

Private Sub Form_Activate()
    Form1.Hide
    
    NumeroRandom
    Juego
    End
    
End Sub

Private Sub NumeroRandom()

    Randomize
    
    azar = CInt(Rnd() * 3) + 1
    
End Sub

Private Sub Juego()
    
    Do
        elec = InputBox("Ingrese algunas de las opciones para empezar a jugar", "Juego: Piedra, Papel y Tijera")
    
        If (azar = 1 Or azar = 2) And LCase(elec) = "tijera" Then
            
            ganadas = ganadas + 1
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


