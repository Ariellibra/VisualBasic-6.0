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
Dim ganadas, contador, azar As Integer
Dim elec, nombre As String
Dim adivino As Boolean

Private Sub Form_Activate()
    Form1.Hide
    
    Juego
    MsgBox nombre & " tu resultado final fue:" & vbCrLf & "Intentos: " & contador & " Ganadas: " & ganadas, vbInformation, "Termino el juego"
    End
    
End Sub

Private Sub NumeroRandom()

    Randomize
    
    azar = CInt(Rnd() * 3) + 1
    
End Sub

Private Sub Juego()

' 1 = tijera, 2= piedra, 3= papel
    
    nombre = InputBox("Ingrese su nombre", "Nombre")
    
    Do
        elec = InputBox(nombre & " Ingrese algunas de las opciones para empezar a jugar" _
                & vbCrLf & "Intentos: " & contador & " Ganadas: " & ganadas & _
                vbCrLf & "Salir: N", "Juego: Piedra, Papel y Tijera")
        
        NumeroRandom
                
        If UCase(elec) = "N" Then
            
            adivino = True
    
        ElseIf azar = 1 Then
            
            If LCase(elec) = "tijera" Then
                
                contador = contador + 1
                MsgBox "Empate, vuelvalo a intentar", vbExclamation, "Empate"
            
            ElseIf LCase(elec) = "piedra" Then
                
                ganadas = ganadas + 1
                contador = contador + 1
                MsgBox "Ganaste", vbInformation, "Ganador"
                
            ElseIf LCase(elec) = "papel" Then
            
                contador = contador + 1
                MsgBox "Perdiste, vuelvalo a intentar", vbCritical, "Perdedor"
            
            End If
            
        ElseIf azar = 2 Then
            
            If LCase(elec) = "tijera" Then
            
                contador = contador + 1
                MsgBox "Perdiste, vuelvalo a intentar", vbCritical, "Perdedor"
            
            ElseIf LCase(elec) = "piedra" Then
                
                contador = contador + 1
                MsgBox "Empate, vuelvalo a intentar", vbExclamation, "Empate"
            
            ElseIf LCase(elec) = "papel" Then
                
                ganadas = ganadas + 1
                contador = contador + 1
                MsgBox "Ganaste", vbInformation, "Ganador"
            
            End If
            
        ElseIf azar = 3 Then
            
            If LCase(elec) = "tijera" Then
                
                ganadas = ganadas + 1
                contador = contador + 1
                MsgBox "Ganaste", vbInformation, "Ganador"
            
            ElseIf LCase(elec) = "piedra" Then
                
                contador = contador + 1
                MsgBox "Perdiste, vuelvalo a intentar", vbCritical, "Perdedor"
            
            ElseIf LCase(elec) = "papel" Then
            
                contador = contador + 1
                MsgBox "Empate, vuelvalo a intentar", vbExclamation, "Empate"
            
            End If
            
        End If
    
    Loop Until adivino = True
    
End Sub


