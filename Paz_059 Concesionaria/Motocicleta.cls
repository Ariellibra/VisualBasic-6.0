VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Motocicleta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private esNueva As Boolean
Private color As String
Private matricula As String
Private cilindrada As Double
Private combustibleActual As Double
Private capacidadTanque As Double
Private numeroRuedas As Integer
Private marca As String
Private modelo As String
Private fechaFabricacion As String
Private velocidadPunta As Integer
Private peso As Integer
Private motor As Boolean

Public Sub moto(colorC As String, matriculaC As String, cilindradaC As Double, capacidadTanqueC As Double, numeroRuedasC As Integer, marcaC As String, modeloC As String, fechaFabricacionC As String, velocidadPuntaC As Integer, pesoC As Integer)
 
esNueva = True
color = colorC
matricula = matriculaC
cilindrada = cilindradaC
combustibleActual = capacidadTanqueC
capacidadTanque = capacidadTanqueC
numeroRuedas = numeroRuedasC
marca = marcaC
modelo = modeloC
fechaFabricacion = fechaFabricacionC
velocidadPunta = velocidadPuntaC
peso = pesoC
motor = False

End Sub

Public Sub arrancar()

    If motor = False Then
        
        If combustibleActual <= 0 Then
            
            MsgBox "No tenes mas Nafta, anda a la Ypf y carga nafta", vbCritical, "Sin Nafta"
        
        Else
        
            motor = True
            MsgBox "Prendiendo el Motor", vbInformation, "Arrancar"
        
            If cilindrada > 0 And cilindrada <= 110 Then
            
                If (combustibleActual - 3) < 0 Then
                    
                    MsgBox "No tenes mas Nafta, anda a la Ypf y carga nafta", vbCritical, "Sin Nafta"
                Else
                    combustibleActual = combustibleActual - 3
                    
                End If
            ElseIf cilindrada > 125 And cilindrada <= 250 Then
            
                If (combustibleActual - 8) < 0 Then
                    
                    MsgBox "No tenes mas Nafta, anda a la Ypf y carga nafta", vbCritical, "Sin Nafta"
                Else
                    combustibleActual = combustibleActual - 8
                    
                End If
        
            ElseIf cilindrada > 250 Then
            
                If (combustibleActual - 10) < 0 Then
                    
                    MsgBox "No tenes mas Nafta, anda a la Ypf y carga nafta", vbCritical, "Sin Nafta"
                Else
                    combustibleActual = combustibleActual - 10
                    
                End If
        
            End If
            
            If combustibleActual = 0 Then
            
            MsgBox "No tenes mas Nafta, anda a la Ypf y carga nafta", vbCritical, "Sin Nafta"
        
        Else
        
        End If
    
    Else
    
        MsgBox "El motor ya esta encendido", vbCritical, "Motor Encendido"
        
    End If

End Sub

Public Sub detener()
    
    If motor = True Then
        
        motor = False
        MsgBox "Apagando el motor", vbInformation, "Detener"
    Else
        
        MsgBox "El motor ya esta apagado", vbCritical, "Motor Apagado"
    
    End If

End Sub

Public Sub cargarNafta()

    combustibleActual = capacidadTanque
    MsgBox "El tanque de combustible está lleno.", vbInformation, "Combustible recargado"
    
End Sub

Public Function GetCombustibleActual() As Double
    
    GetCombustibleActual = combustibleActual

End Function

Public Function GetColor() As String

    GetColor = color
    
End Function

Public Function GetMatricula() As String

    GetMatricula = matricula
    
End Function

Public Function GetCilindrada() As Double

    GetCilindrada = cilindrada
    
End Function

Public Function GetCapacidadTanque() As Double

    GetCapacidadTanque = capacidadTanque
    
End Function

Public Function GetNumeroRuedas() As Integer

    GetNumeroRuedas = numeroRuedas
    
End Function

Public Function GetMarca() As String

    GetMarca = marca
    
End Function

Public Function GetModelo() As String

    GetModelo = modelo
    
End Function

Public Function GetFechaFabricacion() As String

    GetFechaFabricacion = fechaFabricacion
    
End Function

Public Function GetVelocidadPunta() As Integer

    GetVelocidadPunta = velocidadPunta
    
End Function

Public Function GetPeso() As Double

    GetPeso = peso
    
End Function

