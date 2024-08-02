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
Dim fecha As String
Dim i As Integer
Dim dia As String
Dim mes As String
Dim año As String
Dim diaC, mesC, añoC As Integer
Dim mov As Integer
Dim añoNum As Integer

Dim añoBisiestoVar As Boolean
Dim fechaCorrectaVar As Boolean


Private Sub ValidarFecha()
    
    fecha = InputBox("Ingrese la Fecha", "Validar Fecha")
    
    MsgBox (fecha)
    
    If Len(fecha) > 10 Then
    
        MsgBox "Ingrese una Fecha valida, cerrando el programa", vbCritical, "Error"
        End
    
    End If
    
    For i = 1 To Len(fecha)
        
        If Mid(fecha, i, 1) = "/" Then
            
            If diaC <= 2 Then
                diaC = diaC + 2
                i = i + 1
            ElseIf mesC <= 2 Then
                mesC = mesC + 2
                i = i + 1
            ElseIf añoC <= 4 Then
                añoC = añoC + 1
                i = i + 1
            End If

        End If
        
        If diaC < 2 Then
            
            dia = dia + Mid(fecha, i, 1)
            diaC = diaC + 1
        
        ElseIf mesC < 2 Then
            
            mes = mes + Mid(fecha, i, 1)
            mesC = mesC + 1
            
        ElseIf añoC < 4 Then
            
            año = año + Mid(fecha, i, 1)
            añoC = añoC + 1
            
        End If
            
    
    Next i
    
    If CInt(dia) > 31 Then
                
            MsgBox "Ingrese una Fecha valida, cerrando el programa", vbCritical, "Error Dia"
            End
                
    End If
    
    If CInt(mes) > 12 Then
                
            MsgBox "Ingrese una Fecha valida, cerrando el programa", vbCritical, "Error Mes"
            End
                
    End If
            
    If CInt(año) <= 1910 Then
                
        MsgBox "Ingrese una Fecha valida, cerrando el programa", vbCritical, "Error Año"
        End
        
    ElseIf CInt(año) >= 2100 Then
        
        MsgBox "Ingrese una Fecha valida, cerrando el programa", vbCritical, "Error Año"
        End
    
    End If
    
    'MsgBox (dia)
    'MsgBox (mes)
    'MsgBox (año)
    
    añoNum = CInt(año)
    
    añoBisiestoVar = añoBisiesto(CInt(añoNum))
    
    If añoBisiestoVar = True Then
        
        MsgBox año & " Es año Bisiesto", vbInformation, "Año Bisiesto"
    Else
        
        MsgBox año & " No es año Bisiesto", vbInformation, "Año Bisiesto"
    End If
    
    fechaCorrectaVar = diaMes(CInt(dia), CInt(mes), añoBisiestoVar)
    
    If fechaCorrectaVar = True Then
        
        MsgBox "La fecha " & fecha & " es una fecha valida", vbInformation, "Fecha Valida"
        
    Else
    
        MsgBox "La fecha " & fecha & " no es una fecha valida", vbInformation, "Fecha no Valida"
    
    End If
    

End Sub

Function añoBisiesto(año As Integer) As Boolean
    Dim siEs As Boolean
    Dim n As Integer
    
    For n = 1912 To 2100
    
        If año = n Then
            
            siEs = True
            Exit For
'            MsgBox año & " Es año Bisiesto", vbInformation, "Año Bisiesto"
            
        Else
            n = n + 3
        
        End If
    
    Next n
    
    añoBisiesto = siEs
        
End Function

Function diaMes(diaF As Integer, mesF As Integer, añoF As Boolean) As Boolean
    Dim fechaCorrecta As Boolean
    
    If mesF = 1 And diaF <= 31 Or mesF = 3 And diaF <= 31 Or mesF = 5 And diaF <= 31 Or mesF = 7 And diaF <= 31 Or mesF = 8 And diaF <= 31 Or mesF = 10 And diaF <= 31 Or mesF = 12 And diaF <= 31 Then
        
        fechaCorrecta = True
    ElseIf mesF = 4 And diaF <= 30 Or mesF = 6 And diaF <= 30 Or mesF = 9 And diaF <= 30 Or mesF = 11 And diaF <= 30 Then
        
        fechaCorrecta = True
    
    ElseIf mesF = 2 Then
        
        If diaF <= 28 Then
            
            fechaCorrecta = True
            
        ElseIf añoF = True And diaF <= 29 Then
            
            fechaCorrecta = True
        
        End If
        
    Else
        
        fechaCorrecta = False
    End If
    
    diaMes = fechaCorrecta

End Function


Private Sub Form_Activate()
    
    Form1.Hide
    ValidarFecha
    End
    
End Sub

