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

Private Sub ValidarFecha()
    
    fecha = InputBox("Ingrese la Fecha", "Validar Fecha")
    
    MsgBox (fecha)
    
    If Len(fecha) > 10 Then
    
        MsgBox "Ingrese una Fecha valida", vbCritical, "Error 1"
    
    End If
    
    For i = 1 To Len(fecha)
        
        If Mid(fecha, i, 1) = "/" Then
            
            i = i + 1
            
            If diaC < 2 Then
                diaC = diaC + 1
            ElseIf mesC < 2 Then
                mesC = mesC + 1
            ElseIf añoC < 4 Then
                añoC = añoC + 1
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
                
            MsgBox "Ingrese una Fecha valida", vbCritical, "Error Dia"
                
    End If
    
    If CInt(mes) > 12 Then
                
            MsgBox "Ingrese una Fecha valida", vbCritical, "Error Mes"
                
    End If
            
    If CInt(año) < 1910 And CInt(año) > 2100 And Len(año) < 4 Then
                
            MsgBox "Ingrese una Fecha valida", vbCritical, "Error Año"
                
    End If
    
    MsgBox (dia)
    MsgBox (mes)
    MsgBox (año)

End Sub



Private Sub Form_Activate()
    
    ValidarFecha
    
End Sub

