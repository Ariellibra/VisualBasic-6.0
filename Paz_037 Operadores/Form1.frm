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

Dim user1, user1Psw As String
Dim user2, user2Psw As String

Dim entrada, contra As String

Private Sub Form_Activate()
    
    Form1.Hide
    
    user1 = "admin"
    user1Psw = "0000"
    
    user2 = "libra"
    user2Psw = "nashe"
    
    entrada = InputBox("Ingrese el Usuario", "Formulario 1: Login")
    
    If LCase(entrada) = user1 Then
        
        contra = InputBox("Ingrese la Contraseña", "Formulario 1: Login")
        
        If LCase(contra) = user1Psw Then
            
            MsgBox "Bienvenido " & entrada, vbExclamation, "Formulario 1: Login"
            
            Unload Me
            Form2.Show
            
        ElseIf contra = "" Then
        
            MsgBox "No puede ingresar espacios vacios" & vbCrLf & "Cerrando el Programa", vbCritical, "Formulario 1: Error"
            End
            
        Else
        
            MsgBox "Contraseña Incorrecta" & vbCrLf & "Cerrando el Programa", vbCritical, "Formulario 1: Error"
            End
            
        End If
    
    ElseIf LCase(entrada) = user2 Then
        
        contra = InputBox("Ingrese la Contraseña", "Formulario 1: Login")
        
        If LCase(contra) = user2Psw Then
            
            MsgBox "Bienvenido " & entrada, vbExclamation, "Formulario 1: Login"
            
            Unload Me
            Form2.Show
            
        ElseIf contra = "" Then
        
            MsgBox "No puede ingresar espacios vacios" & vbCrLf & "Cerrando el Programa", vbCritical, "Formulario 1: Error"
            End
            
        Else
        
            MsgBox "Contraseña Incorrecta" & vbCrLf & "Cerrando el Programa", vbCritical, "Formulario 1: Error"
            End
            
        End If
    
    ElseIf entrada = "" Then
        
        MsgBox "No puede ingresar espacios vacios" & vbCrLf & "Cerrando el Programa", vbCritical, "Formulario 1: Error"
        End
        
    Else
        
        MsgBox "Usuario Incorrecto" & vbCrLf & "Cerrando el Programa", vbCritical, "Formulario 1: Error"
        End
        
    End If
    
    
End Sub

