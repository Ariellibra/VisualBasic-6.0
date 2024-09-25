VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15525
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   15525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   2760
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objMascotas1 As New Mascotas
Dim objMascotas2 As New Mascotas
Dim boton As CommandButton
Dim grande As Boolean


Private Sub Command1_Click()
    
    Print (objMascotas1.GetName)
    Print (objMascotas1.edad)
    Print (objMascotas1.peso)
    Print (objMascotas1.raza)
    
    Print (objMascotas2.GetName)
    Print (objMascotas2.edad)
    Print (objMascotas2.peso)
    Print (objMascotas2.raza)
    
    grande = objMascotas1.esGrande
    
    If grande = True Then
    
        Print ("Es Mayor de edad")
    Else
        
        Print ("Es Menor de edad")
    End If
    
End Sub

Private Sub Form_Activate()
        
    objMascotas1.SetName "Morgana"
    objMascotas1.edad = 5
    objMascotas1.peso = 9
    objMascotas1.raza = "Callejero"
    
    objMascotas2.Constructor "Moona", 3, 7.5, "Callejera"
    
    
End Sub

