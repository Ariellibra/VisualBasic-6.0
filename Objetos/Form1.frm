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

Dim objMascotas1 As Mascotas
Dim boton As CommandButton


Private Sub Command1_Click()
    
    Print (objMascotas1)
    
End Sub

Private Sub Form_Activate()
        
    objMascotas1.name = "Morgana"
    objMascotas1.edad = 5
    objMascotas1.peso = 7.5
    objMascotas1.raza = "Callejero"
    
End Sub

