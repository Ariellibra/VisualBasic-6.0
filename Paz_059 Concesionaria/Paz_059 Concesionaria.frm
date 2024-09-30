VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   ScaleHeight     =   11370
   ScaleWidth      =   14535
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moto1 As New Motocicleta


Private Sub Form_Activate()
    
    moto1.moto True, "Roja", "JTU703", 150#, 50#, 2, "Honda", "CRT500", "17/09/2024", 250, 70
    

End Sub

