VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23430
   LinkTopic       =   "Form1"
   ScaleHeight     =   12195
   ScaleWidth      =   23430
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Cargar Nafta"
      Height          =   735
      Left            =   11280
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detener"
      Height          =   735
      Left            =   9120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arrancar"
      Height          =   735
      Left            =   6840
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   2175
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim moto1 As New Motocicleta


Private Sub Command1_Click()

    moto1.arrancar
    Label1.Caption = moto1.GetCombustibleActual

End Sub

Private Sub Command2_Click()
    
    moto1.detener
    Label1.Caption = moto1.GetCombustibleActual
    
End Sub

Private Sub Command3_Click()
    
    moto1.cargarNafta
    Label1.Caption = moto1.GetCombustibleActual

End Sub

Private Sub Form_Activate()
    
    moto1.moto True, "Roja", "JTU703", 150#, 50#, 2, "Honda", "CRT500", "17/09/2024", 250, 70
    
    Label1.Caption = moto1.GetCombustibleActual

End Sub

