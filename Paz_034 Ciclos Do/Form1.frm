VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num2 As String
Dim num1 As Integer

Private Sub Command1_Click()
    
    Unload Form1
    Form2.Show
    
End Sub

Private Sub Form_Activate()
    
    
    Form1.Hide
    
    Do
        num2 = InputBox("Ingrese un numero para sumar", "Formulario 1")
        num1 = num1 + CInt(num2)
    Loop Until num2 = 0
    MsgBox "La suma de los numeroes es: " & num1, vbInformation, "Formulario 1"
    
    Unload Form1
    Form2.Show
    
End Sub

