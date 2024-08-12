VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   21300
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   21300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   13320
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Label3"
      Height          =   2655
      Left            =   15360
      TabIndex        =   3
      Top             =   7320
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label2"
      Height          =   2415
      Left            =   15120
      TabIndex        =   2
      Top             =   4440
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   2415
      Left            =   15000
      TabIndex        =   1
      Top             =   1560
      Width           =   5295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    
    Form1.Line (2500, 2000)-(3000, 3000) '(diagonal derecha triangulo)
    Form1.Line (2000, 3000)-(2500, 2000)
    Form1.Line (2000, 3000)-(3000, 3000)
    
    Form1.Line (3000, 3000)-(4500, 3000)
    Form1.Line (2500, 2000)-(4500, 2000)
    Form1.Line (4500, 2000)-(4500, 4500)
    
    Form1.Line (2000, 3000)-(2000, 4500)
    Form1.Line (3000, 3000)-(3000, 4500)
    
    Form1.Line (2000, 4500)-(4500, 4500)
    
    Form1.Line (2350, 3800)-(2650, 3800)
    Form1.Line (2350, 3800)-(2350, 4500)
    Form1.Line (2650, 3800)-(2650, 4500)
    
    Form1.Circle (2500, 2600), 250
    
    
    'Form1.Line (1500, 2000)-(3000, 3000)
    'Form1.Line (1000, 1000)-(3000, 3000) '(recta triangulo)
    'Form1.Line (1000, 1000)-(1500, 3000)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Label1 = X
    Label2 = Y
    Label3 = Button
    
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label1 = "Me movi we"
    Label1 = X

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Label2 = Y
    
End Sub

