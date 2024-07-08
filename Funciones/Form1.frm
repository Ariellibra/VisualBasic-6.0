VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   855
      Left            =   6600
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   4200
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   1095
      Left            =   3720
      TabIndex        =   1
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   975
      Left            =   960
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num1, num2 As Integer

Private Sub Op_Mat()
    
    Dim signo As String
    
    signo = Text3
    
    If signo = "+" Then
        Label1 = Sumas(CInt(num1), CInt(num2))
    Else
        Restas
    End If
    
    
End Sub

Private Function Sumas(var1 As Integer, var2 As Integer) As Integer
    
    var1 = CInt(Text1)
    
    var2 = CInt(Text2)

    
    Sumas = var1 + var2
    
End Function

Private Sub Restas()
    
    num1 = CInt(Text1)
    
    num2 = CInt(Text2)
    
    Label1 = num1 - num2
    
End Sub

Private Sub Command1_Click()
    
    Op_Mat
    
End Sub

