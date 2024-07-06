VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Multiplicar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Factorizar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6960
      TabIndex        =   8
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese 2 Numeros para multiplicar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3600
      TabIndex        =   7
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese 1 Numero para factorizar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num1 As Integer

Dim res As Integer

Dim num2 As Integer

Dim num3 As Integer


Private Sub Command1_Click()
    
    If IsNumeric(Text1.Text) = False Then
        Label3.Caption = "Solo se aceptan numeros"
    
    ElseIf VarType(num1) = vbInteger Then
    
        num1 = CInt(Text1.Text)
        
        Select Case num1
            Case Is > 6
                Label3.Caption = "EL numero no puede ser mayor a 6"
        
            Case Is < 1
                Label3.Caption = "EL numero no puede ser menor a 1"
    
            Case 1
                Label3.Caption = "El factorial de 1 es: " & num1
            
            Case 2
                res = 1 * 2
                Label3.Caption = "El factorial de 2 es: " & res
            
            Case 3
                res = 1 * 2 * 3
                Label3.Caption = "El factorial de 3 es: " & res
            
            Case 4
                res = 1 * 2 * 3 * 4
                Label3.Caption = "El factorial de 4 es: " & res
        
            Case 5
                res = 1 * 2 * 3 * 4 * 5
                Label3.Caption = "El factorial de 5 es: " & res
        
            Case 6
                res = 1 * 2 * 3 * 4 * 5 * 6
                Label3.Caption = "El factorial de 6 es: " & res
        
            End Select
        Else
            Label3.Caption = "Solo se aceptan numeros enteros del 1 al 6"
    End If
    
    Text1.Text = ""
    
    
End Sub

Private Sub Command2_Click()
    
    If IsNumeric(Text2.Text) = False Or IsNumeric(Text3.Text) = False Then
        Label3.Caption = "Solo se aceptan numeros"
    Else
        num2 = CInt(Text2.Text)
        num3 = CInt(Text3.Text)
        
        res = num2 * num3
        
        Label3.Caption = "El " & num2 & " multiplicado por " & num3 & " da como resultado: " & res
    End If
    
    Text2.Text = ""
    Text3.Text = ""
    
    
End Sub

Private Sub Command3_Click()
    
    End

End Sub

Private Sub Text1_GotFocus()

    Text1.BackColor = &HC0FFC0

End Sub

Private Sub Text1_LostFocus()

    Text1.BackColor = &HFFFFFF

End Sub

Private Sub Text2_GotFocus()

    Text2.BackColor = &HC0FFC0

End Sub

Private Sub Text2_LostFocus()

    Text2.BackColor = &HFFFFFF

End Sub

Private Sub Text3_GotFocus()

    Text3.BackColor = &HC0FFC0

End Sub

Private Sub Text3_LostFocus()

    Text3.BackColor = &HFFFFFF

End Sub
