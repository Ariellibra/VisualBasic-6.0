VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Aceptar"
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
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
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
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
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1695
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
      Height          =   1935
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4440
      TabIndex        =   5
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim res, suma, resta, multi, divi, vals As String
Dim num, total, cont As Integer

Private Sub Cuentas()

    num = CInt(Text1)
    res = Text2
    
    If cont < 4 Then
    
        If LCase(res) = "suma" Then
    
            Sumas
        
        ElseIf LCase(res) = "resta" Then
            
            Restas
    
        ElseIf LCase(res) = "multi" Then
            
            Multis
        
        ElseIf LCase(res) = "division" Then
            
            Divis
        
        End If
    
    End If
    
End Sub

Private Sub Sumas()
    
    If cont = 0 Then
        total = num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    Else
        total = total + num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    End If

End Sub
Private Sub Restas()
    
    If cont = 0 Then
        total = num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    Else
        total = total - num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    End If

End Sub

Private Sub Multis()
    
    If cont = 0 Then
        total = num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    Else
        total = total * num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    End If

End Sub
Private Sub Divis()
    
    If cont = 0 Then
        total = num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    Else
        total = total / num
        vals = vals & num & ", "
        cont = cont + 1
        Resultado
    End If

End Sub

Private Sub Resultado()

    Label3 = _
    "Valores Ingresados: " & vals & vbCrLf & _
    "Operacion Realizada: " & res & vbCrLf & _
    "Resultado: " & total & vbCrLf & _
    "Cantidad de Valores: " & cont
    
End Sub

Private Sub Command1_Click()
    
    Cuentas
    
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Activate()
    
    Label1 = _
    "Ingrese hasta 4 numeros e indique que operacion desea realizar"
    
    Label2 = _
    "- Suma" & vbCrLf & _
    "- Resta" & vbCrLf & _
    "- Multi" & vbCrLf & _
    "- Division"

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Text1_Change()

End Sub
