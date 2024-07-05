VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10455
   LinkTopic       =   "Form1"
   ScaleHeight     =   9495
   ScaleWidth      =   10455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
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
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ingresar"
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
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Reiniciar"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Reiniciar"
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
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ingresar"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
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
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese un numero del 1 al 100 para comenzar la cuenta regresiva"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   6240
      TabIndex        =   11
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Height          =   315
      Left            =   6240
      TabIndex        =   10
      Top             =   2880
      Width           =   3915
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "La suma de los numeros del 1 al 100 es : "
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   9
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
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
      Height          =   1455
      Left            =   360
      TabIndex        =   8
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese un numero mayor a 0 para calcular su factorial"
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
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim numIngresar, n As Integer
Dim factorial As Integer

Dim i As Integer
Dim sum100 As Integer

Dim j As Integer
Dim cuentaAtras As String
Dim numIngresar2, num100 As Integer


Private Sub Command1_Click()
        
        numIngresar = CInt(Text1.Text)
        
        If numIngresar > 0 Then
            For n = 1 To numIngresar
                factorial = factorial * n
                
            Next n
            
            Label2.Caption = _
                "El resultado del factorial de " & numIngresar & "! es: " & factorial
            Command1.Enabled = False
            Text1.Enabled = False
        Else
            Label2.Caption = "Ingrese un dato correto"
            Command1.Enabled = False
            Text1.Enabled = False
            
        End If
        
End Sub

Private Sub Command2_Click()
    
    n = 1
    numIngresar = 0
    factorial = 1
    
    Label2.Caption = ""
    
    
    Text1.Text = ""
    Command1.Enabled = True
    Text1.Enabled = True
    Text1.SetFocus
    
End Sub

Private Sub Command3_Click()
    
    cuentaAtras = ""
    num100 = 100
    numIngresar2 = 0
    
    Text2.Text = ""
    Text2.Enabled = True
    Command4.Enabled = True
    Label4.Caption = ""
    
    Text2.SetFocus
    
End Sub

Private Sub Command4_Click()
        
    numIngresar2 = CInt(Text2.Text)
    
    For j = 1 To 100
        
        If numIngresar2 = num100 Then
            cuentaAtras = cuentaAtras & num100
            
            Text2.Enabled = False
            Command4.Enabled = False
            
            Exit For
        Else
            cuentaAtras = cuentaAtras & num100 & ", "
            num100 = num100 - 1
        End If
        
    Next j
    
    Label4.Caption = cuentaAtras
    
End Sub

Private Sub Command5_Click()
    
    End
    
End Sub

Private Sub Form_Activate()
    
    j = 1
    n = 1
    factorial = 1
    num100 = 100
    
    For i = 1 To 100
        
        sum100 = sum100 + i
        i = i + 1
        
    Next i
    
    Label3.Caption = "La suma de los numeros impares del 1 al 100 es: " & sum100
    
    
End Sub

