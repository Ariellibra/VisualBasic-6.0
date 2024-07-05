VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   9255
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
      Height          =   615
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
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
      Left            =   7200
      Style           =   1  'Graphical
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
      Left            =   5280
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
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
      Left            =   840
      TabIndex        =   0
      Top             =   1560
      Width           =   1455
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Escriba un numero para ver los impares"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   4095
   End
   Begin VB.Label Label3 
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
      Height          =   435
      Left            =   4800
      TabIndex        =   6
      Top             =   2520
      Width           =   4035
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
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
      Height          =   435
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   4095
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Cual es su edad?"
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
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n As Integer
Dim edad As Integer
Dim edad_str As String
Dim j As Integer
Dim numImpar As Integer
Dim numImpar_str As String



Private Sub Command1_Click()
    
    edad_str = ""
    edad = Text1.Text
    
    For n = 1 To edad
        
        If n = edad Then
            edad_str = edad_str & n
            
        Else
            edad_str = edad_str & n & ", "
            
        End If
        
    Next n
    
    Label2.Caption = edad_str
    
End Sub

Private Sub Command2_Click()
    
    numImpar = Text2.Text
    
    For j = 1 To numImpar
        
        numImpar_str = numImpar_str & j & ", "
        j = j + 1
    
    Next j
    
    Label3.Caption = numImpar_str
    
End Sub

Private Sub Command3_Click()
    
    End
    
End Sub

