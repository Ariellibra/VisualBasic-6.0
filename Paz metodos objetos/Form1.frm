VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   10050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hace Magia"
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   2295
      Left            =   1080
      TabIndex        =   0
      Top             =   2280
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim datosA As String
Dim chara As String
Dim datosB() As String
Dim datosImprimir() As String
Dim lista As New ArrayL

Private Sub Command1_Click()
    
    datosA = Text1
    chara = Text2
    
    'datosB = DividirDatos(datosA, "*")
    
    'MostrarDatos DividirDatos(datosA, "*")
    
    lista.toSplit datosA, chara
    
    lista.toPrint
    
    
End Sub

Private Function MostrarDatos(datos() As String)
    Dim a As Integer
    
    For a = LBound(datos()) To UBound(datos())
        
        Print datos(a)
        
    Next a
    

End Function

Private Function DividirDatos(datos As String, evaluador As String) As String()

    Dim res() As String
    
    res = Split(datos, evaluador)
    
    DividirDatos = res
    
End Function

Private Sub Form_Load()
    
    Text1 = ""
    Text2 = ""
    
End Sub

Private Sub label1_Click()

End Sub
