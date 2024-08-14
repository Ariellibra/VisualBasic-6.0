VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C000C0&
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22590
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   22590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   20280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Traducir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Codificar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Height          =   6315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   21915
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n, i As Integer
Dim oracion As String
Dim codigoArray() As String


Private Sub Codificar()
    i = 0
    Label1 = ""
    
    ReDim codigoArray(Len(oracion))
    
    For n = 1 To Len(oracion)
    
        codigoArray(i) = Mid(oracion, n, 1)
        i = i + 1
    
    Next n
    
    For n = 0 To Len(oracion) - 1
        
        codigoArray(n) = Asc(codigoArray(n))
        If n = 50 Or n = 100 Or n = 150 Or n = 200 Or n = 250 Or n = 300 Or n = 350 Or n = 400 Or n = 450 Then
            Label1 = Label1 & vbCrLf
        Else
            Label1 = Label1 & codigoArray(n)
        End If
    Next n
    
End Sub

Private Sub Traducir()
    
    Label1 = ""

    For n = 0 To Len(oracion) - 1
        
        Label1 = Label1 & Chr(CInt(codigoArray(n)))
        
    Next n
    
End Sub

Private Sub Command1_Click()
    
    Codificar

End Sub

Private Sub Command2_Click()
    
    Traducir
    
End Sub

Private Sub Command3_Click()
    
    End
    
End Sub

Private Sub Form_Activate()
    
    oracion = "A lo largo de la historia, han surgido muchas narraciones capaces de inspirar a generaciones enteras durante décadas o incluso siglos. En este sentido, es posible encontrar historias cortas que, incluso cuando han nacido como tradición oral, se han popularizado tanto que su formato en papel se ha extendido rápidamente a lo largo y ancho de países enteros porque . Es por ello que este es uno de los tipos de literatura con más relevancia cultural."
    Label1 = oracion
    
End Sub

Private Sub Label1_Click()

End Sub
