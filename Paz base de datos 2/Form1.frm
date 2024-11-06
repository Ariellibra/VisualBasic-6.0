VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   13620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Mostrar datos"
      Height          =   615
      Left            =   2280
      TabIndex        =   8
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Mostrar datos"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Datos"
      Height          =   615
      Left            =   2280
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Height          =   2175
      Left            =   4080
      TabIndex        =   6
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim base As Database
Dim registro As Recordset

Private Sub Command1_Click()
        
        
        registro.AddNew
        registro.Fields("NombrePerro") = Text1(0)
        registro.Fields("NombreDueño") = Text1(1)
        registro.Fields("Edad") = CInt(Text1(2))
        registro.Fields("Direccion") = Text1(3)
        registro.Fields("Observaciones") = Text1(4)
        registro.Update
        
        
End Sub

Private Sub Command2_Click()
    

    If registro.EOF = True Then
    
        registro.MoveFirst
        mostrar
        registro.MoveNext
    
    Else
        
        mostrar
        registro.MoveNext
        
    End If
    

End Sub

Private Sub mostrar()
    
    Label1.Caption = _
    "Nombre Perro: " & registro.Fields("NombrePerro").Value & vbCrLf & _
    "Nombre Dueño: " & registro.Fields("NombreDueño").Value & vbCrLf & _
    "Edad: " & registro.Fields("Edad").Value & vbCrLf & _
    "Direccion: " & registro.Fields("Direccion").Value & vbCrLf & _
    "Observaciones: " & registro.Fields("Observaciones").Value

    
End Sub

Private Sub Command3_Click()
    
    registro.Delete
    
End Sub

Private Sub Form_Activate()
    
    Set base = OpenDatabase(App.Path & "\bd2.mdb")
    Set registro = base.OpenRecordset("Perros", dbOpenTable)
    
End Sub

