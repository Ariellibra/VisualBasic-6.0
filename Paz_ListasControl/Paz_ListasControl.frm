VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   18630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "Agregar"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Remover"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Modificar"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3720
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
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
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Borrar"
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
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cargar"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   7560
      TabIndex        =   13
      Top             =   480
      Width           =   10695
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
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
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
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   3360
      Width           =   2055
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
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   2055
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
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
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
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
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
      Left            =   14880
      TabIndex        =   16
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      Index           =   3
      Left            =   5160
      TabIndex        =   12
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      Index           =   2
      Left            =   5160
      TabIndex        =   11
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      Index           =   1
      Left            =   5160
      TabIndex        =   10
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      Index           =   0
      Left            =   5160
      TabIndex        =   9
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Curso"
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
      Index           =   3
      Left            =   2760
      TabIndex        =   8
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Edad"
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
      Index           =   2
      Left            =   2760
      TabIndex        =   7
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Apellido"
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
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Nombre"
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
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer

Private Sub Command1_Click()
    
    For n = 0 To 3
        
        Label2(n) = Text1(n)
        
    Next n
    
End Sub

Private Sub Command2_Click()
  
    List1.AddItem ("El Alumno/a " & Label2(0) & " " & Label2(1) & _
    " que tiene " & Label2(2) & " años de edad esta en el curso " & Label2(3))
    
    Label3 = "Cantidad de usuarios cargados " & List1.ListCount
        
End Sub

Private Sub Command3_Click()

    List1.Clear '// borra todo
    'Print (List1.ListCount) '// cuenta la cantidad de espacios
    'Print (List1.ListIndex) '// me dice donde estoy parado
    'Print (List1.SelCount) '// cantidad de datos seleccionados
    
    Label3 = "Cantidad de usuarios cargados " & List1.ListCount
    
End Sub

Private Sub Command4_Click()
    
    End
    
End Sub

Private Sub Command5_Click()
    
    List1.List(List1.ListIndex) = "El Alumno/a " & Label2(0) & " " & Label2(1) & _
    " que tiene " & Label2(2) & " años de edad esta en el curso " & Label2(3)

End Sub

Private Sub Command6_Click()
    
    List1.RemoveItem (List1.ListIndex)
    
End Sub

Private Sub Command7_Click()
    
    List1.AddItem "Nose", 1
    
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If Index = 0 Then
        
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            
            KeyAscii = 0
            
        End If
    
    End If
    
    If Index = 1 Then
        
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            
            KeyAscii = 0
            
        End If
    
    End If
    
    If Index = 2 Then
        
        If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
            
            KeyAscii = KeyAscii
        Else
        
            KeyAscii = 0
            
        End If
    
    End If
    
    If Index = 3 Then
        
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            
            KeyAscii = 0
            
        End If
    
    End If
    
End Sub
