VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22530
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   22530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Nueva Orden"
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
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   18000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
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
      Left            =   12360
      TabIndex        =   5
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Elejir"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Elejir"
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
      Left            =   12480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   12360
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Elejir"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   4440
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   7320
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   14400
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7320
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ingrediente As String

Private Sub Command1_Click()

    If IsNumeric(Text1.Text) = True Then
        
        If Text1.Text = 1 Then
            Label2.Caption = "Ingrese el nombre del ingrediente para su Pizza Vegetariana: " & vbCrLf & _
            "Solo puede elejir un ingrediente"
            
            Label2.Visible = True
            Text2.Visible = True
            Command2.Visible = True
            
            Label3.Visible = False
            Text3.Visible = False
            Command3.Visible = False
            
            
        ElseIf Text1.Text = 2 Then
        
            Label3.Caption = "Ingrese el nombre del ingrediente para su Pizza No Vegetariana: " & vbCrLf & _
            "Solo puede elejir un ingrediente"
            
            Label2.Visible = False
            Text2.Visible = False
            Command2.Visible = False
            
            Label3.Visible = True
            Text3.Visible = True
            Command3.Visible = True
        
        Else
        
            Label2.Visible = True
            Label2.Caption = "No ingreso nada"
            
            Text2.Visible = False
            Command2.Visible = False
            Label3.Visible = False
            Text3.Visible = False
            Command3.Visible = False
            
        End If
            
    Else
                
        If LCase(Text1.Text) = "vegetariana" Then
            
            Label2.Caption = "Ingrese el nombre del ingrediente para su Pizza Vegetariana: " & vbCrLf & _
            "Solo puede elejir un ingrediente"
            
            Label2.Visible = True
            Text2.Visible = True
            Command2.Visible = True
            
            Label3.Visible = False
            Text3.Visible = False
            Command3.Visible = False
    
    
        ElseIf LCase(Text1.Text) = "no vegetariana" Then
            
            Label3.Caption = "Ingrese el nombre del ingrediente para su Pizza No Vegetariana: " & vbCrLf & _
            "Solo puede elejir un ingrediente"
            
            Label2.Visible = False
            Text2.Visible = False
            Command2.Visible = False
            
            Label3.Visible = True
            Text3.Visible = True
            Command3.Visible = True
            
        Else
            Label2.Visible = True
            Label2.Caption = "No ingreso nada"
            
            Text2.Visible = False
            Command2.Visible = False
            Label3.Visible = False
            Text3.Visible = False
            Command3.Visible = False
        
        End If
    
    End If

End Sub

Private Sub Command2_Click()

    ingrediente = Text2.Text
    
    If Not ingrediente = "" Then
    
        Label4.Caption = "La Pizza elejida es Vegetariana y los ingredientes son: " & vbCrLf & _
            "El ingrediente elejido fue:" & vbCrLf & _
            "1 - [" & ingrediente & "]" & vbCrLf & _
            "Los ingredientes bases son:" & vbCrLf & _
            "1 - [Mozzarella]" & vbCrLf & _
            "3 - [Tomate]"
            
            Label4.Visible = True
    Else
        
        Label4.Visible = True
        Label4.Caption = "No elijio un ingrediente"
        
    End If

'    If IsNumeric(Text2.Text) = True Then
'
'        If Text2.Text = 1 Then
'            c
'
'            Label4.Visible = True
'
'
'        ElseIf Text2.Text = 2 Then
'            Label4.Caption = "La Pizza elejida es Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Tofu]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'        Else
'            Label4.Visible = True
'            Label4.Caption = "Opcion Incorrecta"
'
'        End If
'
'    Else
'
'        If LCase(Text2.Text) = "pimiento" Then
'
'            Label4.Caption = "La Pizza elejida es Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Pimiento]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'
'        ElseIf LCase(Text2.Text) = "tofu" Then
'
'            Label4.Caption = "La Pizza elejida es Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Tofu]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'        Else
'            Label4.Visible = True
'            Label4.Caption = "Opcion Incorrecta"
'
'        End If
'
'    End If
    

End Sub

Private Sub Command3_Click()

    ingrediente = Text3.Text
    
    If Not ingrediente = "" Then
    
        Label4.Caption = "La Pizza elejida es No Vegetariana y los ingredientes son: " & vbCrLf & _
            "El ingrediente elejido fue:" & vbCrLf & _
            "1 - [" & ingrediente & "]" & vbCrLf & _
            "Los ingredientes bases son:" & vbCrLf & _
            "1 - [Mozzarella]" & vbCrLf & _
            "3 - [Tomate]"
            
            Label4.Visible = True
    Else
        
        Label4.Visible = True
        Label4.Caption = "No elijio un ingrediente"
        
    End If
    
    
'    If IsNumeric(Text3.Text) = True Then
'
'        If Text3.Text = 1 Then
'            Label4.Caption = "La Pizza elejida es No Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Peperoni]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'
'        ElseIf Text3.Text = 2 Then
'            Label4.Caption = "La Pizza elejida es No Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Jamon]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'        ElseIf Text3.Text = 3 Then
'            Label4.Caption = "La Pizza elejida es No Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Salmon]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'        Else
'            Label4.Visible = True
'            Label4.Caption = "Opcion Incorrecta"
'
'        End If
'
'    Else
'
'        If LCase(Text3.Text) = "peperoni" Then
'
'            Label4.Caption = "La Pizza elejida es No Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Peperoni]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'
'        ElseIf LCase(Text3.Text) = "jamon" Then
'
'            Label4.Caption = "La Pizza elejida es No Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Jamon]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'        ElseIf LCase(Text3.Text) = "salmon" Then
'
'            Label4.Caption = "La Pizza elejida es No Vegetariana y los ingredientes son: " & vbCrLf & _
'            "1 - [Salmon]" & vbCrLf & _
'            "1 - [Mozzarella]" & vbCrLf & _
'            "3 - [Tomate]"
'
'            Label4.Visible = True
'
'        Else
'            Label4.Visible = True
'            Label4.Caption = "Opcion Incorrecta"
'
'        End If
'
'    End If
    
End Sub

Private Sub Command4_Click()
    
    End
    
End Sub

Private Sub Command5_Click()
    
    Label2.Visible = False
    Text2.Visible = False
    Command2.Visible = False
            
    Label3.Visible = False
    Text3.Visible = False
    Command3.Visible = False
    
    Label4.Visible = False
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
    Label4.Caption = ""
    
    ingrediente = ""

    
End Sub

Private Sub Form_Activate()
    
    Label1.Caption = _
    "Pizzeria Bella Napoli, como quiere su Pizza? " & vbCrLf & _
    "1 - Vegetariana: " & vbCrLf & _
    "2 - No Vegetariana: "
    
End Sub

