VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Cerrar sesion"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000080FF&
      Caption         =   "Calcular"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text4 
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
      Left            =   3480
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
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
      Left            =   2280
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
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
      Height          =   735
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Height          =   735
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   1455
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
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   2295
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
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   4680
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    End

End Sub

Private Sub Command2_Click()

    If Text1.Text = "Libra" Then
        If Text2.Text = "46214621" Then
            Label2.Caption = "Sesion Iniciada: " & vbCrLf & "Usuario: " & Text1.Text
            
            Text1.Text = ""
            Text2.Text = ""
            
            Command4.Visible = True
            Label2.Visible = True
            Label3.Visible = True
            Text3.Visible = True
            Text4.Visible = True
            Command3.Visible = True
            
            Label1.Visible = False
            Text1.Visible = False
            Text2.Visible = False
            Command2.Visible = False
            
        Else
            Label2.Visible = True
            Label2.Caption = "Contraseña incorrecta"
        End If
    Else
        Label2.Visible = True
        Label2.Caption = "Usuario o Contraseña incorrecto"
    End If
    
        
End Sub

Private Sub Command3_Click()
    
    If CInt(Text4.Text) = 0 Then
        Text3.Text = ""
        Text4.Text = ""
        Label4.Height = 975
        Label4.Visible = True
        Label4.Caption = "No se puede dividir entre 0, vuelva a hacer otra division"
        
    Else
        Label4.Caption = "El resultado de dividir" & vbCrLf & Text3.Text & " entre " & Text4.Text & " es: " & vbCrLf & (CLng(Text3.Text) / CLng(Text4.Text))
        Text3.Text = ""
        Text4.Text = ""
        Label4.Height = 975
        Label4.Visible = True

    End If
    
End Sub

Private Sub Command4_Click()

    Label1.Visible = True
    Text1.Visible = True
    Text2.Visible = True
    Command2.Visible = True
    Label2.Visible = False
    Label3.Visible = False
    Label4.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Command3.Visible = False
    Command4.Visible = False
    
End Sub

Private Sub Form_Load()

    Label1.Caption = "Ingrese al programa" & vbCrLf & "Usuario" & vbCrLf & "Contraseña"
    
    Label3.Caption = "Ingrese el Divisor y el Dividendo"

End Sub

Private Sub Text1_GotFocus()

    Text1.BackColor = &H80FF80
    
End Sub

Private Sub Text1_LostFocus()

    Text1.BackColor = &H80000005
    
End Sub

Private Sub Text2_GotFocus()

    Text2.BackColor = &H80FF80
    
End Sub

Private Sub Text2_LostFocus()

    Text2.BackColor = &H80000005
    
End Sub

Private Sub Text3_GotFocus()

    Text3.BackColor = &H80FF80
    
End Sub

Private Sub Text3_LostFocus()

    Text3.BackColor = &H80000005
    
End Sub

Private Sub Text4_GotFocus()

    Text4.BackColor = &H80FF80
    
End Sub

Private Sub Text4_LostFocus()

    Text4.BackColor = &H80000005
    
End Sub

