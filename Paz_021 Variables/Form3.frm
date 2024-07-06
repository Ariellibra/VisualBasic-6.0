VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00008000&
   Caption         =   "Form3"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form3"
   ScaleHeight     =   6360
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Atras"
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
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text3 
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Multiplicar"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sumar"
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1575
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
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      Top             =   1800
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
      Height          =   735
      Left            =   720
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
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
      Height          =   975
      Left            =   5160
      TabIndex        =   10
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese 1 numero para multiplicar todo"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
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
      Height          =   975
      Left            =   720
      TabIndex        =   8
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese 2 numeros a sumar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num1, num2, num3, res As Integer

Private Sub Form_Activate()

    num1 = 0
    num2 = 0
    num3 = 0
    res = 0
    
End Sub


Private Sub Command1_Click()

    num1 = CInt(Text1.Text)
    num2 = CInt(Text2.Text)
    
    res = num1 + num2
    
    Label2.Visible = True
    Label2.Caption = "El resultado de la suna es " & res
    
    Label3.Visible = True
    Text3.Visible = True
    Command3.Visible = True
    
    Text1.Text = ""
    Text2.Text = ""

End Sub

Private Sub Command2_Click()
    
    End
    
End Sub

Private Sub Command3_Click()
    
    num3 = CInt(Text3.Text)
    
    res = res * num3
    
    Label4.Caption = "El resultado de la multiplicacion es " & res
    Label4.Visible = True
    
    Text3.Text = ""
    
End Sub

Private Sub Command4_Click()
    
    Unload Form3
    Form2.Show

End Sub


Private Sub Text1_GotFocus()
    
    Text1.BackColor = &HC0FFC0
    
End Sub


Private Sub Text1_LostFocus()

    Text1.BackColor = &H80000005

End Sub
Private Sub Text2_GotFocus()
    
    Text2.BackColor = &HC0FFC0
    
End Sub


Private Sub Text2_LostFocus()

    Text2.BackColor = &H80000005

End Sub
Private Sub Text3_GotFocus()
    
    Text3.BackColor = &HC0FFC0
    
End Sub


Private Sub Text3_LostFocus()

    Text3.BackColor = &H80000005

End Sub

