VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2910
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1719.324
   ScaleMode       =   0  'User
   ScaleWidth      =   4521.024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   585
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   2325
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1920
      Width           =   1380
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
      Height          =   585
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Usuario:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Contraseña:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1800
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

    End
    
End Sub

Private Sub Command1_Click()
    If Text1.Text = "programador" Then
        
        If Text2.Text = "2024" Then
        
            Unload Me
            Form2.Show
        
        Else
            MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
            Text2.SetFocus
            
        End If
        
    Else
        MsgBox "El usuario no es válido. Vuelva a intentarlo", , "Inicio de sesión"
        Text1.SetFocus
        
    End If
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
