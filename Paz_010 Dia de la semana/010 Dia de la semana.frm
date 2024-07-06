VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6270
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprobar"
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
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
      Height          =   1095
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ingrese el numero del dia del 1 al 7:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    If Text1.Text = 1 Then
        Label2.Caption = "Hoy es Lunes"
    End If
    If Text1.Text = 2 Then
        Label2.Caption = "Hoy es Martes"
    End If
    If Text1.Text = 3 Then
        Label2.Caption = "Hoy es Miercoles"
    End If
    If Text1.Text = 4 Then
        Label2.Caption = "Hoy es Jueves"
    End If
    If Text1.Text = 5 Then
        Label2.Caption = "Hoy es Viernes"
    End If
    If Text1.Text = 6 Then
        Label2.Caption = "Hoy es Sabado"
    End If
    If Text1.Text = 7 Then
        Label2.Caption = "Hoy es Domingo"
    End If
    
    
    
End Sub

Private Sub Command2_Click()

    End
    
End Sub

Private Sub Text1_GotFocus()

    Text1.BackColor = &H80FF80

End Sub

Private Sub Text1_LostFocus()

    Text1.BackColor = &HFFFFFF

End Sub
