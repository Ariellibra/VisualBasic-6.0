VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   11640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Salir"
      Height          =   615
      Left            =   10080
      TabIndex        =   9
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Division"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   7
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Multiplicacion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   6
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Resta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Suma"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   2
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5640
      TabIndex        =   8
      Top             =   2400
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Label1.Caption = "Se realizo la operacion SUMA entre los valores: " & " " & Text1 & ", " & Text2 & ", " & Text3 & ", " & Text4 & " y el resultado es: " & (CInt(Text1) + CInt(Text2) + CInt(Text3) + CInt(Text4))
    
    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = " "
    Text4.Text = " "

End Sub

Private Sub Command2_Click()

    Label1.Caption = "Se realizo la operacion RESTA entre los valores " & " " & Text1 & ", " & Text2 & ", " & Text3 & ", " & Text4 & " y el resultado es: " & (CInt(Text1) - CInt(Text2) - CInt(Text3) - CInt(Text4))
    
    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = " "
    Text4.Text = " "

End Sub

Private Sub Command3_Click()

    Label1.Caption = "Se realizo la operacion MULTIPLICACION entre los valores " & " " & Text1 & ", " & Text2 & ", " & Text3 & ", " & Text4 & " y el resultado es: " & (CLng(Text1) * CLng(Text2) * CLng(Text3) * CLng(Text4))
    
    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = " "
    Text4.Text = " "

End Sub

Private Sub Command4_Click()

    Label1.Caption = "Se realizo la operacion DIVISION entre los valores " & " " & Text1 & ", " & Text2 & ", " & Text3 & ", " & Text4 & " y el resultado es: " & (CInt(Text1) / CInt(Text2) / CInt(Text3) / CInt(Text4))
    
    Text1.Text = " "
    Text2.Text = " "
    Text3.Text = " "
    Text4.Text = " "

End Sub

Private Sub Command5_Click()
    End
    
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Form_Load()

End Sub

Private Sub Text1_GotFocus()
    Text1.BackColor = RGB(51, 255, 172)

End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = RGB(255, 255, 255)

End Sub

Private Sub Text2_GotFocus()
    Text2.BackColor = RGB(51, 255, 172)

End Sub

Private Sub Text2_LostFocus()
    Text2.BackColor = RGB(255, 255, 255)

End Sub

Private Sub Text3_GotFocus()
    Text3.BackColor = RGB(51, 255, 172)
    
End Sub

Private Sub Text3_LostFocus()
    Text3.BackColor = RGB(255, 255, 255)

End Sub

Private Sub Text4_GotFocus()
    Text4.BackColor = RGB(51, 255, 172)

End Sub

Private Sub Text4_LostFocus()
    Text4.BackColor = RGB(255, 255, 255)

End Sub
