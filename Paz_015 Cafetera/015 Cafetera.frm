VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   3615
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
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
      Height          =   615
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1935
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
      Height          =   1215
      Left            =   360
      TabIndex        =   3
      Top             =   1920
      Width           =   2415
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
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If IsNumeric(Text1.Text) = True Then
        If Text1.Text = 1 Then

            Label2.Caption = "Ya he preparado su 'Cafe'. Disfrutelo"
        ElseIf Text1.Text = 2 Then

            Label2.Caption = "Ya he preparado su 'Cappuccino'. Disfrutelo"
        ElseIf Text1.Text = 3 Then

            Label2.Caption = "Ya he preparado su 'Lagrima'. Disfrutelo"
        Else
            Label2.Caption = "Opcion Incorrecta"
        End If
            
    Else
                
        If LCase(Text1.Text) = "cafe" Then
            
            Label2.Caption = "Ya he preparado su 'Cafe'. Disfrutelo"
    
    
        ElseIf LCase(Text1.Text) = "cappuccino" Then
            
            Label2.Caption = "Ya he preparado su 'Cappuccino'. Disfrutelo"
    
    
        ElseIf LCase(Text1.Text) = "lagrima" Then
            
            Label2.Caption = "Ya he preparado su 'Lagrima'. Disfrutelo"
        
        Else
            Label2.Caption = "Opcion Incorrecta"
        
        End If
    
    End If
    
    Text1.SetFocus
    
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Activate()
    
    Label1.Caption = _
    "Menu: " & vbCrLf & _
    "1 - Cafe: " & vbCrLf & _
    "2 - Cappuccino: " & vbCrLf & _
    "3 - Lagrima: "

End Sub

Private Sub Text1_GotFocus()

    Text1.BackColor = &H80FF80

End Sub
Private Sub Text1_LostFocus()
    
    Text1.BackColor = &HFFFFC0

End Sub
