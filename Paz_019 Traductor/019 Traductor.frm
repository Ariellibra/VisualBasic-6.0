VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   5055
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
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Borrar"
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
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
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
      Left            =   360
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
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
      Height          =   615
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim palabra As String


Private Sub Command1_Click()
    
    palabra = Text1.Text
    
    If LCase(palabra) = "lunes" Then
        Label2.Caption = "Monday"
        
    ElseIf LCase(palabra) = "martes" Then
        Label2.Caption = "Tuesday"
        
    ElseIf LCase(palabra) = "miercoles" Then
        Label2.Caption = "Wednesday"
    
    ElseIf LCase(palabra) = "jueves" Then
        Label2.Caption = "Thursday"
        
    ElseIf LCase(palabra) = "viernes" Then
        Label2.Caption = "Friday"
        
    ElseIf LCase(palabra) = "sabado" Then
        Label2.Caption = "Saturday"
    
    ElseIf LCase(palabra) = "domingo" Then
        Label2.Caption = "Sunday"
        
    ElseIf LCase(palabra) = "monday" Then
        Label2.Caption = "Lunes"
        
    ElseIf LCase(palabra) = "tuesday" Then
        Label2.Caption = "Martes"
            
    ElseIf LCase(palabra) = "wednesday" Then
        Label2.Caption = "Miercoles"
    
    ElseIf LCase(palabra) = "thursday" Then
        Label2.Caption = "Jueves"
    
    ElseIf LCase(palabra) = "friday" Then
        Label2.Caption = "Viernes"
        
    ElseIf LCase(palabra) = "saturday" Then
        Label2.Caption = "Sabado"
        
    ElseIf LCase(palabra) = "Sunday" Then
        Label2.Caption = "Domingo"
    
    Else
        Label2.Caption = "Traduccion no disponible"
    
    End If
    
    Command1.Enabled = False
    
End Sub

Private Sub Command2_Click()
    
    Text1.Text = ""
    Label2.Caption = ""
    
    Command1.Enabled = True
    
End Sub

Private Sub Command3_Click()
    
    End
    
End Sub

Private Sub Form_Activate()
    
    Label1.Caption = _
    "Traductor Online" & vbCrLf & _
    "Ingles - Español" & vbCrLf & _
    "Español - Ingles" & vbCrLf & _
    "Version: Dias de la Semana"
    
    Text1.SetFocus
    
End Sub


Private Sub Text1_GotFocus()

    Text1.BackColor = &H80FF80

End Sub
Private Sub Text1_LostFocus()
    
    Text1.BackColor = &HFFFFC0

End Sub
