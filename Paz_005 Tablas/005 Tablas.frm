VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00004000&
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
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
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command12 
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
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
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
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400040&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400040&
      Caption         =   "Escriba el numero y aprete el boton para ver la tabla de multiplicar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Label2.Caption = _
    Text1.Text & " x 0 = " & CInt(Text1.Text) * 0 & vbCrLf & _
    Text1.Text & " x 1 = " & CInt(Text1.Text) * 1 & vbCrLf & _
    Text1.Text & " x 2 = " & CInt(Text1.Text) * 2 & vbCrLf & _
    Text1.Text & " x 3 = " & CInt(Text1.Text) * 3 & vbCrLf & _
    Text1.Text & " x 4 = " & CInt(Text1.Text) * 4 & vbCrLf & _
    Text1.Text & " x 5 = " & CInt(Text1.Text) * 5 & vbCrLf & _
    Text1.Text & " x 6 = " & CInt(Text1.Text) * 6 & vbCrLf & _
    Text1.Text & " x 7 = " & CInt(Text1.Text) * 7 & vbCrLf & _
    Text1.Text & " x 8 = " & CInt(Text1.Text) * 8 & vbCrLf & _
    Text1.Text & " x 9 = " & CInt(Text1.Text) * 9 & vbCrLf & _
    Text1.Text & " x 10 = " & CInt(Text1.Text) * 10
    
End Sub

Private Sub Command12_Click()
    
    End
    
End Sub

