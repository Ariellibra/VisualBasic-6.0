VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00004000&
   Caption         =   "Form4"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12615
   LinkTopic       =   "Form4"
   ScaleHeight     =   5295
   ScaleWidth      =   12615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Limpiar lista"
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
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Ver Viaje"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Guardar Datos"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
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
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
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
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1815
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
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   9840
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   8400
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   6600
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0FF&
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
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim locali As String
Dim km As String
Dim fuel As String

Dim localiAcum As String
Dim kmAcum As String
Dim fuelAcum As String

Dim localiTotal As Integer
Dim kmTotal As Long
Dim fuelTotal As Double


Private Sub Command1_Click()

    locali = Text1.Text
    km = Text2.Text
    fuel = Text3.Text
    
    localiAcum = localiAcum + vbCrLf + locali
    kmAcum = kmAcum + vbCrLf + km
    fuelAcum = fuelAcum + vbCrLf + fuel
    
    localiTotal = localiTotal + 1
    kmTotal = kmTotal + CLng(km)
    fuelTotal = fuelTotal + CDbl(fuel)
    
    Label2.Caption = _
    "Ultimo Datos de Viaje: " & vbCrLf & _
    "Localidad: " & locali & "   " & _
    "Km: " & km & "   " & _
    "Nafta: " & fuel
    

End Sub

Private Sub Command2_Click()
    
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    
    Command3.Visible = True
    
    Label3.Caption = "Localidades" & vbCrLf & vbCrLf & localiAcum & vbCrLf & "Total: " & localiTotal
    Label4.Caption = "Kilometros" & vbCrLf & vbCrLf & kmAcum & vbCrLf & "Total: " & kmTotal
    Label5.Caption = "Nafta" & vbCrLf & vbCrLf & fuelAcum & vbCrLf & "Total: " & fuelTotal

End Sub

Private Sub Command3_Click()
    
    Label3.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    
    Command3.Visible = False
    
    locali = "Merlo"
    fuel = 0
    km = 0
    
    localiAcum = "Merlo"
    fuelAcum = 0
    kmAcum = 0
    
    localiTotal = 0
    kmTotal = 0
    fuelTotal = 0
    
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    
    Label2.Caption = _
    "Ultimo Datos de Viaje: " & vbCrLf & _
    "Localidad: " & locali & "   " & _
    "Km: " & km & "   " & _
    "Nafta: " & fuel
    
    
End Sub

Private Sub Command4_Click()

    End

End Sub

Private Sub Form_Activate()

    locali = "Merlo"
    fuel = 0
    km = 0
    
    localiAcum = "Merlo"
    fuelAcum = 0
    kmAcum = 0
    
    localiTotal = 0
    kmTotal = 0
    fuelTotal = 0
    
    Label1.Caption = _
    "Ingrese los datos del Viaje, Partimos desde Merlo" & vbCrLf & _
    "Destino: " & _
    " Kilometraje: " & _
    " Nafta Gastada: "
    
    Label2.Caption = _
    "Ultimo Datos de Viaje: " & vbCrLf & _
    "Localidad: " & locali & "   " & _
    "Km: " & km & "   " & _
    "Nafta: " & fuel
    
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
