VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
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
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
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
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
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
      Height          =   3615
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese el vuelto, para saber que billetes entregar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vuelto, cont1000, cont500, cont200, cont100, cont50, cont20, cont10 As Integer


Private Sub Monedero()
    
    vuelto = CInt(Text1)
    
    Mil
    Quinientos
    Doscientos
    Cien
    Cincuenta
    Veinte
    Dies
    
    Label2 = _
    "Billetes de $ 1000 = " & cont1000 & vbCrLf & _
    "Billetes de $ 500 = " & cont500 & vbCrLf & _
    "Billetes de $ 200 = " & cont200 & vbCrLf & _
    "Billetes de $ 100 = " & cont100 & vbCrLf & _
    "Billetes de $ 50 = " & cont50 & vbCrLf & _
    "Billetes de $ 20 = " & cont20 & vbCrLf & _
    "Billetes de $ 10 = " & cont10 & vbCrLf & _
    "Y sobro $ " & vuelto
    
    Limpiar
    

End Sub

Private Sub Mil()

    If vuelto >= 1000 Then
        
        Do While vuelto >= 1000
        
            vuelto = vuelto - 1000
            cont1000 = cont1000 + 1
        Loop
        
    End If

End Sub
Private Sub Quinientos()

    If vuelto >= 500 Then
        
        Do While vuelto >= 500
        
            vuelto = vuelto - 500
            cont500 = cont500 + 1
        Loop
        
    End If

End Sub
Private Sub Doscientos()

    If vuelto >= 200 Then
        
        Do While vuelto >= 200
        
            vuelto = vuelto - 200
            cont200 = cont200 + 1
        Loop
        
    End If
End Sub
Private Sub Cien()

    If vuelto >= 100 Then
        
        Do While vuelto >= 100
        
            vuelto = vuelto - 100
            cont100 = cont100 + 1
        Loop
        
    End If

End Sub

Private Sub Cincuenta()

    If vuelto >= 50 Then
        
        Do While vuelto >= 50
        
            vuelto = vuelto - 50
            cont50 = cont50 + 1
        Loop
        
    End If

End Sub
Private Sub Veinte()

    If vuelto >= 20 Then
        
        Do While vuelto >= 20
        
            vuelto = vuelto - 20
            cont20 = cont20 + 1
        Loop
        
    End If


End Sub
Private Sub Dies()

    If vuelto >= 10 Then
        
        Do While vuelto >= 10
        
            vuelto = vuelto - 10
            cont10 = cont10 + 1
        Loop
        
    End If

End Sub

Private Sub Limpiar()
    
    cont1000 = 0
    cont500 = 0
    cont200 = 0
    cont100 = 0
    cont50 = 0
    cont20 = 0
    cont10 = 0
    
End Sub

Private Sub Command1_Click()
    
    Monedero
    
End Sub

Private Sub Command2_Click()
    
    End
    
End Sub

Private Sub Form_Activate()

    cont1000 = 0
    cont500 = 0
    cont200 = 0
    cont100 = 0
    cont50 = 0
    cont20 = 0
    cont10 = 0
    
End Sub

