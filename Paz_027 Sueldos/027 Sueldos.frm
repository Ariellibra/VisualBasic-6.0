VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Limpiar Datos"
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
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cargar Sueldo"
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
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
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
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Puede agregar hasta 5 Sueldos:"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sueldo1, sueldo2, sueldo3, sueldo4, sueldo5 As Integer

Dim mayor, menor As Integer
Dim promedio As Double

Dim contador As Integer


Private Sub Command1_Click()
    
    If contador = 1 Then
        
        sueldo1 = CInt(Text1.Text)
        mayor = sueldo1
        menor = sueldo1
        
        Label1.Caption = "Puede agregar hasta 5 Sueldos: " & vbCrLf & "Sueldos cargados: " & contador & vbCrLf & _
        "Sueldo 1: $ " & sueldo1
        
        Text1.Text = ""
        
    ElseIf contador = 2 Then
        
        sueldo2 = CInt(Text1.Text)
        If menor < sueldo2 Then
            If mayor < sueldo2 Then
                mayor = sueldo2
            End If
        ElseIf sueldo2 < menor Then
            menor = sueldo2
        End If
        
        Label1.Caption = "Puede agregar hasta 5 Sueldos: " & vbCrLf & "Sueldos cargados: " & contador & vbCrLf & _
        "Sueldo 1: $ " & sueldo1 & vbCrLf & _
        "Sueldo 2: $ " & sueldo2
        
        Text1.Text = ""
        
    ElseIf contador = 3 Then
    
        sueldo3 = CInt(Text1.Text)
        If menor < sueldo3 Then
            If mayor < sueldo3 Then
                mayor = sueldo3
            End If
        ElseIf sueldo3 < menor Then
            menor = sueldo3
        End If
        
        Label1.Caption = "Puede agregar hasta 5 Sueldos: " & vbCrLf & "Sueldos cargados: " & contador & vbCrLf & _
        "Sueldo 1: $ " & sueldo1 & vbCrLf & _
        "Sueldo 2: $ " & sueldo2 & vbCrLf & _
        "Sueldo 3: $ " & sueldo3
        
        Text1.Text = ""
        
    ElseIf contador = 4 Then
    
        sueldo4 = CInt(Text1.Text)
        If menor < sueldo4 Then
            If mayor < sueldo4 Then
                mayor = sueldo4
            End If
        ElseIf sueldo4 < menor Then
            menor = sueldo4
        End If
        
        Label1.Caption = "Puede agregar hasta 5 Sueldos: " & vbCrLf & "Sueldos cargados: " & contador & vbCrLf & _
        "Sueldo 1: $ " & sueldo1 & vbCrLf & _
        "Sueldo 2: $ " & sueldo2 & vbCrLf & _
        "Sueldo 3: $ " & sueldo3 & vbCrLf & _
        "Sueldo 4: $ " & sueldo4
        
        Text1.Text = ""
        
        
    ElseIf contador = 5 Then
    
        sueldo5 = CInt(Text1.Text)
        If menor < sueldo5 Then
            If mayor < sueldo5 Then
                mayor = sueldo5
            End If
        ElseIf sueldo5 < menor Then
            menor = sueldo5
        End If
        
        Label1.Caption = "Puede agregar hasta 5 Sueldos: " & vbCrLf & "Sueldos cargados: " & contador & vbCrLf & _
        "Sueldo 1: $ " & sueldo1 & vbCrLf & _
        "Sueldo 2: $ " & sueldo2 & vbCrLf & _
        "Sueldo 3: $ " & sueldo3 & vbCrLf & _
        "Sueldo 4: $ " & sueldo4 & vbCrLf & _
        "Sueldo 5: $ " & sueldo5
        
        Command1.Enabled = False
        Text1.Enabled = False
        
        promedio = (sueldo1 + sueldo2 + sueldo3 + sueldo4 + sueldo5) / 5
        
        Label1.Caption = "Puede agregar hasta 5 Sueldos: " & vbCrLf & "Sueldos cargados: " & contador & vbCrLf & _
        "Sueldo 1: $ " & sueldo1 & vbCrLf & _
        "Sueldo 2: $ " & sueldo2 & vbCrLf & _
        "Sueldo 3: $ " & sueldo3 & vbCrLf & _
        "Sueldo 4: $ " & sueldo4 & vbCrLf & _
        "Sueldo 5: $ " & sueldo5 & vbCrLf & _
        "Datos de los sueldos" & vbCrLf & _
        "Sueldo mas Alto: $ " & mayor & vbCrLf & _
        "Sueldo mas Bajo: $ " & menor & vbCrLf & _
        "Promedio: $ " & promedio
        
        Text1.Text = ""
        
    End If
    
    contador = contador + 1
    
    
    
End Sub

Private Sub Command2_Click()
    
    Command1.Enabled = True
    Text1.Enabled = True
    
    Label1.Caption = ""
    
    sueldo1 = 0
    sueldo2 = 0
    sueldo3 = 0
    sueldo4 = 0
    sueldo5 = 0
    
    promedio = 0
    contador = 1
    mayor = 0
    menor = 0

End Sub

Private Sub Command3_Click()
    End
End Sub

Private Sub Form_Activate()

    sueldo1 = 0
    sueldo2 = 0
    sueldo3 = 0
    sueldo4 = 0
    sueldo5 = 0
    
    promedio = 0
    contador = 1
    mayor = 0
    menor = 0

End Sub

