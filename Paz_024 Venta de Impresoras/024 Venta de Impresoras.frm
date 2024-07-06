VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
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
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Vale de Regalo"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "T. de Credito"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Efectivo"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
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
      Left            =   6600
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Height          =   2775
      Left            =   840
      TabIndex        =   7
      Top             =   2760
      Width           =   7695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Forma de pago"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese cuantas impresoras quiere"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cantImpre As Integer
Dim precioImpre As Double
Dim precioVenta As Double
Dim descuento As Double

Private Sub Command1_Click()
    
    End

End Sub

Private Sub Command2_Click()
    
    cantImpre = CInt(Text1.Text)
    precioVenta = ((precioImpre * 1.21) * cantImpre) * 0.9
    
    Label3.Caption = _
    "Factura: " & vbCrLf & _
    "Cantidad de Impresoras: " & cantImpre & vbCrLf & _
    "Precio Unitario con IVA: $ " & (precioImpre * 1.21) & vbCrLf & _
    "Total sin descuento: $ " & ((precioImpre * 1.21) * cantImpre) & vbCrLf & _
    "Forma de Pago: Efectivo" & vbCrLf & _
    "Descuento: $ " & ((precioImpre * 1.21) * cantImpre) * 0.1 & vbCrLf & _
    "Total a Pagar: $ " & precioVenta
    
    Text1.Text = ""

End Sub

Private Sub Command3_Click()

    cantImpre = CInt(Text1.Text)
    precioVenta = ((precioImpre * 1.21) * cantImpre) * 0.95
    
    Label3.Caption = _
    "Factura: " & vbCrLf & _
    "Cantidad de Impresoras: " & cantImpre & vbCrLf & _
    "Precio Unitario con IVA: $ " & (precioImpre * 1.21) & vbCrLf & _
    "Total sin descuento: $ " & ((precioImpre * 1.21) * cantImpre) & vbCrLf & _
    "Forma de Pago: Tarjeta de Credito" & vbCrLf & _
    "Descuento: $ " & ((precioImpre * 1.21) * cantImpre) * 0.05 & vbCrLf & _
    "Total a Pagar: $ " & precioVenta
    
    Text1.Text = ""

End Sub

Private Sub Command4_Click()

    cantImpre = CInt(Text1.Text)
    precioVenta = ((precioImpre * 1.21) * cantImpre) * 0.85
    
    Label3.Caption = _
    "Factura: " & vbCrLf & _
    "Cantidad de Impresoras: " & cantImpre & vbCrLf & _
    "Precio Unitario con IVA: $ " & (precioImpre * 1.21) & vbCrLf & _
    "Total sin descuento: $ " & ((precioImpre * 1.21) * cantImpre) & vbCrLf & _
    "Forma de Pago: Vale de Regalo" & vbCrLf & _
    "Descuento: $ " & ((precioImpre * 1.21) * cantImpre) * 0.15 & vbCrLf & _
    "Total a Pagar: $ " & precioVenta
    
    Text1.Text = ""

End Sub

Private Sub Form_Activate()
    
    cantImpre = 0
    precioImpre = 27080

End Sub

Private Sub Text1_GotFocus()
    
    Text1.BackColor = &HC0FFC0
    
End Sub

Private Sub Text1_LostFocus()
    
    Text1.BackColor = &H80000005
    
End Sub
