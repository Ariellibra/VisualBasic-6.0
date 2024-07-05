VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   9375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Comprar"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
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
      Left            =   4200
      TabIndex        =   10
      Top             =   1800
      Width           =   975
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
      TabIndex        =   9
      Top             =   1800
      Width           =   975
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
      Left            =   600
      TabIndex        =   8
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton Command6 
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Nueva Venta"
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
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver Resumen"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label9 
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
      Height          =   1575
      Left            =   6120
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Height          =   1575
      Left            =   4680
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label7 
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
      Height          =   1575
      Left            =   3240
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Height          =   1575
      Left            =   240
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Premium 60$"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Gold 45$"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2040
      TabIndex        =   5
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Normal 30$"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
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
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Video Edit 2.0"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cantVentas, cantLicN, cantLicG, cantLicP As Integer
Dim totalVenta, totalVentaTotal, totalVentaN, totalVentaG, totalVentaP As Integer
Dim totalVentLicN, totalVentLicG, totalVentLicP As Integer

Private Sub Command1_Click()
        
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
        
    Text1.Enabled = False
    Text2.Enabled = False
    Text3.Enabled = False
    
    Command1.Enabled = False
    
    If IsNumeric(Text1.Text) = True Then
        If Text1.Text > 0 Then
            
            totalVenta = totalVenta + (CInt(Text1.Text) * 30)
            
            totalVentaN = totalVentaN + (CInt(Text1.Text) * 30)
            cantLicN = cantLicN + CInt(Text1.Text)
            totalVentLicN = totalVentLicN + 1
            
        End If
    End If
    
    If IsNumeric(Text2.Text) = True Then
        If Text2.Text > 0 Then
            
            totalVenta = totalVenta + (CInt(Text2.Text) * 45)
            
            totalVentaG = totalVentaG + (CInt(Text2.Text) * 45)
            cantLicG = cantLicG + CInt(Text2.Text)
            totalVentLicG = totalVentLicG + 1
            
        End If
        
    End If
    
    If IsNumeric(Text3.Text) = True Then
        If Text3.Text > 0 Then
            
            totalVenta = totalVenta + (CInt(Text3.Text) * 60)
            
            totalVentaP = totalVentaP + (CInt(Text3.Text) * 60)
            cantLicP = cantLicP + CInt(Text3.Text)
            totalVentLicP = totalVentLicP + 1
            
        End If
        
    End If
    
    cantVentas = cantVentas + 1
    totalVentaTotal = totalVentaTotal + totalVenta
    
    Label2.Visible = True
    
    Label2.Caption = _
    "Total de la Venta: $" & totalVenta
    
    
End Sub

Private Sub Command4_Click()
        
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True

    Label6.Caption = _
    "Resumen" & vbCrLf & _
    "Cant. Lic. Vendidas: " & vbCrLf & _
    "Ventas realizadas x Lic: " & vbCrLf & _
    "Total Recaudado x Lic: " & vbCrLf & _
    "Total Recaudado: $ " & totalVentaTotal
    
    Label7.Caption = _
    "Normales" & vbCrLf & _
    cantLicN & vbCrLf & _
    totalVentLicN & vbCrLf & _
    "$ " & totalVentaN
    
    Label8.Caption = _
    "Gold" & vbCrLf & _
    cantLicG & vbCrLf & _
    totalVentLicG & vbCrLf & _
    "$ " & totalVentaG
    
    Label9.Caption = _
    "Premium" & vbCrLf & _
    cantLicP & vbCrLf & _
    totalVentLicP & vbCrLf & _
    "$ " & totalVentaP
    
End Sub

Private Sub Command5_Click()
                
    Command1.Enabled = True
        
    totalVenta = 0
    
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    
    Text1.Enabled = True
    Text2.Enabled = True
    Text3.Enabled = True
    
    Label2.Caption = ""
    Label2.Visible = False
    
    Label6.Visible = False
    Label7.Visible = False
    Label8.Visible = False
    Label9.Visible = False
    
End Sub

Private Sub Command6_Click()
    
    End
    
End Sub

