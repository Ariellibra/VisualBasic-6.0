VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00808080&
   Caption         =   "Form2"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   8535
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command17 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H0080C0FF&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H0080FF80&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H008080FF&
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H0080C0FF&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H0080C0FF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H0080C0FF&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim num1 As Long
Dim num2 As Long

Dim num1_str As String
Dim num2_str As String

Dim op As String
Dim opSumaT, opRestaT, opMultiT, opDivT As Boolean

Dim estaVacio As Boolean

Dim receptor As Boolean

Dim resultado As Double


Private Sub Command10_Click()
    
    If receptor = True Then
        num1_str = num1_str + "0"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "0"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command11_Click()
    
    op = num1_str + "+"
    
    opSumaT = True
    opRestaT = False
    opDivT = False
    opMultiT = False
    
    receptor = False
    estaVacio = True
    
    Label1.Caption = op
    
    Command11.Enabled = False
    Command12.Enabled = True
    Command13.Enabled = True
    Command16.Enabled = True
    
    Command11.BackColor = &HFFC0FF
    Command12.BackColor = &H80C0FF
    Command13.BackColor = &H80C0FF
    Command16.BackColor = &H80C0FF
    
    num2_str = ""
    num2 = 0
    
End Sub

Private Sub Command12_Click()
    
    op = num1_str + "-"
    
    opSumaT = False
    opRestaT = True
    opDivT = False
    opMultiT = False
    
    receptor = False
    estaVacio = True
    
    Label1.Caption = op
    
    Command11.Enabled = True
    Command12.Enabled = False
    Command13.Enabled = True
    Command16.Enabled = True
    
    Command11.BackColor = &H80C0FF
    Command12.BackColor = &HFFC0FF
    Command13.BackColor = &H80C0FF
    Command16.BackColor = &H80C0FF
    
    num2_str = ""
    num2 = 0
    
End Sub

Private Sub Command13_Click()
    
    op = num1_str + "/"
    
    opSumaT = False
    opRestaT = False
    opDivT = True
    opMultiT = False
    
    receptor = False
    estaVacio = True
    
    Label1.Caption = op
    
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = False
    Command16.Enabled = True
    
    Command11.BackColor = &H80C0FF
    Command12.BackColor = &H80C0FF
    Command13.BackColor = &HFFC0FF
    Command16.BackColor = &H80C0FF
    
    num2_str = ""
    num2 = 0
    
End Sub

Private Sub Command14_Click()
        
    receptor = True
    num1 = 0
    num2 = 0
    
    num1_str = ""
    num2_str = ""
    
    Label1.Caption = ""
    
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    Command16.Enabled = True
    
    Command11.BackColor = &H80C0FF
    Command12.BackColor = &H80C0FF
    Command13.BackColor = &H80C0FF
    Command16.BackColor = &H80C0FF
    
    estaVacio = True

End Sub

Private Sub Command15_Click()
        
    If estaVacio = True Then
        
        Label1.Caption = "Math Error"
        receptor = True
        num1 = 0
        num2 = 0
    
        num1_str = ""
        num2_str = ""
        
        Command11.Enabled = True
        Command12.Enabled = True
        Command13.Enabled = True
        Command16.Enabled = True
        
        Command11.BackColor = &H80C0FF
        Command12.BackColor = &H80C0FF
        Command13.BackColor = &H80C0FF
        Command16.BackColor = &H80C0FF
    
        estaVacio = True
        
    ElseIf opSumaT = True Then
        
        resultado = num1 + num2
        Label1.Caption = resultado
        
        num2_str = ""
        num2 = 0
    
        num1_str = resultado
        num1 = resultado
        
        op = num1_str
        
    ElseIf opRestaT = True Then
        
        resultado = num1 - num2
        Label1.Caption = resultado
        
        num2_str = ""
        num2 = 0
        
        num1_str = resultado
        num1 = resultado
        
        op = num1_str
    
    ElseIf opMultiT = True Then
        
        resultado = num1 * num2
        Label1.Caption = resultado
        
        num2_str = ""
        num2 = 0
        
        num1_str = resultado
        num1 = resultado
        
        op = num1_str
    
    ElseIf opDivT = True Then
        If Not num2 = 0 Then
            
            resultado = num1 / num2
            Label1.Caption = resultado
            
            num2_str = ""
            num2 = 0
            
            num1_str = resultado
            num1 = resultado
            
            op = num1_str
            
        Else
            
            Label1.Caption = "Math Error"
            
            num2_str = ""
            num2 = 0
            
            num1_str = ""
            num1 = 0
            
            receptor = True
        End If
            
    End If
    
    opSumaT = False
    opRestaT = False
    opDivT = False
    opMultiT = False
    
    estaVacio = True
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    Command16.Enabled = True
    
    Command11.BackColor = &H80C0FF
    Command12.BackColor = &H80C0FF
    Command13.BackColor = &H80C0FF
    Command16.BackColor = &H80C0FF
    
End Sub

Private Sub Command16_Click()
    
    op = num1_str + "*"
    
    opSumaT = False
    opRestaT = False
    opDivT = False
    opMultiT = True
    
    receptor = False
    estaVacio = True
    
    Label1.Caption = op
    
    Command11.Enabled = True
    Command12.Enabled = True
    Command13.Enabled = True
    Command16.Enabled = False
    
    Command11.BackColor = &H80C0FF
    Command12.BackColor = &H80C0FF
    Command13.BackColor = &H80C0FF
    Command16.BackColor = &HFFC0FF
    
    num2_str = ""
    num2 = 0
    
End Sub

Private Sub Command17_Click()

    End
    
End Sub

Private Sub Command1_Click()
    
    If receptor = True Then
        num1_str = num1_str + "1"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "1"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command2_Click()
    
    If receptor = True Then
        num1_str = num1_str + "2"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "2"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command3_Click()
    
    If receptor = True Then
        num1_str = num1_str + "3"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "3"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command4_Click()
    
    If receptor = True Then
        num1_str = num1_str + "4"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "4"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command5_Click()

    If receptor = True Then
        num1_str = num1_str + "5"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "5"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command6_Click()
    
    If receptor = True Then
        num1_str = num1_str + "6"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "6"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
        
End Sub

Private Sub Command7_Click()

    If receptor = True Then
        num1_str = num1_str + "7"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "7"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command8_Click()
    
    If receptor = True Then
        num1_str = num1_str + "8"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "8"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Command9_Click()
    
    If receptor = True Then
        num1_str = num1_str + "9"
        num1 = CLng(num1_str)
        
        Label1.Caption = num1_str
        
        estaVacio = False
        
    ElseIf receptor = False Then
        num2_str = num2_str + "9"
        num2 = CLng(num2_str)
        
        Label1.Caption = op + num2_str
        
        estaVacio = False
        
    End If
    
End Sub

Private Sub Form_Activate()
    
    receptor = True
    
    opSumaT = False
    opRestaT = False
    opDivT = False
    opMultiT = False
    
    estaVacio = True
    
    num1 = 0
    num2 = 0
    
End Sub


