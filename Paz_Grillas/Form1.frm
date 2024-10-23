VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16620
   LinkTopic       =   "Form1"
   ScaleHeight     =   10605
   ScaleWidth      =   16620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8880
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8760
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8493
      _Version        =   393216
      GridLines       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Grilla.AddItem vbTab & "lunes" & vbTab & "martes" & vbTab & "miercoles", 1
    Grilla.AddItem vbTab & "jueves" & vbTab & "viernes" & vbTab & "sabado", 2
    Grilla.Width = 20000
    Grilla.ColWidth(1) = 1500
    Grilla.ColWidth(2) = 2500
    Grilla.ColWidth(3) = 3500
    Grilla.ColAlignment(3) = 4
    
    Grilla.Col = 2
    Grilla.Row = 1
    Grilla.Text = "domingo"
    
    Grilla.TextMatrix(2, 3) = "programacion"
    
    
End Sub

Private Sub Command2_Click()
    
    Print Grilla.ColSel

End Sub

Private Sub Form_Activate()
    
    Grilla.Cols = 4
    Grilla.Rows = 4

End Sub

