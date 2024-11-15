VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22665
   LinkTopic       =   "Form1"
   ScaleHeight     =   11445
   ScaleWidth      =   22665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Frame1"
      Height          =   9375
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   21375
      Begin MSFlexGridLib.MSFlexGrid GrillaClientes 
         Height          =   2175
         Left            =   5400
         TabIndex        =   5
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   3836
         _Version        =   393216
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Index           =   2
         Left            =   360
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   720
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    
    GrillaClientes.Cols = 5
    
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
