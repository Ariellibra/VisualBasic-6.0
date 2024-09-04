VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   12555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   3600
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   9
      Left            =   8400
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   8
      Left            =   6360
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   7
      Left            =   4320
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   6
      Left            =   2280
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   4
      Left            =   8400
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   3
      Left            =   6360
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   2
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Label1"
      Height          =   855
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim nameAlu As String
    Dim n As Integer
    
    Open ("E:\GitHub_@Ariellibra\Paz_ArchivosPlanos\nombresAlumnos.txt") For Input As #1
    
    Do Until EOF(1)
    
        Input #1, nameAlu
        
        Label1(n).Caption = nameAlu
            
    Loop
    
    Close #1
    
End Sub

