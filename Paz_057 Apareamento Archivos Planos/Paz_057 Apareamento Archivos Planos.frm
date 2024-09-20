VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23475
   LinkTopic       =   "Form1"
   ScaleHeight     =   12285
   ScaleWidth      =   23475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command1"
      Height          =   735
      Left            =   15960
      TabIndex        =   5
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox List3 
      Height          =   6690
      Left            =   15960
      TabIndex        =   4
      Top             =   1560
      Width           =   6975
   End
   Begin VB.ListBox List2 
      Height          =   6690
      Left            =   8640
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   735
      Left            =   8640
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   6975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n As Integer
Dim dato1a, dato2a, dato3a, dato4a, dato5a As String
Dim dato1b, dato2b, dato3b, dato4b, dato5b As String


Private Sub Command1_Click()

    
    Open App.Path & "\Ariel Paz - Archivo de alumnos  1.txt" For Input As #1
    
    Do Until EOF(1)
    
        dato1a = ""
        Input #1, dato1a
        
        List1.AddItem (dato1a)
    
    Loop
    
    Close #1

End Sub

Private Sub Command2_Click()

    Open App.Path & "\Ariel Paz - archivo actualizador de alumnos.txt" For Input As #2
    
    Do Until EOF(2)
    
        Input #2, dato1a
        
        List2.AddItem (dato1a)
    
    Loop
    
    Close #2
    
End Sub

Private Sub Command3_Click()

    Open App.Path & "\Ariel Paz - Archivo de alumnos  1.txt" For Output As #1
    
    For n = 0 To List2.ListCount - 1
        
        Print #1, List2.List(n)
        
    Next n
    
    Close #1
    
    Open App.Path & "\Ariel Paz - Archivo de alumnos  1.txt" For Input As #1
    
    Do Until EOF(1)
    
        Input #1, dato1a
        
        List3.AddItem (dato1a)
    
    Loop
    
    Close #1

    
    

End Sub

