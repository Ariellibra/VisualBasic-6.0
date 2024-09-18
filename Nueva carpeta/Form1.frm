VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22185
   LinkTopic       =   "Form1"
   ScaleHeight     =   11355
   ScaleWidth      =   22185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   8280
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   4020
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   6.985
      ScaleMode       =   0  'User
      ScaleWidth      =   9.737
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   5580
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   5040
      TabIndex        =   3
      Top             =   5400
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   12840
      TabIndex        =   2
      Text            =   "Seleccionar una img"
      Top             =   600
      Width           =   4095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   615
      Left            =   6000
      TabIndex        =   1
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   3855
      Left            =   6120
      Top             =   240
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ban As Boolean

Private Sub Check1_Click()
    
    If Check1.Value = 1 Then
        Picture1.Visible = True
    Else
        Picture1.Visible = False
    End If
    
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Image1.Visible = True
    Else
        Image1.Visible = False
    End If
    
End Sub

Private Sub Combo1_Change()

'    If Combo1.ListIndex = 0 Then
'
'        Image1.Picture = LoadPicture("C:\Documents and Settings\All Users\Documentos\Mis imágenes\Imágenes de muestra\Colinas azules.jpg")
'
'    End If


End Sub

Private Sub Combo1_Click()

    If Combo1.ListIndex = 0 Then

        Image1.Picture = LoadPicture("C:\Documents and Settings\All Users\Documentos\Mis imágenes\Imágenes de muestra\Colinas azules.jpg")

    ElseIf Combo1.ListIndex = 1 Then

        Image1.Picture = LoadPicture("C:\Documents and Settings\All Users\Documentos\Mis imágenes\Imágenes de muestra\Invierno.jpg")

    ElseIf Combo1.ListIndex = 2 Then

        Image1.Picture = LoadPicture("C:\Documents and Settings\All Users\Documentos\Mis imágenes\Imágenes de muestra\Nenúfares.jpg")

    ElseIf Combo1.ListIndex = 3 Then

        Image1.Picture = LoadPicture("C:\Documents and Settings\All Users\Documentos\Mis imágenes\Imágenes de muestra\Puesta de sol.jpg")


    End If

End Sub

Private Sub Command2_Click()

    Print (Combo1.ListIndex)

End Sub

Private Sub Form_Activate()
    
    ban = False
    
    Combo1.AddItem ("Colinas azules")
    Combo1.AddItem ("Invierno")
    Combo1.AddItem ("Nenufares")
    Combo1.AddItem ("Puesta de sol")
    
End Sub

