VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12300
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   23265
   LinkTopic       =   "Form1"
   ScaleHeight     =   12300
   ScaleWidth      =   23265
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   2400
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cargar Notas"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n, contListaNota As Integer
Dim nota(1 To 5) As String
Dim flag As Boolean

Private Sub Command1_Click()
    
    nota(0, 0) = ReDimPreserve(nota, contListaNota, 5)
    
    flag = True
    
    Do
        For n = 0 To 4
        
            nota(contListaNota, n) = InputBox("Ingrese la nota: " & n + 1, "Cargar Notas")
            If n = 0 Then
                List1.AddItem ("")
                List1.List(contListaNota) = List1.List(contListaNota) & "[ " & nota(contListaNota, n) & ", "
        
            ElseIf n = 4 Then
                List1.List(contListaNota) = List1.List(contListaNota) & nota(contListaNota, n) & " ]"
            Else
                List1.List(contListaNota) = List1.List(contListaNota) & nota(contListaNota, n) & ", "
            End If
        
        Next n
        flag = False
    Loop Until flag = False
    
    contListaNota = contListaNota + 1
    
    
    
End Sub

'redim preserve both dimensions for a multidimension array *ONLY
Public Function ReDimPreserve(aArrayToPreserve, nNewFirstUBound, nNewLastUBound)
    ReDimPreserve = False
    Dim nOldFirstUBound, nOldLastUBound, nFirst, nLast As Integer
    'check if its in array first
    If IsArray(aArrayToPreserve) Then
        'create new array
        ReDim aPreservedArray(nNewFirstUBound, nNewLastUBound)
        'get old lBound/uBound
        nOldFirstUBound = UBound(aArrayToPreserve, 1)
        nOldLastUBound = UBound(aArrayToPreserve, 2)
        'loop through first
        For nFirst = LBound(aArrayToPreserve, 1) To nNewFirstUBound
            For nLast = LBound(aArrayToPreserve, 2) To nNewLastUBound
                'if its in range, then append to new array the same way
                If nOldFirstUBound >= nFirst And nOldLastUBound >= nLast Then
                    aPreservedArray(nFirst, nLast) = aArrayToPreserve(nFirst, nLast)
                End If
            Next
        Next
        'return the array redimmed
        If IsArray(aPreservedArray) Then ReDimPreserve = aPreservedArray
    End If
End Function


Private Sub Form_Activate()
    
    contListaNota = 1
    
End Sub

