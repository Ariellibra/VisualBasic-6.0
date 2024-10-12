VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   14550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   ScaleHeight     =   14550
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Conbinar y Guardar Datos"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7440
      Width           =   2655
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   240
      TabIndex        =   4
      Top             =   8520
      Width           =   14775
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   8880
      TabIndex        =   3
      Top             =   1320
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cargar Datos Actualizados"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Cargar Datos Originales"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   8295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim n, i, m, k, comas As Integer
Dim dato1a, dato2a, dato3a, dato4a, dato5a As String
Dim dato1b, dato2b, dato3b, dato4b, dato5b As String
Dim linea As String
Dim encontrado As Boolean

Private Sub Command1_Click()

    Open App.Path & "\Ariel Paz - Archivo de alumnos  1.txt" For Input As #1
    
    Do Until EOF(1)
        Input #1, dato1a
        List1.AddItem (dato1a)
    Loop
    
    Close #1
    
    For n = 0 To List1.ListCount - 1
    
        List1.List(n) = Replace(List1.List(n), ";", ",")
        
    Next n
    
    Open App.Path & "\Ariel Paz - Archivo de alumnos  1.txt" For Output As #1
    
    For n = 0 To List1.ListCount - 1
    
        Write #1, List1.List(n)
        
    Next n
    
    Close #1
    
End Sub

Private Sub Command2_Click()

    Open App.Path & "\Ariel Paz - archivo actualizador de alumnos.txt" For Input As #2
    
    Do Until EOF(2)
    
        Input #2, dato1a
        List2.AddItem (dato1a)
        
    Loop
    
    Close #2
    
    For n = 0 To List2.ListCount - 1
    
        List2.List(n) = Replace(List2.List(n), ";", ",")
        
    Next n
    
    Open App.Path & "\Ariel Paz - archivo actualizador de alumnos.txt" For Output As #2
    
    For n = 0 To List2.ListCount - 1
    
        Write #2, List2.List(n)
        
    Next n
    
    Close #2
    
End Sub

Private Sub Command3_Click()

    List3.Clear

    For n = 0 To List1.ListCount - 1
        linea = List1.List(n)
        comas = 0
        dato1a = ""
        dato2a = ""
        dato3a = ""
        dato4a = ""
        dato5a = ""
        
        For i = 1 To Len(linea)
            If Mid(linea, i, 1) = "," Then
                comas = comas + 1
                
                Select Case comas
                    Case 1
                        dato1a = Trim(dato1a)
                    Case 2
                        dato2a = Trim(dato2a)
                    Case 3
                        dato3a = Trim(dato3a)
                    Case 4
                        dato4a = Trim(dato4a)
                    Case 5
                        dato5a = Trim(dato5a)
                End Select
            Else
                Select Case comas
                    Case 0
                        dato1a = dato1a & Mid(linea, i, 1)
                    Case 1
                        dato2a = dato2a & Mid(linea, i, 1)
                    Case 2
                        dato3a = dato3a & Mid(linea, i, 1)
                    Case 3
                        dato4a = dato4a & Mid(linea, i, 1)
                    Case 4
                        dato5a = dato5a & Mid(linea, i, 1)
                End Select
            End If
        Next i
        
        Select Case comas
            Case 0
                dato1a = Trim(dato1a)
            Case 1
                dato2a = Trim(dato2a)
            Case 2
                dato3a = Trim(dato3a)
            Case 3
                dato4a = Trim(dato4a)
            Case 4
                dato5a = Trim(dato5a)
        End Select
        
        encontrado = False
        
        For m = 0 To List2.ListCount - 1
            linea = List2.List(m)
            comas = 0
            dato1b = ""
            dato2b = ""
            dato3b = ""
            dato4b = ""
            dato5b = ""
            
            For k = 1 To Len(linea)
            
                If Mid(linea, k, 1) = "," Then
                
                    comas = comas + 1
                    
                    Select Case comas
                        Case 1
                            dato1b = Trim(dato1b)
                        Case 2
                            dato2b = Trim(dato2b)
                        Case 3
                            dato3b = Trim(dato3b)
                        Case 4
                            dato4b = Trim(dato4b)
                        Case 5
                            dato5b = Trim(dato5b)
                    End Select
                    
                Else
                
                    Select Case comas
                        Case 0
                            dato1b = dato1b & Mid(linea, k, 1)
                        Case 1
                            dato2b = dato2b & Mid(linea, k, 1)
                        Case 2
                            dato3b = dato3b & Mid(linea, k, 1)
                        Case 3
                            dato4b = dato4b & Mid(linea, k, 1)
                        Case 4
                            dato5b = dato5b & Mid(linea, k, 1)
                    End Select
                End If
            Next k
            
            Select Case comas
                Case 0
                    dato1b = Trim(dato1b)
                Case 1
                    dato2b = Trim(dato2b)
                Case 2
                    dato3b = Trim(dato3b)
                Case 3
                    dato4b = Trim(dato4b)
                Case 4
                    dato5b = Trim(dato5b)
            End Select
            
            If dato1a = dato1b Then
            
                dato3a = dato2b
                dato4a = dato4b
                List3.AddItem dato1a & "," & dato2a & "," & dato2b & "," & dato4b & "," & dato5a & "," & dato3b & "," & dato5b
                encontrado = True
                
                Exit For
                
            End If
            
        Next m
        
        If Not encontrado Then
        
            List3.AddItem dato1a & "," & dato2a & "," & dato3a & "," & dato4a & "," & dato5a
            
        End If
    Next n
    
    Open App.Path & "\Ariel Paz - Archivo de alumnos  1.txt" For Output As #1
    
    For n = 0 To List3.ListCount - 1
    
        Print #1, List3.List(n)
        
    Next n
    
    Close #1
    
End Sub


Private Sub Command4_Click()

    End

End Sub
