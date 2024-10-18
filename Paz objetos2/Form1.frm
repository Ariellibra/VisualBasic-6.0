VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   9630
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   615
      Index           =   4
      Left            =   4080
      TabIndex        =   6
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   615
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   615
      Index           =   2
      Left            =   4080
      TabIndex        =   4
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   615
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Label1"
      Height          =   615
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Alumnos() As New Alumno
Dim n, j As Integer
Dim i, m, k, comas As Integer
Dim dato1a, dato2a, dato3a, dato4a, dato5a As String
Dim linea As String
Dim encontrado As Boolean


Private Sub Command1_Click()

    traerDatos
    
End Sub

Private Sub Command2_Click()

    Label1(0).Caption = Alumnos(j).getApellido
    Label1(1).Caption = Alumnos(j).getNombre
    Label1(2).Caption = Alumnos(j).getFechaNac
    Label1(3).Caption = Alumnos(j).getTelefono
    Label1(4).Caption = Alumnos(j).getMail
    
    j = j + 1
    
    If j > UBound(Alumnos()) Then
        
        j = 0
        
    End If

End Sub

Private Sub Form_Activate()
    
    n = 0
    k = 0
        
End Sub

Private Sub traerDatos()

    Open App.Path & "\Ariel Paz - Archivo de alumnos  1.txt" For Input As #1
    
    Do Until EOF(1)
    
        ReDim Preserve Alumnos(k)
    
        Input #1, linea
        
        comas = 0
        dato1a = ""
        dato2a = ""
        dato3a = ""
        dato4a = 0
        dato5a = ""
        
        For i = 1 To Len(linea)
            
            If Mid(linea, i, 1) = ";" Then
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
              
        Next i
        
        Alumnos(k).Alumno dato1a, dato2a, dato3a, CLng(dato4a), dato5a
        
        k = k + 1
                                        
    Loop
    
    Close #1
    
End Sub
