VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18165
   LinkTopic       =   "Form1"
   ScaleHeight     =   10980
   ScaleWidth      =   18165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Normal"
      Height          =   615
      Left            =   11160
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Al Revez"
      Height          =   615
      Left            =   11160
      TabIndex        =   1
      Top             =   1920
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   5895
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10398
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim n, j As Integer
Dim i, m, k, comas As Integer
Dim dato1a, dato2a, dato3a, dato4a, dato5a As String
Dim linea As String

Private Sub Command1_Click()
    
    
    Grilla.Clear
    formateaDatos
    k = 20
    traerDatosAlrevez

End Sub

Private Sub traerDatosAlrevez()

    Open App.Path & "\Ariel Paz - archivo actualizador de alumnos.txt" For Input As #1
    
    Do Until EOF(1)
    
        Input #1, linea
        
        comas = 0
        dato1a = ""
        dato2a = ""
        dato3a = ""
        dato4a = ""
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
        
        Grilla.AddItem k & vbTab & dato1a & vbTab & dato2a & vbTab & dato3a & vbTab & dato4a & vbTab & dato5a, k
        
        k = k - 1
        
        If k = 0 Then
        
        Else
            Grilla.RemoveItem (k)
        End If
                                        
    Loop
    
    Close #1
    
End Sub

Private Sub Command2_Click()
    
    Grilla.Clear
    formateaDatos
    k = 1
    traerDatos
    
End Sub

Private Sub Form_Activate()
    
    k = 0
    
    formateaDatos
    
    
End Sub

Private Sub formateaDatos()
    
    Grilla.Cols = 6
    Grilla.Rows = 21
    
    Grilla.TextMatrix(0, 1) = "Nombre"
    Grilla.TextMatrix(0, 2) = "Fecha"
    Grilla.TextMatrix(0, 3) = "Nota"
    Grilla.TextMatrix(0, 4) = "Telefono"
    Grilla.TextMatrix(0, 5) = "Materia"
    
End Sub

Private Sub traerDatos()

    Open App.Path & "\Ariel Paz - archivo actualizador de alumnos.txt" For Input As #1
    
    Do Until EOF(1)
    
        Input #1, linea
        
        comas = 0
        dato1a = ""
        dato2a = ""
        dato3a = ""
        dato4a = ""
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
        
        Grilla.AddItem k & vbTab & dato1a & vbTab & dato2a & vbTab & dato3a & vbTab & dato4a & vbTab & dato5a, k
        
        k = k + 1
        
        If k = 0 Then
        
        Else
        
            Grilla.RemoveItem (k)
            
        End If
        
                                        
    Loop
    
    Close #1
    
End Sub

