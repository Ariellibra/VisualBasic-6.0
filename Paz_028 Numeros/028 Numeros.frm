VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cargar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim num1 As Integer

Dim cont1, cont2, cont3, cont4 As Integer

Dim cont1str, cont2str, cont3str, cont4str As String

Private Sub Command1_Click()

    num1 = CInt(Text1.Text)
    
'   If num1 >= 100 And num1 <= 900 Then
    If num1 >= 100 Then
        If num1 <= 900 Then
            
            If num1 <= 300 Then
                If num1 >= 100 Then
                    cont1 = cont1 + 1
                    cont1str = cont1str + Text1.Text + ", "
                End If
                
            ElseIf num1 <= 500 Then
                If num1 >= 301 Then
                    cont2 = cont2 + 1
                    cont2str = cont2str + Text1.Text + ", "
                End If
                
            ElseIf num1 <= 700 Then
                If num1 >= 501 Then
                    cont3 = cont3 + 1
                    cont3str = cont3str + Text1.Text + ", "
                End If
            
            ElseIf num1 <= 900 Then
                If num1 >= 701 Then
                    cont4 = cont4 + 1
                    cont4str = cont4str + Text1.Text + ", "
                End If
            
            End If
            
        End If
        
                            
                
        
'        If num1 <= 900 And num1 >= 701 Then
'            cont4 = cont4 + 1
'            cont4str = cont4str + Text1.Text + ", "
'
'        ElseIf num1 <= 700 And num1 >= 501 Then
'            cont3 = cont3 + 1
'            cont3str = cont3str + Text1.Text + ", "
'
'        ElseIf num1 <= 500 And num1 >= 301 Then
'            cont2 = cont2 + 1
'            cont2str = cont2str + Text1.Text + ", "
'
'        ElseIf num1 <= 300 And num1 >= 100 Then
'            cont1 = cont1 + 1
'            cont1str = cont1str + Text1.Text + ", "
'
        End If

        num1 = 0
        Text1.Text = ""

        Label1.Caption = _
        "Los numeros del 100 al 300 son: " & vbCrLf & "Cantidad: " & cont1 & vbCrLf & "Numeros: " & cont1str & vbCrLf & _
        "Los numeros del 301 al 500 son: " & vbCrLf & "Cantidad: " & cont2 & vbCrLf & "Numeros: " & cont2str & vbCrLf & _
        "Los numeros del 501 al 700 son: " & vbCrLf & "Cantidad: " & cont3 & vbCrLf & "Numeros: " & cont3str & vbCrLf & _
        "Los numeros del 701 al 900 son: " & vbCrLf & "Cantidad: " & cont4 & vbCrLf & "Numeros: " & cont4str

'    End If
    
End Sub

Private Sub Command2_Click()

    End
    
End Sub

Private Sub Form_Activate()

    num1 = 0

    cont1 = 0
    cont2 = 0
    cont3 = 0
    cont4 = 0

    cont1str = ""
    cont2str = ""
    cont3str = ""
    cont4str = ""
    
    Label1.Caption = _
        "Los numeros del 100 al 300 son: " & vbCrLf & "Cantidad: " & cont1 & vbCrLf & "Numeros: " & cont1str & vbCrLf & _
        "Los numeros del 301 al 500 son: " & vbCrLf & "Cantidad: " & cont2 & vbCrLf & "Numeros: " & cont2str & vbCrLf & _
        "Los numeros del 501 al 700 son: " & vbCrLf & "Cantidad: " & cont3 & vbCrLf & "Numeros: " & cont3str & vbCrLf & _
        "Los numeros del 701 al 900 son: " & vbCrLf & "Cantidad: " & cont4 & vbCrLf & "Numeros: " & cont4str

    
End Sub

