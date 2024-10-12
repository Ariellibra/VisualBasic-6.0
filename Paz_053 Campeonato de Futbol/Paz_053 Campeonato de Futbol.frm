VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C000&
   Caption         =   "Form1"
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20790
   LinkTopic       =   "Form1"
   ScaleHeight     =   10350
   ScaleWidth      =   20790
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Empezar Campeonato"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   6600
      TabIndex        =   6
      Top             =   1440
      Width           =   8655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3375
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   480
      TabIndex        =   3
      Top             =   7680
      Width           =   5775
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   5775
      Left            =   6600
      TabIndex        =   2
      Top             =   2520
      Width           =   8655
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3240
      TabIndex        =   1
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF80FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim equipos(7) As String
Dim grupoA(3) As String
Dim grupoB(3) As String
Dim golesEquipos(7) As Integer
Dim puntosEquipos(7) As Integer
Dim resultados() As String
Dim i, j As Integer

Private Sub Command2_Click()

    End
    
End Sub

Private Sub Form_Load()

    equipos(0) = "Boca Juniors"
    equipos(1) = "River Plate"
    equipos(2) = "Racing Club"
    equipos(3) = "Independiente"
    equipos(4) = "San Lorenzo"
    equipos(5) = "Huracán"
    equipos(6) = "Vélez Sarsfield"
    equipos(7) = "Lanús"
    
    For i = 0 To 7
    
        golesEquipos(i) = 0
        puntosEquipos(i) = 0
        
    Next i
    
    Randomize
    
    DividirEnGrupos

    MostrarGrupos

End Sub

Private Sub DividirEnGrupos()

    Dim randIndex As Integer
    Dim temp As String

    For i = 0 To 7
    
        randIndex = Int(Rnd * 8)
        temp = equipos(i)
        equipos(i) = equipos(randIndex)
        equipos(randIndex) = temp
        
    Next i

    For i = 0 To 3
    
        grupoA(i) = equipos(i)
        grupoB(i) = equipos(i + 4)
        
    Next i
    
End Sub

Private Sub MostrarGrupos()

    Label1.Caption = "Grupo A:" & vbCrLf
    
    For i = 0 To 3
    
        Label1.Caption = Label1.Caption & grupoA(i) & vbCrLf
        
    Next i
    
    Label2.Caption = "Grupo B:" & vbCrLf
    
    For i = 0 To 3
    
        Label2.Caption = Label2.Caption & grupoB(i) & vbCrLf
        
    Next i
    
End Sub

Private Sub SimularPartidos(grupo() As String, ganador As String)

    Dim golesA, golesB As Integer
    Dim indiceA, indiceB As Integer

    For i = 0 To 3
        For j = i + 1 To 3
        
            golesA = Int(Rnd * 5)
            golesB = Int(Rnd * 5)
            
            For indiceA = 0 To 7
            
                If equipos(indiceA) = grupo(i) Then
                    
                    Exit For
                    
                End If
                
            Next indiceA
            
            For indiceB = 0 To 7
            
                If equipos(indiceB) = grupo(j) Then
                
                    Exit For
                    
                End If
                
            Next indiceB
            
            golesEquipos(indiceA) = golesEquipos(indiceA) + golesA
            golesEquipos(indiceB) = golesEquipos(indiceB) + golesB
            
            If golesA > golesB Then
        
                puntosEquipos(indiceA) = puntosEquipos(indiceA) + 3
                Label3.Caption = Label3.Caption & grupo(i) & " " & golesA & " - " & grupo(j) & " " & golesB & " ->>>> Ganador " & grupo(i) & vbCrLf
                
            ElseIf golesB > golesA Then
            
                puntosEquipos(indiceB) = puntosEquipos(indiceB) + 3
                Label3.Caption = Label3.Caption & grupo(i) & " " & golesA & " - " & grupo(j) & " " & golesB & " ->>>> Ganador " & grupo(j) & vbCrLf
                
            Else
            
                puntosEquipos(indiceA) = puntosEquipos(indiceA) + 1
                puntosEquipos(indiceB) = puntosEquipos(indiceB) + 1
                Label3.Caption = Label3.Caption & grupo(i) & " " & golesA & " - " & grupo(j) & " " & golesB & " ->>>> Empate" & vbCrLf
                
            End If
        Next j
    Next i
    
    Dim maxPuntos, maxGoles, diffGoles As Integer
    
    maxPuntos = -1
    maxGoles = -1
    
    For i = 0 To 3
    
        For indiceA = 0 To 7
        
            If equipos(indiceA) = grupo(i) Then Exit For
            
        Next indiceA
        
        If puntosEquipos(indiceA) > maxPuntos Or (puntosEquipos(indiceA) = maxPuntos And golesEquipos(indiceA) - golesEquipos(Abs(indiceA - 1)) > diffGoles) Then
            
            maxPuntos = puntosEquipos(indiceA)
            maxGoles = golesEquipos(indiceA)
            diffGoles = golesEquipos(indiceA) - golesEquipos(Abs(indiceA - 1))
            ganador = grupo(i)
            
        End If
        
    Next i
    
End Sub

Private Sub MostrarTablaPuntos()

    Dim tempEquipo As String
    Dim tempPuntos, tempGoles As Integer
    
    For i = 0 To 6
    
        For j = i + 1 To 7
        
            If puntosEquipos(i) < puntosEquipos(j) Or (puntosEquipos(i) = puntosEquipos(j) And golesEquipos(i) - golesEquipos(j) < 0) Then

                tempPuntos = puntosEquipos(i)
                puntosEquipos(i) = puntosEquipos(j)
                puntosEquipos(j) = tempPuntos
                
                tempGoles = golesEquipos(i)
                golesEquipos(i) = golesEquipos(j)
                golesEquipos(j) = tempGoles
                
                tempEquipo = equipos(i)
                equipos(i) = equipos(j)
                equipos(j) = tempEquipo
            End If
            
        Next j
        
    Next i
    
    Label5.Caption = "Tabla de Puntos:" & vbCrLf & _
                     "Equipo - Puntos - Goles" & vbCrLf
    
    For i = 0 To 7
    
        Label5.Caption = Label5.Caption & equipos(i) & " - " & puntosEquipos(i) & " - " & golesEquipos(i) & vbCrLf
        
    Next i
    
End Sub
Private Sub Command1_Click()

    Label3.Caption = ""

    Dim ganadorA As String
    Call SimularPartidos(grupoA, ganadorA)

    Dim ganadorB As String
    Call SimularPartidos(grupoB, ganadorB)

    Dim golesFinalA, golesFinalB As Integer
    
    golesFinalA = Int(Rnd * 5)
    golesFinalB = Int(Rnd * 5)

    Dim indiceA, indiceB As Integer

    For indiceA = 0 To 7
    
        If equipos(indiceA) = ganadorA Then
        
            Exit For
        
        End If
        
    Next indiceA
    
    For indiceB = 0 To 7
    
        If equipos(indiceB) = ganadorB Then
        
            Exit For
        
        End If
        
    Next indiceB

    golesEquipos(indiceA) = golesEquipos(indiceA) + golesFinalA
    golesEquipos(indiceB) = golesEquipos(indiceB) + golesFinalB

    Dim campeon As String
    
    If golesFinalA > golesFinalB Then
    
        campeon = ganadorA
        Label3.Caption = Label3.Caption & vbCrLf & "Final:" & vbCrLf & ganadorA & " " & golesFinalA & " - " & ganadorB & " " & golesFinalB & " ->>>> Ganador " & ganadorA & vbCrLf
        
    ElseIf golesFinalB > golesFinalA Then
    
        campeon = ganadorB
        Label3.Caption = Label3.Caption & vbCrLf & "Final:" & vbCrLf & ganadorA & " " & golesFinalA & " - " & ganadorB & " " & golesFinalB & " ->>>> Ganador " & ganadorB & vbCrLf
        
    Else

        campeon = ganadorA
        Label3.Caption = Label3.Caption & vbCrLf & "Final:" & vbCrLf & ganadorA & " " & golesFinalA & " - " & ganadorB & " " & golesFinalB & " ->>>> Ganador por penales " & ganadorA & vbCrLf
        
    End If

    Dim maxGoles, minGoles As Integer
    
    Dim equipoMaxGoles, equipoMinGoles As String
    
    maxGoles = -1
    minGoles = 999

    Dim i As Integer
    
    For i = 0 To 7
    
        If golesEquipos(i) > maxGoles Then
        
            maxGoles = golesEquipos(i)
            equipoMaxGoles = equipos(i)
            
        End If
        
        If golesEquipos(i) < minGoles Then
        
            minGoles = golesEquipos(i)
            equipoMinGoles = equipos(i)
            
        End If
    Next i

    Label4.Caption = "Equipo con más goles: " & equipoMaxGoles & " (" & maxGoles & " goles)" & vbCrLf & "Equipo con menos goles: " & equipoMinGoles & " (" & minGoles & " goles)"

    MostrarTablaPuntos
End Sub

