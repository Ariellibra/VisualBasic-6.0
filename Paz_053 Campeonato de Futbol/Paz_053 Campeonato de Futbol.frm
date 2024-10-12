VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22740
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   22740
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1080
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Label1"
      Height          =   3495
      Left            =   15840
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Label1"
      Height          =   6135
      Left            =   9600
      TabIndex        =   2
      Top             =   1320
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   3495
      Left            =   5160
      TabIndex        =   1
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   4095
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
Dim resultados() As String
Dim i, j As Integer

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

Private Sub SimularPartidos(grupo() As String, ByRef ganador As String)

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
                Label3.Caption = Label3.Caption & grupo(i) & " " & golesA & " - " & grupo(j) & " " & golesB & " ->>>> Ganador " & grupo(i) & vbCrLf
            ElseIf golesB > golesA Then
                Label3.Caption = Label3.Caption & grupo(i) & " " & golesA & " - " & grupo(j) & " " & golesB & " ->>>> Ganador " & grupo(j) & vbCrLf
            Else
                Label3.Caption = Label3.Caption & grupo(i) & " " & golesA & " - " & grupo(j) & " " & golesB & " ->>>> Empate" & vbCrLf
            End If
        Next j
    Next i
    
    Dim maxGoles As Integer
    
    maxGoles = -1
    
    For i = 0 To 3
        For indiceA = 0 To 7
            If equipos(indiceA) = grupo(i) Then
                Exit For
            End If
        Next indiceA
        
        If golesEquipos(indiceA) > maxGoles Then
            maxGoles = golesEquipos(indiceA)
            ganador = grupo(i)
        End If
    Next i
End Sub


Private Sub Command1_Click()
    ' Limpiar resultados previos
    Label3.Caption = ""
    
    ' Simular partidos en grupo A
    Dim ganadorA As String
    Call SimularPartidos(grupoA, ganadorA)
    
    ' Simular partidos en grupo B
    Dim ganadorB As String
    Call SimularPartidos(grupoB, ganadorB)
    
    ' Simular la final
    Dim golesFinalA, golesFinalB As Integer
    golesFinalA = Int(Rnd * 5)
    golesFinalB = Int(Rnd * 5)
    
    Dim indiceA, indiceB As Integer
    ' Encontrar índices de los ganadores
    For indiceA = 0 To 7
        If equipos(indiceA) = ganadorA Then Exit For
    Next indiceA
    For indiceB = 0 To 7
        If equipos(indiceB) = ganadorB Then Exit For
    Next indiceB
    
    golesEquipos(indiceA) = golesEquipos(indiceA) + golesFinalA
    golesEquipos(indiceB) = golesEquipos(indiceB) + golesFinalB
    Dim campeon As String
    campeon = IIf(golesFinalA > golesFinalB, ganadorA, ganadorB)
    Label3.Caption = Label3.Caption & vbCrLf & "Final:" & vbCrLf & ganadorA & " " & golesFinalA & " - " & ganadorB & " " & golesFinalB & " ->>>> Ganador " & campeon & vbCrLf
    
    ' Determinar equipo con más y menos goles
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
    
    ' Mostrar equipos con más y menos goles
    Label4.Caption = "Equipo con más goles: " & equipoMaxGoles & " (" & maxGoles & " goles)" & vbCrLf & "Equipo con menos goles: " & equipoMinGoles & " (" & minGoles & " goles)"
End Sub

