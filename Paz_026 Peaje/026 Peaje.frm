VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
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
      Height          =   855
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Omnibuses"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Camionetas"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Camiones"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Autos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim autoCont As Integer
Dim autoPasa As Integer

Dim camionCont As Integer
Dim camionPasa As Integer

Dim camionetaCont As Integer
Dim camionetaPasa As Integer

Dim bondiCont As Integer
Dim bondiPasa As Integer



Private Sub Command1_Click()

    autoCont = autoCont + 1
    autoPasa = autoPasa + 5
    
    Label1.Caption = _
    "Autos: " & autoCont & " Pasajeros: " & autoPasa & vbCrLf & _
    "Camiones: " & camionCont & " Pasajeros: " & camionPasa & vbCrLf & _
    "Camionetas: " & camionetaCont & " Pasajeros: " & camionetaPasa & vbCrLf & _
    "Omnibuses: " & bondiCont & " Pasajeros: " & bondiPasa

End Sub

Private Sub Command2_Click()
    
    camionCont = camionCont + 1
    camionPasa = camionPasa + 1
    
    Label1.Caption = _
    "Autos: " & autoCont & " Pasajeros: " & autoPasa & vbCrLf & _
    "Camiones: " & camionCont & " Pasajeros: " & camionPasa & vbCrLf & _
    "Camionetas: " & camionetaCont & " Pasajeros: " & camionetaPasa & vbCrLf & _
    "Omnibuses: " & bondiCont & " Pasajeros: " & bondiPasa
    
End Sub

Private Sub Command3_Click()
    
    camionetaCont = camionetaCont + 1
    camionetaPasa = camionetaPasa + 3
    
    Label1.Caption = _
    "Autos: " & autoCont & " Pasajeros: " & autoPasa & vbCrLf & _
    "Camiones: " & camionCont & " Pasajeros: " & camionPasa & vbCrLf & _
    "Camionetas: " & camionetaCont & " Pasajeros: " & camionetaPasa & vbCrLf & _
    "Omnibuses: " & bondiCont & " Pasajeros: " & bondiPasa

End Sub

Private Sub Command4_Click()

    bondiCont = bondiCont + 1
    bondiPasa = bondiPasa + 10
    
    Label1.Caption = _
    "Autos: " & autoCont & " Pasajeros: " & autoPasa & vbCrLf & _
    "Camiones: " & camionCont & " Pasajeros: " & camionPasa & vbCrLf & _
    "Camionetas: " & camionetaCont & " Pasajeros: " & camionetaPasa & vbCrLf & _
    "Omnibuses: " & bondiCont & " Pasajeros: " & bondiPasa

End Sub

Private Sub Command5_Click()
    
    End
    
End Sub

Private Sub Form_Activate()

    autoCont = 0
    autoPasa = 0

    camionCont = 0
    camionPasa = 0

    camionetaCont = 0
    amionetaPasa = 0

    bondiCont = 0
    bondiPasa = 0
    
    Label1.Caption = _
    "Autos: " & autoCont & " Pasajeros: " & autoPasa & vbCrLf & _
    "Camiones: " & camionCont & " Pasajeros: " & camionPasa & vbCrLf & _
    "Camionetas: " & camionetaCont & " Pasajeros: " & camionetaPasa & vbCrLf & _
    "Omnibuses: " & bondiCont & " Pasajeros: " & bondiPasa
    
End Sub

