VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "E:\GitHub_@Ariellibra\Paz base de datos\bd1.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "Label1"
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FF80&
      Caption         =   "Label1"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "Label1"
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    

    If Data1.Recordset.EOF = True Then
    
        Data1.Recordset.MoveFirst
        mostrar
    
    Else
    
        mostrar
        Data1.Recordset.MoveNext
    
    End If

    'Print Data1.Recordset.RecordCount
    
    'Print Data1.Recordset.EOF
    
End Sub
Private Sub mostrar()
    
    Label1.Caption = Data1.Recordset.Fields("nombre").Value
    Label2.Caption = Data1.Recordset.Fields("apellido").Value
    Label3.Caption = Data1.Recordset.Fields("edad").Value
    
End Sub


Private Sub Command2_Click()
    
    If Data1.Recordset.BOF = True Then
    
        Data1.Recordset.MoveLast
        mostrar
    
    Else
    
        mostrar
        Data1.Recordset.MovePrevious
    
    End If

End Sub

Private Sub Form_Load()

    Data1.DatabaseName = App.Path & "\bd1.mdb"
    Data1.RecordSource = "alumnos"
    Data1.Refresh
    
End Sub


