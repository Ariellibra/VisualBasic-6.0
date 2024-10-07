VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Barco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private hundido As Boolean
Private posci As Integer
Private dueño As String

Public Sub Constructor(ByVal dueñoC As String, ByVal posciC As Integer, ByVal hundidoC As Boolean)

    dueño = dueñoC
    posci = posciC
    hundido = hundidoC

End Sub

Public Function GetPosci() As Integer

    GetPosci = posci
    
End Function

Public Sub SetHundido(bool As Boolean)
    
    hundido = bool

End Sub

 
