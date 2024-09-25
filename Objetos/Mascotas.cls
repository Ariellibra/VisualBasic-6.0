VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mascotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private name As String
Public edad As Integer
Public peso As Double
Public raza As String

Public Sub Constructor(n As String, e As Integer, p As Double, r As String)
    
    name = n
    edad = e
    peso = p
    raza = r

End Sub

Public Function GetName() As String

    GetName = name
        
End Function

Public Function SetName(x As String)

    name = x

End Function

Public Function esGrande() As Boolean
    Dim grande As Boolean
    
    If edad > 8 Then
        
        grande = True
    
    Else
        
        grande = False
        
    End If
    
    esGrande = grande
    
End Function
