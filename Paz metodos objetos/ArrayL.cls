VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private datos() As String

Public Sub toMake(datosC As String)
    
    datos() = datosC

End Sub

Public Sub SetDatos(datosC() As String)

    datos() = datosC()
    
End Sub


Public Function getDatos()

    getDatos = datos()
    
End Function

Public Function toSplit(var As String, char As String)
        
    datos() = Split(var, char)
    
End Function

Public Function toPrint()
    Dim a As Integer
    
    For a = LBound(datos()) To UBound(datos())
        
        Form1.Label1 = Form1.Label1 & datos(a) & vbCrLf
        
    Next a
    

End Function
