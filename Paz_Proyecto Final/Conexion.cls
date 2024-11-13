VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private base As Database
Private registro As Recordset

Public Function Conectar(nombreTabla As String)
    
    Set base = OpenDatabase(App.Path & "/bd1.mdb")
    Set registro = base.OpenRecordset(nombreTabla, dbOpenTable)
    
End Function

Public Function Chequeo(usu As String, pass As String) As Boolean
    
    registro.Index = "indexEmpleado"
    registro.Seek "=", usu
    
    If registro.NoMatch = False Then
        
        If registro.Fields("contraseña") = pass Then
            
            Chequeo = True
        
        End If
        
    End If
    
End Function
