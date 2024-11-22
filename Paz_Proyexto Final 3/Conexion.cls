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
Public registro As Recordset
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
Public Function altaCliente(dato1 As String, dato2 As String, dato3 As String, dato4 As Long)
    
    registro.AddNew
    registro.Fields("nombre") = LCase(dato1)
    registro.Fields("apellido") = LCase(dato2)
    registro.Fields("direccion") = LCase(dato3)
    registro.Fields("cuit") = CLng(dato4)
    registro.Update
    
End Function
Public Function altaProducto(dato1 As String, dato2 As String, dato3 As String, dato4 As String, dato5 As String)
    
    registro.AddNew
    registro.Fields("codigoProducto") = LCase(dato1)
    registro.Fields("nombreProducto") = LCase(dato2)
    registro.Fields("costro") = LCase(dato3)
    registro.Fields("venta") = LCase(dato4)
    registro.Fields("stock") = LCase(dato5)
    registro.Update
    
End Function
