VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Alumno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private apellido As String
Private nombre As String
Private fechaNac As String
Private telefono As Long
Private mail As String

Public Sub Alumno(ByVal ape As String, ByVal nom As String, ByVal feNac As String, ByVal telef As Long, ByVal ml As String)
    apellido = ape
    nombre = nom
    fechaNac = feNac
    telefono = telef
    mail = ml
End Sub

Public Function getApellido() As String
    getApellido = apellido
End Function

Public Sub setApellido(ape As String)
    apellido = ape
End Sub

Public Function getNombre() As String
    getNombre = nombre
End Function

Public Sub setNombre(nom As String)
    nombre = nom
End Sub

Public Function getFechaNac() As String
    getFechaNac = fechaNac
End Function

Public Sub setFechaNac(feNac As String)
    fechaNac = feNac
End Sub

Public Function getTelefono() As Long
    getTelefono = telefono
End Function

Public Sub setTelefono(telef As Long)
    telefono = telef
End Sub

Public Function getMail() As String
    getMail = mail
End Function

Public Sub setMail(ml As String)
    mail = ml
End Sub

