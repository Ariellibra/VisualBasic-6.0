VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n As Integer
Dim galon As Double

Private Sub Form_Activate()
    
    Form3.Hide
    
    galon = 3.785
    
    For n = 1 To 20
        
        If n = 1 Then
            MsgBox n & " Galon = " & galon * n & " Litros", vbInformation, "Formulario 3"
        Else
            MsgBox n & " Galones = " & galon * n & " Litros", vbInformation, "Formulario 3"
        End If
    
    Next n
    
    End
    
End Sub

