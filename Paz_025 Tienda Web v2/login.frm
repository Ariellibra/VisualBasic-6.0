VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   3015
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1781.361
   ScaleMode       =   0  'User
   ScaleWidth      =   4464.688
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1980
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1980
      Width           =   1260
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Nombre de usuario:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFC0FF&
      Caption         =   "&Contraseña:"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establecer la variable global a false
    'para indicar un inicio de sesión fallido
    LoginSucceeded = False
    Me.Hide
    
    End
    
End Sub

Private Sub cmdOK_Click()
    'comprobar si la contraseña es correcta
    If txtUserName = "admin" Then
        If txtPassword = "libra" Then
        'colocar código aquí para pasar al sub
        'que llama si la contraseña es correcta
        'lo más fácil es establecer una variable global
            LoginSucceeded = True
            Me.Hide
            Form2.Show
        Else
            MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
        End If
    Else
        MsgBox "El usuario no es válido. Vuelva a intentarlo", , "Inicio de sesión"
        txtUserName.SetFocus
        SendKeys "{Home}+{End}"
    End If
End Sub

