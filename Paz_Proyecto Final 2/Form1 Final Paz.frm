VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   22875
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   22875
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C000&
      Caption         =   "Facturacion"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   12015
      Left            =   1080
      TabIndex        =   39
      Top             =   1440
      Visible         =   0   'False
      Width           =   21855
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   2280
         TabIndex        =   62
         Top             =   4440
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   10
         Left            =   2280
         TabIndex        =   61
         Top             =   3480
         Width           =   2415
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00FF8080&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   3000
         Top             =   0
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   10560
         Width           =   1695
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   9720
         Width           =   1695
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   8880
         Width           =   1695
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H0080FF80&
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4320
         Width           =   1695
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FF8080&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   2280
         TabIndex        =   40
         Top             =   1680
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid GridFact 
         Height          =   7095
         Left            =   7080
         TabIndex        =   46
         Top             =   1800
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   12515
         _Version        =   393216
         Cols            =   5
         BackColorBkg    =   16744703
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   49
         Top             =   4320
         Width           =   4335
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   59
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label Label24 
         BackColor       =   &H00808080&
         Caption         =   "Datos Producto"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2535
         Left            =   360
         TabIndex        =   57
         Top             =   2760
         Width           =   6495
      End
      Begin VB.Label Label23 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   855
         Left            =   4920
         TabIndex        =   56
         Top             =   720
         Width           =   9375
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Fecha: "
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14520
         TabIndex        =   55
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFC0FF&
         Caption         =   "NumFactura: "
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   18360
         TabIndex        =   54
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   51
         Top             =   9600
         Width           =   4335
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFC0FF&
         Caption         =   "SubTotal"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   50
         Top             =   8760
         Width           =   4335
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cuit"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   48
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   600
         TabIndex        =   47
         Top             =   10440
         Width           =   4335
      End
      Begin VB.Label Label20 
         BackColor       =   &H00808080&
         Caption         =   "Importante: El programa usa el 'CODIGO' para las consultas en la base de datos, recuerde que el CODIGO' es un numero UNICO."
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2055
         Left            =   15000
         TabIndex        =   53
         Top             =   9240
         Width           =   6495
      End
      Begin VB.Label Label19 
         BackColor       =   &H00808080&
         Caption         =   "Datos Cliente"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1815
         Left            =   360
         TabIndex        =   52
         Top             =   720
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C000&
      Caption         =   "Facturacion"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   840
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF80FF&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   1080
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   19095
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   2280
         TabIndex        =   36
         Top             =   4800
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   2280
         TabIndex        =   29
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   2280
         TabIndex        =   28
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   5
         Left            =   2280
         TabIndex        =   27
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   2280
         TabIndex        =   26
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FF8080&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080FF80&
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4080
         Width           =   1695
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   5760
         Top             =   240
      End
      Begin MSFlexGridLib.MSFlexGrid GridProd 
         Height          =   7095
         Left            =   7080
         TabIndex        =   25
         Top             =   720
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12515
         _Version        =   393216
         Cols            =   5
         BackColorBkg    =   16744703
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Stock"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   37
         Top             =   4680
         Width           =   4335
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   34
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   33
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Costo"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   32
         Top             =   3000
         Width           =   4335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Venta"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   31
         Top             =   3840
         Width           =   4335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00808080&
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   4815
         Left            =   360
         TabIndex        =   35
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808080&
         Caption         =   "Importante: El programa usa el 'CODIGO' para las consultas en la base de datos, recuerde que el CODIGO' es un numero UNICO."
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2055
         Left            =   360
         TabIndex        =   30
         Top             =   5760
         Width           =   6495
      End
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   1080
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   19095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   5760
         Top             =   240
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H008080FF&
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080C0FF&
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   720
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid GridClientes 
         Height          =   6255
         Left            =   7080
         TabIndex        =   6
         Top             =   720
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   11033
         _Version        =   393216
         Cols            =   5
         BackColorBkg    =   33023
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Franklin Gothic Medium"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   2280
         TabIndex        =   5
         Top             =   3960
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   2280
         TabIndex        =   4
         Top             =   3120
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         Top             =   2280
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2280
         TabIndex        =   2
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808080&
         Caption         =   "Importante: El programa usa el 'CUIT' para las consultas en la base de datos, recuerde que el 'CUIT' es un numero UNICO."
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   2055
         Left            =   360
         TabIndex        =   17
         Top             =   4920
         Width           =   6495
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   16
         Top             =   3840
         Width           =   4335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Apellido"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   15
         Top             =   3000
         Width           =   4335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   14
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cuit"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   13
         Top             =   1320
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   3975
         Left            =   360
         TabIndex        =   12
         Top             =   720
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label25 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Cuit"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   58
      Top             =   4320
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dataCliente As New Conexion
Dim dataProd As New Conexion
Dim contFact, contFactRow As Integer

Private Sub Command1_Click()
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    Text1(0).SetFocus
End Sub

Private Sub Command10_Click()
    modificarProducto
End Sub

Private Sub Command11_Click()
    altaProducto
End Sub

Private Sub Command12_Click()
    buscarProducto
End Sub

Private Sub Command13_Click()
    Frame3.Visible = True
    Frame1.Visible = False
    Frame2.Visible = False
    Text1(9).SetFocus
End Sub

Private Sub Command14_Click()
    buscaClienteFactura
End Sub

Private Sub Command15_Click()
    sumaCantidad
End Sub

Private Sub Command19_Click()
    buscaProductoFactura
End Sub

Private Sub Command2_Click()
    buscarCliente
End Sub

Private Sub Command3_Click()
    altaCliente
End Sub

Private Sub Command4_Click()
    modificarCliente
End Sub

Private Sub Command5_Click()
    eliminarCliente
End Sub

Private Sub Command6_Click()
    Limpiar
End Sub

Private Sub Command7_Click()
    Frame2.Visible = True
    Frame1.Visible = False
    Frame3.Visible = False
    Text1(4).SetFocus
End Sub

Private Sub Command8_Click()
    Limpiar
End Sub

Private Sub Command9_Click()
    eliminarProducto
End Sub

Private Sub Form_Activate()
    
    dataCliente.Conectar "Clientes"
    dataProd.Conectar "Productos"
    cargaClientes
    cargaProd
    cargaFact
    contFact = 1
    contFactRow = 1
    Label22.Caption = "Fecha: " & Date

End Sub
Private Sub altaCliente()
    If Text1(0) = "" Or Text1(1) = "" Or Text1(2) = "" Or Text1(3) = "" Then
        MsgBox "Error, no se puede cargar datos vacios", vbCritical, "Error"
        Limpiar
        Exit Sub
    ElseIf Not IsNumeric(Text1(0)) Then
        MsgBox "Error, el 'CUIT', solamente acepta numeros", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    clienteSeek
    If dataCliente.registro.NoMatch Then
        dataCliente.altaCliente Text1(1), Text1(2), Text1(3), CLng(Text1(0))
        MsgBox "Cliente agregado exitosamente", vbInformation, "Éxito"
    Else
        MsgBox "El Cuit ya esta cargado en la base de datos", vbCritical, "Error"
    End If
    Limpiar
    cargaClientes
    
End Sub
Private Sub modificarCliente()
    If Text1(0) = "" Then
        MsgBox "Error, no se puede cargar datos vacios", vbCritical, "Error"
        Limpiar
        Exit Sub
    ElseIf Not IsNumeric(Text1(0)) Then
        MsgBox "Error, el 'CUIT', solamente acepta numeros", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    clienteSeek
    If dataCliente.registro.NoMatch Then
        MsgBox "El 'CUIT' no existe", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    dataCliente.registro.Edit
    If Text1(1) <> "" Then
        dataCliente.registro.Fields("nombre").Value = UCase(Text1(1))
    End If
    If Text1(2) <> "" Then
        dataCliente.registro.Fields("apellido").Value = UCase(Text1(2))
    End If
    If Text1(3) <> "" Then
        dataCliente.registro.Fields("direccion").Value = UCase(Text1(3))
    End If
    
    dataCliente.registro.Update
    MsgBox "Cliente modificado exitosamente", vbInformation, "Éxito"
    Limpiar
    cargaClientes
    
End Sub
Private Sub eliminarCliente()
    Dim seguro As Integer

    If Text1(0) = "" Then
        MsgBox "Error, el 'CUIT' no puede estar vacio", vbCritical, "Error"
        Limpiar
        Exit Sub
    ElseIf Not IsNumeric(Text1(0)) Then
        MsgBox "Error, el 'CUIT', solamente acepta numeros", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    clienteSeek
    If dataCliente.registro.NoMatch Then
        MsgBox "El 'CUIT' no existe", vbCritical, "Error"
        Limpiar
        Exit Sub
    Else
        seguro = MsgBox("Esta seguro que quiera eliminar el registro, esta accion es 'PERMANENTE' " & vbCrLf & _
        "CUIT: " & dataCliente.registro.Fields("cuit") & vbCrLf & _
        "Nombre: " & dataCliente.registro.Fields("nombre") & vbCrLf & _
        "Apellido: " & dataCliente.registro.Fields("apellido") & vbCrLf & _
        "Direccion: " & dataCliente.registro.Fields("direccion"), vbYesNo + vbExclamation, "Eliminar")
        
        If seguro = vbYes Then
            dataCliente.registro.Delete
            MsgBox "Registro Borrado con Exito", vbInformation, "Eliminado"
        End If
    End If
    
    Limpiar
    cargaClientes
    
End Sub
Private Sub buscarCliente()
    Dim n, i As Integer
    
    For i = 0 To 3
    If Text1(i) <> "" Then
        For n = 1 To GridClientes.Rows - 1
            If GridClientes.TextMatrix(n, i + 1) = UCase(Text1(i)) Then
                GridClientes.Col = i + 1
                GridClientes.Row = n
                GridClientes.SetFocus
                GridClientes.CellBackColor = &H80FF80
                Timer1.Enabled = True
                Exit For
            End If
        Next n
    End If
    Next i
    
End Sub
Private Sub clienteSeek()
    dataCliente.registro.Index = "indexClientes"
    dataCliente.registro.Seek "=", CLng(Text1(0))
End Sub
Private Sub cargaClientes()
    Dim n As Integer
    
    With GridClientes
    .Cols = 5
    .Rows = 1
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "Cuit"
    .TextMatrix(0, 2) = "Nombre"
    .TextMatrix(0, 3) = "Apellido"
    .TextMatrix(0, 4) = "Direccion"
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 2500
    .ColWidth(3) = 2500
    .ColWidth(4) = 3500
    End With
    
    dataCliente.registro.MoveFirst
    n = 1
    
    Do While Not dataCliente.registro.EOF
        With GridClientes
        .Rows = .Rows + 1
        .TextMatrix(n, 0) = n
        .TextMatrix(n, 1) = dataCliente.registro.Fields("cuit").Value
        .TextMatrix(n, 2) = dataCliente.registro.Fields("nombre").Value
        .TextMatrix(n, 3) = dataCliente.registro.Fields("apellido").Value
        .TextMatrix(n, 4) = dataCliente.registro.Fields("direccion").Value
       End With
       
       dataCliente.registro.MoveNext
       n = n + 1
    Loop
    
End Sub
Private Sub altaProducto()
    If Text1(4) = "" Or Text1(5) = "" Or Text1(6) = "" Or Text1(7) = "" Or Text1(8) = "" Then
        MsgBox "Error, no se puede cargar datos vacios", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    productoSeek
    
    If dataProd.registro.NoMatch Then
        dataProd.altaProducto Text1(4), Text1(5), Text1(6), Text1(7), Text1(8)
        MsgBox "Producto agregado exitosamente", vbInformation, "Éxito"
    Else
        MsgBox "El producto ya está cargado en la base de datos", vbCritical, "Error"
    End If
    Limpiar
    cargaProd
    
End Sub
Private Sub modificarProducto()
    If Text1(4) = "" Then
        MsgBox "Error, el 'Código' no puede estar vacío", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If

    productoSeek
    
    If dataProd.registro.NoMatch Then
        MsgBox "El 'Código' del producto no existe", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    dataProd.registro.Edit
    If Text1(5) <> "" Then
        dataProd.registro.Fields("codigoProducto").Value = UCase(Text1(5))
    End If
    If Text1(6) <> "" Then
        dataProd.registro.Fields("codigoProducto").Value = UCase(Text1(6))
    End If
    If Text1(7) <> "" Then
        dataProd.registro.Fields("codigoProducto").Value = UCase(Text1(4))
    End If
    If Text1(8) <> "" Then
        dataProd.registro.Fields("codigoProducto").Value = UCase(Text1(8))
    End If
    dataProd.registro.Update

    MsgBox "Producto modificado exitosamente", vbInformation, "Éxito"
    Limpiar
    cargaProd
    
End Sub
Private Sub eliminarProducto()
    Dim seguro As Integer

    If Text1(4) = "" Then
        MsgBox "Error, el 'Código' no puede estar vacío", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    productoSeek
    If dataProd.registro.NoMatch Then
        MsgBox "El 'Código' del producto no existe", vbCritical, "Error"
        Limpiar
        Exit Sub
    Else
        seguro = MsgBox("¿Está seguro que quiere eliminar el registro? Esta acción es 'PERMANENTE'" & vbCrLf & _
                        "Código: " & dataProd.registro.Fields("codigoProducto") & vbCrLf & _
                        "Nombre: " & dataProd.registro.Fields("nombreProducto") & vbCrLf & _
                        "Costo: " & dataProd.registro.Fields("costo") & vbCrLf & _
                        "Venta: " & dataProd.registro.Fields("venta") & vbCrLf & _
                        "Stock: " & dataProd.registro.Fields("stock"), vbYesNo + vbExclamation, "Eliminar")
        
        If seguro = vbYes Then
            dataProd.registro.Delete
            MsgBox "Registro borrado con éxito", vbInformation, "Eliminado"
        End If
    End If
    
    Limpiar
    cargaProd
    
End Sub
Private Sub buscarProducto()
    Dim n, i As Integer
    
    For i = 4 To 8
        If Text1(i) <> "" Then
            For n = 1 To GridProd.Rows - 1
                If GridProd.TextMatrix(n, i - 3) = UCase(Text1(i)) Then
                    GridProd.Col = i - 3
                    GridProd.Row = n
                    GridProd.SetFocus
                    GridProd.CellBackColor = &H80FF80
                    
                    Timer1.Enabled = True
                    Exit For
                End If
            Next n
        End If
    Next i
End Sub
Private Sub productoSeek()
    dataProd.registro.Index = "indexProducto"
    dataProd.registro.Seek "=", Text1(4)
End Sub
Private Sub cargaProd()
    Dim n As Integer
    
    With GridProd
    .Cols = 6
    .Rows = 1
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "Codigo"
    .TextMatrix(0, 2) = "Nombre"
    .TextMatrix(0, 3) = "Costo"
    .TextMatrix(0, 4) = "Venta"
    .TextMatrix(0, 5) = "Stock"
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 2500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    End With
    
    dataProd.registro.MoveFirst
    n = 1
    
    Do While Not dataProd.registro.EOF
        With GridProd
        .Rows = .Rows + 1
        .TextMatrix(n, 0) = n
        .TextMatrix(n, 1) = dataProd.registro.Fields("codigoProducto").Value
        .TextMatrix(n, 2) = dataProd.registro.Fields("nombreProducto").Value
        .TextMatrix(n, 3) = dataProd.registro.Fields("costo").Value
        .TextMatrix(n, 4) = dataProd.registro.Fields("venta").Value
        .TextMatrix(n, 5) = dataProd.registro.Fields("stock").Value
       End With
       
       dataProd.registro.MoveNext
       n = n + 1
    Loop
    
End Sub
Private Sub buscaClienteFactura()
    If Text1(9) = "" Then
        MsgBox "Error, no se puede cargar datos vacios", vbCritical, "Error"
        Limpiar
        Exit Sub
    ElseIf Not IsNumeric(Text1(9)) Then
        MsgBox "Error, el 'CUIT', solamente acepta numeros", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If
    
    clienteSeekFactura
    If dataCliente.registro.NoMatch Then
        MsgBox "El 'CUIT' no existe", vbCritical, "Error"
        Limpiar
        Exit Sub
    Else
        Label23.Caption = _
        "Nombre: " & dataCliente.registro.Fields("nombre").Value & vbCrLf & _
        "Apellido: " & dataCliente.registro.Fields("apellido").Value & vbCrLf & _
        "Direccion: " & dataCliente.registro.Fields("direccion").Value
        Command14.Enabled = False
        Text1(9).Enabled = False
    End If
    
End Sub
Private Sub buscaProductoFactura()
    If Text1(10) = "" Then
        MsgBox "Error, el 'Código' no puede estar vacío", vbCritical, "Error"
        Limpiar
        Exit Sub
    End If

    productoSeekFactura
    
    If dataProd.registro.NoMatch Then
        MsgBox "El 'Código' del producto no existe", vbCritical, "Error"
        Limpiar
        Exit Sub
    Else
        With GridFact
        .Rows = .Rows + 1
        .TextMatrix(contFactRow, 0) = contFactRow
        .TextMatrix(contFactRow, 1) = dataProd.registro.Fields("codigoProducto").Value
        .TextMatrix(contFactRow, 2) = dataProd.registro.Fields("nombreProducto").Value
        .TextMatrix(contFactRow, 3) = dataProd.registro.Fields("stock").Value
        .TextMatrix(contFactRow, 5) = dataProd.registro.Fields("venta").Value
        .TextMatrix(contFactRow, 7) = "X"
        End With
        Text1(10).Enabled = False
        Command19.Enabled = False
    End If
    
    'MsgBox "Producto modificado exitosamente", vbInformation, "Éxito"
    
End Sub
Private Sub sumaCantidad()
    
    If Text1(11) = "" Then
        MsgBox "Error, la cantidad no puede estar vacía", vbCritical, "Error"
        Limpiar
        Exit Sub
    
    ElseIf Not IsNumeric(Text1(11)) Then
        MsgBox "Error, la cantidad tiene que ser un numero", vbCritical, "Error"
        Limpiar
        Exit Sub
    Else
        With GridFact
        .TextMatrix(contFactRow, 4) = Text1(11)
        .TextMatrix(contFactRow, 6) = CLng(Text1(11)) * CLng(.TextMatrix(contFactRow, 5))
        End With
        contFactRow = contFactRow + 1
        Text1(10).Enabled = True
        Command19.Enabled = True
    End If
    
    
    
End Sub
Private Sub clienteSeekFactura()
    dataCliente.registro.Index = "indexClientes"
    dataCliente.registro.Seek "=", CLng(Text1(9))
End Sub
Private Sub productoSeekFactura()
    dataProd.registro.Index = "indexProducto"
    dataProd.registro.Seek "=", Text1(10)
End Sub
Private Sub cargaFact()

    With GridFact
    .Cols = 8
    .Rows = 1
    .TextMatrix(0, 0) = "ID"
    .TextMatrix(0, 1) = "Codigo"
    .TextMatrix(0, 2) = "Nombre"
    .TextMatrix(0, 3) = "Stock"
    .TextMatrix(0, 4) = "Cantidad"
    .TextMatrix(0, 5) = "Precio"
    .TextMatrix(0, 6) = "SubTotal"
    .TextMatrix(0, 7) = "Borrar"
    .ColWidth(0) = 500
    .ColWidth(1) = 1500
    .ColWidth(2) = 2500
    .ColWidth(3) = 1000
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    .ColWidth(6) = 2000
    .ColWidth(7) = 1500
    End With

End Sub
Private Sub Limpiar()
    Dim n As Integer
    For n = 0 To Text1.Count - 1
        Text1(n).Text = ""
    Next n
End Sub

Private Sub GridFact_Click()
    GridFact.RemoveItem (GridFact.RowSel)
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = &H80FF80
End Sub
Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = vbWhite
End Sub

Private Sub Timer1_Timer()
    cargaClientes
    cargaProd
    Me.SetFocus
    Timer1.Enabled = False
End Sub
