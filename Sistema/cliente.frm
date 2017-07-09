VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form prgcliente 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Clientes"
   ClientHeight    =   9075
   ClientLeft      =   930
   ClientTop       =   405
   ClientWidth     =   10995
   LinkTopic       =   "Form2"
   ScaleHeight     =   9075
   ScaleWidth      =   10995
   Begin VB.Frame PantaContacto 
      Caption         =   "Contactos"
      Height          =   1815
      Left            =   6960
      TabIndex        =   72
      Top             =   6840
      Visible         =   0   'False
      Width           =   2535
      Begin VB.CommandButton CierraContactos 
         Caption         =   "Menu F10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4560
         MouseIcon       =   "cliente.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Salida"
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox EmailIII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   89
         Text            =   " "
         Top             =   3840
         Width           =   6135
      End
      Begin VB.TextBox TelefonoIII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   87
         Text            =   " "
         Top             =   3480
         Width           =   6135
      End
      Begin VB.TextBox NombreIII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   85
         Top             =   3120
         Width           =   6135
      End
      Begin VB.TextBox EmailII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   83
         Text            =   " "
         Top             =   2400
         Width           =   6135
      End
      Begin VB.TextBox TelefonoII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   81
         Text            =   " "
         Top             =   2040
         Width           =   6135
      End
      Begin VB.TextBox NombreII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   79
         Top             =   1680
         Width           =   6135
      End
      Begin VB.TextBox EmailI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   77
         Text            =   " "
         Top             =   960
         Width           =   6135
      End
      Begin VB.TextBox TelefonoI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   75
         Text            =   " "
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox NombreI 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   73
         Top             =   240
         Width           =   6135
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   9360
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9360
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9480
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label23 
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   90
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   88
         Top             =   3480
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   86
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   84
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label18 
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   82
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   80
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "E-Mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   74
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   10320
      TabIndex        =   102
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox LocalidadII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   98
      Text            =   " "
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox PostalII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7800
      MaxLength       =   15
      TabIndex        =   97
      Text            =   " "
      Top             =   5160
      Width           =   1935
   End
   Begin VB.ComboBox ProvinciaII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   96
      Text            =   " "
      Top             =   5160
      Width           =   3135
   End
   Begin VB.CommandButton Historial 
      Caption         =   "HISTORIAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   95
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton ClieLista 
      Caption         =   "LISTAS DE PRECIOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   94
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Bonifica 
      Caption         =   "BONIFICACIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7200
      TabIndex        =   93
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2055
      Left            =   840
      TabIndex        =   6
      Top             =   6720
      Visible         =   0   'False
      Width           =   5895
      Begin VB.CommandButton Acepta 
         Caption         =   "Confirma F11"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3720
         MouseIcon       =   "cliente.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela F12"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         MouseIcon       =   "cliente.frx":1298
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":15A2
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Hasta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   12
         Text            =   " "
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox Desde 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   11
         Text            =   " "
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Vercontactos 
      Caption         =   "CONTACTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   92
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox PorceIva 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   70
      Text            =   " "
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox DireccionII 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   66
      Top             =   4440
      Width           =   3135
   End
   Begin VB.TextBox Fantasia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   64
      Top             =   4080
      Width           =   3135
   End
   Begin VB.TextBox Condicion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   61
      Text            =   " "
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox NroLista 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   59
      Text            =   " "
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton DatosAdicinales 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8520
      MouseIcon       =   "cliente.frx":19E4
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":1CEE
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Pedidos de Clientes"
      Top             =   1560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TipoClie 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      MaxLength       =   6
      TabIndex        =   55
      Text            =   " "
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MouseIcon       =   "cliente.frx":25B8
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":28C2
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borra  F2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      MouseIcon       =   "cliente.frx":3104
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":340E
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Elimina el Registro"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpia F3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2040
      MouseIcon       =   "cliente.frx":3C50
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":3F5A
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta F4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      MouseIcon       =   "cliente.frx":479C
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":4AA6
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Consulta de Datos"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Lista 
      Caption         =   "Listado F9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7800
      MouseIcon       =   "cliente.frx":52E8
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":55F2
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Impresion "
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      MouseIcon       =   "cliente.frx":5E34
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":613E
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Salida"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primer F5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3960
      MouseIcon       =   "cliente.frx":6980
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":6C8A
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Primer Registro"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "Anterior F6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      MouseIcon       =   "cliente.frx":70CC
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":73D6
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Registro Anterior"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Siguiente 
      Caption         =   "Siguien. F7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5880
      MouseIcon       =   "cliente.frx":7818
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":7B22
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Registro Siguiente"
      Top             =   5640
      Width           =   855
   End
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo F8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      MouseIcon       =   "cliente.frx":7F64
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":826E
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Salida"
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox Expreso 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   40
      Text            =   " "
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   39
      Top             =   6840
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.ComboBox Provincia 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   38
      Text            =   " "
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox fax 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   36
      Text            =   " "
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox email 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      MaxLength       =   100
      TabIndex        =   35
      Text            =   " "
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox Observaciones 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   26
      Text            =   " "
      Top             =   3000
      Width           =   4695
   End
   Begin VB.Frame Frame3 
      Caption         =   "Condicion de Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   5640
      TabIndex        =   23
      Top             =   0
      Width           =   4095
      Begin VB.OptionButton Iva6 
         Caption         =   "Exterior"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   42
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton Iva5 
         Caption         =   "Monotributo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Iva4 
         Caption         =   "Exento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Iva3 
         Caption         =   "Cons. Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2280
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Iva2 
         Caption         =   "No Inscripto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton Iva1 
         Caption         =   "Inscripto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Cuit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      MaxLength       =   13
      TabIndex        =   22
      Text            =   " "
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox Telefono 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   21
      Text            =   " "
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox Postal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   15
      TabIndex        =   20
      Text            =   " "
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Localidad 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   19
      Text            =   " "
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox Direccion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   14
      Text            =   " "
      Top             =   840
      Width           =   3135
   End
   Begin VB.TextBox Cliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1575
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9480
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "cliente.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Clientes"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9240
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Razon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   480
      Width           =   3135
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1200
      TabIndex        =   24
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "cliente.frx":86B0
      Left            =   120
      List            =   "cliente.frx":86B7
      TabIndex        =   4
      Top             =   7200
      Visible         =   0   'False
      Width           =   9735
   End
   Begin MSMask.MaskEdBox FechaAlta 
      Height          =   285
      Left            =   6720
      TabIndex        =   69
      Top             =   3720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label26 
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   101
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label25 
      Caption         =   "Codigo Postal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5640
      TabIndex        =   100
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label24 
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   99
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "% Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5640
      TabIndex        =   71
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label11 
      Caption         =   "Fecha Alta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   68
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Direccion Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   67
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Denominacion Comercial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   65
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label DesCondicion 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   63
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label17 
      Caption         =   "Condicion de Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   62
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Nro Lista"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5640
      TabIndex        =   60
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label DesTipoClie 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7680
      TabIndex        =   57
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label12 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   5640
      TabIndex        =   56
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Expreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   41
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Provincia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label20 
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5640
      TabIndex        =   33
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Cuit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5640
      TabIndex        =   18
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Telefono"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo Postal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Poblaci 
      Caption         =   "Localidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Razon Social"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "prgcliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZProvincia As Integer
Private Auxi As String
Dim ResultadoCuit As String

Dim ZZAyudaCli(10000) As String
Dim ZZLugarCli As Integer

Sub Imprime_Descripcion()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoClie"
    ZSql = ZSql + " Where TipoClie.Codigo = " + "'" + TipoClie.Text + "'"
    spTipoClie = ZSql
    Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoClie.RecordCount > 0 Then
        DesTipoClie.Caption = rstTipoClie!Descripcion
        rstTipoClie.Close
            Else
        DesTipoClie.Caption = ""
    End If

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CondPago"
    ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Condicion.Text + "'"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstCondPago.RecordCount > 0 Then
        DesCondicion.Caption = rstCondPago!Nombre
        rstCondPago.Close
            Else
        DesCondicion.Caption = ""
    End If
    
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        Cliente.Text = Trim(rstCliente!Cliente)
        Razon.Text = Trim(rstCliente!Razon)
        Direccion.Text = Trim(rstCliente!Direccion)
        Localidad.Text = Trim(rstCliente!Localidad)
        Postal.Text = Trim(rstCliente!Postal)
        Telefono.Text = Trim(rstCliente!Telefono)
        Observaciones.Text = Trim(rstCliente!Observaciones)
        Cuit.Text = Trim(rstCliente!Cuit)
        email.Text = Trim(rstCliente!email)
        fax.Text = Trim(rstCliente!fax)
        PorceIva.Text = Str$(rstCliente!PorceIva)
        Iva1.Value = False
        Iva2.Value = False
        Iva3.Value = False
        Iva4.Value = False
        Iva5.Value = False
        Iva6.Value = False
        Provincia.ListIndex = Val(rstCliente!Provincia)
        Select Case Val(rstCliente!Iva)
            Case 1
                Iva1.Value = True
            Case 2
                Iva2.Value = True
            Case 3
                Iva3.Value = True
            Case 4
                Iva4.Value = True
            Case 5
                Iva5.Value = True
            Case 6
                Iva6.Value = True
            Case Else
        End Select
        Expreso.Text = Trim(rstCliente!Expreso)
        TipoClie.Text = rstCliente!TipoClie
        NroLista.Text = Str$(rstCliente!NroLista)
        Condicion.Text = rstCliente!Condicion
        Fantasia.Text = rstCliente!Fantasia
        DireccionII.Text = rstCliente!DireccionII
        FechaAlta.Text = rstCliente!FechaAlta
        LocalidadII.Text = IIf(IsNull(rstCliente!LocalidadII), "", rstCliente!LocalidadII)
        ZZProvinciaII = IIf(IsNull(rstCliente!ProvinciaII), "0", rstCliente!ProvinciaII)
        ProvinciaII.ListIndex = ZZProvinciaII
        PostalII.Text = IIf(IsNull(rstCliente!PostalII), "", rstCliente!PostalII)
        
        NombreI.Text = Trim(rstCliente!NombreI)
        TelefonoI.Text = Trim(rstCliente!TelefonoI)
        EmailI.Text = Trim(rstCliente!EmailI)
        
        NombreII.Text = Trim(rstCliente!NombreII)
        TelefonoII.Text = Trim(rstCliente!TelefonoII)
        EmailII.Text = Trim(rstCliente!EmailII)
        
        NombreIII.Text = Trim(rstCliente!NombreIII)
        TelefonoIII.Text = Trim(rstCliente!TelefonoIII)
        EmailIII.Text = Trim(rstCliente!EmailIII)
        
        
        rstCliente.Close
        Call Format_datos
        Call Imprime_Descripcion
    End If

End Sub

Private Sub Acepta_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Cliente SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Clientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Cliente.Cliente, Cliente.Razon, Cliente.Direccion, Cliente.Telefono, Cliente.Cuit, " _
            + "Auxiliar.Nombre " _
            + "From " _
            + DSQ + ".dbo.Cliente Cliente, " _
            + DSQ + ".dbo.Auxiliar Auxiliar " _
            + "Where " _
            + "Cliente.CodigoEmpresa = Auxiliar.Empresa AND " _
            + "Cliente.Cliente >= '" + Desde.Text + "' AND " _
            + "Cliente.Cliente <= '" + Hasta.Text + "'"
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Cliente.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.SelectionFormula = "{Cliente.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Cliente.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Bonifica_Click()
    ZZPasaCliente = Cliente
    ZZPasaProceso = 0
    PrgClienteBonifica.Show
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
    Cliente.SetFocus
End Sub

Private Sub CierraContactos_Click()
    PantaContacto.Visible = False
End Sub

Private Sub ClieLista_Click()
    ZZPasaCliente = Cliente
    PrgClienteLista.Show
End Sub

Private Sub cmdAdd_Click()

    Call Verifica_datos
    
    WIva = "0"
    If Iva1.Value = True Then
        WIva = "1"
    End If
    If Iva2.Value = True Then
        WIva = "2"
    End If
    If Iva3.Value = True Then
        WIva = "3"
    End If
    If Iva4.Value = True Then
        WIva = "4"
    End If
    If Iva5.Value = True Then
        WIva = "5"
    End If
    If Iva6.Value = True Then
        WIva = "6"
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        rstCliente.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Cliente SET "
        ZSql = ZSql + " Razon = " + "'" + Razon.Text + "',"
        ZSql = ZSql + " Direccion = " + "'" + Direccion.Text + "',"
        ZSql = ZSql + " Localidad = " + "'" + Localidad.Text + "',"
        ZSql = ZSql + " Postal = " + "'" + Postal.Text + "',"
        ZSql = ZSql + " Telefono = " + "'" + Telefono.Text + "',"
        ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
        ZSql = ZSql + " NombreI = " + "'" + NombreI.Text + "',"
        ZSql = ZSql + " TelefonoI = " + "'" + TelefonoI.Text + "',"
        ZSql = ZSql + " EmailI = " + "'" + EmailI.Text + "',"
        ZSql = ZSql + " NombreII = " + "'" + NombreII.Text + "',"
        ZSql = ZSql + " TelefonoII = " + "'" + TelefonoII.Text + "',"
        ZSql = ZSql + " EmailII = " + "'" + EmailII.Text + "',"
        ZSql = ZSql + " NombreIII = " + "'" + NombreIII.Text + "',"
        ZSql = ZSql + " TelefonoIII = " + "'" + TelefonoIII.Text + "',"
        ZSql = ZSql + " EmailIII = " + "'" + EmailIII.Text + "',"
        ZSql = ZSql + " Fantasia = " + "'" + Fantasia.Text + "',"
        ZSql = ZSql + " DireccionII = " + "'" + DireccionII.Text + "',"
        ZSql = ZSql + " LocalidadII = " + "'" + LocalidadII.Text + "',"
        ZSql = ZSql + " PostalII = " + "'" + PostalII.Text + "',"
        ZSql = ZSql + " Cuit = " + "'" + Cuit.Text + "',"
        ZSql = ZSql + " Email = " + "'" + email.Text + "',"
        ZSql = ZSql + " Fax = " + "'" + fax.Text + "',"
        ZSql = ZSql + " PorceIva = " + "'" + PorceIva.Text + "',"
        ZSql = ZSql + " Provincia = " + "'" + Mid$(Str$(Provincia.ListIndex), 2, 2) + "',"
        ZSql = ZSql + " ProvinciaII = " + "'" + Mid$(Str$(ProvinciaII.ListIndex), 2, 2) + "',"
        ZSql = ZSql + " Iva = " + "'" + WIva + "',"
        ZSql = ZSql + " Expreso = " + "'" + Expreso.Text + "',"
        ZSql = ZSql + " TipoClie = " + "'" + TipoClie.Text + "',"
        ZSql = ZSql + " NroLista = " + "'" + NroLista.Text + "',"
        ZSql = ZSql + " Condicion = " + "'" + Condicion.Text + "'"
        ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
                Else
            
        ZZFechaAlta = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Cliente ("
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Telefono ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "NombreI ,"
        ZSql = ZSql + "TelefonoI ,"
        ZSql = ZSql + "EmailI ,"
        ZSql = ZSql + "NombreII ,"
        ZSql = ZSql + "TelefonoII ,"
        ZSql = ZSql + "EmailII ,"
        ZSql = ZSql + "NombreIII ,"
        ZSql = ZSql + "TelefonoIII ,"
        ZSql = ZSql + "EmailIII ,"
        ZSql = ZSql + "Fantasia ,"
        ZSql = ZSql + "DireccionII ,"
        ZSql = ZSql + "LocalidadII ,"
        ZSql = ZSql + "PostalII ,"
        ZSql = ZSql + "FechaAlta ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "Email ,"
        ZSql = ZSql + "Fax ,"
        ZSql = ZSql + "PorceIva ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "ProvinciaII ,"
        ZSql = ZSql + "Iva ,"
        ZSql = ZSql + "Expreso ,"
        ZSql = ZSql + "TipoClie ,"
        ZSql = ZSql + "NroLista ,"
        ZSql = ZSql + "Condicion )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Cliente.Text + "',"
        ZSql = ZSql + "'" + Razon.Text + "',"
        ZSql = ZSql + "'" + Direccion.Text + "',"
        ZSql = ZSql + "'" + Localidad.Text + "',"
        ZSql = ZSql + "'" + Postal.Text + "',"
        ZSql = ZSql + "'" + Telefono.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + NombreI.Text + "',"
        ZSql = ZSql + "'" + TelefonoI.Text + "',"
        ZSql = ZSql + "'" + EmailI.Text + "',"
        ZSql = ZSql + "'" + NombreII.Text + "',"
        ZSql = ZSql + "'" + TelefonoII.Text + "',"
        ZSql = ZSql + "'" + EmailII.Text + "',"
        ZSql = ZSql + "'" + NombreIII.Text + "',"
        ZSql = ZSql + "'" + TelefonoIII.Text + "',"
        ZSql = ZSql + "'" + EmailIII.Text + "',"
        ZSql = ZSql + "'" + Fantasia.Text + "',"
        ZSql = ZSql + "'" + DireccionII.Text + "',"
        ZSql = ZSql + "'" + LocalidadII.Text + "',"
        ZSql = ZSql + "'" + PostalII.Text + "',"
        ZSql = ZSql + "'" + ZZFechaAlta + "',"
        ZSql = ZSql + "'" + Cuit.Text + "',"
        ZSql = ZSql + "'" + email.Text + "',"
        ZSql = ZSql + "'" + fax.Text + "',"
        ZSql = ZSql + "'" + PorceIva.Text + "',"
        ZSql = ZSql + "'" + Mid$(Str$(Provincia.ListIndex), 2, 2) + "',"
        ZSql = ZSql + "'" + Mid$(Str$(ProvinciaII.ListIndex), 2, 2) + "',"
        ZSql = ZSql + "'" + WIva + "',"
        ZSql = ZSql + "'" + Expreso.Text + "',"
        ZSql = ZSql + "'" + TipoClie.Text + "',"
        ZSql = ZSql + "'" + NroLista.Text + "',"
        ZSql = ZSql + "'" + Condicion.Text + "')"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    
    Rem Call CmdLimpiar_Click
    
    
    m$ = "Grabacion realizada"
    aaaaaa% = MsgBox(m$, 0, "Archivo de Clientes")
    
    
    
    Cliente.SetFocus
    
End Sub

Private Sub cmdDelete_Click()
    If Cliente.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            rstCliente.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE Cliente"
                ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
                spCliente = ZSql
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Cliente.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    email.Text = ""
    fax.Text = ""
    PorceIva.Text = ""
    Expreso.Text = ""
    TipoClie.Text = ""
    DesTipoClie.Caption = ""
    Condicion.Text = ""
    DesCondicion.Caption = ""
    NroLista.Text = ""
    Fantasia.Text = ""
    DireccionII.Text = ""
    Localidad.Text = ""
    PostalII.Text = ""
    FechaAlta.Text = "  /  /    "
    
    NombreI.Text = ""
    TelefonoI.Text = ""
    EmailI.Text = ""
    NombreII.Text = ""
    TelefonoII.Text = ""
    EmailII.Text = ""
    NombreIII.Text = ""
    TelefonoIII.Text = ""
    EmailIII.Text = ""
    
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    
    Provincia.ListIndex = 0
    ProvinciaII.ListIndex = 0

    Cliente.SetFocus
    
End Sub

Private Sub CmdClose_Click()
    prgcliente.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command1_Click()
        ZSql = ""
        ZSql = ZSql + "UPDATE Cliente SET "
        ZSql = ZSql + " ProvinciaII = Provincia"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
End Sub

Private Sub DatosAdicinales_Click()
    ZZPasaCliente = Cliente.Text
    PrgClienteAdi.Show
End Sub

Private Sub Historial_Click()
    ZZPasaCliente = Cliente
    PrgHistorialClienteAuto.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Razon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Direccion.SetFocus
    End If
    If KeyAscii = 27 Then
        Razon.Text = ""
    End If
End Sub

Private Sub Direccion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Localidad.SetFocus
    End If
    If KeyAscii = 27 Then
        Direccion.Text = ""
    End If
End Sub

Private Sub Localidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Provincia.SetFocus
    End If
    If KeyAscii = 27 Then
        Localidad.Text = ""
    End If
End Sub

Private Sub Provincia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NroLista.SetFocus
    End If
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NroLista_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Postal.SetFocus
    End If
    If KeyAscii = 27 Then
        NroLista.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Postal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuit.SetFocus
    End If
    If KeyAscii = 27 Then
        Postal.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuit.Text <> "" Then
            Rem Call verifica_cuit(Cuit.Text, ResultadoCuit)
            ResultadoCuit = "S"
            If ResultadoCuit = "S" Then
                Telefono.SetFocus
                    Else
                Cuit.SetFocus
            End If
                Else
            Telefono.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cuit.Text = ""
    End If
End Sub

Private Sub Telefono_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        email.SetFocus
    End If
    If KeyAscii = 27 Then
        Telefono.Text = ""
    End If
End Sub

Private Sub EMail_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fax.SetFocus
    End If
    If KeyAscii = 27 Then
        email.Text = ""
    End If
End Sub

Private Sub Fax_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorceIva.SetFocus
    End If
    If KeyAscii = 27 Then
        fax.Text = ""
    End If
End Sub

Private Sub PorceIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        PorceIva.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Expreso.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub EXPRESO_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoClie.SetFocus
    End If
    If KeyAscii = 27 Then
        Expreso.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TipoClie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoClie"
        ZSql = ZSql + " Where TipoClie.Codigo = " + "'" + TipoClie.Text + "'"
        spTipoClie = ZSql
        Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoClie.RecordCount > 0 Then
            DesTipoClie.Caption = rstTipoClie!Descripcion
            rstTipoClie.Close
            Condicion.SetFocus
                Else
            TipoClie.SetFocus
            DesTipoClie.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        TipoClie.Text = ""
        DesTipoClie.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Condicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CondPago"
        ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Condicion.Text + "'"
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstCondPago.RecordCount > 0 Then
            DesCondicion.Caption = rstCondPago!Nombre
            rstCondPago.Close
            Fantasia.SetFocus
                Else
            Condicion.SetFocus
            DesCondicion.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        Condicion.Text = ""
        DesCondicion.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fantasia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DireccionII.SetFocus
    End If
    If KeyAscii = 27 Then
        Fantasia.Text = ""
    End If
End Sub

Private Sub DireccionII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LocalidadII.SetFocus
    End If
    If KeyAscii = 27 Then
        DireccionII.Text = ""
    End If
End Sub

Private Sub LocalidadII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PostalII.SetFocus
    End If
    If KeyAscii = 27 Then
        LocalidadII.Text = ""
    End If
End Sub

Private Sub PostalII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ProvinciaII.SetFocus
    End If
    If KeyAscii = 27 Then
        PostalII.Text = ""
    End If
End Sub

Private Sub ProvinciaII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Razon.SetFocus
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            rstCliente.Close
            Call Imprime_Datos
                Else
            WCliente = Cliente.Text
            CmdLimpiar_Click
        Cliente.Text = WCliente
        End If
        Razon.SetFocus
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
    End If
End Sub


Private Sub NombreI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TelefonoI.SetFocus
    End If
    If KeyAscii = 27 Then
        NombreI.Text = ""
    End If
End Sub

Private Sub TelefonoI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EmailI.SetFocus
    End If
    If KeyAscii = 27 Then
        TelefonoI.Text = ""
    End If
End Sub

Private Sub EmailI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NombreII.SetFocus
    End If
    If KeyAscii = 27 Then
        EmailI.Text = ""
    End If
End Sub



Private Sub NombreII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TelefonoII.SetFocus
    End If
    If KeyAscii = 27 Then
        NombreII.Text = ""
    End If
End Sub

Private Sub TelefonoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EmailII.SetFocus
    End If
    If KeyAscii = 27 Then
        TelefonoII.Text = ""
    End If
End Sub

Private Sub EmailII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NombreIII.SetFocus
    End If
    If KeyAscii = 27 Then
        EmailII.Text = ""
    End If
End Sub




Private Sub NombreIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TelefonoIII.SetFocus
    End If
    If KeyAscii = 27 Then
        NombreIII.Text = ""
    End If
End Sub

Private Sub TelefonoIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EmailIII.SetFocus
    End If
    If KeyAscii = 27 Then
        TelefonoIII.Text = ""
    End If
End Sub

Private Sub EmailIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NombreI.SetFocus
    End If
    If KeyAscii = 27 Then
        EmailIII.Text = ""
    End If
End Sub






Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Private Sub Consulta_Click()

    Opcion.Clear
    Opcion.AddItem "Cliente"
    Opcion.AddItem "TipoClie"
    Opcion.AddItem "Condicion de Pago"
    Opcion.Visible = True
    
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Cliente + " " + !Fantasia
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoClie"
            ZSql = ZSql + " Order by TipoClie.Codigo"
            spTipoClie = ZSql
            Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoClie.RecordCount > 0 Then
                With rstTipoClie
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoClie.Close
            End If
        
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Order by CondPago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
            End If
        
        
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            TipoClie.Text = WIndice.List(Indice)
            Call TipoClie_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            Condicion.Text = WIndice.List(Indice)
            Call Condicion_KeyPress(13)
                    
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    email.Text = ""
    fax.Text = ""
    PorceIva.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    Expreso.Text = ""
    TipoClie.Text = ""
    DesTipoClie.Caption = ""
    NroLista.Text = ""
    Condicion.Text = ""
    DesCondicion.Caption = ""
    Fantasia.Text = ""
    DireccionII.Text = ""
    LocalidadII.Text = ""
    PostalII.Text = ""
    FechaAlta.Text = "  /  /    "
    
    NombreI.Text = ""
    TelefonoI.Text = ""
    EmailI.Text = ""
    NombreII.Text = ""
    TelefonoII.Text = ""
    EmailII.Text = ""
    NombreIII.Text = ""
    TelefonoIII.Text = ""
    EmailIII.Text = ""
    
    Provincia.Clear
    
    Provincia.AddItem "0 - Capital Federal"
    Provincia.AddItem "1 - Buenos Aires"
    Provincia.AddItem "2 - Catamarca"
    Provincia.AddItem "3 - Cordoba"
    Provincia.AddItem "4 - Corrientes"
    Provincia.AddItem "5 - Chaco"
    Provincia.AddItem "6 - Chubut"
    Provincia.AddItem "7 - Entre Rios"
    Provincia.AddItem "8 - Formosa"
    Provincia.AddItem "9 - Jujuy"
    Provincia.AddItem "10 - La Pampa"
    Provincia.AddItem "11 - La Rioja"
    Provincia.AddItem "12 - Mendoza"
    Provincia.AddItem "13 - Misiones"
    Provincia.AddItem "14 - Neuquen"
    Provincia.AddItem "15 - Rio Negro"
    Provincia.AddItem "16 - Salta"
    Provincia.AddItem "17 - San Juan"
    Provincia.AddItem "18 - San Luis"
    Provincia.AddItem "19 - Santa Cruz"
    Provincia.AddItem "20 - Santa Fe"
    Provincia.AddItem "21 - Santiago del Estero"
    Provincia.AddItem "22 - Tucuman"
    Provincia.AddItem "23 - Tierra del Fuego"
    Provincia.AddItem "24 - Exterior"
    Provincia.AddItem ""
    
    Provincia.ListIndex = 0
    
    
    ProvinciaII.Clear
    
    ProvinciaII.AddItem "Capital Federal"
    ProvinciaII.AddItem "Buenos Aires"
    ProvinciaII.AddItem "Catamarca"
    ProvinciaII.AddItem "Cordoba"
    ProvinciaII.AddItem "Corrientes"
    ProvinciaII.AddItem "Chaco"
    ProvinciaII.AddItem "Chubut"
    ProvinciaII.AddItem "Entre Rios"
    ProvinciaII.AddItem "Formosa"
    ProvinciaII.AddItem "Jujuy"
    ProvinciaII.AddItem "La Pampa"
    ProvinciaII.AddItem "La Rioja"
    ProvinciaII.AddItem "Mendoza"
    ProvinciaII.AddItem "Misiones"
    ProvinciaII.AddItem "Neuquen"
    ProvinciaII.AddItem "Rio Negro"
    ProvinciaII.AddItem "Salta"
    ProvinciaII.AddItem "San Juan"
    ProvinciaII.AddItem "San Luis"
    ProvinciaII.AddItem "Santa Cruz"
    ProvinciaII.AddItem "Santa Fe"
    ProvinciaII.AddItem "Santiago del Estero"
    ProvinciaII.AddItem "Tucuman"
    ProvinciaII.AddItem "Tierra del Fuego"
    ProvinciaII.AddItem "Exterior"
    ProvinciaII.AddItem ""
    
    ProvinciaII.ListIndex = 0
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    Pantalla.Clear
    WIndice.Clear
    
    If KeyAscii > 31 Then
        ZAyuda = Ayuda.Text + Chr$(KeyAscii)
            Else
        Select Case KeyAscii
            Case 27
                Ayuda.Text = ""
                ZAyuda = ""
            Case 8
                WEspacios = Len(Ayuda.Text)
                If WEspacios > 0 Then
                    WEspacios = WEspacios - 1
                    ZAyuda = Left$(Ayuda.Text, WEspacios)
                End If
            Case Else
                ZAyuda = Ayuda.Text
        End Select
    End If
    WEspacios = Len(ZAyuda)
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Erase ZZAyudaCli
            ZZLugarCli = 0
        
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.fantasia LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Fantasia
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            ZZLugarCli = ZZLugarCli + 1
                            ZZAyudaCli(ZZLugarCli) = rstCliente!Cliente
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.razon LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZZEntra = "S"
                            For Ciclo = 1 To ZZLugarCli
                                If UCase(ZZAyudaCli(Ciclo)) = UCase(rstCliente!Cliente) Then
                                    ZZEntra = "N"
                                    Exit For
                                End If
                            Next Ciclo
                            If ZZEntra = "S" Then
                                IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstCliente!Cliente
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoClie"
            ZSql = ZSql + " Where TipoClie.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by TipoClie.Codigo"
            spTipoClie = ZSql
            Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoClie.RecordCount > 0 Then
                With rstTipoClie
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoClie.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Condpago"
            ZSql = ZSql + " Where Condpago.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Condpago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
            End If
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Cliente_DblClick()

    Opcion.Clear
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Expreso"
    Opcion.AddItem "TipoClie"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub TipoClie_DblClick()

    Opcion.Clear
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Expreso"
    Opcion.AddItem "TipoClie"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Condicion_DblClick()

    Opcion.Clear
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Expreso"
    Opcion.AddItem "TipoClie"
    Opcion.AddItem "Condicion"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Razon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Direccion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Localidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Provincia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Postal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Telefono_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Email_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Expreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cuit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fax_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PorceIva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fantasia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DireccionII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub LocalidadII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PostalII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ProvinciaII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NombreI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TelefonoI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub EmailI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NombreII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TelefonoII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub EmailII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NombreIII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TelefonoIII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TipoClie_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NroLista_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Condicion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Panta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impresora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 113
            Call cmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call Primer_Click
        Case 117
            Call Anterior_Click
        Case 118
            Call Siguiente_Click
        Case 119
            Call Ultimo_Click
        Case 120
            Call Lista_Click
        Case 121
            Call CmdClose_Click
        Case 122
            Call Acepta_Click
        Case 123
            Call Cancela_Click
        Case Else
    End Select
End Sub


Private Sub Anterior_Click()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente < " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " Order by Cliente.Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveLast
            Cliente.Text = rstCliente!Cliente
        End With
        rstCliente.Close
        Call Imprime_Datos
        Cliente.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Clientes")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Cliente) as [ClienteMenor]"
    ZSql = ZSql + " FROM Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        rstCliente.MoveFirst
        ZUltimo = IIf(IsNull(rstCliente!ClienteMenor), "", rstCliente!ClienteMenor)
        Cliente.Text = ZUltimo
        rstCliente.Close
        Call Imprime_Datos
        Cliente.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Cliente) as [ClienteMayor]"
    ZSql = ZSql + " FROM Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        rstCliente.MoveLast
        ZUltimo = IIf(IsNull(rstCliente!ClienteMayor), "", rstCliente!ClienteMayor)
        Cliente.Text = ZUltimo
        rstCliente.Close
        Call Imprime_Datos
        Cliente.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente > " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " Order by Cliente.Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Cliente.Text = rstCliente!Cliente
        End With
        rstCliente.Close
        Call Imprime_Datos
        Cliente.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Clientes")
    End If
End Sub

Private Sub Vercontactos_Click()

    PantaContacto.Height = 5895
    PantaContacto.Left = 120
    PantaContacto.Top = 840
    PantaContacto.Width = 9375
    
    PantaContacto.Visible = True
    
    NombreI.SetFocus

End Sub
