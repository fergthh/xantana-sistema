VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form prgArticuloII 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Articulos"
   ClientHeight    =   8130
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   11790
   Visible         =   0   'False
   Begin VB.TextBox Embarque 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   110
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox DescripcionII 
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   109
      Top             =   720
      Width           =   5535
   End
   Begin VB.TextBox CodigoProveedor 
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
      Left            =   8880
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   106
      Text            =   " "
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   855
      Left            =   11280
      TabIndex        =   105
      Top             =   360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Descuento 
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
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   103
      Top             =   3600
      Width           =   1095
   End
   Begin VB.ComboBox ListaPrecio 
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
      Left            =   8760
      TabIndex        =   101
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ComboBox Iva 
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
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   100
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Frame MuestraVenta 
      Height          =   1335
      Left            =   4800
      TabIndex        =   82
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
      Begin VB.TextBox Venta6 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   93
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Venta5 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   91
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox Venta4 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   89
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox Venta3 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   87
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Venta2 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   85
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Venta1 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   83
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label37 
         Caption         =   "Venta 6"
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
         Left            =   480
         TabIndex        =   94
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label36 
         Caption         =   "Venta 5"
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
         Left            =   480
         TabIndex        =   92
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label35 
         Caption         =   "Venta 4"
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
         Left            =   480
         TabIndex        =   90
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label34 
         Caption         =   "Venta 3"
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
         Left            =   480
         TabIndex        =   88
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label33 
         Caption         =   "Venta 2"
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
         Left            =   480
         TabIndex        =   86
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "Venta 1"
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
         Left            =   480
         TabIndex        =   84
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox PosicionII 
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   97
      Text            =   " "
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Cif 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   9960
      MaxLength       =   10
      TabIndex        =   95
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Precio 
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
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   80
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Comision 
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   78
      Text            =   " "
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Posicion 
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
      Left            =   4320
      MaxLength       =   10
      TabIndex        =   77
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton ImpreVenta 
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
      Left            =   11160
      MouseIcon       =   "articuloII.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Pedidos de Clientes"
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox StockAnterior 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   7920
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   74
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Stock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   72
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox Costo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   63
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox CostoFuturo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   5520
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   62
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox CostoAnterior 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   59
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox Fob 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   9960
      MaxLength       =   10
      TabIndex        =   57
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox MargenFuturo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   8160
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   55
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox UnidadCaja 
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
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   53
      Text            =   " "
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox MinimoVenta 
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
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   51
      Text            =   " "
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Color 
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
      Left            =   8520
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   50
      Text            =   " "
      Top             =   0
      Width           =   2775
   End
   Begin VB.TextBox Margen 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   5520
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   48
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   45
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Salidas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   5520
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   43
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Despacho 
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
      Left            =   8760
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   42
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox CodigoBarra 
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
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   39
      Text            =   " "
      Top             =   3240
      Width           =   3015
   End
   Begin VB.TextBox Entradas 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   38
      Top             =   2400
      Width           =   1215
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
      MouseIcon       =   "articuloII.frx":0BD4
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borra  F2"
      Enabled         =   0   'False
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
      Left            =   960
      MouseIcon       =   "articuloII.frx":1720
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Elimina el Registro"
      Top             =   3960
      Width           =   735
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
      Left            =   1800
      MouseIcon       =   "articuloII.frx":226C
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":2576
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3960
      Width           =   735
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
      Left            =   2640
      MouseIcon       =   "articuloII.frx":2DB8
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":30C2
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Consulta de Datos"
      Top             =   3960
      Width           =   735
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
      Left            =   11160
      MouseIcon       =   "articuloII.frx":3904
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":3C0E
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Impresion "
      Top             =   360
      Visible         =   0   'False
      Width           =   735
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
      Left            =   6840
      MouseIcon       =   "articuloII.frx":4450
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":475A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Salida"
      Top             =   3960
      Width           =   735
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
      Left            =   3480
      MouseIcon       =   "articuloII.frx":4F9C
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":52A6
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Primer Registro"
      Top             =   3960
      Width           =   735
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
      Left            =   4320
      MouseIcon       =   "articuloII.frx":56E8
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":59F2
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Registro Anterior"
      Top             =   3960
      Width           =   735
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
      Left            =   5160
      MouseIcon       =   "articuloII.frx":5E34
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":613E
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Registro Siguiente"
      Top             =   3960
      Width           =   735
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
      Left            =   6000
      MouseIcon       =   "articuloII.frx":6580
      MousePointer    =   99  'Custom
      Picture         =   "articuloII.frx":688A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Salida"
      Top             =   3960
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   5
      Top             =   5160
      Visible         =   0   'False
      Width           =   5775
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
         Left            =   4680
         MouseIcon       =   "articuloII.frx":6CCC
         MousePointer    =   99  'Custom
         Picture         =   "articuloII.frx":6FD6
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   960
         Width           =   855
      End
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
         Left            =   3600
         MouseIcon       =   "articuloII.frx":7418
         MousePointer    =   99  'Custom
         Picture         =   "articuloII.frx":7722
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Desde1 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   20
         Text            =   " "
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Hasta1 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   19
         Text            =   " "
         Top             =   1560
         Width           =   975
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
         TabIndex        =   11
         Text            =   " "
         Top             =   840
         Width           =   1455
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
         TabIndex        =   10
         Text            =   " "
         Top             =   480
         Width           =   1455
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
         Left            =   1800
         TabIndex        =   9
         Top             =   2160
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
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Desde Familia"
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
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Hasta Familia"
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
         TabIndex        =   21
         Top             =   1560
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
         Left            =   480
         TabIndex        =   7
         Top             =   840
         Width           =   1215
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
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   240
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   360
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Familia 
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
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   17
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Minimo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Left            =   10560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   13
      Text            =   " "
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Codigo 
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1335
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11400
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Articulo.rpt"
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
      Left            =   10920
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
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
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   0
      Width           =   5535
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
      Height          =   2220
      ItemData        =   "articuloII.frx":7B64
      Left            =   240
      List            =   "articuloII.frx":7B6B
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   7335
   End
   Begin MSMask.MaskEdBox FechaCostoAnterior 
      Height          =   285
      Left            =   2640
      TabIndex        =   61
      Top             =   1560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox FechaCosto 
      Height          =   285
      Left            =   2640
      TabIndex        =   65
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox FechaCierre 
      Height          =   285
      Left            =   840
      TabIndex        =   66
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox FechaUltimaEntrada 
      Height          =   285
      Left            =   6600
      TabIndex        =   68
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin MSMask.MaskEdBox FechaUltimaSalida 
      Height          =   285
      Left            =   9480
      TabIndex        =   70
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      Enabled         =   0   'False
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
   Begin VB.Label Label22 
      Caption         =   "Embarque"
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
      Left            =   2760
      TabIndex        =   111
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Desc. Ingles"
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
      TabIndex        =   108
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Codigo Proveedor"
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
      Left            =   7200
      TabIndex        =   107
      Top             =   720
      Width           =   2175
   End
   Begin VB.Image Foto 
      Height          =   3975
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   11520
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11520
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label42 
      Caption         =   "Dto."
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
      Left            =   2760
      TabIndex        =   104
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label41 
      Caption         =   "Lista"
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
      Left            =   7920
      TabIndex        =   102
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label40 
      Caption         =   "Iva"
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
      Left            =   5520
      TabIndex        =   99
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label39 
      Caption         =   "Posicion Ped."
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
      TabIndex        =   98
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label38 
      Caption         =   "Cif"
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
      Left            =   9480
      TabIndex        =   96
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label31 
      Caption         =   "Precio"
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
      Left            =   6840
      TabIndex        =   81
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label30 
      Caption         =   "Comision"
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
      TabIndex        =   79
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "Stock Ant."
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
      Left            =   6840
      TabIndex        =   75
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label27 
      Caption         =   "Stock"
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
      TabIndex        =   73
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label26 
      Caption         =   "Ult. Egreso"
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
      Left            =   8160
      TabIndex        =   71
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label25 
      Caption         =   "Ult. Ing."
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
      Left            =   5400
      TabIndex        =   69
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "Cierre"
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
      TabIndex        =   67
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label21 
      Caption         =   "Costo "
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
      TabIndex        =   64
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "Costo Anterior"
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
      TabIndex        =   60
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Fob"
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
      Left            =   9000
      TabIndex        =   58
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Margen Futuro"
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
      Left            =   6840
      TabIndex        =   56
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label16 
      Caption         =   "Unid.x Caja"
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
      TabIndex        =   54
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Minimo Fact."
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
      Left            =   2760
      TabIndex        =   52
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "Margen"
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
      Left            =   4680
      TabIndex        =   49
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Proveedor"
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
      TabIndex        =   47
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label DesProveedor 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   2400
      TabIndex        =   46
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label13 
      Caption         =   "Salidas"
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
      Left            =   4680
      TabIndex        =   44
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Despacho"
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
      Left            =   7080
      TabIndex        =   41
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Codigo Barra"
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
      Left            =   2760
      TabIndex        =   40
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Entradas"
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
      Left            =   2400
      TabIndex        =   25
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label11 
      Caption         =   "Costo Futuro"
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
      Left            =   4320
      TabIndex        =   24
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label DesFamilia 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   7200
      TabIndex        =   18
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "Grupo"
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
      Left            =   5160
      TabIndex        =   16
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Stock Minimo"
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
      Left            =   9120
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo "
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
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "prgArticuloII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZPrecio  As Double
Dim ZZMargen As Double
Dim ZZFoto As Image

Dim ZZCodAnt As String

Dim WMovi(20000, 3) As String


Sub Imprime_Descripcion()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Familia"
    ZSql = ZSql + " Where Familia.Codigo= " + "'" + Familia.Text + "'"
    spFamilia = ZSql
    Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
    If rstFamilia.RecordCount > 0 Then
        DesFamilia.Caption = rstFamilia!Descripcion
        rstFamilia.Close
            Else
        DesFamilia.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = rstProveedor!Nombre
        rstProveedor.Close
            Else
        DesProveedor.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Despacho"
    ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Despacho.Text + "'"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        Rem DesDespacho.Caption = IIf(IsNull(rstDespacho!Descripcion), "", rstDespacho!Descripcion)
        rstDespacho.Close
            Else
        Rem DesDespacho.Caption = ""
    End If
    
    
    
    ZEmbarque = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM OrdenImportacion"
    ZSql = ZSql + " Where OrdenImportacion.Articulo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " and OrdenImportacion.Estado = 0"
    spOrdenImportacion = ZSql
    Set rstOrdenImportacion = db.OpenRecordset(spOrdenImportacion, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrdenImportacion.RecordCount > 0 Then
    
        With rstOrdenImportacion
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZEmbarque = ZEmbarque + rstOrdenImportacion!Cantidad
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstOrdenImportacion.Close
    End If
    
    Embarque.Text = Str$(ZEmbarque)
    Embarque.Text = Pusing("###,###", Embarque.Text)
    
    
    
    
    
    WPaso = "F:\Fotos\" + Codigo.Text + ".Jpg"
    MiArchivo = Dir(WPaso)
    
    Foto.Picture = LoadPicture()
    
    If Trim(MiArchivo) <> "" Then
        Foto.Picture = LoadPicture(WPaso)
    End If
    
End Sub

Sub Verifica_datos()
    If Val(Familia.Text) = 0 Then
        Familia.Text = "0"
    End If
    If Val(Proveedor.Text) = 0 Then
        Proveedor.Text = "0"
    End If
    If Val(Margen.Text) = 0 Then
        Margen.Text = "0"
    End If
    If Val(MargenFuturo.Text) = 0 Then
        MargenFuturo.Text = "0"
    End If
    If Val(Fob.Text) = 0 Then
        Fob.Text = "0"
    End If
    If Val(Cif.Text) = 0 Then
        Cif.Text = "0"
    End If
    If Val(CostoAnterior.Text) = 0 Then
        CostoAnterior.Text = "0"
    End If
    If Val(Costo.Text) = 0 Then
        Costo.Text = "0"
    End If
    If Val(CostoFuturo.Text) = 0 Then
        CostoFuturo.Text = "0"
    End If
    If Val(Minimo.Text) = 0 Then
        Minimo.Text = "0"
    End If
    If Val(Entradas.Text) = 0 Then
        Entradas.Text = "0"
    End If
    If Val(Salidas.Text) = 0 Then
        Salidas.Text = "0"
    End If
    If Val(Stock.Text) = 0 Then
        Stock.Text = "0"
    End If
    If Val(StockAnterior.Text) = 0 Then
        StockAnterior.Text = "0"
    End If
    If Val(Venta1.Text) = 0 Then
        Venta1.Text = "0"
    End If
    If Val(Venta2.Text) = 0 Then
        Venta2.Text = "0"
    End If
    If Val(Venta3.Text) = 0 Then
        Venta3.Text = "0"
    End If
    If Val(Venta4.Text) = 0 Then
        Venta4.Text = "0"
    End If
    If Val(Venta5.Text) = 0 Then
        Venta5.Text = "0"
    End If
    If Val(Venta6.Text) = 0 Then
        Venta6.Text = "0"
    End If
    If Val(Comision.Text) = 0 Then
        Comision.Text = "0"
    End If
    If Val(Precio.Text) = 0 Then
        Precio.Text = "0"
    End If
End Sub

Sub Format_datos()

    If Val(Descuento.Text) <> 0 Then
        Descuento.Text = Pusing("###,###.##", Descuento.Text)
    End If
    If Val(Margen.Text) <> 0 Then
        Margen.Text = Pusing("###,###.##", Margen.Text)
    End If
    If Val(MargenFuturo.Text) <> 0 Then
        MargenFuturo.Text = Pusing("###,###.##", MargenFuturo.Text)
    End If
    If Val(Fob.Text) <> 0 Then
        Fob.Text = Pusing("###,###.###", Fob.Text)
    End If
    If Val(Cif.Text) <> 0 Then
        Cif.Text = Pusing("###,###.###", Cif.Text)
    End If
    If Val(Costo.Text) <> 0 Then
        Costo.Text = Pusing("###,###.##", Costo.Text)
    End If
    If Val(CostoAnterior.Text) <> 0 Then
        CostoAnterior.Text = Pusing("###,###.##", CostoAnterior.Text)
    End If
    If Val(CostoFuturo.Text) <> 0 Then
        CostoFuturo.Text = Pusing("###,###.##", CostoFuturo.Text)
    End If
    If Val(Venta1.Text) <> 0 Then
        Venta1.Text = Pusing("###,###.##", Venta1.Text)
    End If
    If Val(Venta2.Text) <> 0 Then
        Venta2.Text = Pusing("###,###.##", Venta2.Text)
    End If
    If Val(Venta3.Text) <> 0 Then
        Venta3.Text = Pusing("###,###.##", Venta3.Text)
    End If
    If Val(Venta4.Text) <> 0 Then
        Venta4.Text = Pusing("###,###.##", Venta4.Text)
    End If
    If Val(Venta5.Text) <> 0 Then
        Venta5.Text = Pusing("###,###.##", Venta5.Text)
    End If
    If Val(Venta6.Text) <> 0 Then
        Venta6.Text = Pusing("###,###.##", Venta6.Text)
    End If
    If Val(Comision.Text) <> 0 Then
        Comision.Text = Pusing("###,###.##", Comision.Text)
    End If
    If Val(Precio.Text) <> 0 Then
        Precio.Text = Pusing("###,###.##", Precio.Text)
    End If
    
End Sub

Sub Imprime_Datos()

    ZZCodAnt = Codigo.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Codigo.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Codigo.Text = Trim(rstArticulo!Codigo)
        Descripcion.Text = Trim(rstArticulo!Descripcion)
        Color.Text = Trim(rstArticulo!Color)
        Familia.Text = Str$(rstArticulo!Grupo)
        Proveedor.Text = Str$(rstArticulo!Proveedor)
        MinimoVenta.Text = Trim(rstArticulo!MinimoVenta)
        Minimo.Text = Str$(rstArticulo!Minimo)
        UnidadCaja.Text = Trim(rstArticulo!UnidadCaja)
        Margen.Text = Str$(rstArticulo!Margen)
        MargenFuturo.Text = Str$(rstArticulo!MargenFuturo)
        Fob.Text = Str$(rstArticulo!Fob)
        Cif.Text = Str$(rstArticulo!Cif)
        Costo.Text = Str$(rstArticulo!Costo)
        CostoAnterior.Text = Str$(rstArticulo!CostoAnterior)
        FechaCosto.Text = rstArticulo!FechaCosto
        FechaCostoAnterior.Text = rstArticulo!FechaCostoAnterior
        CostoFuturo.Text = Str$(rstArticulo!CostoFuturo)
        FechaCierre.Text = rstArticulo!FechaCierre
        FechaUltimaSalida.Text = rstArticulo!FechaUltimaSalida
        FechaUltimaEntrada.Text = rstArticulo!FechaUltimaEntrada
        Stock.Text = Str$(rstArticulo!Stock)
        Entradas.Text = Str$(rstArticulo!Entradas)
        Salidas.Text = Str$(rstArticulo!Salidas)
        StockAnterior.Text = Str$(rstArticulo!StockAnterior)
        Venta1.Text = Str$(rstArticulo!Venta1)
        Venta2.Text = Str$(rstArticulo!Venta2)
        Venta3.Text = Str$(rstArticulo!Venta3)
        Venta4.Text = Str$(rstArticulo!Venta4)
        Venta5.Text = Str$(rstArticulo!Venta5)
        Venta6.Text = Str$(rstArticulo!Venta6)
        Posicion.Text = Str$(rstArticulo!Posicion)
        PosicionII.Text = Str$(rstArticulo!PosicionII)
        Despacho.Text = Str$(rstArticulo!Despacho)
        Comision.Text = Str$(rstArticulo!Comision)
        CodigoBarra.Text = Trim(rstArticulo!CodigoBarra)
        Precio.Text = Str$(rstArticulo!Precio)
        Descuento.Text = Str$(rstArticulo!Descuento)
        Iva.ListIndex = rstArticulo!Iva
        ListaPrecio.ListIndex = rstArticulo!ListaPrecio
        DescripcionII.Text = IIf(IsNull(rstArticulo!DescripcionII), "", rstArticulo!DescripcionII)
        DescripcionII.Text = Trim(DescripcionII.Text)
        CodigoProveedor.Text = IIf(IsNull(rstArticulo!CodigoProveedor), "", rstArticulo!CodigoProveedor)
        Call Calcula_Precio
        rstArticulo.Close
        Call Format_datos
        Call Imprime_Descripcion
    End If
    
End Sub

Private Sub Acepta_Click()
    
    If Val(Desde1.Text) = 0 Then
         Desde1.Text = "0"
    End If
    If Val(Hasta1.Text) = 0 Then
         Hasta1.Text = "0"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Articulo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Linea, " _
                + "Auxiliar.Nombre, " _
                + "Lineas.Nombre " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Lineas Lineas " _
                + "Where " _
                + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Linea = Lineas.Linea AND " _
                + "Articulo.Codigo >= '" + Desde.Text + "' AND " _
                + "Articulo.Codigo <= '" + Hasta.Text + "' AND " _
                + "Articulo.Linea >= " + Desde1.Text + " AND " _
                + "Articulo.Linea <= " + Hasta1.Text
    
    Uno = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Dos = " and {Articulo.Linea} in " + Desde1.Text + " to " + Hasta1.Text
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
    Codigo.SetFocus
End Sub

Private Sub cmdAdd_Click()

    If Codigo.Text <> "" Then

        Call Verifica_datos
        Call Calcula_Precio
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Codigo.Text + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            ZZCosto = rstArticulo!Costo
            If ZZCosto <> Val(Costo.Text) Then
                FechaCostoAnterior.Text = FechaCosto.Text
                CostoAnterior.Text = Str$(ZZCosto)
                FechaCosto.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            End If
            
            ZZOrdFechaCosto = Right$(FechaCosto.Text, 4) + Mid$(FechaCosto.Text, 4, 2) + Left$(FechaCosto.Text, 2)
        
            rstArticulo.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
            ZSql = ZSql + " Color = " + "'" + Color.Text + "',"
            ZSql = ZSql + " Grupo = " + "'" + Familia.Text + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " MinimoVenta = " + "'" + MinimoVenta.Text + "',"
            ZSql = ZSql + " UnidadCaja = " + "'" + UnidadCaja.Text + "',"
            ZSql = ZSql + " Margen = " + "'" + Margen.Text + "',"
            ZSql = ZSql + " MargenFuturo = " + "'" + MargenFuturo.Text + "',"
            ZSql = ZSql + " Fob = " + "'" + Fob.Text + "',"
            ZSql = ZSql + " Cif = " + "'" + Cif.Text + "',"
            ZSql = ZSql + " CostoAnterior = " + "'" + CostoAnterior.Text + "',"
            ZSql = ZSql + " FechaCostoAnterior = " + "'" + FechaCostoAnterior.Text + "',"
            ZSql = ZSql + " Costo = " + "'" + Costo.Text + "',"
            ZSql = ZSql + " FechaCosto = " + "'" + FechaCosto.Text + "',"
            ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
            ZSql = ZSql + " CostoFuturo = " + "'" + CostoFuturo.Text + "',"
            ZSql = ZSql + " FechaCierre = " + "'" + FechaCierre.Text + "',"
            ZSql = ZSql + " FechaUltimaEntrada = " + "'" + FechaUltimaEntrada.Text + "',"
            ZSql = ZSql + " FechaUltimaSalida = " + "'" + FechaUltimaSalida.Text + "',"
            ZSql = ZSql + " Minimo = " + "'" + Minimo.Text + "',"
            ZSql = ZSql + " Entradas = " + "'" + Entradas.Text + "',"
            ZSql = ZSql + " Salidas = " + "'" + Salidas.Text + "',"
            ZSql = ZSql + " Stock = " + "'" + Stock.Text + "',"
            ZSql = ZSql + " StockAnterior = " + "'" + StockAnterior.Text + "',"
            ZSql = ZSql + " Iva = " + "'" + Str$(Iva.ListIndex) + "',"
            ZSql = ZSql + " Venta1 = " + "'" + Venta1.Text + "',"
            ZSql = ZSql + " Venta2 = " + "'" + Venta2.Text + "',"
            ZSql = ZSql + " Venta3 = " + "'" + Venta3.Text + "',"
            ZSql = ZSql + " Venta4 = " + "'" + Venta4.Text + "',"
            ZSql = ZSql + " Venta5 = " + "'" + Venta5.Text + "',"
            ZSql = ZSql + " Venta6 = " + "'" + Venta6.Text + "',"
            ZSql = ZSql + " Posicion = " + "'" + Posicion.Text + "',"
            ZSql = ZSql + " PosicionII = " + "'" + PosicionII.Text + "',"
            ZSql = ZSql + " Comision = " + "'" + Comision.Text + "',"
            ZSql = ZSql + " Despacho = " + "'" + Despacho.Text + "',"
            ZSql = ZSql + " CodigoBarra = " + "'" + CodigoBarra.Text + "',"
            ZSql = ZSql + " CodigoProveedor = " + "'" + CodigoProveedor.Text + "',"
            ZSql = ZSql + " DescripcionII = " + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + " ListaPrecio = " + "'" + Str$(ListaPrecio.ListIndex) + "',"
            ZSql = ZSql + " Descuento = " + "'" + Descuento.Text + "',"
            ZSql = ZSql + " Precio = " + "'" + Precio.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            If Trim(CodigoBarra.Text) <> "" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.CodigoBarra = " + "'" + CodigoBarra.Text + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZCodigo = rstArticulo!Codigo
                    rstArticulo.Close
                    m$ = "Codigo de Barra Duplicado (" + ZZCodigo + ")"
                    a% = MsgBox(m$, 0, "Archivo de Articulos")
                    Exit Sub
                End If
            End If
            
            FechaCosto.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZZOrdFechaCosto = Right$(FechaCosto.Text, 4) + Mid$(FechaCosto.Text, 4, 2) + Left$(FechaCosto.Text, 2)
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Articulo ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Color ,"
            ZSql = ZSql + "Grupo ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "MinimoVenta ,"
            ZSql = ZSql + "UnidadCaja ,"
            ZSql = ZSql + "Margen ,"
            ZSql = ZSql + "MargenFuturo ,"
            ZSql = ZSql + "Fob ,"
            ZSql = ZSql + "Cif ,"
            ZSql = ZSql + "CostoAnterior ,"
            ZSql = ZSql + "FechaCostoAnterior ,"
            ZSql = ZSql + "Costo ,"
            ZSql = ZSql + "FechaCosto ,"
            ZSql = ZSql + "OrdFechaCosto ,"
            ZSql = ZSql + "CostoFuturo ,"
            ZSql = ZSql + "FechaCierre ,"
            ZSql = ZSql + "FechaUltimaEntrada ,"
            ZSql = ZSql + "FechaUltimaSalida ,"
            ZSql = ZSql + "Minimo ,"
            ZSql = ZSql + "Entradas ,"
            ZSql = ZSql + "Salidas ,"
            ZSql = ZSql + "Stock ,"
            ZSql = ZSql + "StockAnterior ,"
            ZSql = ZSql + "Iva ,"
            ZSql = ZSql + "Venta1 ,"
            ZSql = ZSql + "Venta2 ,"
            ZSql = ZSql + "Venta3 ,"
            ZSql = ZSql + "Venta4 ,"
            ZSql = ZSql + "Venta5 ,"
            ZSql = ZSql + "Venta6 ,"
            ZSql = ZSql + "Posicion ,"
            ZSql = ZSql + "PosicionII ,"
            ZSql = ZSql + "Comision ,"
            ZSql = ZSql + "Despacho ,"
            ZSql = ZSql + "CodigoBarra ,"
            ZSql = ZSql + "CodigoProveedor ,"
            ZSql = ZSql + "DescripcionII ,"
            ZSql = ZSql + "ListaPrecio ,"
            ZSql = ZSql + "Descuento ,"
            ZSql = ZSql + "Precio )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + Color.Text + "',"
            ZSql = ZSql + "'" + Familia.Text + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + MinimoVenta.Text + "',"
            ZSql = ZSql + "'" + UnidadCaja.Text + "',"
            ZSql = ZSql + "'" + Margen.Text + "',"
            ZSql = ZSql + "'" + MargenFuturo.Text + "',"
            ZSql = ZSql + "'" + Fob.Text + "',"
            ZSql = ZSql + "'" + Cif.Text + "',"
            ZSql = ZSql + "'" + CostoAnterior.Text + "',"
            ZSql = ZSql + "'" + FechaCostoAnterior.Text + "',"
            ZSql = ZSql + "'" + Costo.Text + "',"
            ZSql = ZSql + "'" + FechaCosto.Text + "',"
            ZSql = ZSql + "'" + ZZFechaCosto + "',"
            ZSql = ZSql + "'" + CostoFuturo.Text + "',"
            ZSql = ZSql + "'" + FechaCierre.Text + "',"
            ZSql = ZSql + "'" + FechaUltimaEntrada.Text + "',"
            ZSql = ZSql + "'" + FechaUltimaSalida.Text + "',"
            ZSql = ZSql + "'" + Minimo.Text + "',"
            ZSql = ZSql + "'" + Entradas.Text + "',"
            ZSql = ZSql + "'" + Salidas.Text + "',"
            ZSql = ZSql + "'" + Stock.Text + "',"
            ZSql = ZSql + "'" + StockAnterior.Text + "',"
            ZSql = ZSql + "'" + Str$(Iva.ListIndex) + "',"
            ZSql = ZSql + "'" + Venta1.Text + "',"
            ZSql = ZSql + "'" + Venta2.Text + "',"
            ZSql = ZSql + "'" + Venta3.Text + "',"
            ZSql = ZSql + "'" + Venta4.Text + "',"
            ZSql = ZSql + "'" + Venta5.Text + "',"
            ZSql = ZSql + "'" + Venta6.Text + "',"
            ZSql = ZSql + "'" + Posicion.Text + "',"
            ZSql = ZSql + "'" + PosicionII.Text + "',"
            ZSql = ZSql + "'" + Comision.Text + "',"
            ZSql = ZSql + "'" + Despacho.Text + "',"
            ZSql = ZSql + "'" + CodigoBarra.Text + "',"
            ZSql = ZSql + "'" + CodigoProveedor.Text + "',"
            ZSql = ZSql + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + "'" + Str$(ListaPrecio.ListIndex) + "',"
            ZSql = ZSql + "'" + Descuento.Text + "',"
            ZSql = ZSql + "'" + Precio.Text + "')"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
End Sub

Private Sub CmdDelete_Click()

    If Codigo.Text <> "" Then
        
        If Val(Stock.Text) = 0 Then
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Codigo.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                T$ = "Borrar Registro"
                m$ = "Desea Borrar el Registro "
                Respuesta% = MsgBox(m$, 32 + 4, T$)
                If Respuesta% = 6 Then
            
                    ZSql = ""
                    ZSql = ZSql + "DELETE Articulo"
                    ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Call CmdLimpiar_Click
                End If
            End If
            
                Else
                
            m$ = "No se puede dar de baja ya que posee stock"
            a% = MsgBox(m$, 0, "Archivo de Articulos")
            
        End If
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Descripcion.Text = ""
    Color.Text = ""
    Familia.Text = ""
    DesFamilia.Caption = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    MinimoVenta.Text = ""
    UnidadCaja.Text = ""
    Margen.Text = ""
    MargenFuturo.Text = ""
    Fob.Text = ""
    Cif.Text = ""
    CostoAnterior.Text = ""
    FechaCostoAnterior.Text = "  /  /    "
    Costo.Text = ""
    FechaCosto.Text = "  /  /    "
    CostoFuturo.Text = ""
    FechaCierre.Text = "  /  /    "
    FechaUltimaEntrada.Text = "  /  /    "
    FechaUltimaSalida.Text = "  /  /    "
    Minimo.Text = "1"
    Entradas.Text = ""
    Salidas.Text = ""
    Stock.Text = ""
    StockAnterior.Text = ""
    Venta1.Text = ""
    Venta2.Text = ""
    Venta3.Text = ""
    Venta4.Text = ""
    Venta5.Text = ""
    Venta6.Text = ""
    Posicion.Text = ""
    PosicionII.Text = ""
    Despacho.Text = ""
    Rem DesDespacho.Caption = ""
    Comision.Text = "10.00"
    CodigoBarra.Text = ""
    Precio.Text = ""
    Descuento.Text = ""
    CodigoProveedor.Text = ""
    DescripcionII.Text = ""
    Embarque.Text = ""
    
    Iva.ListIndex = 0
    ListaPrecio.ListIndex = 0
    
    Codigo.SetFocus
End Sub

Private Sub cmdClose_Click()
    prgArticuloII.Hide
    Unload Me
    Menu2.Show
End Sub

Private Sub Command1_Click()
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " ListaPrecio = " + "'" + "0" + "',"
    ZSql = ZSql + " Descuento = " + "'" + "0" + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
End Sub

Private Sub Command2_Click()



    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Entradas = " + "'" + "0" + "',"
    ZSql = ZSql + " Salidas = " + "'" + "0" + "',"
    ZSql = ZSql + " Stock = " + "'" + "0" + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)









    ZLugar = 0
    Erase WMovi



    ZPasa = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Order by Estadistica.Articulo"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstEstadistica!ordfecha >= "20090304" And rstEstadistica!ordfecha <= "20090331" Then
                        If ZPasa = 0 Then
                            ZPasa = 1
                            ZCorte = rstEstadistica!Articulo
                            ZCantidad = 0
                        End If
                        
                        If ZCorte <> rstEstadistica!Articulo Then
                        
                            ZLugar = ZLugar + 1
                            WMovi(ZLugar, 1) = ZCorte
                            WMovi(ZLugar, 2) = ""
                            WMovi(ZLugar, 3) = Str$(ZCantidad)
                        
                            ZCorte = rstEstadistica!Articulo
                            ZCantidad = 0
                            
                        End If
                        
                        ZCantidad = ZCantidad + rstEstadistica!Cantidad
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    If ZPasa <> 0 Then
    
        ZLugar = ZLugar + 1
        WMovi(ZLugar, 1) = ZCorte
        WMovi(ZLugar, 2) = ""
        WMovi(ZLugar, 3) = Str$(ZCantidad)
        
    End If
    
    
    



    ZPasa = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM MOvstk"
    ZSql = ZSql + " Order by MOvstk.Articulo"
    spMovStk = ZSql
    Set rstMovstk = db.OpenRecordset(spMovStk, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovstk.RecordCount > 0 Then
        With rstMovstk
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstMovstk!ordfecha >= "20090304" And rstMovstk!ordfecha <= "20090331" Then
                
                        If ZPasa = 0 Then
                            ZPasa = 1
                            ZCorte = rstMovstk!Articulo
                            ZEntrada = 0
                            ZSalida = 0
                        End If
                        
                        If ZCorte <> rstMovstk!Articulo Then
                        
                            ZLugar = ZLugar + 1
                            WMovi(ZLugar, 1) = ZCorte
                            WMovi(ZLugar, 2) = Str$(ZEntrada)
                            WMovi(ZLugar, 3) = Str$(ZSalida)
                        
                            ZCorte = rstMovstk!Articulo
                            ZEntrada = 0
                            ZSalida = 0
                            
                        End If
                        
                        If rstMovstk!Cantidad > 0 Then
                            ZEntrada = ZEntrada + rstMovstk!Cantidad
                                Else
                            ZSalida = ZSalida + Abs(rstMovstk!Cantidad)
                        End If
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstMovstk.Close
    End If
    
    If ZPasa <> 0 Then
    
        ZLugar = ZLugar + 1
        WMovi(ZLugar, 1) = ZCorte
        WMovi(ZLugar, 2) = Str$(ZEntrada)
        WMovi(ZLugar, 3) = Str$(ZSalida)
        
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZZCodigo = WMovi(Ciclo, 1)
        ZZEntrada = WMovi(Ciclo, 2)
        ZZSalida = WMovi(Ciclo, 3)
        ZZCantidad = Str$(Val(ZZEntrada) - Val(ZZSalida))
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Entradas = Entradas + " + "'" + ZZEntrada + "',"
        ZSql = ZSql + " Salidas = Salidas + " + "'" + ZZSalida + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZZCodigo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    


    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Stock = StockAnterior + Entradas - Salidas"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

    
    
    
    
    Stop

End Sub

Private Sub Command9_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Prueba"
    ZSql = ZSql + " Where Prueba.Codigo = " + "'" + Codigo.Text + "'"
    spPrueba = ZSql
    Set rstPrueba = db.OpenRecordset(spPrueba, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueba.RecordCount > 0 Then
    
        rstPrueba.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Prueba SET "
        ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
        ZSql = ZSql + " Cantidad = " + "'" + "100" + "',"
        ZSql = ZSql + " Precio = " + "'" + "2.30" + "',"
        ZSql = ZSql + " Foto = " + "'" + ZZFoto + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Prueba ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Foto )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + Descripcion.Text + "',"
        ZSql = ZSql + "'" + "100" + "',"
        ZSql = ZSql + "'" + "2.20" + "',"
        ZSql = ZSql + "'" + imgX + "')"
        spPrueba = ZSql
        Set rstPrueba = db.OpenRecordset(spPrueba, dbOpenSnapshot, dbSQLPassThrough)
        
    End If

End Sub

Private Sub ImpreVenta_Click()
    If MuestraVenta.Visible = False Then
        MuestraVenta.Height = 2775
        MuestraVenta.Left = 2520
        MuestraVenta.Top = 1080
        MuestraVenta.Width = 4215
        MuestraVenta.Visible = True
            Else
        MuestraVenta.Visible = False
    End If
End Sub

Private Sub MuestraVenta_Click()
    MuestraVenta.Visible = False
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Desde1.Text = ""
    Hasta1.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Color.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Color_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Proveedor.SetFocus
    End If
    If KeyAscii = 27 Then
        Color.Text = ""
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = rstProveedor!Nombre
            Familia.SetFocus
                Else
            DesProveedor.Caption = ""
            Proveedor.Text = ""
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Familia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Familia"
        ZSql = ZSql + " Where Familia.Codigo = " + "'" + Familia.Text + "'"
        spFamilia = ZSql
        Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
        If rstFamilia.RecordCount > 0 Then
            DesFamilia.Caption = rstFamilia!Descripcion
            DescripcionII.SetFocus
                Else
            DesFamilia.Caption = ""
            Familia.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Familia.Text = ""
        DesFamilia.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DescripcionII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CodigoProveedor.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionII.Text = ""
    End If
End Sub

Private Sub Costo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Precio
        If Val(Costo.Text) <> 0 Then
            Costo.Text = Pusing("###,###.##", Costo.Text)
        End If
        Margen.SetFocus
    End If
    If KeyAscii = 27 Then
        Costo.Text = ""
        Call Calcula_Precio
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Margen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Calcula_Precio
        If Val(Margen.Text) <> 0 Then
            Margen.Text = Pusing("###,###.##", Margen.Text)
        End If
        Fob.SetFocus
    End If
    If KeyAscii = 27 Then
        Margen.Text = ""
        Call Calcula_Precio
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CostoFuturo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(CostoFuturo.Text) <> 0 Then
            CostoFuturo.Text = Pusing("###,###.##", CostoFuturo.Text)
        End If
        MargenFuturo.SetFocus
    End If
    If KeyAscii = 27 Then
        CostoFuturo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub MargenFuturo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(MargenFuturo.Text) <> 0 Then
            MargenFuturo.Text = Pusing("###,###.##", MargenFuturo.Text)
        End If
        Cif.SetFocus
    End If
    If KeyAscii = 27 Then
        MargenFuturo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fob_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Fob.Text) <> 0 Then
            Fob.Text = Pusing("###,###.###", Fob.Text)
        End If
        CostoFuturo.SetFocus
    End If
    If KeyAscii = 27 Then
        Fob.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cif_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Cif.Text) <> 0 Then
            Cif.Text = Pusing("###,###.###", Cif.Text)
        End If
        Minimo.SetFocus
    End If
    If KeyAscii = 27 Then
        Cif.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Minimo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Minimo.Text) <> 0 Then
            Minimo.Text = Pusing("###,###.##", Minimo.Text)
        End If
        UnidadCaja.SetFocus
    End If
    If KeyAscii = 27 Then
        Minimo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub UnidadCaja_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MinimoVenta.SetFocus
    End If
    If KeyAscii = 27 Then
        UnidadCaja.Text = ""
    End If
End Sub

Private Sub MinimoVenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Despacho.SetFocus
    End If
    If KeyAscii = 27 Then
        MinimoVenta.Text = ""
    End If
End Sub

Private Sub Despacho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Despacho"
        ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Despacho.Text + "'"
        spDespacho = ZSql
        Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
        If rstDespacho.RecordCount > 0 Then
            Rem DesDespacho.Caption = rstDespacho!Descripcion
            Comision.SetFocus
                Else
            Rem DesDespacho.Caption = ""
            Despacho.Text = ""
            Comision.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Comision_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Comision.Text) <> 0 Then
            Comision.Text = Pusing("###,###.##", Comision.Text)
        End If
        CodigoBarra.SetFocus
    End If
    If KeyAscii = 27 Then
        Comision.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CodigoBarra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PosicionII.SetFocus
    End If
    If KeyAscii = 27 Then
        CodigoBarra.Text = ""
    End If
End Sub

Private Sub CodigoProveedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo.SetFocus
    End If
    If KeyAscii = 27 Then
        CodigoProveedor.Text = ""
    End If
End Sub

Private Sub PosicionII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento.SetFocus
    End If
    If KeyAscii = 27 Then
        PosicionII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Descuento.Text) <> 0 Then
            Descuento.Text = Pusing("###,###.##", Descuento.Text)
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Descuento.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Posicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PosicionII.SetFocus
    End If
    If KeyAscii = 27 Then
        Posicion.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub



Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
        
            If Trim(Codigo.Text) <> "" Then
                ZZVeri = UCase(Left$(Codigo.Text, 1))
                If ZZVeri < "A" Or ZZVeri > "Z" Then
                    ZZVeri = Left$(ZZCodAnt, 1)
                    Codigo.Text = ZZVeri + Codigo.Text
                End If
            End If
        
            Auxi = UCase(Left$(Codigo.Text, 1))
            Auxi1 = Mid$(Codigo.Text, 2, 5)
            Call Ceros(Auxi1, 5)
            Codigo.Text = Auxi + Auxi1
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Codigo.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                Call Imprime_Datos
                    Else
                WArticulo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WArticulo
            End If
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
End Sub
    
Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde1.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta1.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Articulos"
     Opcion.AddItem "Grupos"
     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Despachos"

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
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Familia"
            ZSql = ZSql + " Order by Familia.Codigo"
            spFamilia = ZSql
            Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
            If rstFamilia.RecordCount > 0 Then
                With rstFamilia
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstFamilia.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Order by Proveedor.Proveedor"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Proveedor) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Despacho"
            ZSql = ZSql + " Order by Despacho.Codigo"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
            If rstDespacho.RecordCount > 0 Then
                With rstDespacho
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstDespacho.Close
            End If
        
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    Rem Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Familia.Text = WIndice.List(Indice)
            Call Familia_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
            
        Case 3
            Indice = Pantalla.ListIndex
            Despacho.Text = WIndice.List(Indice)
            Call Despacho_KeyPress(13)
                    
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    Codigo.Text = ""
    Descripcion.Text = ""
    Color.Text = ""
    Familia.Text = ""
    DesFamilia.Caption = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    MinimoVenta.Text = ""
    UnidadCaja.Text = ""
    Margen.Text = ""
    MargenFuturo.Text = ""
    Fob.Text = ""
    Cif.Text = ""
    CostoAnterior.Text = ""
    FechaCostoAnterior.Text = "  /  /    "
    Costo.Text = ""
    FechaCosto.Text = "  /  /    "
    CostoFuturo.Text = ""
    FechaCierre.Text = "  /  /    "
    FechaUltimaEntrada.Text = "  /  /    "
    FechaUltimaSalida.Text = "  /  /    "
    Minimo.Text = "1"
    Entradas.Text = ""
    Salidas.Text = ""
    Stock.Text = ""
    StockAnterior.Text = ""
    Venta1.Text = ""
    Venta2.Text = ""
    Venta3.Text = ""
    Venta4.Text = ""
    Venta5.Text = ""
    Venta6.Text = ""
    Posicion.Text = ""
    PosicionII.Text = ""
    Despacho.Text = ""
    Rem DesDespacho.Caption = ""
    Comision.Text = "10.00"
    CodigoBarra.Text = ""
    Precio.Text = ""
    Descuento.Text = ""
    CodigoProveedor.Text = ""
    DescripcionII.Text = ""
    Embarque.Text = ""
    
    Iva.Clear
    
    Iva.AddItem "21%"
    Iva.AddItem "10.5%"
    
    Iva.ListIndex = 0
    
    ListaPrecio.Clear
    
    ListaPrecio.AddItem ""
    ListaPrecio.AddItem "No Imprime"
    
    ListaPrecio.ListIndex = 0
    
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
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Familias"
            ZSql = ZSql + " Where Familias.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Familias.Familia"
            spFamilia = ZSql
            Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
            If rstFamilia.RecordCount > 0 Then
                With rstFamilia
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Familia) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Familia
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstFamilia.Close
            End If
            
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Proveedor.Proveedor"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                With rstProveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Proveedor) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Despacho"
            ZSql = ZSql + " Where Despacho.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Despacho.Codigo"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
            If rstDespacho.RecordCount > 0 Then
                With rstDespacho
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstDespacho.Close
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

Private Sub Codigo_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Linea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Despacho_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub SubLinea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"
    Opcion.AddItem "SubLineas de Ventas"

    Opcion.ListIndex = 3
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Color_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Familia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Costo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Margen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CostoFuturo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub MargenFuturo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fob_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cif_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Minimo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub UnidadCaja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub MinimoVenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Comision_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CodigoBarra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CodigoProveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DescripcionII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PosicionII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descuento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub




Private Sub Stock_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub StockII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub StockIII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub StockIV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Clasificacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle6_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Talle10_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CodPrv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Despacho_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ListaPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Desde1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta1_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call CmdDelete_Click
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
            Call cmdClose_Click
        Case 122
            Call Acepta_Click
        Case 123
            Call Cancela_click
        Case Else
    End Select
End Sub

Private Sub Anterior_Click()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveLast
            Codigo.Text = rstArticulo!Codigo
        End With
        rstArticulo.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de Articulos")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM Articulo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.MoveFirst
        ZUltimo = IIf(IsNull(rstArticulo!CodigoMenor), "", rstArticulo!CodigoMenor)
        Codigo.Text = ZUltimo
        rstArticulo.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Articulo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.MoveLast
        ZUltimo = IIf(IsNull(rstArticulo!CodigoMayor), "", rstArticulo!CodigoMayor)
        Codigo.Text = ZUltimo
        rstArticulo.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Codigo.Text = rstArticulo!Codigo
        End With
        rstArticulo.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Articulos")
    End If
End Sub


Private Sub Calcula_Precio()
    
    ZZPrecio = 0
    If Val(Costo.Text) <> 0 And Val(Margen.Text) <> 0 Then
        ZZMargen = Val(Costo.Text) * (Val(Margen.Text) / 100)
        Call Redondeo(ZZMargen)
        ZZPrecio = Val(Costo.Text) + ZZMargen
    End If
    Precio.Text = Str$(ZZPrecio)
    Precio.Text = Pusing("###,###.##", Precio.Text)
    
End Sub















































