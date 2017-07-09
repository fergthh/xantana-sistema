VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form prgInsumo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   9030
   ClientLeft      =   180
   ClientTop       =   1065
   ClientWidth     =   12480
   LinkTopic       =   "Form2"
   ScaleHeight     =   9030
   ScaleWidth      =   12480
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   9240
      TabIndex        =   60
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox StockVI 
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
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   58
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox StockIII 
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
      Left            =   7320
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   56
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox StockV 
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
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   54
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox StockII 
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
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   52
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox StockIV 
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
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   50
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox StockI 
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
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   48
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox Capacidad 
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
      Left            =   9000
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   46
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Envase"
      Height          =   375
      Left            =   5640
      TabIndex        =   45
      Top             =   2880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   5400
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   5400
      Width           =   375
   End
   Begin VB.ComboBox Moneda 
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
      Left            =   6240
      TabIndex        =   41
      Top             =   2160
      Width           =   2415
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
      TabIndex        =   38
      Top             =   2880
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
      MaxLength       =   10
      TabIndex        =   35
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Proveedor 
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
      MaxLength       =   8
      TabIndex        =   32
      Top             =   360
      Width           =   1695
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
      Left            =   240
      MouseIcon       =   "insumo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   4680
      Width           =   735
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
      MouseIcon       =   "insumo.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Elimina el Registro"
      Top             =   4680
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
      Left            =   1920
      MouseIcon       =   "insumo.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Limpia la pantalla"
      Top             =   4680
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
      Left            =   2760
      MouseIcon       =   "insumo.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Consulta de Datos"
      Top             =   4680
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
      Left            =   6960
      MouseIcon       =   "insumo.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Impresion "
      Top             =   4680
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
      Left            =   7800
      MouseIcon       =   "insumo.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Salida"
      Top             =   4680
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
      Left            =   3600
      MouseIcon       =   "insumo.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":46D2
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Primer Registro"
      Top             =   4680
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
      Left            =   4440
      MouseIcon       =   "insumo.frx":4B14
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":4E1E
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Registro Anterior"
      Top             =   4680
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
      Left            =   5280
      MouseIcon       =   "insumo.frx":5260
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":556A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Registro Siguiente"
      Top             =   4680
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
      Left            =   6120
      MouseIcon       =   "insumo.frx":59AC
      MousePointer    =   99  'Custom
      Picture         =   "insumo.frx":5CB6
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Salida"
      Top             =   4680
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
      Left            =   480
      TabIndex        =   5
      Top             =   6240
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
         MouseIcon       =   "insumo.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "insumo.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   31
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
         MouseIcon       =   "insumo.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "insumo.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   960
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
         MaxLength       =   16
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
         MaxLength       =   16
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
      TabIndex        =   19
      Top             =   5760
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
      Left            =   1080
      TabIndex        =   14
      Top             =   6240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Linea 
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
      Left            =   7320
      MaxLength       =   4
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
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
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   13
      Text            =   " "
      Top             =   2880
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
      MaxLength       =   16
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   2055
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11400
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Insumos.rpt"
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
      Left            =   3600
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
      Height          =   2700
      ItemData        =   "insumo.frx":6F90
      Left            =   240
      List            =   "insumo.frx":6F97
      TabIndex        =   3
      Top             =   6120
      Visible         =   0   'False
      Width           =   7335
   End
   Begin MSMask.MaskEdBox FechaCosto 
      Height          =   285
      Left            =   2640
      TabIndex        =   37
      Top             =   2160
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4455
      Left            =   8640
      TabIndex        =   42
      Top             =   4560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   7858
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label12 
      Caption         =   "en Terceros"
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
      Left            =   6000
      TabIndex        =   59
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Deposito III"
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
      Left            =   6000
      TabIndex        =   57
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "MK"
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
      Left            =   3120
      TabIndex        =   55
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Produccion"
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
      Left            =   3120
      TabIndex        =   53
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   11400
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label7 
      Caption         =   "De Cliente"
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
      Left            =   240
      TabIndex        =   51
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "General"
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
      Left            =   240
      TabIndex        =   49
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Capacidad"
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
      Left            =   7800
      TabIndex        =   47
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Moneda"
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
      TabIndex        =   40
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   11520
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   11520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   11520
      Y1              =   1920
      Y2              =   1920
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
      TabIndex        =   39
      Top             =   2880
      Width           =   855
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
      TabIndex        =   36
      Top             =   2160
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
      TabIndex        =   34
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
      Left            =   3600
      TabIndex        =   33
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label DesLinea 
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
      Left            =   8280
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
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
      Left            =   2760
      TabIndex        =   12
      Top             =   2880
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
Attribute VB_Name = "prgInsumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZCostoAnterior As Double
Dim ZZPorce(100, 2) As String
Rem para el vector

Dim WWWPasaCodigo As String
Dim WWWPasaCosto As String

Private WFecha As String
Private WPlazo1 As Integer
Private WVencimiento As String


Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Sub Imprime_Descripcion()

    
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
            
    
End Sub

Sub Verifica_datos()
    If Val(Linea.Text) = 0 Then
        Linea.Text = "0"
    End If
    If Val(Costo.Text) = 0 Then
        Costo.Text = "0"
    End If
    If Val(Minimo.Text) = 0 Then
        Minimo.Text = "0"
    End If
    If Val(Stock.Text) = 0 Then
        Stock.Text = "0"
    End If
    If Val(StockI.Text) = 0 Then
        StockI.Text = "0"
    End If
    If Val(StockII.Text) = 0 Then
        StockII.Text = "0"
    End If
    If Val(StockIII.Text) = 0 Then
        StockIII.Text = "0"
    End If
    If Val(StockIV.Text) = 0 Then
        StockIV.Text = "0"
    End If
    If Val(StockV.Text) = 0 Then
        StockV.Text = "0"
    End If
    If Val(StockVI.Text) = 0 Then
        StockVI.Text = "0"
    End If
End Sub

Sub Format_datos()
    If Val(Costo.Text) <> 0 Then
        Costo.Text = Pusing("###,###.###", Costo.Text)
    End If
End Sub

Sub Imprime_Datos()

    ZZCodAnt = Codigo.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Insumo"
    ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Codigo.Text + "'"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumo.RecordCount > 0 Then
        Codigo.Text = Trim(rstInsumo!Codigo)
        Descripcion.Text = Trim(rstInsumo!Descripcion)
        Linea.Text = Trim(rstInsumo!Linea)
        Proveedor.Text = Trim(rstInsumo!Proveedor)
        Minimo.Text = Str$(rstInsumo!Minimo)
        Costo.Text = Str$(rstInsumo!Costo)
        If Trim(rstInsumo!FechaCosto) <> "" Then
            FechaCosto.Text = rstInsumo!FechaCosto
                Else
            FechaCosto.Text = "  /  /    "
        End If
        Stock.Text = Str$(rstInsumo!Stock)
        
        ZStockI = IIf(IsNull(rstInsumo!StockI), "0", rstInsumo!StockI)
        ZStockII = IIf(IsNull(rstInsumo!StockII), "0", rstInsumo!StockII)
        ZStockIII = IIf(IsNull(rstInsumo!StockIII), "0", rstInsumo!StockIII)
        ZStockIV = IIf(IsNull(rstInsumo!StockIV), "0", rstInsumo!StockIV)
        ZStockV = IIf(IsNull(rstInsumo!StockV), "0", rstInsumo!StockV)
        ZStockVI = IIf(IsNull(rstInsumo!StockVI), "0", rstInsumo!StockVI)
        
        StockI.Text = Str$(ZStockI)
        StockII.Text = Str$(ZStockII)
        StockIII.Text = Str$(ZStockIII)
        StockIV.Text = Str$(ZStockIV)
        StockV.Text = Str$(ZStockV)
        StockVI.Text = Str$(ZStockVI)
        
        Moneda.ListIndex = rstInsumo!Moneda
        
        
        rstInsumo.Close
        
        Call LeeHistorial
        
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
    ZSql = ZSql + "UPDATE Insumo SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + "1" + "'"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Insumo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Insumo.Codigo, Insumo.Descripcion, Insumo.Linea, Insumo.Stock " _
            + "From " _
            + DSQ + ".dbo.Insumo Insumo " _
            + "Where " _
            + "Insumo.Codigo >= '" + Desde.Text + "' AND " _
            + "Insumo.Codigo <= '" + Hasta.Text + "'"
    
    Listado.Connect = Connect()
    
    Uno = "{Insumo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub


Private Sub Cancela_Click()
    Frame2.Visible = False
    Codigo.SetFocus
End Sub

Private Sub cmdAdd_Click()


    If Moneda.ListIndex = 0 Then
        m$ = "Se debe informar moneda"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Insumos")
        Exit Sub
    End If
        

    If Codigo.Text <> "" Then

        ZZActualizaCosto = "N"

        Call Verifica_datos
        Codigo.Text = UCase(Trim(Codigo.Text))
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Insumo"
        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Codigo.Text + "'"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        If rstInsumo.RecordCount > 0 Then
        
            ZZCosto = rstInsumo!Costo
            rstInsumo.Close
        
            Rem If ZZCosto <> Val(Costo.Text) Then
                
                ZZActualizaCosto = "S"
                FechaCosto.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(FechaCosto.Text, 4) + Mid$(FechaCosto.Text, 4, 2) + Left$(FechaCosto.Text, 2)
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO InsumoHistorial ("
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Costo )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + Codigo.Text + "',"
                ZSql = ZSql + "'" + FechaCosto.Text + "',"
                ZSql = ZSql + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + "'" + Str$(ZZCosto) + "')"
                spInsumoHistorial = ZSql
                Set rstInsumoHistorial = db.OpenRecordset(spInsumoHistorial, dbOpenSnapshot, dbSQLPassThrough)
            
            Rem End If
            
            FechaCosto.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZZOrdFechaCosto = Right$(FechaCosto.Text, 4) + Mid$(FechaCosto.Text, 4, 2) + Left$(FechaCosto.Text, 2)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Insumo SET "
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
            ZSql = ZSql + " Linea = " + "'" + Linea.Text + "',"
            ZSql = ZSql + " Proveedor = " + "'" + Proveedor.Text + "',"
            ZSql = ZSql + " Costo = " + "'" + Costo.Text + "',"
            ZSql = ZSql + " MOneda = " + "'" + Str$(Moneda.ListIndex) + "',"
            ZSql = ZSql + " FechaCosto = " + "'" + FechaCosto.Text + "',"
            ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
            ZSql = ZSql + " Minimo = " + "'" + Minimo.Text + "',"
            ZSql = ZSql + " Stock = " + "'" + Stock.Text + "',"
            ZSql = ZSql + " StockI = " + "'" + StockI.Text + "',"
            ZSql = ZSql + " StockII = " + "'" + StockII.Text + "',"
            ZSql = ZSql + " StockIII = " + "'" + StockIII.Text + "',"
            ZSql = ZSql + " StockIV = " + "'" + StockIV.Text + "',"
            ZSql = ZSql + " StockV = " + "'" + StockV.Text + "',"
            ZSql = ZSql + " StockVI = " + "'" + StockVI.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZZActualizaCosto = "S"
            FechaCosto.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZZOrdFechaCosto = Right$(FechaCosto.Text, 4) + Mid$(FechaCosto.Text, 4, 2) + Left$(FechaCosto.Text, 2)
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Insumo ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Costo ,"
            ZSql = ZSql + "Moneda ,"
            ZSql = ZSql + "FechaCosto ,"
            ZSql = ZSql + "OrdFechaCosto ,"
            ZSql = ZSql + "Minimo ,"
            ZSql = ZSql + "Stock ,"
            ZSql = ZSql + "StockI ,"
            ZSql = ZSql + "StockII ,"
            ZSql = ZSql + "StockIII ,"
            ZSql = ZSql + "StockIV ,"
            ZSql = ZSql + "StockV,"
            ZSql = ZSql + "StockVI )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + Linea.Text + "',"
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + Costo.Text + "',"
            ZSql = ZSql + "'" + Str$(Moneda.ListIndex) + "',"
            ZSql = ZSql + "'" + FechaCosto.Text + "',"
            ZSql = ZSql + "'" + ZZOrdFechaCosto + "',"
            ZSql = ZSql + "'" + Minimo.Text + "',"
            ZSql = ZSql + "'" + Stock.Text + "',"
            ZSql = ZSql + "'" + StockI.Text + "',"
            ZSql = ZSql + "'" + StockII.Text + "',"
            ZSql = ZSql + "'" + StockIII.Text + "',"
            ZSql = ZSql + "'" + StockIV.Text + "',"
            ZSql = ZSql + "'" + StockV.Text + "',"
            ZSql = ZSql + "'" + StockVI.Text + "')"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        If ZZActualizaCosto = "S" Then
            
            WWWPasaCodigo = Codigo.Text
            WWWPasaCosto = Costo.Text
            Call Actualiza_Articulos
            
            ZZTipo = Right(Codigo.Text, 2)
            ZZDilu = Right(Codigo.Text, 1)
            ZZCostoDilu = 0
            ZZCostoEsencia = Val(Costo.Text)
                    
            If ZZTipo = "TP" Or ZZTipo = "TB" Or ZZTipo = "TI" Or ZZTipo = "TA" Then
                        
                ZZLargo = Len(Codigo.Text) - 2
                ZZBase = Left$(Codigo.Text, ZZLargo)
                        
                Select Case ZZTipo
                    Case "TP"
                        ZZDiluyente = "Dilu-P"
                        ZZDilu = "P"
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Insumo"
                        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZZDiluyente + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstInsumo.RecordCount > 0 Then
                            ZZCostoDilu = rstInsumo!Costo
                        End If
                        
                    Case "TB"
                        ZZDiluyente = "Dilu-B"
                        ZZDilu = "B"
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Insumo"
                        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZZDiluyente + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstInsumo.RecordCount > 0 Then
                            ZZCostoDilu = rstInsumo!Costo
                        End If
                    
                    Case "TI"
                        ZZDiluyente = "Dilu-I"
                        ZZDilu = "I"
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Insumo"
                        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZZDiluyente + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstInsumo.RecordCount > 0 Then
                            ZZCostoDilu = rstInsumo!Costo
                        End If
                    
                    Case "TA"
                        ZZDiluyente = "Dilu-A"
                        ZZDilu = "A"
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Insumo"
                        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZZDiluyente + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstInsumo.RecordCount > 0 Then
                            ZZCostoDilu = rstInsumo!Costo
                        End If
                    
                    Case Else
                    
                End Select
                
                If ZZCostoDilu = 0 Then
                    m$ = "NO se puede actualizar los precios por que el costo del diluyente es igaual a 0"
                    aaaaaa% = MsgBox(m$, 0, "Archivo de Insumos")
                        Else
                    For Ciclo = 1 To 19
                        
                        ZZAuxiliar = ZZBase + ZZPorce(Ciclo, 1) + ZZDilu
                        
                        ZZCantiEsencia = Val(ZZPorce(Ciclo, 2))
                        ZZCantiDilu = 100 - ZZCantiEsencia
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Insumo"
                        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZZAuxiliar + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstInsumo.RecordCount > 0 Then
                        
                            ZZCostoAnterior = rstInsumo!Costo
                            ZZCosto = (ZZCostoEsencia * (ZZCantiEsencia / 100)) + (ZZCostoDilu * (ZZCantiDilu / 100))
                        
                            rstInsumo.Close
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE Insumo SET "
                            ZSql = ZSql + " Costo = " + "'" + Str$(ZZCosto) + "',"
                            ZSql = ZSql + " FechaCosto = " + "'" + FechaCosto.Text + "',"
                            ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "'"
                            ZSql = ZSql + " Where Codigo = " + "'" + ZZAuxiliar + "'"
                            spInsumo = ZSql
                            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                                    
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO InsumoHistorial ("
                            ZSql = ZSql + "Codigo ,"
                            ZSql = ZSql + "Fecha ,"
                            ZSql = ZSql + "OrdFecha ,"
                            ZSql = ZSql + "Costo )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + ZZAuxiliar + "',"
                            ZSql = ZSql + "'" + FechaCosto.Text + "',"
                            ZSql = ZSql + "'" + ZZOrdFechaCosto + "',"
                            ZSql = ZSql + "'" + Str$(ZZCosto) + "')"
                            spInsumoHistorial = ZSql
                            Set rstInsumoHistorial = db.OpenRecordset(spInsumoHistorial, dbOpenSnapshot, dbSQLPassThrough)
                            
                            WWWPasaCodigo = ZZAuxiliar
                            WWWPasaCosto = Str$(ZZCosto)
                            Call Actualiza_Articulos
                            
                        End If
                        
                    Next Ciclo
                                
                End If
                
            End If
            
        End If
        
        Rem Call CmdLimpiar_Click
    
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Insumos")
    
        
        Codigo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Codigo.Text <> "" Then
        
        If Val(Stock.Text) = 0 Then
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Insumo"
            ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Codigo.Text + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstInsumo.RecordCount > 0 Then
                rstInsumo.Close
                T$ = "Borrar Registro"
                m$ = "Desea Borrar el Registro "
                Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
                If Respuestaaaaaa% = 6 Then
            
                    ZSql = ""
                    ZSql = ZSql + "DELETE Insumo"
                    ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                    spInsumo = ZSql
                    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Call CmdLimpiar_Click
                End If
            End If
            
                Else
                
            m$ = "No se puede dar de baja ya que posee stock"
            aaaaaa% = MsgBox(m$, 0, "Archivo de Insumos")
            
        End If
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Descripcion.Text = ""
    Linea.Text = ""
    DesLinea.Caption = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Costo.Text = ""
    FechaCosto.Text = "  /  /    "
    Minimo.Text = ""
    Stock.Text = ""
    StockI.Text = ""
    StockII.Text = ""
    StockIII.Text = ""
    StockIV.Text = ""
    StockV.Text = ""
    StockVI.Text = ""
    Moneda.ListIndex = 0

    Call Limpia_Vector
        
    Codigo.SetFocus
End Sub

Private Sub CmdClose_Click()
    prgInsumo.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Command1_Click()
    
   ZSql = ""
    ZSql = ZSql + "UPDATE Insumo SET "
    ZSql = ZSql + " StockI = Stock - StockII - StockIII - StockIV- StockV - StockVI"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)

Stop
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Insumo SET "
    ZSql = ZSql + " StockI = Stock " + ","
    ZSql = ZSql + " StockII = 0 " + ","
    ZSql = ZSql + " StockIII = 0 " + ","
    ZSql = ZSql + " StockIV = 0 " + ","
    ZSql = ZSql + " StockV = 0 " + ","
    ZSql = ZSql + " StockVI = 0 "
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Proveedor.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
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
            Costo.SetFocus
                Else
            DesProveedor.Caption = ""
            Proveedor.Text = ""
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub LInea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo.SetFocus
    End If
    If KeyAscii = 27 Then
        Linea.Text = ""
        DesLinea.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Costo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Costo.Text) <> 0 Then
            Costo.Text = Pusing("###,###.###", Costo.Text)
        End If
        Minimo.SetFocus
    End If
    If KeyAscii = 27 Then
        Costo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Minimo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Minimo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Insumo"
            ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Codigo.Text + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstInsumo.RecordCount > 0 Then
                rstInsumo.Close
                Call Imprime_Datos
                    Else
                WInsumo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WInsumo
            End If
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
    
Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Insumos"
     Opcion.AddItem "Proveedores"

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
            ZSql = ZSql + " FROM Insumo"
            ZSql = ZSql + " Order by Insumo.Codigo"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstInsumo.RecordCount > 0 Then
                With rstInsumo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Trim(Proveedor.Text) = "" Or Trim(UCase(Proveedor.Text)) = Trim(UCase(rstInsumo!Proveedor)) Then
                                IngresaItem = rstInsumo!Codigo + " " + rstInsumo!Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstInsumo!Codigo
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstInsumo.Close
            End If
        
        Case 1
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
                            IngresaItem = !Proveedor + " " + !Nombre
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
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
                    
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    Moneda.Clear
    
    Moneda.AddItem ""
    Moneda.AddItem "Pesos"
    Moneda.AddItem "Dolares"

    Moneda.ListIndex = 0

    Codigo.Text = ""
    Descripcion.Text = ""
    Linea.Text = ""
    DesLinea.Caption = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Costo.Text = ""
    FechaCosto.Text = "  /  /    "
    Minimo.Text = ""
    Stock.Text = ""
    StockI.Text = ""
    StockII.Text = ""
    StockIII.Text = ""
    StockIV.Text = ""
    StockV.Text = ""
    StockVI.Text = ""
                    
    ZZPorce(1, 1) = "A"
    ZZPorce(1, 2) = "5"
            
    ZZPorce(2, 1) = "B"
    ZZPorce(2, 2) = "10"
            
    ZZPorce(3, 1) = "C"
    ZZPorce(3, 2) = "15"
            
    ZZPorce(4, 1) = "D"
    ZZPorce(4, 2) = "20"
            
    ZZPorce(5, 1) = "E"
    ZZPorce(5, 2) = "25"
            
    ZZPorce(6, 1) = "F"
    ZZPorce(6, 2) = "30"
            
    ZZPorce(7, 1) = "G"
    ZZPorce(7, 2) = "35"
            
    ZZPorce(8, 1) = "H"
    ZZPorce(8, 2) = "40"
            
    ZZPorce(9, 1) = "I"
    ZZPorce(9, 2) = "45"
            
    ZZPorce(10, 1) = "J"
    ZZPorce(10, 2) = "50"
            
    ZZPorce(11, 1) = "K"
    ZZPorce(11, 2) = "55"
            
    ZZPorce(12, 1) = "L"
    ZZPorce(12, 2) = "60"
            
    ZZPorce(13, 1) = "M"
    ZZPorce(13, 2) = "65"
            
    ZZPorce(14, 1) = "N"
    ZZPorce(14, 2) = "70"
            
    ZZPorce(15, 1) = "O"
    ZZPorce(15, 2) = "75"
            
    ZZPorce(16, 1) = "P"
    ZZPorce(16, 2) = "80"
            
    ZZPorce(17, 1) = "Q"
    ZZPorce(17, 2) = "85"
            
    ZZPorce(18, 1) = "R"
    ZZPorce(18, 2) = "90"
            
    ZZPorce(19, 1) = "S"
    ZZPorce(19, 2) = "95"

    Call Limpia_Vector
        
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
            ZSql = ZSql + " FROM Insumo"
            ZSql = ZSql + " Where Insumo.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Insumo.Codigo"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstInsumo.RecordCount > 0 Then
                With rstInsumo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Trim(Proveedor.Text) = "" Or Trim(UCase(Proveedor.Text)) = Trim(UCase(rstInsumo!Proveedor)) Then
                                IngresaItem = !Codigo + " " + !Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Codigo
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstInsumo.Close
            End If
            
        Case 1
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
                            IngresaItem = !Proveedor + " " + !Nombre
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
    Opcion.AddItem "Insumos"
    Opcion.AddItem "Grupos"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub LInea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Insumos"
    Opcion.AddItem "Grupos"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Insumos"
    Opcion.AddItem "Grupos"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 1
    
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

Private Sub Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Minimo_KeyDown(KeyCode As Integer, Shift As Integer)
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
    ZSql = ZSql + " FROM Insumo"
    ZSql = ZSql + " Where Insumo.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Insumo.Codigo"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumo.RecordCount > 0 Then
        With rstInsumo
            .MoveLast
            Codigo.Text = rstInsumo!Codigo
        End With
        rstInsumo.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Insumos")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM Insumo"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumo.RecordCount > 0 Then
        rstInsumo.MoveFirst
        ZUltimo = IIf(IsNull(rstInsumo!CodigoMenor), "", rstInsumo!CodigoMenor)
        Codigo.Text = ZUltimo
        rstInsumo.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub


Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Insumo"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumo.RecordCount > 0 Then
        rstInsumo.MoveLast
        ZUltimo = IIf(IsNull(rstInsumo!CodigoMayor), "", rstInsumo!CodigoMayor)
        Codigo.Text = ZUltimo
        rstInsumo.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Insumo"
    ZSql = ZSql + " Where Insumo.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Insumo.Codigo"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumo.RecordCount > 0 Then
        With rstInsumo
            .MoveFirst
            Codigo.Text = rstInsumo!Codigo
        End With
        rstInsumo.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Insumos")
    End If
End Sub


Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 3
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Costo"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
       End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub LeeHistorial()
        
    Renglon = 0
    Call Limpia_Vector
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM InsumoHistorial"
    ZSql = ZSql + " Where InsumoHistorial.Codigo = " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by InsumoHistorial.OrdFecha"
        
    spInsumoHistorial = ZSql
    Set rstInsumoHistorial = db.OpenRecordset(spInsumoHistorial, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumoHistorial.RecordCount > 0 Then
    
        With rstInsumoHistorial
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                    WVector1.TextMatrix(Renglon, 1) = rstInsumoHistorial!Fecha
                    WVector1.TextMatrix(Renglon, 2) = Pusing("###,###.###", Str$(rstInsumoHistorial!Costo))
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstInsumoHistorial.Close
    End If
End Sub

Private Sub Actualiza_Articulos()

    Rem Dim WWWPasaCodigo As String
    Rem Dim WWWPasaCosto As String

    Dim WWWVector(1000) As String
    Erase WWWVector
    WWWLugar = 0


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Insumo = " + "'" + WWWPasaCodigo + "'"
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    If UCase(Trim(!Linea)) = UCase(Trim("ESC")) And UCase(Trim(!Tamano)) = UCase(Trim("7")) Then
                    
                        Select Case UCase(Trim(!Tipo))
                            Case "PE", "HO", "JA", "VE", "PR", "CR", "ES", "AM", "PI", "LI", "AH", "SM"
                                WWWLugar = WWWLugar + 1
                                WWWVector(WWWLugar) = !Codigo
                            Case Else
                        End Select
                        
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If

    For Ciclo = 1 To WWWLugar
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WWWVector(Ciclo) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            WWWLinea = rstArticulo!Linea
            WWWTipo = rstArticulo!Tipo
            WWWFragancia = rstArticulo!Fragancia
            WWWCalidad = rstArticulo!Calidad
            WWWTamano = rstArticulo!Tamano
        
            Select Case UCase(Trim(rstArticulo!Tipo))
                Case "PE"
                    Select Case UCase(Trim(rstArticulo!Linea))
                        Case "SOL"
                            Select Case UCase(Trim(rstArticulo!Calidad))
                                Case "U", "4"
                                    WWWCosto1 = (((Val(WWWPasaCosto) * 2) * 1.4) * 0.1) + 3.2
                                    WWWCosto2 = 0
                                    WWWCosto3 = 0
                                    WWWCosto4 = 0
                                Case "0"
                                    WWWCosto1 = (((Val(WWWPasaCosto) * 2) * 1.4) * 0.15) + 3.5
                                    WWWCosto2 = 0
                                    WWWCosto3 = 0
                                    WWWCosto4 = 0
                                Case "6"
                                    WWWCosto1 = (((Val(WWWPasaCosto) * 2) * 1.4) * 0.2) + 3.5
                                    WWWCosto2 = 0
                                    WWWCosto3 = 0
                                    WWWCosto4 = 0
                                Case "5"
                                    WWWCosto1 = (((Val(WWWPasaCosto) * 2) * 1.4) * 0.05) + 3
                                    WWWCosto2 = 0
                                    WWWCosto3 = 0
                                    WWWCosto4 = 0
                                Case "3"
                                    WWWCosto1 = (((Val(WWWPasaCosto) * 2) * 1.4) * 0.18) + 3.5
                                    WWWCosto2 = 0
                                    WWWCosto3 = 0
                                    WWWCosto4 = 0
                                Case "A"
                                    WWWCosto1 = (((Val(WWWPasaCosto) * 2) * 1.4) * 0.12) + 3.5
                                    WWWCosto2 = 0
                                    WWWCosto3 = 0
                                    WWWCosto4 = 0
                                Case "8"
                                    WWWCosto1 = (((Val(WWWPasaCosto) * 2) * 1.4) * 0.16) + 3.5
                                    WWWCosto2 = 0
                                    WWWCosto3 = 0
                                    WWWCosto4 = 0
                                Case Else
                                    WWWCosto4 = Val(WWWPasaCosto) * 2
                                    WWWCosto1 = WWWCosto4 * 1.4
                                    WWWCosto2 = WWWCosto4 * 1.2
                                    WWWCosto3 = WWWCosto4 * 1.1
                            End Select
                                
                        Case Else
                            WWWCosto4 = Val(WWWPasaCosto) * 2
                            WWWCosto1 = WWWCosto4 * 1.4
                            WWWCosto2 = WWWCosto4 * 1.2
                            WWWCosto3 = WWWCosto4 * 1.1
                    End Select
                Case "HO", "JA", "VE", "CR", "ES", "AH"
                    WWWCosto4 = Val(WWWPasaCosto) * 2
                    WWWCosto1 = WWWCosto4 * 1.4
                    WWWCosto2 = WWWCosto4 * 1.2
                    WWWCosto3 = WWWCosto4 * 1.1
                Case "PR"
                    WWWCosto4 = Val(WWWPasaCosto) * 2 * 1.2
                    WWWCosto1 = WWWCosto4 * 1.4
                    WWWCosto2 = WWWCosto4 * 1.2
                    WWWCosto3 = WWWCosto4 * 1.1
                Case "AM"
                    WWWCosto2 = Val(WWWPasaCosto) * 1.7
                    WWWCosto1 = WWWCosto2 * 1.2
                    WWWCosto3 = 0
                    WWWCosto4 = 0
                Case "LI"
                    WWWCosto2 = Val(WWWPasaCosto) * 2
                    WWWCosto1 = WWWCosto2 * 1.2
                    WWWCosto3 = 0
                    WWWCosto4 = 0
                Case "SM", "PI"
                    WWWCosto1 = Val(WWWPasaCosto) * 2
                    WWWCosto2 = 0
                    WWWCosto3 = 0
                    WWWCosto4 = 0
                Case Else
            End Select
            
            rstArticulo.Close
            
            
            WWWLista = "0"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Precios"
            ZSql = ZSql + " Where Precios.Lista = " + "'" + WWWLista + "'"
            ZSql = ZSql + " and Precios.LInea = " + "'" + WWWLinea + "'"
            ZSql = ZSql + " and Precios.Tipo = " + "'" + WWWTipo + "'"
            ZSql = ZSql + " and Precios.fragancia = " + "'" + WWWFragancia + "'"
            ZSql = ZSql + " and Precios.Calidad = " + "'" + WWWCalidad + "'"
            ZSql = ZSql + " and Precios.Tamano = " + "'" + WWWTamano + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                
                ZZClave = rstPrecios!Clave
                ZZCodigo = rstPrecios!Codigo
                ZZLinea = rstPrecios!Linea
                ZZTipo = rstPrecios!Tipo
                ZZFragancia = rstPrecios!Fragancia
                ZZCalidad = rstPrecios!Calidad
                ZZTamano = rstPrecios!Tamano
                ZZLista = rstPrecios!Lista
                ZZDesde = rstPrecios!Desde
                ZZHasta = rstPrecios!Hasta
                ZZOrdDesde = rstPrecios!OrdDesde
                ZZOrdHasta = rstPrecios!OrdHasta
                ZZTope1 = rstPrecios!Tope1
                ZZValor1 = rstPrecios!Valor1
                ZZTope2 = rstPrecios!Tope2
                ZZValor2 = rstPrecios!Valor2
                ZZTope3 = rstPrecios!Tope3
                ZZValor3 = rstPrecios!Valor3
                ZZTope4 = rstPrecios!Tope4
                ZZValor4 = rstPrecios!Valor4
                rstPrecios.Close
                
                If WWWCosto1 > ZZValor1 Then
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO PreciosHistorial ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "Codigo ,"
                    ZSql = ZSql + "Linea ,"
                    ZSql = ZSql + "Tipo ,"
                    ZSql = ZSql + "Fragancia ,"
                    ZSql = ZSql + "Calidad ,"
                    ZSql = ZSql + "Tamano ,"
                    ZSql = ZSql + "Lista ,"
                    ZSql = ZSql + "Desde ,"
                    ZSql = ZSql + "Hasta ,"
                    ZSql = ZSql + "OrdDesde ,"
                    ZSql = ZSql + "OrdHasta ,"
                    ZSql = ZSql + "Tope1 ,"
                    ZSql = ZSql + "Valor1 ,"
                    ZSql = ZSql + "Tope2 ,"
                    ZSql = ZSql + "Valor2 ,"
                    ZSql = ZSql + "Tope3 ,"
                    ZSql = ZSql + "Valor3 ,"
                    ZSql = ZSql + "Tope4 ,"
                    ZSql = ZSql + "Valor4 )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + ZZClave + "',"
                    ZSql = ZSql + "'" + ZZCodigo + "',"
                    ZSql = ZSql + "'" + ZZLinea + "',"
                    ZSql = ZSql + "'" + ZZTipo + "',"
                    ZSql = ZSql + "'" + ZZFragancia + "',"
                    ZSql = ZSql + "'" + ZZCalidad + "',"
                    ZSql = ZSql + "'" + ZZTamano + "',"
                    ZSql = ZSql + "'" + ZZLista + "',"
                    ZSql = ZSql + "'" + ZZDesde + "',"
                    ZSql = ZSql + "'" + ZZHasta + "',"
                    ZSql = ZSql + "'" + ZZOrdDesde + "',"
                    ZSql = ZSql + "'" + ZZOrdHasta + "',"
                    ZSql = ZSql + "'" + Str$(ZZTope1) + "',"
                    ZSql = ZSql + "'" + Str$(ZZValor1) + "',"
                    ZSql = ZSql + "'" + Str$(ZZTope2) + "',"
                    ZSql = ZSql + "'" + Str$(ZZValor2) + "',"
                    ZSql = ZSql + "'" + Str$(ZZTope3) + "',"
                    ZSql = ZSql + "'" + Str$(ZZValor3) + "',"
                    ZSql = ZSql + "'" + Str$(ZZTope4) + "',"
                    ZSql = ZSql + "'" + Str$(ZZValor4) + "')"
                    spPreciosHistorial = ZSql
                    Set rstPreciosHistorial = db.OpenRecordset(spPreciosHistorial, dbOpenSnapshot, dbSQLPassThrough)
            
                    WWWDesde = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    
                    WPlazo1 = 61
                    WFecha = WWWDesde
                    Call Calcula_vencimiento(WFecha, WPlazo1, WVencimiento)
                    WWWHasta = WVencimiento
                    
                    ZZOrdDesde = Right$(WWWDesde, 4) + Mid$(WWWDesde, 4, 2) + Left$(WWWDesde, 2)
                    ZZOrdHasta = Right$(WWWHasta, 4) + Mid$(WWWHasta, 4, 2) + Left$(WWWHasta, 2)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Precios SET "
                    ZSql = ZSql + " Desde = " + "'" + WWWDesde + "',"
                    ZSql = ZSql + " Hasta = " + "'" + WWWHasta + "',"
                    ZSql = ZSql + " OrdDesde = " + "'" + ZZOrdDesde + "',"
                    ZSql = ZSql + " OrdHasta = " + "'" + ZZOrdHasta + "',"
                    ZSql = ZSql + " Valor1 = " + "'" + Str$(WWWCosto1) + "',"
                    ZSql = ZSql + " Valor2 = " + "'" + Str$(WWWCosto2) + "',"
                    ZSql = ZSql + " Valor3 = " + "'" + Str$(WWWCosto3) + "',"
                    ZSql = ZSql + " Valor4 = " + "'" + Str$(WWWCosto4) + "'"
                    ZSql = ZSql + " Where Precios.Lista = " + "'" + WWWLista + "'"
                    ZSql = ZSql + " and Precios.LInea = " + "'" + WWWLinea + "'"
                    ZSql = ZSql + " and Precios.Tipo = " + "'" + WWWTipo + "'"
                    ZSql = ZSql + " and Precios.fragancia = " + "'" + WWWFragancia + "'"
                    ZSql = ZSql + " and Precios.Calidad = " + "'" + WWWCalidad + "'"
                    ZSql = ZSql + " and Precios.Tamano = " + "'" + WWWTamano + "'"
                    spPrecios = ZSql
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
                        Else
        
        
        
        
        
                    WWWDesde = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                    
                    WPlazo1 = 61
                    WFecha = WWWDesde
                    Call Calcula_vencimiento(WFecha, WPlazo1, WVencimiento)
                    WWWHasta = WVencimiento
                    
                    ZZOrdDesde = Right$(WWWDesde, 4) + Mid$(WWWDesde, 4, 2) + Left$(WWWDesde, 2)
                    ZZOrdHasta = Right$(WWWHasta, 4) + Mid$(WWWHasta, 4, 2) + Left$(WWWHasta, 2)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Precios SET "
                    ZSql = ZSql + " Desde = " + "'" + WWWDesde + "',"
                    ZSql = ZSql + " Hasta = " + "'" + WWWHasta + "',"
                    ZSql = ZSql + " OrdDesde = " + "'" + ZZOrdDesde + "',"
                    ZSql = ZSql + " OrdHasta = " + "'" + ZZOrdHasta + "'"
                    ZSql = ZSql + " Where Precios.Lista = " + "'" + WWWLista + "'"
                    ZSql = ZSql + " and Precios.LInea = " + "'" + WWWLinea + "'"
                    ZSql = ZSql + " and Precios.Tipo = " + "'" + WWWTipo + "'"
                    ZSql = ZSql + " and Precios.fragancia = " + "'" + WWWFragancia + "'"
                    ZSql = ZSql + " and Precios.Calidad = " + "'" + WWWCalidad + "'"
                    ZSql = ZSql + " and Precios.Tamano = " + "'" + WWWTamano + "'"
                    spPrecios = ZSql
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
                End If
        
            End If
                    
            
        End If
        
    Next Ciclo
End Sub

