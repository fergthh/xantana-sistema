VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgExpreso 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Expreso"
   ClientHeight    =   8070
   ClientLeft      =   930
   ClientTop       =   405
   ClientWidth     =   10095
   LinkTopic       =   "Form2"
   ScaleHeight     =   8070
   ScaleWidth      =   10095
   Begin VB.ComboBox Estado 
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
      TabIndex        =   54
      Text            =   " "
      Top             =   3360
      Width           =   3135
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
      MouseIcon       =   "expreso.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   3840
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
      Left            =   1200
      MouseIcon       =   "expreso.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Elimina el Registro"
      Top             =   3840
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
      Left            =   2160
      MouseIcon       =   "expreso.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3840
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
      Left            =   3120
      MouseIcon       =   "expreso.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Consulta de Datos"
      Top             =   3840
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
      Left            =   7920
      MouseIcon       =   "expreso.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Impresion "
      Top             =   3840
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
      Left            =   8880
      MouseIcon       =   "expreso.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Salida"
      Top             =   3840
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
      Left            =   4080
      MouseIcon       =   "expreso.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":46D2
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Primer Registro"
      Top             =   3840
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
      Left            =   5040
      MouseIcon       =   "expreso.frx":4B14
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":4E1E
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Registro Anterior"
      Top             =   3840
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
      Left            =   6000
      MouseIcon       =   "expreso.frx":5260
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":556A
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Registro Siguiente"
      Top             =   3840
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
      Left            =   6960
      MouseIcon       =   "expreso.frx":59AC
      MousePointer    =   99  'Custom
      Picture         =   "expreso.frx":5CB6
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Salida"
      Top             =   3840
      Width           =   855
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
      Top             =   5040
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
      Left            =   6960
      MaxLength       =   20
      TabIndex        =   36
      Text            =   " "
      Top             =   2640
      Width           =   2055
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
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   35
      Text            =   " "
      Top             =   2640
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
      Width           =   5775
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
      Top             =   120
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
         TabIndex        =   40
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
      Left            =   6960
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
      MaxLength       =   20
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
   Begin VB.TextBox Codigo 
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
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9480
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Expreso.rpt"
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
      Left            =   9120
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Nombre 
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
      Left            =   3000
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   5415
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
         Left            =   3240
         MouseIcon       =   "expreso.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "expreso.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   52
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
         Left            =   4200
         MouseIcon       =   "expreso.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "expreso.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Hasta 
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
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Desde 
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
         MaxLength       =   6
         TabIndex        =   11
         Text            =   " "
         Top             =   360
         Width           =   855
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
      Height          =   1620
      Left            =   2160
      TabIndex        =   24
      Top             =   5520
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
      Height          =   2460
      ItemData        =   "expreso.frx":6F90
      Left            =   120
      List            =   "expreso.frx":6F97
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   9735
   End
   Begin VB.Label Label8 
      Caption         =   "Estado"
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
      TabIndex        =   53
      Top             =   3360
      Width           =   1695
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
      Left            =   5520
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
      Left            =   120
      TabIndex        =   33
      Top             =   2640
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
      Caption         =   "Cuit / DNI"
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
      Left            =   5520
      TabIndex        =   18
      Top             =   1920
      Width           =   975
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1815
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
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgExpreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Dim ResultadoCuit As String

Sub Imprime_Descripcion()
End Sub

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
         Codigo.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Descuento.Text = Pusing("###,###.##", Descuento.Text)
    Rem MInimo.Text = Pusing("###,###.##", MInimo.Text)
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Codigo.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
    
        Codigo.Text = rstExpreso!Codigo
        Nombre.Text = Trim(rstExpreso!Nombre)
        Direccion.Text = Trim(rstExpreso!Direccion)
        Localidad.Text = Trim(rstExpreso!Localidad)
        Postal.Text = Trim(rstExpreso!Postal)
        Telefono.Text = Trim(rstExpreso!Telefono)
        Observaciones.Text = Trim(rstExpreso!Observaciones)
        Cuit.Text = Trim(rstExpreso!Cuit)
        email.Text = Trim(rstExpreso!email)
        fax.Text = Trim(rstExpreso!fax)
        Iva1.Value = False
        Iva2.Value = False
        Iva3.Value = False
        Iva4.Value = False
        Iva5.Value = False
        Iva6.Value = False
        Provincia.ListIndex = Val(rstExpreso!Provincia)
        Estado.ListIndex = rstExpreso!Estado
        Select Case Val(rstExpreso!Iva)
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
        
        rstExpreso.Close
        Call Format_datos
        Call Imprime_Descripcion
    End If

End Sub

Private Sub Acepta_Click()

    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    
    If Val(Hasta.Text) = 0 And Val(Hasta.Text) = 0 Then
         Hasta.Text = "9999"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Expreso SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Expresos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Expreso.Codigo, Expreso.Nombre, Expreso.Direccion, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Expreso Expreso, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Expreso.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Expreso.Codigo >= " + Desde.Text + " AND " _
                + "Expreso.Codigo <= " + Hasta.Text
    
    Listado.GroupSelectionFormula = "{Expreso.Codigo} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{Expreso.Codigo} in " + Desde.Text + " to " + Hasta.Text
    
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

    If Trim(Cuit.Text) = "" Then
        m$ = "Se debe informar numero de cuit"
        a% = MsgBox(m$, 0, "Archivo de Expresos")
        Exit Sub
    End If

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
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Codigo.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        rstExpreso.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Expreso SET "
        ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "',"
        ZSql = ZSql + " Direccion = " + "'" + Direccion.Text + "',"
        ZSql = ZSql + " Localidad = " + "'" + Localidad.Text + "',"
        ZSql = ZSql + " Postal = " + "'" + Postal.Text + "',"
        ZSql = ZSql + " Telefono = " + "'" + Telefono.Text + "',"
        ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
        ZSql = ZSql + " Cuit = " + "'" + Cuit.Text + "',"
        ZSql = ZSql + " Email = " + "'" + email.Text + "',"
        ZSql = ZSql + " Fax = " + "'" + fax.Text + "',"
        ZSql = ZSql + " Provincia = " + "'" + Str$(Provincia.ListIndex) + "',"
        ZSql = ZSql + " Estado = " + "'" + Str$(Estado.ListIndex) + "',"
        ZSql = ZSql + " Iva = " + "'" + WIva + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        spExpreso = ZSql
        Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Expreso ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Telefono ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "Email ,"
        ZSql = ZSql + "Fax ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "Iva )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + Nombre.Text + "',"
        ZSql = ZSql + "'" + Direccion.Text + "',"
        ZSql = ZSql + "'" + Localidad.Text + "',"
        ZSql = ZSql + "'" + Postal.Text + "',"
        ZSql = ZSql + "'" + Telefono.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + Cuit.Text + "',"
        ZSql = ZSql + "'" + email.Text + "',"
        ZSql = ZSql + "'" + fax.Text + "',"
        ZSql = ZSql + "'" + Str$(Provincia.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Estado.ListIndex) + "',"
        ZSql = ZSql + "'" + WIva + "')"
        spExpreso = ZSql
        Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
        
    Call CmdLimpiar_Click
    Codigo.SetFocus
    
End Sub

Private Sub CmdDelete_Click()
    If Val(Codigo.Text) <> 0 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Expreso"
        ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Codigo.Text + "'"
        spExpreso = ZSql
        Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
        If rstExpreso.RecordCount > 0 Then
            rstExpreso.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE Expreso"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spExpreso = ZSql
                Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    email.Text = ""
    fax.Text = ""
    
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    Provincia.ListIndex = 0
    Estado.ListIndex = 0

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Expreso"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        rstExpreso.MoveLast
        ZUltimo = IIf(IsNull(rstExpreso!CodigoMayor), "0", rstExpreso!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstExpreso.Close
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgExpreso.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Direccion.SetFocus
    End If
    If KeyAscii = 27 Then
        Nombre.Text = ""
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
        Postal.SetFocus
    End If
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
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

Private Sub Cuit_KeyPress(KeyAscii As Integer)
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

Private Sub Telefono_KeyPress(KeyAscii As Integer)
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
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        fax.Text = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Expreso"
            ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Codigo.Text + "'"
            spExpreso = ZSql
            Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
            If rstExpreso.RecordCount > 0 Then
                rstExpreso.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
            Nombre.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    Opcion.Clear
    Opcion.AddItem "Expreso"
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
            ZSql = ZSql + " FROM Expreso"
            ZSql = ZSql + " Order by Expreso.Codigo"
            spExpreso = ZSql
            Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
            If rstExpreso.RecordCount > 0 Then
                With rstExpreso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstExpreso.Close
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
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    Codigo.Text = ""
    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    email.Text = ""
    fax.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    
    Provincia.Clear
    
    Provincia.AddItem "Capital Federal"
    Provincia.AddItem "Buenos Aires"
    Provincia.AddItem "Catamarca"
    Provincia.AddItem "Cordoba"
    Provincia.AddItem "Corrientes"
    Provincia.AddItem "Chaco"
    Provincia.AddItem "Chubut"
    Provincia.AddItem "Entre Rios"
    Provincia.AddItem "Formosa"
    Provincia.AddItem "Jujuy"
    Provincia.AddItem "La Pampa"
    Provincia.AddItem "La Rioja"
    Provincia.AddItem "Mendoza"
    Provincia.AddItem "Misiones"
    Provincia.AddItem "Neuquen"
    Provincia.AddItem "Rio Negro"
    Provincia.AddItem "Salta"
    Provincia.AddItem "San Juan"
    Provincia.AddItem "San Luis"
    Provincia.AddItem "Santa Cruz"
    Provincia.AddItem "Santa Fe"
    Provincia.AddItem "Santiago del Estero"
    Provincia.AddItem "Tucuman"
    Provincia.AddItem "Tierra del Fuego"
    Provincia.AddItem "Exterior"
    Provincia.AddItem ""

    Estado.Clear
    
    Estado.AddItem "Activo"
    Estado.AddItem "Inactivo"
    
    Estado.ListIndex = 0
    Provincia.ListIndex = 0

 
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Expreso"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        rstExpreso.MoveLast
        ZUltimo = IIf(IsNull(rstExpreso!CodigoMayor), "0", rstExpreso!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstExpreso.Close
    End If
    
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
            ZSql = ZSql + " FROM Expreso"
            ZSql = ZSql + " Where Expreso.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Expreso.Codigo"
            spExpreso = ZSql
            Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
            If rstExpreso.RecordCount > 0 Then
                With rstExpreso
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstExpreso!Codigo) + " " + rstExpreso!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstExpreso!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstExpreso.Close
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

    Opcion.Clear
    Opcion.AddItem "Expreso"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Estado_KeyDown(KeyCode As Integer, Shift As Integer)
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
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Expreso.Codigo"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        With rstExpreso
            .MoveLast
            Codigo.Text = rstExpreso!Codigo
        End With
        rstExpreso.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de Expresos")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM Expreso"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        rstExpreso.MoveFirst
        ZUltimo = IIf(IsNull(rstExpreso!CodigoMenor), "0", rstExpreso!CodigoMenor)
        Codigo.Text = ZUltimo
        rstExpreso.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Expreso"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        rstExpreso.MoveLast
        ZUltimo = IIf(IsNull(rstExpreso!CodigoMayor), "0", rstExpreso!CodigoMayor)
        Codigo.Text = ZUltimo
        rstExpreso.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Expreso.Codigo"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        With rstExpreso
            .MoveFirst
            Codigo.Text = rstExpreso!Codigo
        End With
        rstExpreso.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Expresos")
    End If
End Sub
































