VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form prgArticuloPrecio 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios de Articulos"
   ClientHeight    =   8130
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   11790
   Visible         =   0   'False
   Begin VB.TextBox WTitulo2 
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
      Index           =   10
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   5040
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Index           =   9
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Index           =   8
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   5520
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Index           =   7
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Index           =   6
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Index           =   5
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Index           =   4
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4800
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Index           =   3
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4920
      Width           =   375
   End
   Begin VB.ComboBox Activo 
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
      TabIndex        =   31
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox Lista 
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
      TabIndex        =   28
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Valor4 
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
      Left            =   8880
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   27
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Tope4 
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
      Left            =   7680
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   26
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Valor3 
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
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   24
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Tope3 
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
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   23
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Valor2 
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
      Left            =   4080
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   21
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Tope2 
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
      Left            =   2760
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   20
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Valor1 
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
      TabIndex        =   18
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Tope1 
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
      Left            =   240
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   17
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox WTitulo2 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox WTitulo2 
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox Tamano 
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
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   12
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Calidad 
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
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   11
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Fragancia 
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
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   10
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Tipo 
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
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   9
      Text            =   " "
      Top             =   0
      Width           =   975
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
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   8
      Top             =   720
      Width           =   5535
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
      MouseIcon       =   "articuloprecio.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "articuloprecio.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   3240
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
      Left            =   1200
      MouseIcon       =   "articuloprecio.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "articuloprecio.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Consulta de Datos"
      Top             =   3240
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
      Left            =   2280
      MouseIcon       =   "articuloprecio.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "articuloprecio.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Linea 
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
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   240
      Top             =   600
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
      Left            =   8400
      TabIndex        =   3
      Top             =   0
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   2
      Top             =   360
      Width           =   5535
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   4440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5953
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Desde 
      Height          =   285
      Left            =   8400
      TabIndex        =   45
      Top             =   3000
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
   Begin MSMask.MaskEdBox Hasta 
      Height          =   285
      Left            =   8400
      TabIndex        =   46
      Top             =   3480
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
   Begin VB.Label Label11 
      Caption         =   "Hasta"
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
      TabIndex        =   48
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Desde"
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
      TabIndex        =   47
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Left            =   9000
      TabIndex        =   36
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   6480
      TabIndex        =   35
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Left            =   4080
      TabIndex        =   34
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
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
      Left            =   1440
      TabIndex        =   33
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label41 
      Caption         =   "Activo"
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
      Left            =   5280
      TabIndex        =   32
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   10200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label DesLista 
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
      TabIndex        =   30
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label7 
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
      Left            =   240
      TabIndex        =   29
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   10200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label9 
      Caption         =   "Hasta"
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
      TabIndex        =   25
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Hasta"
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
      Left            =   5280
      TabIndex        =   22
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Hasta"
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
      Left            =   2880
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Hasta"
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
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   10200
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1935
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
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "prgArticuloPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZPrecio  As Double
Dim ZZMargen As Double
Dim ZZFoto As Image
Dim ZZTextil As Integer
Dim ZZCodAnt As String

Dim WMovi(20000, 3) As String


Sub Imprime_Descripcion()

End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
End Sub

Sub Imprime_Datos()
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.LInea = " + "'" + Linea.Text + "'"
    ZSql = ZSql + " and Articulo.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and Articulo.fragancia = " + "'" + Fragancia.Text + "'"
    ZSql = ZSql + " and Articulo.Calidad = " + "'" + Calidad.Text + "'"
    ZSql = ZSql + " and Articulo.Tamano = " + "'" + Tamano.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        
        Descripcion.Text = Trim(rstArticulo!Descripcion)
        DescripcionII.Text = Trim(rstArticulo!DescripcionII)
        
        Activo.ListIndex = rstArticulo!Activo
        Rem FechaInactivo.Text = rstrticulo!FechaInactivo
        
        rstArticulo.Close
        Call Format_datos
        Call Imprime_Descripcion
    End If
    
End Sub


Private Sub cmdAdd_Click()

    Call Verifica_datos
    
    ZZCodigo = Trim(Linea.Text) + "-" + Trim(Tipo.Text) + "-" + Trim(Fragancia.Text) + "-" + Trim(Calidad.Text) + "-" + Trim(Tamano.Text)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Lista = " + "'" + Lista.Text + "'"
    ZSql = ZSql + " and Precios.LInea = " + "'" + Linea.Text + "'"
    ZSql = ZSql + " and Precios.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and Precios.fragancia = " + "'" + Fragancia.Text + "'"
    ZSql = ZSql + " and Precios.Calidad = " + "'" + Calidad.Text + "'"
    ZSql = ZSql + " and Precios.Tamano = " + "'" + Tamano.Text + "'"
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
        ZZTope1 = Str$(rstPrecios!Tope1)
        ZZValor1 = Str$(rstPrecios!Valor1)
        ZZTope2 = Str$(rstPrecios!Tope2)
        ZZValor2 = Str$(rstPrecios!Valor2)
        ZZTope3 = Str$(rstPrecios!Tope3)
        ZZValor3 = Str$(rstPrecios!Valor3)
        ZZTope4 = Str$(rstPrecios!Tope4)
        ZZValor4 = Str$(rstPrecios!Valor4)
        rstPrecios.Close

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
        ZSql = ZSql + "'" + ZZTope1 + "',"
        ZSql = ZSql + "'" + ZZValor1 + "',"
        ZSql = ZSql + "'" + ZZTope2 + "',"
        ZSql = ZSql + "'" + ZZValor2 + "',"
        ZSql = ZSql + "'" + ZZTope3 + "',"
        ZSql = ZSql + "'" + ZZValor3 + "',"
        ZSql = ZSql + "'" + ZZTope4 + "',"
        ZSql = ZSql + "'" + ZZValor4 + "')"
        spPreciosHistorial = ZSql
        Set rstPreciosHistorial = db.OpenRecordset(spPreciosHistorial, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    Call cmdClose_Click

End Sub

Private Sub cmdClose_Click()
    prgArticuloPrecio.Hide
    Unload Me
    prgArticulo.Show
End Sub


Private Sub Lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM lista"
        ZSql = ZSql + " Where lista.Codigo = " + "'" + Lista.Text + "'"
        spLista = ZSql
        Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstLista.RecordCount > 0 Then
            DesLista.Caption = rstLista!Descripcion
            rstLista.Close
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Precios"
            ZSql = ZSql + " Where Precios.Lista = " + "'" + Lista.Text + "'"
            ZSql = ZSql + " and Precios.LInea = " + "'" + Linea.Text + "'"
            ZSql = ZSql + " and Precios.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Precios.fragancia = " + "'" + Fragancia.Text + "'"
            ZSql = ZSql + " and Precios.Calidad = " + "'" + Calidad.Text + "'"
            ZSql = ZSql + " and Precios.Tamano = " + "'" + Tamano.Text + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Tope1.Text = Str$(rstPrecios!Tope1)
                Valor1.Text = Str$(rstPrecios!Valor1)
                Tope2.Text = Str$(rstPrecios!Tope2)
                Valor2.Text = Str$(rstPrecios!Valor2)
                Tope3.Text = Str$(rstPrecios!Tope3)
                Valor3.Text = Str$(rstPrecios!Valor3)
                Tope4.Text = Str$(rstPrecios!Tope4)
                Valor4.Text = Str$(rstPrecios!Valor4)
                Desde.Text = rstPrecios!Desde
                Hasta.Text = rstPrecios!Hasta
                rstPrecios.Close
                    Else
                Tope1.Text = ""
                Valor1.Text = ""
                Tope2.Text = ""
                Valor2.Text = ""
                Tope3.Text = ""
                Valor3.Text = ""
                Tope4.Text = ""
                Valor4.Text = ""
                Desde.Text = "  /  /    "
                Hasta.Text = "  /  /    "
            End If
            
            Tope1.Text = Pusing("###,###.##", Tope1.Text)
            Valor1.Text = Pusing("###,###.##", Valor1.Text)
            Tope2.Text = Pusing("###,###.##", Tope2.Text)
            Valor2.Text = Pusing("###,###.##", Valor2.Text)
            Tope3.Text = Pusing("###,###.##", Tope3.Text)
            Valor3.Text = Pusing("###,###.##", Valor3.Text)
            Tope4.Text = Pusing("###,###.##", Tope4.Text)
            Valor4.Text = Pusing("###,###.##", Valor4.Text)
            
            Call LeeHistorial
        
        End If
    End If
    If KeyAscii = 27 Then
        Lista.Text = ""
        DesLista.Caption = ""
        Tope1.Text = ""
        Valor1.Text = ""
        Tope2.Text = ""
        Valor2.Text = ""
        Tope3.Text = ""
        Valor3.Text = ""
        Tope4.Text = ""
        Valor4.Text = ""
        Call LeeHistorial
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Lista de Precios"

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
            ZSql = ZSql + " FROM Lista"
            ZSql = ZSql + " Order by Lista.Descripcion"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
            If rstLista.RecordCount > 0 Then
                With rstLista
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
                rstLista.Close
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
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Lista.Text = WIndice.List(Indice)
            Call Lista_KeyPress(13)
            
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    Linea.Text = ""
    Tipo.Text = ""
    Fragancia.Text = ""
    Calidad.Text = ""
    Tamano.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    Lista.Text = "0"
    DesLista.Caption = ""
    Rem FechaInactivo.Text = "  /  /    "
    
    
    Activo.Clear
    
    Activo.AddItem "Si"
    Activo.AddItem "No"
    
    Activo.ListIndex = 0

    Call Limpia_Vector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Lista"
    ZSql = ZSql + " Where Lista.Codigo = " + "'" + Lista.Text + "'"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstLista.RecordCount > 0 Then
        DesLista.Caption = Trim(rstLista!Descripcion)
        rstLista.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZPasaArticulo + "'"
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Linea.Text = rstArticulo!Linea
            Tipo.Text = rstArticulo!Tipo
            Fragancia.Text = rstArticulo!Fragancia
            Calidad.Text = rstArticulo!Calidad
            Tamano.Text = rstArticulo!Tamano
        End With
        rstArticulo.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.LInea = " + "'" + Linea.Text + "'"
    ZSql = ZSql + " and Articulo.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and Articulo.fragancia = " + "'" + Fragancia.Text + "'"
    ZSql = ZSql + " and Articulo.Calidad = " + "'" + Calidad.Text + "'"
    ZSql = ZSql + " and Articulo.Tamano = " + "'" + Tamano.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Descripcion.Text = Trim(rstArticulo!Descripcion)
        DescripcionII.Text = Trim(rstArticulo!DescripcionII)
        Activo.ListIndex = rstArticulo!Activo
        Rem FechaInactivo.Text = rstArticulo!FechaInactivo
        rstArticulo.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Lista = " + "'" + Lista.Text + "'"
    ZSql = ZSql + " and Precios.LInea = " + "'" + Linea.Text + "'"
    ZSql = ZSql + " and Precios.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and Precios.fragancia = " + "'" + Fragancia.Text + "'"
    ZSql = ZSql + " and Precios.Calidad = " + "'" + Calidad.Text + "'"
    ZSql = ZSql + " and Precios.Tamano = " + "'" + Tamano.Text + "'"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        Tope1.Text = Str$(rstPrecios!Tope1)
        Valor1.Text = Str$(rstPrecios!Valor1)
        Tope2.Text = Str$(rstPrecios!Tope2)
        Valor2.Text = Str$(rstPrecios!Valor2)
        Tope3.Text = Str$(rstPrecios!Tope3)
        Valor3.Text = Str$(rstPrecios!Valor3)
        Tope4.Text = Str$(rstPrecios!Tope4)
        Valor4.Text = Str$(rstPrecios!Valor4)
        Desde.Text = rstPrecios!Desde
        Hasta.Text = rstPrecios!Hasta
        rstPrecios.Close
            Else
        Tope1.Text = ""
        Valor1.Text = ""
        Tope2.Text = ""
        Valor2.Text = ""
        Tope3.Text = ""
        Valor3.Text = ""
        Tope4.Text = ""
        Valor4.Text = ""
        Desde.Text = "  /  /    "
        Hasta.Text = "  /  /    "
    End If
    
    Tope1.Text = Pusing("###,###.##", Tope1.Text)
    Valor1.Text = Pusing("###,###.##", Valor1.Text)
    Tope2.Text = Pusing("###,###.##", Tope2.Text)
    Valor2.Text = Pusing("###,###.##", Valor2.Text)
    Tope3.Text = Pusing("###,###.##", Tope3.Text)
    Valor3.Text = Pusing("###,###.##", Valor3.Text)
    Tope4.Text = Pusing("###,###.##", Tope4.Text)
    Valor4.Text = Pusing("###,###.##", Valor4.Text)
    
    Call LeeHistorial

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
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Where Lineas.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Lineas.Descripcion"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                With rstLinea
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
                rstLinea.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoPro"
            ZSql = ZSql + " Where TipoPro.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by TipoPro.Descripcion"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoPro.RecordCount > 0 Then
                With rstTipoPro
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
                rstTipoPro.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Fragancia"
            ZSql = ZSql + " Where Fragancia.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Fragancia.Descripcion"
            spFragancia = ZSql
            Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
            If rstFragancia.RecordCount > 0 Then
                With rstFragancia
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
                rstFragancia.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Calidad"
            ZSql = ZSql + " Where Calidad.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Calidad.Descripcion"
            spCalidad = ZSql
            Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
            If rstCalidad.RecordCount > 0 Then
                With rstCalidad
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
                rstCalidad.Close
            End If
            
        Case 4
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Tamano"
            ZSql = ZSql + " Where Tamano.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Tamano.Descripcion"
            spTamano = ZSql
            Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
            If rstTamano.RecordCount > 0 Then
                With rstTamano
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
                rstTamano.Close
            End If
            
        Case 5
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
                rstInsumo.Close
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

Private Sub LInea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Tipo_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Fragancia_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub Calidad_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 3
    
    Call Opcion_Click

End Sub

Private Sub Tamano_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 4
    
    Call Opcion_Click

End Sub

Private Sub Insumo_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Insumo"
    Opcion.AddItem "Lineas de Ventas"

    Opcion.ListIndex = 5
    
    Call Opcion_Click

End Sub

Private Sub Sector_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Insumo"
    Opcion.AddItem "Lineas de Ventas"

    Opcion.ListIndex = 6
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub LInea_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fragancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Calidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tamano_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Insumo_KeyDown(KeyCode As Integer, Shift As Integer)
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
End Sub

Private Sub Anterior_Click()
            
    WCodigo = Trim(Linea.Text) + "-" + Trim(Tipo.Text) + "-" + Trim(Fragancia.Text) + "-" + Trim(Calidad.Text) + "-" + Trim(Tamano.Text)
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo < " + "'" + WCodigo + "'"
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveLast
            Linea.Text = rstArticulo!Linea
            Tipo.Text = rstArticulo!Tipo
            Fragancia.Text = rstArticulo!Fragancia
            Calidad.Text = rstArticulo!Calidad
            Tamano.Text = rstArticulo!Tamano
        End With
        rstArticulo.Close
        Rem  Call Tamano_KeyPress(13)
        Linea.SetFocus
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
        rstArticulo.Close
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZUltimo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            Linea.Text = rstArticulo!Linea
            Tipo.Text = rstArticulo!Tipo
            Fragancia.Text = rstArticulo!Fragancia
            Calidad.Text = rstArticulo!Calidad
            Tamano.Text = rstArticulo!Tamano
            rstArticulo.Close
            rme Call Tamano_KeyPress(13)
            Call Imprime_Datos
        End If
        Linea.SetFocus
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

    WCodigo = Trim(Linea.Text) + "-" + Trim(Tipo.Text) + "-" + Trim(Fragancia.Text) + "-" + Trim(Calidad.Text) + "-" + Trim(Tamano.Text)
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo > " + "'" + WCodigo + "'"
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Linea.Text = rstArticulo!Linea
            Tipo.Text = rstArticulo!Tipo
            Fragancia.Text = rstArticulo!Fragancia
            Calidad.Text = rstArticulo!Calidad
            Tamano.Text = rstArticulo!Tamano
        End With
        rstArticulo.Close
        Call Tamano_KeyPress(13)
        Linea.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Articulos")
    End If
End Sub




Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 11
    WVector1.FixedRows = 1
    WVector1.Rows = 10001
    
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
                WVector1.Text = "Desde"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Hasta"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector1.Text = "Tope1"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector1.Text = "Valor1"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector1.Text = "Tope2"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector1.Text = "Valor2"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector1.Text = "Tope3"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 8
                WVector1.Text = "Valor3"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector1.Text = "Tope4"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WVector1.Text = "Valor5"
                WVector1.ColWidth(Ciclo) = 1050
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
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
    Rem modificar el Tamano de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub


Private Sub Busqueda()

    Rem On Error GoTo WError
    
    Call Limpia_Vector
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Descripcion <> ''"
    If Trim(Linea.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Linea = " + "'" + Linea.Text + "'"
    End If
    If Trim(Tipo.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Tipo = " + "'" + Tipo.Text + "'"
    End If
    If Trim(Fragancia.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Fragancia = " + "'" + Fragancia.Text + "'"
    End If
    If Trim(Calidad.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Calidad = " + "'" + Calidad.Text + "'"
    End If
    If Trim(Tamano.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Tamano = " + "'" + Tamano.Text + "'"
    End If
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    ZLugar = ZLugar + 1
                    WVector1.TextMatrix(ZLugar, 1) = !Codigo
                    WVector1.TextMatrix(ZLugar, 2) = !Descripcion
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If

End Sub


Private Sub WVector1_DblClick()

    WVector1.Col = 1
    ZZClave = WVector1.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZClave + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Linea.Text = rstArticulo!Linea
        Tipo.Text = rstArticulo!Tipo
        Fragancia.Text = rstArticulo!Fragancia
        Calidad.Text = rstArticulo!Calidad
        Tamano.Text = rstArticulo!Tamano
        rstArticulo.Close
        Call Tamano_KeyPress(13)
        Linea.SetFocus
    End If
    
End Sub

Private Sub WVector1_Click()

    WVector1.Col = 1
    ZZClave = WVector1.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZClave + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Linea.Text = rstArticulo!Linea
        Tipo.Text = rstArticulo!Tipo
        Fragancia.Text = rstArticulo!Fragancia
        Calidad.Text = rstArticulo!Calidad
        Tamano.Text = rstArticulo!Tamano
        rstArticulo.Close
        Call Tamano_KeyPress(13)
        Linea.SetFocus
    End If
    
End Sub


Private Sub LeeHistorial()
        
    Renglon = 0
    Call Limpia_Vector
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PreciosHistorial"
    ZSql = ZSql + " Where PreciosHistorial.lista = " + "'" + Lista.Text + "'"
    ZSql = ZSql + " and PreciosHistorial.Linea = " + "'" + Linea.Text + "'"
    ZSql = ZSql + " and PreciosHistorial.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and PreciosHistorial.Fragancia = " + "'" + Fragancia.Text + "'"
    ZSql = ZSql + " and PreciosHistorial.calidad = " + "'" + Calidad.Text + "'"
    ZSql = ZSql + " and PreciosHistorial.Tamano = " + "'" + Tamano.Text + "'"
    ZSql = ZSql + " Order by PreciosHistorial.OrdDesde"
        
    spPreciosHistorial = ZSql
    Set rstPreciosHistorial = db.OpenRecordset(spPreciosHistorial, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosHistorial.RecordCount > 0 Then
    
        With rstPreciosHistorial
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                    WVector1.TextMatrix(Renglon, 1) = rstPreciosHistorial!Desde
                    WVector1.TextMatrix(Renglon, 2) = rstPreciosHistorial!Hasta
                    WVector1.TextMatrix(Renglon, 3) = Pusing("###,###.##", Str$(rstPreciosHistorial!Tope1))
                    WVector1.TextMatrix(Renglon, 4) = Pusing("###,###.##", Str$(rstPreciosHistorial!Valor1))
                    WVector1.TextMatrix(Renglon, 5) = Pusing("###,###.##", Str$(rstPreciosHistorial!Tope2))
                    WVector1.TextMatrix(Renglon, 6) = Pusing("###,###.##", Str$(rstPreciosHistorial!Valor2))
                    WVector1.TextMatrix(Renglon, 7) = Pusing("###,###.##", Str$(rstPreciosHistorial!Tope3))
                    WVector1.TextMatrix(Renglon, 8) = Pusing("###,###.##", Str$(rstPreciosHistorial!Valor3))
                    WVector1.TextMatrix(Renglon, 9) = Pusing("###,###.##", Str$(rstPreciosHistorial!Tope4))
                    WVector1.TextMatrix(Renglon, 10) = Pusing("###,###.##", Str$(rstPreciosHistorial!Valor4))
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstPreciosHistorial.Close
    End If
End Sub


