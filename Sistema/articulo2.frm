VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BEC61919-E6C4-11D1-BE7D-C63815000000}#1.0#0"; "FLEXWIZ.OCX"
Begin VB.Form prgArticulo2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Articulos"
   ClientHeight    =   8250
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   12210
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   12210
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   360
      TabIndex        =   13
      Top             =   240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "articulo2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "frameConsulta"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Costos"
      TabPicture(1)   =   "articulo2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "SubWizard1"
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Precios e Impuestos"
      TabPicture(2)   =   "articulo2.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "btnAsignarLista"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "FrameListaPrecios"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).ControlCount=   4
      Begin VB.Frame frameConsulta 
         Caption         =   "Consulta"
         Height          =   4215
         Left            =   -71640
         TabIndex        =   50
         Top             =   960
         Visible         =   0   'False
         Width           =   4815
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
            Height          =   2940
            ItemData        =   "articulo2.frx":0054
            Left            =   240
            List            =   "articulo2.frx":005E
            TabIndex        =   56
            Top             =   480
            Width           =   4335
         End
         Begin VB.ListBox PantallaFiltrada 
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
            ItemData        =   "articulo2.frx":0075
            Left            =   240
            List            =   "articulo2.frx":007C
            TabIndex        =   60
            Top             =   840
            Visible         =   0   'False
            Width           =   4335
         End
         Begin VB.CommandButton btnCerrarConsulta 
            Caption         =   "Cerrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1800
            TabIndex        =   51
            Top             =   3480
            Width           =   1215
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
            Left            =   4080
            MouseIcon       =   "articulo2.frx":008A
            MousePointer    =   99  'Custom
            Picture         =   "articulo2.frx":0394
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Pedidos de Clientes"
            Top             =   3480
            Visible         =   0   'False
            Width           =   495
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
            TabIndex        =   58
            Top             =   480
            Visible         =   0   'False
            Width           =   4335
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
            ItemData        =   "articulo2.frx":0C5E
            Left            =   240
            List            =   "articulo2.frx":0C65
            TabIndex        =   57
            Top             =   840
            Visible         =   0   'False
            Width           =   4335
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
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   55
            Text            =   " "
            Top             =   3480
            Visible         =   0   'False
            Width           =   255
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
            Left            =   3075
            MaxLength       =   10
            TabIndex        =   54
            Text            =   " "
            Top             =   3480
            Visible         =   0   'False
            Width           =   180
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
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   53
            Text            =   " "
            Top             =   3480
            Visible         =   0   'False
            Width           =   255
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
            Left            =   3720
            MaxLength       =   10
            TabIndex        =   52
            Text            =   " "
            Top             =   3480
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin VB.Frame FrameListaPrecios 
         Caption         =   "Listas de Precios Disponibles"
         Height          =   5175
         Left            =   6240
         TabIndex        =   47
         Top             =   600
         Visible         =   0   'False
         Width           =   4455
         Begin VB.CommandButton btnCerrarListasPrecios 
            Caption         =   "Cerrar Listas"
            Height          =   495
            Left            =   1680
            TabIndex        =   49
            Top             =   4560
            Width           =   1215
         End
         Begin VB.ListBox ListasPrecios 
            Height          =   3960
            Left            =   240
            TabIndex        =   48
            Top             =   480
            Width           =   3975
         End
      End
      Begin MSFlexGridWizard.SubWizard SubWizard1 
         Height          =   1455
         Left            =   -71040
         TabIndex        =   35
         Top             =   6720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2566
      End
      Begin VB.Frame Frame7 
         Caption         =   "Proveedores Relacionados"
         Height          =   2295
         Left            =   -73560
         TabIndex        =   34
         Top             =   2160
         Width           =   8655
         Begin MSFlexGridLib.MSFlexGrid ProveedoresRelacionados 
            Height          =   1815
            Left            =   240
            TabIndex        =   36
            Top             =   360
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3201
            _Version        =   393216
            OLEDropMode     =   1
         End
      End
      Begin VB.CommandButton btnAsignarLista 
         Caption         =   "ASIGNAR A LISTA DE PRECIOS"
         Height          =   495
         Left            =   6720
         TabIndex        =   33
         Top             =   4980
         Width           =   3375
      End
      Begin VB.Frame Frame5 
         Caption         =   "I.V.A"
         Height          =   1455
         Left            =   1440
         TabIndex        =   30
         Top             =   540
         Width           =   8655
         Begin VB.ComboBox cmbIva 
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
            Left            =   3480
            TabIndex        =   31
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "IVA (%):"
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
            Left            =   3960
            TabIndex        =   32
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Precios de Venta"
         Height          =   2655
         Left            =   1440
         TabIndex        =   29
         Top             =   2220
         Width           =   8655
         Begin MSMask.MaskEdBox WTexto3 
            Height          =   375
            Left            =   6120
            TabIndex        =   38
            Top             =   840
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   3480
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   46
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   3
            Left            =   4680
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   45
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   4080
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   44
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   2880
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   43
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   2280
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   42
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTexto2 
            Height          =   375
            Left            =   7080
            TabIndex        =   41
            Top             =   840
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTexto1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox WCombo1 
            Height          =   315
            Left            =   7560
            TabIndex        =   39
            Text            =   "Combo2"
            Top             =   480
            Visible         =   0   'False
            Width           =   390
         End
         Begin MSFlexGridLib.MSFlexGrid Wvector1 
            Height          =   2055
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   8175
            _ExtentX        =   14420
            _ExtentY        =   3625
            _Version        =   393216
            OLEDropMode     =   1
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Costo Actual (En Dolares)"
         Height          =   1095
         Left            =   -73680
         TabIndex        =   18
         Top             =   720
         Width           =   8655
         Begin VB.TextBox Costo 
            Height          =   285
            Left            =   3600
            TabIndex        =   19
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "Costo:"
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
            Left            =   2520
            TabIndex        =   20
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Observaciones"
         Height          =   2175
         Left            =   -74400
         TabIndex        =   15
         Top             =   2400
         Width           =   10215
         Begin VB.TextBox Observaciones 
            Height          =   1575
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   360
            Width           =   7455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "General"
         Height          =   1455
         Left            =   -74400
         TabIndex        =   14
         Top             =   720
         Width           =   10215
         Begin VB.TextBox Rubro 
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
            Left            =   5670
            MaxLength       =   8
            TabIndex        =   27
            Top             =   840
            Width           =   735
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
            Left            =   4920
            MaxLength       =   50
            TabIndex        =   23
            Top             =   360
            Width           =   5055
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
            Left            =   2040
            MaxLength       =   25
            TabIndex        =   21
            Top             =   840
            Width           =   2895
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
            Left            =   2040
            MaxLength       =   10
            TabIndex        =   16
            Text            =   " "
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label DesRubro 
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
            Height          =   285
            Left            =   6480
            TabIndex        =   26
            Top             =   840
            Width           =   3495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Rubro"
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
            Left            =   5040
            TabIndex        =   25
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Descripcion Corta"
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
            Left            =   360
            TabIndex        =   24
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Caption         =   "Descripcion"
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
            Left            =   3720
            TabIndex        =   22
            Top             =   360
            Width           =   1095
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
            Left            =   1200
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
      End
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consultar (F4)"
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
      MouseIcon       =   "articulo2.frx":0C73
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":0F7D
      TabIndex        =   12
      ToolTipText     =   "Consulta de Datos"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar (F3)"
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
      MouseIcon       =   "articulo2.frx":17BF
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":1AC9
      TabIndex        =   11
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borrar (F2)"
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
      Left            =   2400
      MouseIcon       =   "articulo2.frx":230B
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":2615
      TabIndex        =   10
      ToolTipText     =   "Elimina el Registro"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Grabar (F1)"
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
      Left            =   1320
      MouseIcon       =   "articulo2.frx":2E57
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":3161
      TabIndex        =   9
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   9360
      TabIndex        =   8
      Top             =   7800
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   10320
      TabIndex        =   7
      Top             =   7560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu (F10)"
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
      Left            =   9960
      MouseIcon       =   "articulo2.frx":39A3
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":3CAD
      TabIndex        =   6
      ToolTipText     =   "Salida"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primer (F5)"
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
      Left            =   5640
      MouseIcon       =   "articulo2.frx":44EF
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":47F9
      TabIndex        =   5
      ToolTipText     =   "Primer Registro"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "Anterior (F6)"
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
      Left            =   6720
      MouseIcon       =   "articulo2.frx":4C3B
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":4F45
      TabIndex        =   4
      ToolTipText     =   "Registro Anterior"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Siguiente 
      Caption         =   "Siguiente (F7)"
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
      MouseIcon       =   "articulo2.frx":5387
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":5691
      TabIndex        =   3
      ToolTipText     =   "Registro Siguiente"
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo (F8)"
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
      MouseIcon       =   "articulo2.frx":5AD3
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":5DDD
      TabIndex        =   2
      ToolTipText     =   "Salida"
      Top             =   6960
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   7200
      TabIndex        =   1
      Top             =   8040
      Width           =   1935
   End
End
Attribute VB_Name = "prgArticulo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZPrecio  As Double
Dim ZZMargen As Double
Dim ZZFoto As Image
Dim ZZTextil As Integer
Dim ZZCodAnt As String

Dim ZStockI As Double
Dim ZStockII As Double
Dim ZStockIII As Double
Dim ZStockIV As Double
Dim ZStockV As Double
Dim ZStockVI As Double

Private WFecha As String
Private WPlazo1 As Integer
Private WVencimiento As String

Dim WMovi(20000, 3) As String
Dim WIva(4) As String

Dim mRow As Integer
Dim mCol As Integer
Dim mColSel As Integer
Dim mText As String

' CONTROLES PARA GRILLA
Private WParametros(4, 5)
Private WFormato(100) As String
Private WControl As String


Sub Imprime_Descripcion()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Sector"
    ZSql = ZSql + " Where Sector.Codigo= " + "'" + Sector.Text + "'"
    spSector = ZSql
    Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
    If rstSector.RecordCount > 0 Then
        DesSector.Caption = rstSector!Descripcion
        rstSector.Close
            Else
        DesSector.Caption = ""
    End If
    
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
        Moneda.ListIndex = rstPrecios!Moneda
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
        Moneda.ListIndex = 0
    End If
    
    Tope1.Text = Pusing("###,###,###.##", Tope1.Text)
    Valor1.Text = Pusing("###,###,###.##", Valor1.Text)
    Tope2.Text = Pusing("###,###,###.##", Tope2.Text)
    Valor2.Text = Pusing("###,###,###.##", Valor2.Text)
    Tope3.Text = Pusing("###,###,###.##", Tope3.Text)
    Valor3.Text = Pusing("###,###,###.##", Valor3.Text)
    Tope4.Text = Pusing("###,###,###.##", Tope4.Text)
    Valor4.Text = Pusing("###,###,###.##", Valor4.Text)
    
    Call LeeHistorial
    
End Sub

Private Function Verifica_datos() As Boolean
    Dim grabar As Boolean
    grabar = True
    
    If Trim(Codigo.Text) = "" Then grabar = False
    If Trim(Descripcion.Text) = "" Then grabar = False
    If Trim(DescripcionII.Text) = "" Then grabar = False
    If Trim(Rubro.Text) = "" Then grabar = False
    
    If Trim(Costo.Text) = "" Then grabar = False
    
    ' Hacer verificacion de valides de Rubro cuando este realizada esta parte.
    Auxi = Trim(Rubro.Text)
    
    Call Ceros(Auxi, 4)
    
    Rubro.Text = Auxi
    
    ZSql = ""
    ZSql = ZSql + "Select Codigo"
    ZSql = ZSql + " FROM TipoPro"
    ZSql = ZSql + " Where Codigo = '" + Rubro.Text + "'"
    spTipoPro = ZSql
    Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstTipoPro
        If .RecordCount > 0 Then
            .Close
        Else
            grabar = False
        End If
    End With
    
    If Val(cmbIva.ListIndex) < 0 Then grabar = False
    
    ' Recorremos la grilla y verificamos que haya datos en las columnas de neto y final.
    For i = 1 To WVector1.Rows
    
        With WVector1
            .Row = i
            .Col = 1
            If Trim(.Text) <> "" Then
            
                ' Chequeamos que haya datos en Neto
                .Col = 3
                
                If Val(.Text) <= 0 Then
                    grabar = False
                    Exit For
                End If
                
                ' Chequeamos que haya datos en Final
                .Col = 4
                
                If Val(.Text) <= 0 Then
                    grabar = False
                    Exit For
                End If
                
            Else
            
                Exit For
            
            End If
        
        End With
    
    Next
    
    
    Verifica_datos = grabar
    
End Function

Sub Format_datos()
End Sub

Sub Imprime_Datos()
    Call Rubro_KeyPress(13)
    ' Faltaria cargar los datos de la lista de precios. Cuando este terminada.
    
End Sub

Sub Imprime_Datos2()
    
    PantaPrecios.Visible = True
    PantaArticulo.Visible = False
    
    Call Limpia_Vector2
    
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
        Sector.Text = rstArticulo!Sector
            
        Activo.ListIndex = rstArticulo!Activo
        Facturable.ListIndex = rstArticulo!Facturable
        Etiqueta.ListIndex = rstArticulo!Etiqueta
        
        Stock.Text = Str$(rstArticulo!Stock)
        
        ZStockI = IIf(IsNull(rstArticulo!StockI), "0", rstArticulo!StockI)
        ZStockII = IIf(IsNull(rstArticulo!StockII), "0", rstArticulo!StockII)
        ZStockIII = IIf(IsNull(rstArticulo!StockIII), "0", rstArticulo!StockIII)
        ZStockIV = IIf(IsNull(rstArticulo!StockIV), "0", rstArticulo!StockIV)
        ZStockV = IIf(IsNull(rstArticulo!StockV), "0", rstArticulo!StockV)
        ZStockVI = IIf(IsNull(rstArticulo!StockVI), "0", rstArticulo!StockVI)
        
        StockI.Text = Str$(ZStockI)
        StockII.Text = Str$(ZStockII)
        StockIII.Text = Str$(ZStockIII)
        StockIV.Text = Str$(ZStockIV)
        StockV.Text = Str$(ZStockV)
        StockVI.Text = Str$(ZStockVI)
        
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
    
    Listado.Connect = Connect()
    
    Uno = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Dos = " and {Articulo.Linea} in " + Desde1.Text + " to " + Hasta1.Text
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
    Linea.SetFocus
End Sub

Private Sub Ayuda_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Ayuda.Text) <> "" Then
        Dim WTextoABuscar, WTexto
        
        WTextoABuscar = Trim(Ayuda.Text)
        
        PantallaFiltrada.Clear
        
        For i = 0 To Pantalla.ListCount
            WTexto = Pantalla.List(i)
            
            If WTexto Like "*" & WTextoABuscar & "*" Or WTexto Like "*" & UCase(WTextoABuscar) & "*" Then
                PantallaFiltrada.AddItem WTexto
            End If
        Next
        
        PantallaFiltrada.Visible = True
    Else
        PantallaFiltrada.Visible = False
    End If
End Sub

Private Sub btnAsignarLista_Click()

    ListasPrecios.Clear
    WIndice.Clear
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Lista"
    Rem ZSql = ZSql + " Where Sector.Codigo = " + "'" + Sector.Text + "'"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    With rstLista
        If .RecordCount > 0 Then
            .MoveFirst
            
            Do While .EOF = False And .BOF = False
                
                IngresaItem = !Codigo + " " + !Descripcion
                ListasPrecios.AddItem IngresaItem
                IngresaItem = !Codigo
                WIndice.AddItem IngresaItem
                
                .MoveNext
            
            Loop
            
        End If
    End With
    
    FrameListaPrecios.Visible = True
End Sub

Private Sub btnCerrarConsulta_Click()
    FrameConsulta.Visible = False
    Rubro.SetFocus
End Sub

Private Sub btnCerrarListasPrecios_Click()
    FrameListaPrecios.Visible = False
    WVector1.Row = 1
    WVector1.Col = 3
    WVector1.SetFocus
End Sub

Private Sub cmbIva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(cmbIva.ListIndex) <= 0 Then Exit Sub
        
        btnAsignarLista.SetFocus
        btnAsignarLista_Click
    End If
    If KeyAscii = 27 Then
        cmbIva.ListIndex = 0
    End If
End Sub

Private Sub cmdAdd_Click()

    Dim WCodigo, WDescripcion, WDescripcionII, WCosto, WCodigoIva, WObservaciones, WPasa As String
    Dim WRubro As Integer
    
    WCodigo = ""
    WDescripcion = ""
    WDescripcionII = ""
    WCosto = ""
    WCodigoIva = ""
    WObservaciones = ""
    WRubro = 0
    
    WPasa = "N"
    
    If Not Verifica_datos Then
        m$ = "Grabacion no se pudo realizar" & Chr(13) & "Hay datos que no son validos."
        aaaaaa% = MsgBox(m$, 0, "Alta de Articulos")
        Exit Sub
    End If
    
    WCodigo = Trim(Codigo.Text)
    WDescripcion = Left$(Trim(Descripcion.Text), 50)
    WDescripcionII = Left$(Trim(DescripcionII.Text), 20)
    WCosto = Val(Costo.Text)
    WCodigoIva = Left$(WIva(cmbIva.ListIndex), 1)
    WObservaciones = Left$(Trim(Observaciones.Text), 200)
    WRubro = Left$(Rubro.Text, 4) ' Verificar bien por el tema de que si seguira o no siendo alfanumerico.
    
    ZSql = ""
    ZSql = ZSql + "Select Codigo"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
    
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        With rstArticulo
            If .RecordCount > 0 Then
                .Close
                ' Actualizamos existente
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + "Descripcion = '" + WDescripcion + "',"
                ZSql = ZSql + "DescripcionII = '" + WDescripcionII + "',"
                ZSql = ZSql + "Costo = " + Str$(WCosto) + ","
                ZSql = ZSql + "Iva = '" + WCodigoIva + "',"
                ZSql = ZSql + "Observaciones = '" + WObservaciones + "',"
                ZSql = ZSql + "Rubro = " + Str$(WRubro) + " "
                ZSql = ZSql + " Where Codigo = '" + WCodigo + "'"
                spArticulo = ZSql
                MsgBox ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                WPasa = "S"
                
            Else
                ' Damos de alta.
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Articulo "
                ZSql = ZSql + "(Codigo, Descripcion, DescripcionII, Costo, Iva, Observaciones, Rubro) "
                ZSql = ZSql + "VALUES ("
                ZSql = ZSql + "'" + WCodigo + "',"
                ZSql = ZSql + "'" + WDescripcion + "',"
                ZSql = ZSql + "'" + WDescripcionII + "',"
                ZSql = ZSql + "'" + WCosto + "',"
                ZSql = ZSql + "'" + WCodigoIva + "',"
                ZSql = ZSql + "'" + WObservaciones + "',"
                ZSql = ZSql + "'" + Str$(WRubro) + "'"
                ZSql = ZSql + ")"
                
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                WPasa = "S"
            End If
        End With
        
        If WPasa = "S" Then
            Call Actualizar_Lista_Precios
        End If
        
    
    m$ = "Grabacion realizada"
    aaaaaa% = MsgBox(m$, 0, "Alta de Articulos")
    
    Call CmdLimpiar_Click
End Sub

Private Sub Actualizar_Lista_Precios()
    Dim WLista, WArticulo, WNeto, WPrecio, WClave, WRenglon, XRenglon
    
    WArticulo = Trim(Codigo.Text)
    WRenglon = 1
    
    ' Borramos la informacion anterior en caso de que hayan, asi no tendremos problemas en los casos en que se elimine alguna lista.
    
    ZSql = ""
    ZSql = ZSql + "DELETE FROM ListaArticulos WHERE Articulo = '" + WArticulo + "'"
    
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    ' Recorremos la grilla y verificamos que haya datos en las columnas de neto y final.
    For i = 1 To WVector1.Rows
    
        WLista = ""
        WNeto = 0
        WPrecio = 0
        WClave = ""
        
        With WVector1
            .Row = i
            .Col = 1
            If Trim(.Text) <> "" Then
            
                .Col = 1
                WLista = Trim(.Text)
                
                .Col = 3
                WNeto = Val(.Text)
                
                .Col = 4
                WPrecio = Val(.Text)
                
                ' Una vez guardada la informacion, la guardamos.
                Auxi = WLista
                Call Ceros(Auxi, 4) ' Solo en caso en que sean numericos las claves de las listas.
                WLista = Auxi
                
                XRenglon = Str$(WRenglon)
                
                Auxi = XRenglon
                Call Ceros(Auxi, 2)
                XRenglon = Auxi
                
                WClave = WArticulo + WLista + XRenglon
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO ListaArticulos "
                ZSql = ZSql + "(Clave, Articulo, Lista, Renglon, Neto, Precio) "
                ZSql = ZSql + "VALUES "
                ZSql = ZSql + "('" + WClave + "','" + WArticulo + "','" + WLista + "','" + XRenglon + "'," + Str$(WNeto) + "," + Str$(WPrecio) + ")"
                
                spLista = ZSql
                Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
                
                WRenglon = WRenglon + 1
                
            Else
            
                Exit For
            
            End If
        
        End With
    
    Next
End Sub

Private Sub cmdDelete_Click()

    If Trim(Codigo.Text) = "" Then Exit Sub

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(Codigo.Text) + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
        T$ = "Borrar Registro"
        m$ = "Desea Borrar el Registro "
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
        
            ZSql = ""
            ZSql = ZSql + "DELETE Articulo"
            ZSql = ZSql + " Where Codigo = " + "'" + Trim(Codigo.Text) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            Call CmdLimpiar_Click
            
        End If
    End If
    'Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    
    Erase WIva
    
    WIva(1) = "21.0"
    WIva(2) = "27.0"
    WIva(3) = "10.5"
    WIva(4) = "0"
    
    With cmbIva
    
        .Clear
        
        .AddItem ""
        .AddItem "% 21.0"
        .AddItem "% 27.0"
        .AddItem "% 10.5"
        .AddItem "% 0     (Exento)"
        
        .ListIndex = 0
    
    End With
    
    FrameConsulta.Visible = False
    FrameListaPrecios.Visible = False
    
    Codigo.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    Observaciones.Text = ""
    Rubro.Text = ""
    Costo.Text = ""
    
    SSTab1.Tab = 0
        
    If Codigo.Visible Then Codigo.SetFocus
    
    Call Limpia_Vector
    
End Sub

Private Sub CmdClose_Click()
    prgArticulo2.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If Trim(Codigo.Text) = "" Then Exit Sub
        
        ZSql = ""
        ZSql = ZSql + "SELECT * FROM Articulo Where Codigo = '" + Trim(Codigo.Text) + "'"
        
        spArticulo = ZSql
        
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        With rstArticulo
            If .RecordCount > 0 Then
                .MoveFirst
                Codigo.Text = Trim(!Codigo)
                Descripcion.Text = Trim(!Descripcion)
                DescripcionII.Text = Trim(!DescripcionII)
                Rubro.Text = Trim(!Rubro)
                
                Observaciones.Text = Trim(!Observaciones)
                
                Costo.Text = Pusing("######.##", Trim(!Costo))
                
                'Call Imprime_Datos
                
                Codigo.SetFocus
                
                cmbIva.ListIndex = Val(!Iva)
                
                .Close
            Else
                Descripcion.SetFocus
            End If
        End With
        
    ElseIf KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Command1_Click()

        ZSql = ""
        ZSql = ZSql + "UPDATE Precios SET "
        ZSql = ZSql + " Hasta = " + "'" + "26/09/2014" + "',"
        ZSql = ZSql + " OrdHasta = " + "'" + "20140926" + "'"
        ZSql = ZSql + " Where Precios.Hasta = " + "'" + "31/12/2099" + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)


End Sub

Private Sub Command2_Click()
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " StockI = Stock - StockII - StockIII - StockIV- StockV - StockVI"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)


Stop

    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " StockI = Stock " + ","
    ZSql = ZSql + " StockII = 0 " + ","
    ZSql = ZSql + " StockIII = 0 " + ","
    ZSql = ZSql + " StockIV = 0 " + ","
    ZSql = ZSql + " StockV = 0 " + ","
    ZSql = ZSql + " StockVI = 0 "
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        With DBGrid1
            
            Select Case .Col
            
            Case 3
                
            
            End Select
            
        End With
        
    ElseIf KeyAscii = 27 Then
        DBGrid1.Text = ""
    End If
    

End Sub

Private Sub Costo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(Costo.Text) = 0 Then Exit Sub
        
        Costo.Text = Pusing("######.##", Trim(Costo.Text))
        
        SSTab1.Tab = 2
        cmbIva.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionII.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub DescripcionII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rubro.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionII.Text = ""
    End If
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
            
            Tope1.Text = Pusing("###,###,###.##", Tope1.Text)
            Valor1.Text = Pusing("###,###,###.##", Valor1.Text)
            Tope2.Text = Pusing("###,###,###.##", Tope2.Text)
            Valor2.Text = Pusing("###,###,###.##", Valor2.Text)
            Tope3.Text = Pusing("###,###,###.##", Tope3.Text)
            Valor3.Text = Pusing("###,###,###.##", Valor3.Text)
            Tope4.Text = Pusing("###,###,###.##", Tope4.Text)
            Valor4.Text = Pusing("###,###,###.##", Valor4.Text)
            
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

Private Sub TipoBusqueda_Click()
    Call Busqueda
End Sub

Private Sub Tope1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tope1.Text = Pusing("###,###,###.##", Tope1.Text)
        Valor1.SetFocus
    End If
    If KeyAscii = 27 Then
        Tope1.Text = ""
    End If
End Sub

Private Sub Valor1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor1.Text = Pusing("###,###,###.##", Valor1.Text)
        Tope2.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor1.Text = ""
    End If
End Sub

Private Sub Tope2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tope2.Text = Pusing("###,###,###.##", Tope2.Text)
        Valor2.SetFocus
    End If
    If KeyAscii = 27 Then
        Tope2.Text = ""
    End If
End Sub

Private Sub Valor2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor2.Text = Pusing("###,###,###.##", Valor2.Text)
        Tope3.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor2.Text = ""
    End If
End Sub

Private Sub Tope3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tope3.Text = Pusing("###,###,###.##", Tope3.Text)
        Valor3.SetFocus
    End If
    If KeyAscii = 27 Then
        Tope3.Text = ""
    End If
End Sub

Private Sub Valor3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor3.Text = Pusing("###,###,###.##", Valor3.Text)
        Tope4.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor3.Text = ""
    End If
End Sub

Private Sub Tope4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tope4.Text = Pusing("###,###,###.##", Tope4.Text)
        Valor4.SetFocus
    End If
    If KeyAscii = 27 Then
        Tope4.Text = ""
    End If
End Sub

Private Sub Valor4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor4.Text = Pusing("###,###,###.##", Valor4.Text)
        Tope1.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor4.Text = ""
    End If
End Sub

Private Sub Sector_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Sector"
        ZSql = ZSql + " Where Sector.Codigo = " + "'" + Sector.Text + "'"
        spSector = ZSql
        Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        If rstSector.RecordCount > 0 Then
            DesSector.Caption = rstSector!Descripcion
            Descripcion.SetFocus
                Else
            DesSector.Caption = ""
            Sector.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Sector.Text = ""
        DesSector.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub LInea_KeyPress(KeyAscii As Integer)
    ' FALTA CONSULTAR EL FORMATO DE ENTRADA.
    If KeyAscii = 13 Then
        Linea.Text = UCase(Linea.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Lineas"
        ZSql = ZSql + " Where Lineas.Codigo = " + "'" + Linea.Text + "'"
        spLinea = ZSql
        Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstLinea.RecordCount > 0 Then
            rstLinea.Close
            Call Tipo_DblClick
            Call Busqueda
            'Tipo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Linea.Text = ""
        Call Busqueda
    End If
End Sub

Private Sub Form_Activate()

    Codigo.SetFocus

End Sub

Private Sub ListasPrecios_Click()
    Dim wrow As Integer
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Lista"
    ZSql = ZSql + " Where Codigo = " + "'" + Trim(WIndice.List(ListasPrecios.ListIndex)) + "'"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    With rstLista
        If .RecordCount > 0 Then
            .MoveFirst
            
            For wrow = 1 To 100
            
                WVector1.Row = wrow
                
                WVector1.Col = 1
                
                If WVector1.Text = Trim(!Codigo) Then
                    
                    .Close
                    Exit Sub
                
                End If
                
                If WVector1.Text = "" Then
                    
                    Exit For
                
                End If
            
            Next
            
            WVector1.Text = Trim(!Codigo)
            
            WVector1.Col = 2
            WVector1.Text = Trim(!Descripcion)
            
            .Close
        End If
    End With
    
End Sub

Private Sub PantallaFiltrada_Click()
    WIndice = PantallaFiltrada.ListIndex
    WTexto = PantallaFiltrada.List(PantallaFiltrada.ListIndex)
    For i = o To Pantalla.ListCount
        If UCase(Pantalla.List(i)) = UCase(WTexto) Then
            
            Pantalla.ListIndex = i
            
            Call Pantalla_Click
            
            PantallaFiltrada.Visible = False
            
            Exit Sub
        End If
    Next
End Sub

Private Sub Rubro_DblClick()
    Opcion.ListIndex = 1
    Opcion_Click
    
    FrameConsulta.Visible = True
End Sub

Private Sub Rubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Trim(Rubro.Text) = "" Then
            DesRubro.Caption = ""
            Call Consulta_Click
            
            Exit Sub
        End If
        
        ' Chequear por el tema de si sigue o no siendo Alfanumerico.
    
        ZSql = ""
        ZSql = ZSql + "Select Codigo, Descripcion"
        ZSql = ZSql + " FROM TipoPro"
        ZSql = ZSql + " Where Codigo = " + "'" + Trim(Rubro.Text) + "'"
        spTipoPro = ZSql
        Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
        
        With rstTipoPro
            If .RecordCount > 0 Then
                Rubro.Text = Trim(!Codigo)
                DesRubro.Caption = Trim(!Descripcion)
                
                Observaciones.SetFocus
                
                .Close
            Else
                Rubro.SetFocus
            End If
        End With
        
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        Call Busqueda
    End If
    ' Comentado hasta que se confirme si es o no alfanumerico.
    'Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Select Case SSTab1.Tab
        Case 0:
            If Codigo.Visible Then Codigo.SetFocus
        Case 1:
            If Costo.Visible Then Costo.SetFocus
        Case 2:
            If cmbIva.Visible Then cmbIva.SetFocus
    End Select
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tipo.Text = UCase(Tipo.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoPro"
        ZSql = ZSql + " Where TipoPro.Codigo = " + "'" + Tipo.Text + "'"
        spTipoPro = ZSql
        Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoPro.RecordCount > 0 Then
            rstTipoPro.Close
            Call Fragancia_DblClick
            Call Busqueda
            'Fragancia.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        Call Busqueda
    End If
End Sub

Private Sub Fragancia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fragancia.Text = UCase(Fragancia.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Fragancia"
        ZSql = ZSql + " Where Fragancia.Codigo = " + "'" + Fragancia.Text + "'"
        spFragancia = ZSql
        Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
        If rstFragancia.RecordCount > 0 Then
            rstFragancia.Close
            Call Calidad_DblClick
            Call Busqueda
            'Calidad.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fragancia.Text = ""
        Call Busqueda
    End If
End Sub

Private Sub Calidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Calidad.Text = UCase(Calidad.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM calidad"
        ZSql = ZSql + " Where calidad.Codigo = " + "'" + Calidad.Text + "'"
        spCalidad = ZSql
        Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
        If rstCalidad.RecordCount > 0 Then
            rstCalidad.Close
            Call Tamano_DblClick
            Call Busqueda
            'Tamano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Calidad.Text = ""
        Call Busqueda
    End If
End Sub

Private Sub Tamano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tamano.Text = UCase(Tamano.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Tamano"
        ZSql = ZSql + " Where Tamano.Codigo = " + "'" + Tamano.Text + "'"
        spTamano = ZSql
        Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
        If rstTamano.RecordCount > 0 Then
            rstTamano.Close
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
                rstArticulo.Close
                Call Imprime_Datos
                Descripcion.SetFocus
                    Else
                WLinea = Linea.Text
                WTipo = Tipo.Text
                WFragancia = Fragancia.Text
                WCalidad = Calidad.Text
                WTamano = Tamano.Text
                CmdLimpiar_Click
                Linea.Text = WLinea
                Tipo.Text = WTipo
                Fragancia.Text = WFragancia
                Calidad.Text = WCalidad
                Tamano.Text = WTamano
                Call Busqueda
            End If
            Descripcion.SetFocus
                Else
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
                rstArticulo.Close
                Call Imprime_Datos
                Descripcion.SetFocus
                    Else
                WLinea = Linea.Text
                WTipo = Tipo.Text
                WFragancia = Fragancia.Text
                WCalidad = Calidad.Text
                WTamano = Tamano.Text
                CmdLimpiar_Click
                Linea.Text = WLinea
                Tipo.Text = WTipo
                Fragancia.Text = WFragancia
                Calidad.Text = WCalidad
                Tamano.Text = WTamano
                Call Busqueda
            End If
            Descripcion.SetFocus
                
        End If
    End If
    If KeyAscii = 27 Then
        Tamano.Text = ""
        Call Busqueda
    End If
End Sub
    
Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            WFecha = Desde.Text
            WPlazo1 = 61
            Call Calcula_vencimiento(WFecha, WPlazo1, WVencimiento)
            Hasta.Text = WVencimiento
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Desde.Text = "  /  /    "
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()
    FrameConsulta.Visible = True
    Pantalla.Visible = False
    Ayuda.Text = ""
    Ayuda.Visible = False
    
    Opcion.Clear
    
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Rubros"
    
    Opcion.Visible = True
    
    'Opcion.ListIndex = 0
    
    'Call Opcion_Click
    
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
            ZSql = ZSql + " Order by Codigo"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                With rstLinea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Trim(!Codigo) + " " + Trim(!Descripcion)
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Trim(!Codigo)
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
            ZSql = ZSql + " Order by Descripcion"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                With rstLinea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Trim(!Codigo) + " " + Trim(!Descripcion)
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Trim(!Codigo)
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLinea.Close
            End If
        Case Else
    End Select
            
    FrameConsulta.Visible = True
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
    
    FrameConsulta.Visible = False
    
    Select Case XIndice
        Case 0
            indice = Pantalla.ListIndex
            Codigo.Text = WIndice.List(indice)
            Call Codigo_KeyPress(13)
                               
        Case 1
            indice = Pantalla.ListIndex
            Rubro.Text = WIndice.List(indice)
            Call Rubro_KeyPress(13)
                    
        Case Else
    End Select
    
End Sub


Sub Form_Load()
    
    Call CmdLimpiar_Click
    
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

Private Sub Sector_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem ""
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
        Call Tamano_KeyPress(13)
        Linea.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Articulos")
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
            Call Tamano_KeyPress(13)
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
        Rem Codigo.Text = ZUltimo
        rstArticulo.Close
        Call Imprime_Datos
        Rem Codigo.SetFocus
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
        aaaaaa% = MsgBox(m$, 0, "Archivo de Articulos")
    End If
End Sub

Private Sub Limpia_VectorII()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 5
    WVector1.FixedRows = 1
    WVector1.Rows = 10
    
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
                WVector1.Text = "ID"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 5000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 25
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Costo"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "######.##"
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

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 5
    WVector1.FixedRows = 1
    WVector1.Rows = 100
    
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
                WVector1.Text = "ID"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 5500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 25
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Costo"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "#####.##"
            Case 4
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "######.##"
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
    PantaPrecios.Visible = False
    PantaArticulo.Visible = True
    ZLugar = 0
    
    If Trim(Linea.Text) = "" And Trim(Tipo.Text) = "" And Trim(Fragancia.Text) = "" And Trim(Calidad.Text) = "" And Trim(Tamano.Text) = "" Then
        Exit Sub
    End If
    
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
                    
                    If TipoBusqueda.Value = 1 Or !Activo = 0 Then
                        ZLugar = ZLugar + 1
                        WVector1.TextMatrix(ZLugar, 1) = !Codigo
                        WVector1.TextMatrix(ZLugar, 2) = !Descripcion
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
    
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1

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


Private Sub Limpia_Vector2()

    WVector2.Clear

    Rem ponga la wvector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 11
    WVector2.FixedRows = 1
    WVector2.Rows = 10001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem wvector2.Text = "Articulo"
    
    Rem Longitud
    Rem wvector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem wvector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
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
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Desde"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector2.Text = "Hasta"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                WVector2.Text = "Tope1"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector2.Text = "Valor1"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 5
                WVector2.Text = "Tope2"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 6
                WVector2.Text = "Valor2"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 7
                WVector2.Text = "Tope3"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 8
                WVector2.Text = "Valor3"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 9
                WVector2.Text = "Tope4"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 10
                WVector2.Text = "Valor5"
                WVector2.ColWidth(Ciclo) = 1000
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
       End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WTitulo2(Ciclo).Text = WVector2.Text
        WTitulo2(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        WTitulo2(Ciclo).Top = WVector2.CellTop + WVector2.Top
        WTitulo2(Ciclo).Width = WVector2.CellWidth
        WTitulo2(Ciclo).Height = WVector2.CellHeight
        WTitulo2(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el Tamano de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub LeeHistorial()
        
    Renglon = 0
    Call Limpia_Vector2
        
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
                    WVector2.TextMatrix(Renglon, 1) = rstPreciosHistorial!Desde
                    WVector2.TextMatrix(Renglon, 2) = rstPreciosHistorial!Hasta
                    WVector2.TextMatrix(Renglon, 3) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Tope1))
                    WVector2.TextMatrix(Renglon, 4) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Valor1))
                    WVector2.TextMatrix(Renglon, 5) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Tope2))
                    WVector2.TextMatrix(Renglon, 6) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Valor2))
                    WVector2.TextMatrix(Renglon, 7) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Tope3))
                    WVector2.TextMatrix(Renglon, 8) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Valor3))
                    WVector2.TextMatrix(Renglon, 9) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Tope4))
                    WVector2.TextMatrix(Renglon, 10) = Pusing("###,###,###.##", Str$(rstPreciosHistorial!Valor4))
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstPreciosHistorial.Close
    End If
End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Mask = ""
                        WTexto3.Text = ""
                        WTexto3.Mask = "##/##/####"
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            If Trim(WTexto1.Text) = "" Then
                WVector1.Text = ""
            Else
                WVector1.Text = WTexto1.Text
            End If
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                If Trim(WTexto2.Text) = "" Then
                    WVector1.Text = ""
                Else
                    WVector1.Text = WTexto2.Text
                End If
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    If Trim(Replace(WTexto3.Text, "/", "")) = "" Then
                        WVector1.Text = ""
                    Else
                        WVector1.Text = WTexto3.Text
                    End If
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" And Trim(WVector1.Text) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
            'WVector1.Text = WVector1.Text
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            'If WControl = "S" Then
                Call Control_wvector1
            'End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
             '   If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
              '  End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
               ' If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                'End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WVector1.Row < WVector1.Rows - 1 Then
                'Call Control_Campo
                'If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                'End If
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Mask = ""
            WTexto3.Text = ""
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                'Call Control_Campo
                'If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                'End If
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
             '   If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
              '  End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
               ' If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
               ' End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 4
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            Rem WVector1.Col = 1
        Case Else
            Rem If WVector1.Col < WVector1.Cols - 1 Then
            Rem     WVector1.Col = WVector1.Col + 1
            Rem End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub
