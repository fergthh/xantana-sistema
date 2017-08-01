VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BEC61919-E6C4-11D1-BE7D-C63815000000}#1.0#0"; "FLEXWIZ.OCX"
Begin VB.Form prgArticulo2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Articulos"
   ClientHeight    =   10080
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   13860
   LinkTopic       =   "Form2"
   ScaleHeight     =   10080
   ScaleWidth      =   13860
   Visible         =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   720
      TabIndex        =   21
      Top             =   240
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11668
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "articulo2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Costos"
      TabPicture(1)   =   "articulo2.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "SubWizard1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Precios e Impuestos"
      TabPicture(2)   =   "articulo2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame6"
      Tab(2).Control(1)=   "Frame5"
      Tab(2).Control(2)=   "btnAsignarLista"
      Tab(2).ControlCount=   3
      Begin MSFlexGridWizard.SubWizard SubWizard1 
         Height          =   1455
         Left            =   3960
         TabIndex        =   44
         Top             =   6720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2566
      End
      Begin VB.Frame Frame7 
         Caption         =   "Proveedores Relacionados"
         Height          =   2295
         Left            =   1320
         TabIndex        =   42
         Top             =   2040
         Width           =   8655
         Begin MSMask.MaskEdBox WTexto3 
            Height          =   375
            Left            =   7800
            TabIndex        =   51
            Top             =   1080
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.ComboBox WCombo1 
            Height          =   315
            Left            =   8160
            TabIndex        =   54
            Text            =   "Combo2"
            Top             =   240
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.TextBox WTexto1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   53
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTexto2 
            Height          =   375
            Left            =   6840
            TabIndex        =   52
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   0
            Left            =   7920
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   720
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   1
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   49
            Top             =   600
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   2
            Left            =   7320
            Locked          =   -1  'True
            TabIndex        =   48
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
            Left            =   6720
            Locked          =   -1  'True
            TabIndex        =   47
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox WTitulo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Index           =   4
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   480
            Visible         =   0   'False
            Width           =   495
         End
         Begin MSFlexGridLib.MSFlexGrid WVector1 
            Height          =   1815
            Left            =   240
            TabIndex        =   45
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
         Left            =   -68280
         TabIndex        =   41
         Top             =   4980
         Width           =   3375
      End
      Begin VB.Frame Frame5 
         Caption         =   "I.V.A"
         Height          =   1455
         Left            =   -73560
         TabIndex        =   38
         Top             =   540
         Width           =   8655
         Begin VB.ComboBox Combo1 
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
            TabIndex        =   39
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
            TabIndex        =   40
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Precios de Venta"
         Height          =   2655
         Left            =   -73560
         TabIndex        =   37
         Top             =   2220
         Width           =   8655
         Begin MSDBGrid.DBGrid DBGrid1 
            Height          =   1935
            Left            =   360
            OleObjectBlob   =   "articulo2.frx":0054
            TabIndex        =   43
            Top             =   480
            Width           =   7935
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Costo Actual (En Dolares)"
         Height          =   1095
         Left            =   1320
         TabIndex        =   26
         Top             =   720
         Width           =   8655
         Begin VB.TextBox Costo 
            Height          =   285
            Left            =   3600
            TabIndex        =   27
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
            TabIndex        =   28
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Observaciones"
         Height          =   2175
         Left            =   -74400
         TabIndex        =   23
         Top             =   2400
         Width           =   10215
         Begin VB.TextBox Observaciones 
            Height          =   1575
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   360
            Width           =   7455
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "General"
         Height          =   1455
         Left            =   -74400
         TabIndex        =   22
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
            TabIndex        =   35
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
            TabIndex        =   31
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
            TabIndex        =   29
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
            TabIndex        =   24
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   30
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
            TabIndex        =   25
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
      Left            =   7320
      MouseIcon       =   "articulo2.frx":0D9F
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":10A9
      TabIndex        =   20
      ToolTipText     =   "Consulta de Datos"
      Top             =   8880
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
      Left            =   6240
      MouseIcon       =   "articulo2.frx":18EB
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":1BF5
      TabIndex        =   19
      ToolTipText     =   "Limpia la pantalla"
      Top             =   8880
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
      Left            =   5160
      MouseIcon       =   "articulo2.frx":2437
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":2741
      TabIndex        =   18
      ToolTipText     =   "Elimina el Registro"
      Top             =   8880
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
      Left            =   4080
      MouseIcon       =   "articulo2.frx":2F83
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":328D
      TabIndex        =   17
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   8880
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   600
      TabIndex        =   16
      Top             =   9240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1560
      TabIndex        =   15
      Top             =   9000
      Visible         =   0   'False
      Width           =   615
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
      Left            =   13080
      MouseIcon       =   "articulo2.frx":3ACF
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":3DD9
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Pedidos de Clientes"
      Top             =   7320
      Width           =   495
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
      Left            =   12720
      MouseIcon       =   "articulo2.frx":46A3
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":49AD
      TabIndex        =   9
      ToolTipText     =   "Salida"
      Top             =   8880
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
      Left            =   8400
      MouseIcon       =   "articulo2.frx":51EF
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":54F9
      TabIndex        =   8
      ToolTipText     =   "Primer Registro"
      Top             =   8880
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
      Left            =   9480
      MouseIcon       =   "articulo2.frx":593B
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":5C45
      TabIndex        =   7
      ToolTipText     =   "Registro Anterior"
      Top             =   8880
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
      Left            =   10560
      MouseIcon       =   "articulo2.frx":6087
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":6391
      TabIndex        =   6
      ToolTipText     =   "Registro Siguiente"
      Top             =   8880
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
      Left            =   11640
      MouseIcon       =   "articulo2.frx":67D3
      MousePointer    =   99  'Custom
      Picture         =   "articulo2.frx":6ADD
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   8880
      Width           =   975
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
      Left            =   9480
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   4335
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   13080
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   975
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
      ItemData        =   "articulo2.frx":6F1F
      Left            =   9480
      List            =   "articulo2.frx":6F26
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   4335
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
      Height          =   1740
      Left            =   9600
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   11640
      MaxLength       =   10
      TabIndex        =   11
      Text            =   " "
      Top             =   8040
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
      Left            =   11955
      MaxLength       =   10
      TabIndex        =   12
      Text            =   " "
      Top             =   8040
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
      Left            =   12240
      MaxLength       =   10
      TabIndex        =   13
      Text            =   " "
      Top             =   8040
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
      Left            =   12600
      MaxLength       =   10
      TabIndex        =   14
      Text            =   " "
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   9600
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

Sub Verifica_datos()
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
End Sub

Sub Imprime_Datos()
    ' Cargamos el rubro
    ZSql = ""
    ZSql = ZSql + "Select * FROM TipoArticulo WHERE Codigo = '" + Trim(Rubro.Text) + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    With rstArticulo
        If .RecordCount > 0 Then
            .MoveFirst
            DesRubro.Caption = IIf(IsNull(!Descripcion), "", Trim(!Descripcion))
            .Close
        Else
            DesRubro.Caption = ""
        End If
        
    End With
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

Private Sub cmdAdd_Click()

    If Trim(Sector.Text) <> "" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Sector"
        ZSql = ZSql + " Where Sector.Codigo = " + "'" + Sector.Text + "'"
        spSector = ZSql
        Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        If rstSector.RecordCount > 0 Then
            rstSector.Close
                Else
            m$ = "Codigo de sector inexistente"
            aaaaaa% = MsgBox(m$, 0, "Archivo de Articulos")
            Exit Sub
        End If
            Else
        m$ = "Se debe informar codigo de sector"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Articulos")
        Exit Sub
    End If


    Call Verifica_datos
    
    ZZCodigo = Trim(Linea.Text) + "-" + Trim(Tipo.Text) + "-" + Trim(Fragancia.Text) + "-" + Trim(Calidad.Text) + "-" + Trim(Tamano.Text)
    
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
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
        ZSql = ZSql + " DescripcionII = " + "'" + DescripcionII.Text + "',"
        ZSql = ZSql + " Stock = " + "'" + Stock.Text + "',"
        ZSql = ZSql + " StockI = " + "'" + StockI.Text + "',"
        ZSql = ZSql + " StockII = " + "'" + StockII.Text + "',"
        ZSql = ZSql + " StockIII = " + "'" + StockIII.Text + "',"
        ZSql = ZSql + " StockIV = " + "'" + StockIV.Text + "',"
        ZSql = ZSql + " StockV = " + "'" + StockV.Text + "',"
        ZSql = ZSql + " StockVI = " + "'" + StockVI.Text + "',"
        ZSql = ZSql + " Sector = " + "'" + Sector.Text + "',"
        ZSql = ZSql + " Activo = " + "'" + Str$(Activo.ListIndex) + "',"
        ZSql = ZSql + " FechaInactivo = " + "'" + FechaInactivo.Text + "',"
        ZSql = ZSql + " Facturable = " + "'" + Str$(Facturable.ListIndex) + "',"
        ZSql = ZSql + " Etiqueta = " + "'" + Str$(Etiqueta.ListIndex) + "'"
        ZSql = ZSql + " Where LInea = " + "'" + Linea.Text + "'"
        ZSql = ZSql + " and Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Fragancia = " + "'" + Fragancia.Text + "'"
        ZSql = ZSql + " and Calidad = " + "'" + Calidad.Text + "'"
        ZSql = ZSql + " and Tamano = " + "'" + Tamano.Text + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Articulo ("
        ZSql = ZSql + "Linea ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Fragancia ,"
        ZSql = ZSql + "Calidad ,"
        ZSql = ZSql + "Tamano ,"
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "DescripcionII ,"
        ZSql = ZSql + "Stock ,"
        ZSql = ZSql + "StockI ,"
        ZSql = ZSql + "StockII ,"
        ZSql = ZSql + "StockIII ,"
        ZSql = ZSql + "StockIV ,"
        ZSql = ZSql + "StockV ,"
        ZSql = ZSql + "StockVI ,"
        ZSql = ZSql + "Sector ,"
        ZSql = ZSql + "Activo ,"
        ZSql = ZSql + "FechaInactivo ,"
        ZSql = ZSql + "Facturable ,"
        ZSql = ZSql + "Etiqueta )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Linea.Text + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + Fragancia.Text + "',"
        ZSql = ZSql + "'" + Calidad.Text + "',"
        ZSql = ZSql + "'" + Tamano.Text + "',"
        ZSql = ZSql + "'" + ZZCodigo + "',"
        ZSql = ZSql + "'" + Descripcion.Text + "',"
        ZSql = ZSql + "'" + DescripcionII.Text + "',"
        ZSql = ZSql + "'" + Stock.Text + "',"
        ZSql = ZSql + "'" + StockI.Text + "',"
        ZSql = ZSql + "'" + StockII.Text + "',"
        ZSql = ZSql + "'" + StockIII.Text + "',"
        ZSql = ZSql + "'" + StockIV.Text + "',"
        ZSql = ZSql + "'" + StockV.Text + "',"
        ZSql = ZSql + "'" + StockVI.Text + "',"
        ZSql = ZSql + "'" + Sector.Text + "',"
        ZSql = ZSql + "'" + Str$(Activo.ListIndex) + "',"
        ZSql = ZSql + "'" + FechaInactivo.Text + "',"
        ZSql = ZSql + "'" + Str$(Facturable.ListIndex) + "',"
        ZSql = ZSql + "'" + Str$(Etiqueta.ListIndex) + "')"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
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
        
        If ZZTope1 <> Val(Tope1.Text) Or ZZValor1 <> Val(Valor1.Text) Or ZZTope2 <> Val(Tope2.Text) Or ZZValor2 <> Val(Valor2.Text) Or ZZTope3 <> Val(Tope3.Text) Or ZZValor3 <> Val(Valor3.Text) Or ZZTope4 <> Val(Tope4.Text) Or ZZValor4 <> Val(Valor4.Text) Then
            
            Desde.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
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

        End If
        
        ZZOrdDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
        ZZOrdHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Precios SET "
        ZSql = ZSql + " MOneda = " + "'" + Str$(Moneda.ListIndex) + "',"
        ZSql = ZSql + " Desde = " + "'" + Desde.Text + "',"
        ZSql = ZSql + " Hasta = " + "'" + Hasta.Text + "',"
        ZSql = ZSql + " OrdDesde = " + "'" + ZZOrdDesde + "',"
        ZSql = ZSql + " OrdHasta = " + "'" + ZZOrdHasta + "',"
        ZSql = ZSql + " Tope1 = " + "'" + Tope1.Text + "',"
        ZSql = ZSql + " Valor1 = " + "'" + Valor1.Text + "',"
        ZSql = ZSql + " Tope2 = " + "'" + Tope2.Text + "',"
        ZSql = ZSql + " Valor2 = " + "'" + Valor2.Text + "',"
        ZSql = ZSql + " Tope3 = " + "'" + Tope3.Text + "',"
        ZSql = ZSql + " Valor3 = " + "'" + Valor3.Text + "',"
        ZSql = ZSql + " Tope4 = " + "'" + Tope4.Text + "',"
        ZSql = ZSql + " Valor4 = " + "'" + Valor4.Text + "'"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + Lista.Text + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + Linea.Text + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + Fragancia.Text + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + Calidad.Text + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + Tamano.Text + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZZClave = Trim(Lista.Text) + ZZCodigo
        Desde.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
        ZZOrdDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
        ZZOrdHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)

        ZSql = ""
        ZSql = ZSql + "INSERT INTO Precios ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Linea ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Fragancia ,"
        ZSql = ZSql + "Calidad ,"
        ZSql = ZSql + "Tamano ,"
        ZSql = ZSql + "Lista ,"
        ZSql = ZSql + "Moneda ,"
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
        ZSql = ZSql + "'" + Linea.Text + "',"
        ZSql = ZSql + "'" + Tipo.Text + "',"
        ZSql = ZSql + "'" + Fragancia.Text + "',"
        ZSql = ZSql + "'" + Calidad.Text + "',"
        ZSql = ZSql + "'" + Tamano.Text + "',"
        ZSql = ZSql + "'" + Lista.Text + "',"
        ZSql = ZSql + "'" + Str$(Moneda.ListIndex) + "',"
        ZSql = ZSql + "'" + Desde.Text + "',"
        ZSql = ZSql + "'" + Hasta.Text + "',"
        ZSql = ZSql + "'" + ZZOrdDesde + "',"
        ZSql = ZSql + "'" + ZZOrdHasta + "',"
        ZSql = ZSql + "'" + Tope1.Text + "',"
        ZSql = ZSql + "'" + Valor1.Text + "',"
        ZSql = ZSql + "'" + Tope2.Text + "',"
        ZSql = ZSql + "'" + Valor2.Text + "',"
        ZSql = ZSql + "'" + Tope3.Text + "',"
        ZSql = ZSql + "'" + Valor3.Text + "',"
        ZSql = ZSql + "'" + Tope4.Text + "',"
        ZSql = ZSql + "'" + Valor4.Text + "')"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)

    End If
    
    
    
    
    
    
    m$ = "Grabacion realizada"
    aaaaaa% = MsgBox(m$, 0, "Archivo de Articulos")
    
    Rem Call CmdLimpiar_Click
    Linea.SetFocus
End Sub

Private Sub cmdDelete_Click()

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
        T$ = "Borrar Registro"
        m$ = "Desea Borrar el Registro "
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
    
    
            T$ = "Borrar Registro"
            m$ = "Usted va a borrar el articulo, usted esta seguro de hacerlo"
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
        
                ZSql = ""
                ZSql = ZSql + "DELETE Articulo"
                ZSql = ZSql + " Where LInea = " + "'" + Linea.Text + "'"
                ZSql = ZSql + " and Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and Fragancia = " + "'" + Fragancia.Text + "'"
                ZSql = ZSql + " and Calidad = " + "'" + Calidad.Text + "'"
                ZSql = ZSql + " and Tamano = " + "'" + Tamano.Text + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
                
            End If
            
        End If
    End If
    Linea.SetFocus
End Sub

Private Sub CmdLimpiar_Click()


    
    Linea.Text = ""
    Tipo.Text = ""
    Fragancia.Text = ""
    Calidad.Text = ""
    Tamano.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    Stock.Text = ""
    StockI.Text = ""
    StockII.Text = ""
    StockIII.Text = ""
    StockIV.Text = ""
    StockV.Text = ""
    StockVI.Text = ""
    Sector.Text = ""
    DesSector.Caption = ""
    FechaInactivo.Text = "  /  /    "
    Lista.Text = "0"
    DesLista.Caption = ""
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Tope1.Text = ""
    Valor1.Text = ""
    Tope2.Text = ""
    Valor2.Text = ""
    Tope3.Text = ""
    Valor3.Text = ""
    Tope4.Text = ""
    Valor4.Text = ""
    
    
    
    Activo.ListIndex = 0
    Facturable.ListIndex = 0
    Etiqueta.ListIndex = 0
    Moneda.ListIndex = 1
    
    Call Limpia_Vector
    Call Limpia_Vector2
    
    PantaPrecios.Visible = False
    PantaArticulo.Visible = True
    
    
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
    
    
    Linea.SetFocus
End Sub

Private Sub CmdClose_Click()
    prgArticulo2.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Codigo.Text) = "" Then: Exit Sub: End If
        
        ZSql = ""
        ZSql = ZSql + "SELECT * FROM Articulo Where Articulo = '" + Trim(Codigo.Text) + "'"
        
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
                
                Costo.Text = Trim(!Costo)
                
                Call Imprime_Datos
                .Close
            Else
                Descripcion.SetFocus
            End If
        End With
        
    ElseIf KeyAscii = 27 Then
        Codigo.Text = ""
    End If
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
        Sector.SetFocus
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
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Lineaa"
     Opcion.AddItem "Tipos"
     Opcion.AddItem "Fragancias"
     Opcion.AddItem "Calidad"
     Opcion.AddItem "Tamano"

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
            ZSql = ZSql + " FROM Lineas"
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
            
        Case 6
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Sector"
            ZSql = ZSql + " Order by Sector.Descripcion"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
            If rstSector.RecordCount > 0 Then
                With rstSector
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
                rstSector.Close
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
            indice = Pantalla.ListIndex
            Linea.Text = WIndice.List(indice)
            Call LInea_KeyPress(13)
            
        Case 1
            indice = Pantalla.ListIndex
            Tipo.Text = WIndice.List(indice)
            Call Tipo_KeyPress(13)
            
        Case 2
            indice = Pantalla.ListIndex
            Fragancia.Text = WIndice.List(indice)
            Call Fragancia_KeyPress(13)
            
        Case 3
            indice = Pantalla.ListIndex
            Calidad.Text = WIndice.List(indice)
            Call Calidad_KeyPress(13)
            
        Case 4
            indice = Pantalla.ListIndex
            Tamano.Text = WIndice.List(indice)
            Call Tamano_KeyPress(13)
            
        Case 6
            indice = Pantalla.ListIndex
            Sector.Text = WIndice.List(indice)
            Call Sector_Keypress(13)
                    
        Case Else
    End Select
    
End Sub


Sub Form_Load()
    
    'Linea.Text = ""
    'Tipo.Text = ""
    'Fragancia.Text = ""
    'Calidad.Text = ""
    'Tamano.Text = ""
    'Descripcion.Text = ""
    'DescripcionII.Text = ""
    'Sector.Text = ""
    'DesSector.Caption = ""
    'Stock.Text = ""
    'StockI.Text = ""
    'StockII.Text = ""
    'StockII.Text = ""
    'StockIV.Text = ""
    'StockV.Text = ""
    'StockVI.Text = ""
    'FechaInactivo.Text = "  /  /    "
    'Lista.Text = "0"
    'DesLista.Caption = ""
    'Desde.Text = "  /  /    "
    'Hasta.Text = "  /  /    "
    'Tope1.Text = ""
    'Valor1.Text = ""
    'Tope2.Text = ""
    'Valor2.Text = ""
    'Tope3.Text = ""
    'Valor3.Text = ""
    'Tope4.Text = ""
    'Valor4.Text = ""
   '
    'Moneda.Clear
    
    'Moneda.AddItem "Pesos"
    'Moneda.AddItem "Dolares"
    
    'Moneda.ListIndex = 1
    
    'TipoBusqueda.Value = 0
    
    'Activo.Clear
    
    'Activo.AddItem "Si"
    'Activo.AddItem "No"
    
    'Activo.ListIndex = 0

    'Facturable.Clear
     
    'Facturable.AddItem "Si"
    'Facturable.AddItem "No"
    
    'Facturable.ListIndex = 0
     
    
    'Etiqueta.Clear
    
    'Etiqueta.AddItem ""
    'Etiqueta.AddItem "Si"
    'Etiqueta.AddItem "No"
    
    'Etiqueta.ListIndex = 0
    
    Call Limpia_Vector
    'Call Limpia_Vector2
    'Call LInea_DblClick
    
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
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 5000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 25
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Costo"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
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
            'If WControl = "S" Then
                Call Control_wvector1
            'End If
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
