VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgHojaProduccion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoja de Packaging"
   ClientHeight    =   9825
   ClientLeft      =   195
   ClientTop       =   375
   ClientWidth     =   14595
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   9825
   ScaleWidth      =   14595
   Visible         =   0   'False
   Begin VB.CommandButton Impresion 
      Caption         =   "Impres. F9"
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
      Left            =   9840
      MouseIcon       =   "HojaProduccion.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "HojaProduccion.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Impresion"
      Top             =   7080
      Width           =   855
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   23
      Top             =   1920
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   8705
      _Version        =   327680
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Articulos"
      TabPicture(0)   =   "HojaProduccion.frx":0B4C
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Panta1"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Materia Prima"
      TabPicture(1)   =   "HojaProduccion.frx":0B68
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Panta2"
      Tab(1).Control(0).Enabled=   0   'False
      Begin VB.Frame Panta2 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   34
         Top             =   480
         Width           =   12015
         Begin VB.TextBox WTituloII 
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
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox WTituloII 
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
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox WTituloII 
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
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   1440
            Width           =   375
         End
         Begin VB.TextBox WTexto22 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF00&
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
            TabIndex        =   38
            Top             =   2520
            Width           =   375
         End
         Begin VB.ComboBox WCombo12 
            Height          =   315
            Left            =   3240
            TabIndex        =   37
            Top             =   1440
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.TextBox WTexto12 
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   36
            Top             =   2040
            Width           =   375
         End
         Begin VB.TextBox WTituloII 
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
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   720
            Width           =   375
         End
         Begin MSMask.MaskEdBox WTexto32 
            Height          =   285
            Left            =   5280
            TabIndex        =   42
            Top             =   2400
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            _Version        =   327680
            BackColor       =   16776960
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
         Begin MSFlexGridLib.MSFlexGrid WVector2 
            Height          =   3855
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   6800
            _Version        =   327680
            BackColor       =   16777152
         End
      End
      Begin VB.Frame Panta1 
         Height          =   4335
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   14055
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
            Index           =   6
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   2280
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
            Index           =   5
            Left            =   6240
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1560
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
            Index           =   4
            Left            =   6360
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   720
            Width           =   375
         End
         Begin VB.TextBox WTexto1 
            BackColor       =   &H00FFFF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3240
            TabIndex        =   30
            Top             =   2040
            Width           =   375
         End
         Begin VB.ComboBox WCombo1 
            Height          =   315
            Left            =   3240
            TabIndex        =   29
            Top             =   1440
            Visible         =   0   'False
            Width           =   390
         End
         Begin VB.TextBox WTexto2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF00&
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
            TabIndex        =   28
            Top             =   2520
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
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   1440
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
            Index           =   2
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   1440
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
            Index           =   3
            Left            =   5400
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   1440
            Width           =   375
         End
         Begin MSMask.MaskEdBox WTexto3 
            Height          =   285
            Left            =   5280
            TabIndex        =   32
            Top             =   2400
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   503
            _Version        =   327680
            BackColor       =   16776960
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
            Height          =   3855
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   13815
            _ExtentX        =   24368
            _ExtentY        =   6800
            _Version        =   327680
            BackColor       =   16777152
         End
      End
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
      Left            =   1440
      TabIndex        =   22
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
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
      ItemData        =   "HojaProduccion.frx":0B84
      Left            =   240
      List            =   "HojaProduccion.frx":0B8B
      TabIndex        =   21
      Top             =   7320
      Visible         =   0   'False
      Width           =   5295
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
      TabIndex        =   20
      Top             =   6960
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.TextBox Pedido 
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   18
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Numero 
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
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
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   10
      Text            =   " "
      Top             =   1200
      Width           =   1575
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1560
      Width           =   8055
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
      Left            =   11040
      MouseIcon       =   "HojaProduccion.frx":0B99
      MousePointer    =   99  'Custom
      Picture         =   "HojaProduccion.frx":0EA3
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Registro Siguiente"
      Top             =   120
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
      Left            =   10200
      MouseIcon       =   "HojaProduccion.frx":12E5
      MousePointer    =   99  'Custom
      Picture         =   "HojaProduccion.frx":15EF
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Registro Anterior"
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   8400
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Graba 
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
      Left            =   6720
      MouseIcon       =   "HojaProduccion.frx":1A31
      MousePointer    =   99  'Custom
      Picture         =   "HojaProduccion.frx":1D3B
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7080
      Width           =   855
   End
   Begin VB.CommandButton Limpia 
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
      Left            =   7800
      MouseIcon       =   "HojaProduccion.frx":257D
      MousePointer    =   99  'Custom
      Picture         =   "HojaProduccion.frx":2887
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7080
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
      Left            =   8880
      MouseIcon       =   "HojaProduccion.frx":30C9
      MousePointer    =   99  'Custom
      Picture         =   "HojaProduccion.frx":33D3
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Consulta de Datos"
      Top             =   7080
      Visible         =   0   'False
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
      Left            =   10800
      MouseIcon       =   "HojaProduccion.frx":3C15
      MousePointer    =   99  'Custom
      Picture         =   "HojaProduccion.frx":3F1F
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Menu Principal"
      Top             =   7080
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      Caption         =   "Nro de Pedido"
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
      TabIndex        =   19
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
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
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Hoja Produccion"
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
      TabIndex        =   16
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label DesCliente 
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
      Left            =   8520
      TabIndex        =   15
      Top             =   1200
      Width           =   4695
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
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
      TabIndex        =   14
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label7 
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
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label DesClienteII 
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
      Left            =   3600
      TabIndex        =   12
      Top             =   1200
      Width           =   4815
   End
End
Attribute VB_Name = "PrgHojaProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private XIndice As Single
Dim Vector(100, 10) As String
Private Auxi As String
Private XColor As String
Private XArticulo As String
Private WTipopro As Integer

Dim ZObserva(100, 2) As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Rem para el vector

Dim WBorraII(1000, 10) As String
Dim WParametrosII(10, 10) As Double
Dim WFormatoII(10) As String
Dim WControlII As String

Private Sub Consulta_Click()

    Opcion.Clear
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Semi-Terminado"
    Opcion.AddItem "Materias Primas"

    Opcion.Visible = True
    
    Opcion.ListIndex = 0
    Rem Call Opcion_Click
     
End Sub


Private Sub Impresion_Click()
    Call WImpresion
End Sub


Private Sub WImpresion()

    
    T$ = "Impresion de Hoja"
    m$ = "Desea realizarf al impresion"
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
    
    
        Listado.WindowTitle = "Impresion de Hoja"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
            
        Listado.SQLQuery = "SELECT Hoja.Numero, Hoja.Renglon, Hoja.Fecha, Hoja.Pedido, Hoja.Cliente, Hoja.Observaciones, Hoja.TipoReg, Hoja.Articulo, Hoja.Insumo, Hoja.SemiTerminado, Hoja.Descripcion, Hoja.Cantidad, Hoja.Envase, Hoja.DesEnvase, Hoja.Observa, " _
                + "Cliente.Razon " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Cliente Cliente " _
                + "Where " _
                + "Hoja.Cliente = Cliente.Cliente AND " _
                + "Hoja.Numero >= " + Numero.Text + " AND " _
                + "Hoja.Numero <= " + Numero.Text
                
        
        Listado.Connect = Connect()
        
        Uno = "{Hoja.Numero} in " + Numero.Text + " to " + Numero.Text
        
        Listado.GroupSelectionFormula = Uno
        Listado.SelectionFormula = Uno
        
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.ReportFileName = "Hoja.rpt"
        
        Listado.Action = 1

    End If

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
        
        Case 2
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
                            IngresaItem = rstInsumo!Codigo + " " + rstInsumo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstInsumo!Codigo
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
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
    
        Case 2
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Insumo"
            ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Claveven$ + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 4
                WVector1.Text = rstArticulo!Codigo
                WVector1.Col = 5
                WVector1.Text = rstArticulo!Descripcion
                WVector1.Col = 6
                rstArticulo.Close
                Call StartEdit
            End If
            Ayuda.Visible = False
                    
        Case Else
    End Select
    
End Sub


Private Sub cmdClose_Click()
    PrgFormula.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub DAA()

    For WRenglon = 1 To 100
    
        Auxi = Numero.Text
        Call Ceros(Auxi1, 8)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Clave = " + "'" + WClave + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            
            Select Case rstHoja!Tiporeg
                Case 1
                    ZTipoReg = rstHoja!Tiporeg
                    ZArticulo = rstHoja!articulo
                    ZCantidad = rstHoja!Cantidad
                
                Case Else
                    ZTipoReg = rstHoja!Tiporeg
                    ZInsumo = rstHoja!Insumo
                    ZSemiterminado = rstHoja!SemiTerminado
                    ZCantidad = rstHoja!Cantidad
            End Select
            
            rstHoja.Close
                
            Select Case ZTipoReg
                Case 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Stock = Stock - " + "'" + Str$(ZCantidad) + "',"
                    ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                Case Else
                    If Trim(ZInsumo) <> "" Then
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Insumo SET "
                        ZSql = ZSql + " Stock = Stock + " + "'" + Str$(ZCantidad) + "',"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZInsumo + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                            Else
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Stock = Stock + " + "'" + Str$(ZCantidad) + "',"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
            End Select
                    
                    
                
            Renglon = Renglon + 1
                
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = Trim(rstHoja!Insumo)
            Auxi1 = Trim(rstHoja!Insumo)
            
            WVector1.Col = 2
            WVector1.Text = Trim(rstHoja!terminado)
            Auxi2 = Trim(rstHoja!terminado)
            
            
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", Str$(rstHoja!Cantidad))
            
            Combo.Text = Trim(rstHoja!Combo)
            
            rstHoja.Close
                
            If Trim(Auxi1) <> "" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Insumo"
                ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Auxi1 + "'"
                spInsumo = ZSql
                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                If rstInsumo.RecordCount > 0 Then
                    WVector1.Col = 3
                    WVector1.Text = rstInsumo!Descripcion
                    rstInsumo.Close
                End If
            End If
                
            If Trim(Auxi2) <> "" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi2 + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 3
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            End If
                
                
        End If
    
    Next WRenglon
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Combo"
    ZSql = ZSql + " Where Combo.Codigo = " + "'" + Combo.Text + "'"
    spCombo = ZSql
    Set rstCombo = db.OpenRecordset(spCombo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCombo.RecordCount > 0 Then
        DesCombo.Caption = rstCombo!Descripcion
        rstCombo.Close
    End If
    
    WVector1.Col = 1
    WVector1.Row = 1


End Sub


Private Sub Graba_Click()

    
    ZSql = ""
    ZSql = ZSql + "DELETE Hoja"
    ZSql = ZSql + " Where Hoja.Numero = " + "'" + Numero.Text + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    

    Renglon = 0
        
    For IRow = 1 To 100
            
        WVector1.Row = IRow
            
        WVector1.Col = 1
        ZZArticulo = WVector1.Text
        
        WVector1.Col = 2
        ZZDescripcion = WVector1.Text
        
        WVector1.Col = 3
        ZZCantidad = Val(WVector1.Text)
        
        WVector1.Col = 4
        ZZEnvase = WVector1.Text
        
        WVector1.Col = 5
        ZZDesEnvase = WVector1.Text
        
        WVector1.Col = 6
        ZZObservaciones = WVector1.Text
        
        If ZZCantidad <> 0 Then
                    
            Renglon = Renglon + 1
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 6)
            
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
                        
            ZZClave = Auxi + Auxi1
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Hoja ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Tiporeg ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Insumo ,"
            ZSql = ZSql + "SemiTerminado ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Envase ,"
            ZSql = ZSql + "DesEnvase ,"
            ZSql = ZSql + "Observa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Str$(Renglon) + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + Pedido.Text + "',"
            ZSql = ZSql + "'" + Cliente.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + "1" + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + Str$(ZZCantidad) + "',"
            ZSql = ZSql + "'" + ZZEnvase + "',"
            ZSql = ZSql + "'" + ZZDesEnvase + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "')"
                            
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
                                       
                                       
                                       
        WVector2.Row = IRow
            
        WVector2.Col = 1
        ZZInsumo = WVector2.Text
        
        WVector2.Col = 2
        ZZSemiTerminado = WVector2.Text
        
        WVector2.Col = 3
        ZZDescripcion = WVector2.Text
        
        WVector2.Col = 4
        ZZCantidad = Val(WVector2.Text)
        
        If ZZCantidad <> 0 Then
                    
            Renglon = Renglon + 1
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 6)
            
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
                        
            ZZClave = Auxi + Auxi1
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Hoja ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Tiporeg ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Insumo ,"
            ZSql = ZSql + "SemiTerminado ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Envase ,"
            ZSql = ZSql + "DesEnvase ,"
            ZSql = ZSql + "Observa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Str$(Renglon) + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + Pedido.Text + "',"
            ZSql = ZSql + "'" + Cliente.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + "2" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + ZZInsumo + "',"
            ZSql = ZSql + "'" + ZZSemiTerminado + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + Str$(ZZCantidad) + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "')"
                            
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
                                       
    Next IRow
    
    
    m$ = "Grabacion realizada"
    aaaaaa% = MsgBox(m$, 0, "Archivo de Familias")
        
    Call WImpresion
        
    Rem Call Limpia_Click
    Numero.SetFocus
        
End Sub


Private Sub Limpia_Click()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    Numero.Text = "1"
    Cliente.Text = ""
    DesCliente.Caption = ""
    DesClienteII.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    ZSql = ZSql + " FROM Hoja"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        rstHoja.MoveLast
        ZUltimo = IIf(IsNull(rstHoja!NumeroMayor), "0", rstHoja!NumeroMayor)
        Numero.Text = ZUltimo + 1
        rstHoja.Close
    End If
    
    Renglon = 0
    Numero.SetFocus

End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Call Limpia_VectorII
    
    Numero.Text = "1"
    Cliente.Text = ""
    DesCliente.Caption = ""
    DesClienteII.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    ZSql = ZSql + " FROM Hoja"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        rstHoja.MoveLast
        ZUltimo = IIf(IsNull(rstHoja!NumeroMayor), "0", rstHoja!NumeroMayor)
        Numero.Text = ZUltimo + 1
        rstHoja.Close
    End If
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    Call Limpia_VectorII

    Renglon = 0
    RenglonII = 0
    
    For WRenglon = 1 To 100
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Clave = " + "'" + WClave + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            
            If rstHoja!Tiporeg = 1 Then
            
                Renglon = Renglon + 1
                    
                WVector1.Row = Renglon
                    
                WVector1.Col = 1
                WVector1.Text = Trim(rstHoja!articulo)
                
                WVector1.Col = 2
                WVector1.Text = Trim(rstHoja!Descripcion)
                
                WVector1.Col = 3
                WVector1.Text = Pusing("###,###.##", Str$(rstHoja!Cantidad))
                
                WVector1.Col = 4
                WVector1.Text = Trim(rstHoja!envase)
                
                WVector1.Col = 5
                WVector1.Text = Trim(rstHoja!DesEnvase)
                
                WVector1.Col = 6
                WVector1.Text = Trim(rstHoja!observa)
                
                    Else
                
                RenglonII = RenglonII + 1
                    
                WVector2.Row = RenglonII
                    
                WVector2.Col = 1
                WVector2.Text = Trim(rstHoja!Insumo)
                
                WVector2.Col = 2
                WVector2.Text = Trim(rstHoja!SemiTerminado)
                
                WVector2.Col = 3
                WVector2.Text = Trim(rstHoja!Descripcion)
                
                WVector2.Col = 4
                WVector2.Text = Pusing("###,###.##", Str$(rstHoja!Cantidad))
                
            End If
            
            rstHoja.Close
                
        End If
    
    Next WRenglon
    
    
    WVector1.Col = 1
    WVector1.Row = 1


    WVector2.Col = 1
    WVector2.Row = 1

End Sub

Sub Numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Numero = " + "'" + Numero.Text + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
    
            Fecha.Text = rstHoja!Fecha
            Cliente.Text = rstHoja!Cliente
            Pedido.Text = rstHoja!Pedido
            Observaciones.Text = rstHoja!Observaciones
            
            rstHoja.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Fantasia
                DesClienteII.Caption = rstCliente!Razon
                rstCliente.Close
            End If
            
            Call Proceso_Click
            
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                
                Else
                    
            Pedido.SetFocus
               
        End If
            
    End If
    
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub



Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Numero = " + "'" + Pedido.Text + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
    
            Cliente.Text = rstPedido!Cliente
            Observaciones.Text = Trim(rstPedido!Observaciones)
            
            rstPedido.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Fantasia
                DesClienteII.Caption = rstCliente!Razon
                rstCliente.Close
            End If
            
            
            Call LeePedido
            Observaciones.SetFocus
               
        End If
            
    End If
    
    If KeyAscii = 27 Then
        Pedido.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
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
            DesCliente.Caption = rstCliente!Fantasia
            DesClienteII.Caption = rstCliente!Razon
            rstCliente.Close
            Observaciones.SetFocus
                Else
            Cliente.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
        DesClienteII.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            aaaaaa% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
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
    
    Rem If XIndice = 0 And KeyAscii <> 13 Then
    Rem     Exit Sub
    Rem End If
    
    Select Case XIndice
        Case 2
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
                        WTexto3.Text = "  /  /    "
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
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
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
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
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

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
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
            WTexto3.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
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
        Case 6
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                rstArticulo.Close
                        Else
                WControl = "N"
            End If
            
            
        Case 4
            If Trim(WVector1.Text) <> "" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Insumo"
                ZSql = ZSql + " Where Insumo.Codigo = " + "'" + WVector1.Text + "'"
                spInsumo = ZSql
                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                If rstInsumo.RecordCount > 0 Then
                    WVector1.Col = 5
                    WVector1.Text = rstInsumo!Descripcion
                    rstInsumo.Close
                            Else
                    WControl = "N"
                End If
                    Else
                WVector1.Col = 5
                WVector1.Text = ""
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        WVector1.Col = 3
        WAuxi2 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
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
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Envase"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 16
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 3000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
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

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 4 Then

        Opcion.Clear
    
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"

        Rem Opcion.Visible = False
    
        Opcion.ListIndex = 2
    
        Call aYUDA_Keypress(13)
    
    End If
    
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 4 Then

        Opcion.Clear
    
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"

        Rem Opcion.Visible = False
    
        Opcion.ListIndex = 2
    
        Call aYUDA_Keypress(13)
    
    End If
    
End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Graba_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub






















Private Sub Busqueda()

    Rem On Error GoTo WError
    
    PantaArticulo.Visible = True
    Call Limpia_VectorII
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
                    
                    ZLugar = ZLugar + 1
                    WVector2.TextMatrix(ZLugar, 1) = !Codigo
                    WVector2.TextMatrix(ZLugar, 2) = !Descripcion
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If

End Sub






Rem
Rem Controles de la WVector2
Rem

Private Sub GridEditTextII(ByVal KeyAscii As Integer)

    XColumna = WVector2.Col
    XTipoDato = WParametrosII(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector2.CellLeft + WVector2.Left
            WTexto12.Top = WVector2.CellTop + WVector2.Top
            WTexto12.Width = WVector2.CellWidth
            WTexto12.Height = WVector2.CellHeight
            WTexto12.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector2.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector2.CellLeft + WVector2.Left
            WTexto22.Top = WVector2.CellTop + WVector2.Top
            WTexto22.Width = WVector2.CellWidth
            WTexto22.Height = WVector2.CellHeight
            WTexto22.MaxLength = WParametrosII(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector2.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector2.CellLeft + WVector2.Left
            WTexto32.Top = WVector2.CellTop + WVector2.Top
            WTexto32.Width = WVector2.CellWidth
            WTexto32.Height = WVector2.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector2.Text) = 10 Then
                        WTexto32.Text = WVector2.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEditII()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector2.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector2.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector2.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector2.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormatoII(WVector2.Col) <> "" Then
            WVector2.Text = Pusing(WFormatoII(WVector2.Col), WVector2.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditComboII()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector2.CellLeft + WVector2.Left
    WCombo12.Top = WVector2.CellTop + WVector2.Top
    WCombo12.Width = WVector2.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto12.Text = WVector2.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto22.Text = WVector2.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            DoEvents
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

Private Sub WTexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto32.Text = WVector2.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector2.SetFocus
            Call Control_CampoII
            If WControlII = "S" Then
                Call Control_WVector2
            End If
            Call StartEditII

        Case vbKeyDown
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row < WVector2.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row + 1
                Rem End If
            End If
            Call StartEditII

        Case vbKeyUp
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.Row > WVector2.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.Row = WVector2.Row - 1
                Rem End If
            End If
            Call StartEditII
        Case 34
            ' Move down 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow < WVector2.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow + 12
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII
            
        Case 33
            ' Move up 1 row.
            WVector2.SetFocus
            DoEvents
            If WVector2.TopRow - 12 > WVector2.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector2.TopRow = WVector2.TopRow - 12
                    WVector2.Row = WVector2.TopRow
                        Else
                    WVector2.TopRow = 1
                    WVector2.Row = WVector2.TopRow
                Rem End If
            End If
            Call StartEditII

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector2.SetFocus
End Sub


Private Sub WVector2_Click()
    StartEditII
End Sub

Private Sub WVector2_LeaveCell()
    EndEditII
End Sub

Private Sub WVector2_GotFocus()
    EndEditII
End Sub

Private Sub WVector2_KeyPress(KeyAscii As Integer)
    XColumna = WVector2.Col
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
        Case Else
            If WParametrosII(2, XColumna) = 0 Then
                GridEditTextII KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEditII()
    Select Case WParametrosII(4, WVector2.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector2.Text
            On Error GoTo 0
            GridEditComboII
        Case Else
            If WParametrosII(2, WVector2.Col) = 0 Then
                GridEditTextII Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_WVector2()
    Select Case WVector2.Col
        Case 4
            If WVector2.Row < WVector2.Rows - 1 Then
                WVector2.Row = WVector2.Row + 1
            End If
            WVector2.Col = 1
        Case Else
            If WVector2.Col < WVector2.Cols - 1 Then
                WVector2.Col = WVector2.Col + 1
            End If
    End Select
    WVector2.SetFocus
    GridEditTextII KeyAscii
End Sub

Private Sub Control_CampoII()
    XColumna = WVector2.Col
    XFila = WVector2.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Trim(WVector2.Text) <> "" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Insumo"
                ZSql = ZSql + " Where Insumo.Codigo = " + "'" + WVector2.Text + "'"
                spInsumo = ZSql
                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                If rstInsumo.RecordCount > 0 Then
                    WVector2.Col = 2
                    WVector2.Text = ""
                    WVector2.Col = 3
                    WVector2.Text = rstInsumo!Descripcion
                    rstInsumo.Close
                            Else
                    WControl = "N"
                End If
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector2.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector2.Col = 1
                WVector2.Text = ""
                WVector2.Col = 3
                WVector2.Text = rstArticulo!Descripcion
                rstArticulo.Close
                        Else
                WControl = "N"
            End If
            
            
        Case Else
            WVector2.Col = XColumna
    End Select
End Sub

Private Sub WVector2_DblClick()

    If WVector2.Col = 0 Or WVector2.Col = 1 Then
    
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False

    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WVector2.Text = ""
    Next Ciclo
    
    Erase WBorraII
    EntraVector = 0
    
    For Ciclo = 1 To WVector2.Rows - 1
        WVector2.Row = Ciclo
        WVector2.Col = 1
        WAuxi1 = WVector2.Text
        WVector2.Col = 3
        WAuxi2 = WVector2.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector2.Cols - 1
                WVector2.Col = Ciclo1
                WBorraII(EntraVector, Ciclo1) = WVector2.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_VectorII
    
    For Ciclo = 1 To EntraVector
        WVector2.Row = Ciclo
        For da = 1 To WVector2.Cols - 1
            WVector2.Col = da
            WVector2.Text = WBorraII(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_VectorII()

    WVector2.Clear

    Rem ponga la WVector2 en negritas
    WVector2.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector2.FontName
    WTexto12.FontSize = WVector2.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector2.FontName
    WTexto22.FontSize = WVector2.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector2.FontName
    WTexto32.FontSize = WVector2.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector2.FontName
    WCombo12.FontSize = WVector2.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la WVector2
    
    WVector2.FixedCols = 1
    WVector2.Cols = 5
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector2.Text = "Articulo"
    
    Rem Longitud
    Rem WVector2.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametrosII(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametrosII(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametrosII(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametrosII(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Insumo"
                WVector2.ColWidth(Ciclo) = 2000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 20
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 2
                WVector2.Text = "SemiTerminado"
                WVector2.ColWidth(Ciclo) = 2500
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 25
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 3
                WVector2.Text = "Descripcion"
                WVector2.ColWidth(Ciclo) = 5000
                WVector2.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametrosII(1, Ciclo) = 50
                WParametrosII(2, Ciclo) = 1
                WParametrosII(3, Ciclo) = 0
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
            Case 4
                WVector2.Text = "Cantidad"
                WVector2.ColWidth(Ciclo) = 1200
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametrosII(1, Ciclo) = 10
                WParametrosII(2, Ciclo) = 0
                WParametrosII(3, Ciclo) = 1
                WParametrosII(4, Ciclo) = 0
                WFormatoII(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WTituloII(Ciclo).Text = WVector2.Text
        WTituloII(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        WTituloII(Ciclo).Top = WVector2.CellTop + WVector2.Top
        WTituloII(Ciclo).Width = WVector2.CellWidth
        WTituloII(Ciclo).Height = WVector2.CellHeight
        WTituloII(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA WVector2
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub WVector2_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub

Private Sub WTexto12_DblClick()

    If WVector2.Col = 1 Then

        Opcion.Clear
    
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"
        Opcion.AddItem "Articulo"

        Rem Opcion.Visible = False
    
        Opcion.ListIndex = 5
    
        Call aYUDA_Keypress(13)
    
    End If
    
End Sub



Private Sub LeePedido()

    Call Limpia_Vector
    Call Limpia_VectorII
    

    Renglon = 0
    
    For WRenglon = 1 To 99
    
        Auxi = Pedido.Text
        Call Ceros(Auxi, 8)
            
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Clave = " + "'" + WClave + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                    
            ZZArticulo = Trim(rstPedido!articulo)
            ZZCantidad = rstPedido!Cantidad
            ZZLinea = rstPedido!Linea
            ZZTipo = rstPedido!Tipo
            ZZFragancia = rstPedido!Fragancia
            ZZCalidad = rstPedido!Calidad
            ZZTamano = rstPedido!Tamano
            ZZFecEntrega = Trim(rstPedido!FecEntrega)
            ZZObserva = Trim(rstPedido!observa)
            ZZFabrica = rstPedido!fabrica
            ZZFacturado = rstPedido!facturado
            ZZFechaFabrica = rstPedido!fechafabrica
            ZZMarca = rstPedido!Marca
            ZZMarcaII = rstPedido!MarcaII
            ZZEntregado = rstPedido!Entregado
            ZZAjuste = rstPedido!Ajuste
            
            ZObserva(1, 1) = rstPedido!Observa11
            ZObserva(1, 2) = rstPedido!Observa1
            
            ZObserva(2, 1) = rstPedido!Observa12
            ZObserva(2, 2) = rstPedido!Observa2
            
            ZObserva(3, 1) = rstPedido!Observa13
            ZObserva(3, 2) = rstPedido!Observa3
            
            ZObserva(4, 1) = rstPedido!Observa14
            ZObserva(4, 2) = rstPedido!Observa4
            
            ZObserva(5, 1) = rstPedido!Observa15
            ZObserva(5, 2) = rstPedido!Observa5
            
            ZObserva(6, 1) = rstPedido!Observa16
            ZObserva(6, 2) = rstPedido!Observa6
            
            ZObserva(7, 1) = rstPedido!Observa17
            ZObserva(7, 2) = rstPedido!Observa7
            
            ZObserva(8, 1) = rstPedido!Observa18
            ZObserva(8, 2) = rstPedido!Observa8
            
            ZObserva(9, 1) = rstPedido!Observa19
            ZObserva(9, 2) = rstPedido!Observa9
            
            ZObserva(10, 1) = rstPedido!Observa20
            ZObserva(10, 2) = rstPedido!Observa10
            
            For ZZCiclo = 1 To 10
                If Val(ZObserva(ZZCiclo, 1)) = Renglon Then
                    ZZObserva = ZObserva(ZZCiclo, 2)
                End If
            Next ZZCiclo
            
            rstPedido.Close
            
            Canti = ZZCantidad
            
            WVector1.Col = 1
            WVector1.Text = Trim(ZZArticulo)
            Auxi1 = ZZArticulo
                
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", Str$(ZZCantidad))
            WWCantidad = ZZCantidad
            
            WVector1.Col = 6
            WVector1.Text = Trim(ZZObserva)
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstArticulo!DescripcionII
                rstArticulo.Close
            End If
            
        End If
    
    Next WRenglon
    
    Call Calcula_Mp
    
    SSTab1.Tab = 1
    SSTab1.Tab = 0
    
End Sub

Private Sub Calcula_Mp()

    Call Limpia_VectorII

    Renglon = 0
    
    For CicloII = 1 To 99
    
        If Trim(WVector1.TextMatrix(CicloII, 1)) <> "" Then
    
            ZZCombo = ""
            ZZCantiCombo = 0
    
            For WRenglon = 1 To 100
            
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                
                ZZCodigo = WVector1.TextMatrix(CicloII, 1)
                WClave = ZZCodigo + Auxi1
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Formula"
                ZSql = ZSql + " Where Formula.Clave = " + "'" + WClave + "'"
                spFormula = ZSql
                Set rstFormula = db.OpenRecordset(spFormula, dbOpenSnapshot, dbSQLPassThrough)
                If rstFormula.RecordCount > 0 Then
                    
                    ZZCantidad = rstFormula!Cantidad * Val(WVector1.TextMatrix(CicloII, 3))
                    ZZInsumo = Trim(rstFormula!Insumo)
                    ZZTerminado = Trim(rstFormula!terminado)
                    
                    ZZEntra = "S"
                    
                    For CicloIII = 1 To Renglon
                        
                        If Trim(ZZInsumo) <> "" Then
                            If UCase(Trim(WVector2.TextMatrix(CicloIII, 1))) = Trim(UCase(ZZInsumo)) Then
                                WVector2.TextMatrix(CicloIII, 4) = Str$(Val(WVector2.TextMatrix(CicloIII, 4)) + ZZCantidad)
                                ZZEntra = "N"
                                Exit For
                            End If
                        End If
                        
                        If Trim(ZZTerminado) <> "" Then
                            If UCase(Trim(WVector2.TextMatrix(CicloIII, 2))) = Trim(UCase(ZZTerminado)) Then
                                WVector2.TextMatrix(CicloIII, 4) = Str$(Val(WVector2.TextMatrix(CicloIII, 4)) + ZZCantidad)
                                ZZEntra = "N"
                                Exit For
                            End If
                        End If
                                            
                    Next CicloIII
                    
                    If ZZEntra = "S" Then
                        
                        Renglon = Renglon + 1
                            
                        WVector2.Row = Renglon
                            
                        WVector2.Col = 1
                        WVector2.Text = Trim(rstFormula!Insumo)
                        Auxi1 = Trim(rstFormula!Insumo)
                        
                        WVector2.Col = 2
                        WVector2.Text = Trim(rstFormula!terminado)
                        Auxi2 = Trim(rstFormula!terminado)
                        
                        
                        WVector2.Col = 4
                        WVector2.Text = Pusing("###,###.##", Str$(rstFormula!Cantidad * Val(WVector1.TextMatrix(CicloII, 3))))
                            
                        If Trim(Auxi1) <> "" Then
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Insumo"
                            ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Auxi1 + "'"
                            spInsumo = ZSql
                            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstInsumo.RecordCount > 0 Then
                                WVector2.Col = 3
                                WVector2.Text = rstInsumo!Descripcion
                                rstInsumo.Close
                            End If
                        End If
                            
                        If Trim(Auxi2) <> "" Then
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Articulo"
                            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi2 + "'"
                            spArticulo = ZSql
                            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstArticulo.RecordCount > 0 Then
                                WVector2.Col = 3
                                WVector2.Text = rstArticulo!Descripcion
                                rstArticulo.Close
                            End If
                        End If
                        
                    End If
                        
                    
                    ZZCombo = Trim(rstFormula!Combo)
                    ZZCantiCombo = Val(WVector1.TextMatrix(CicloII, 3))
                    
                    rstFormula.Close
                        
                        
                End If
            
            Next WRenglon
            
            If Trim(ZZCombo) <> "" Then
                
                For WRenglon = 1 To 100
                
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    
                    ZZCodigo = ZZCombo
                    WClave = ZZCodigo + Auxi1
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Combo"
                    ZSql = ZSql + " Where Combo.Clave = " + "'" + WClave + "'"
                    spCombo = ZSql
                    Set rstCombo = db.OpenRecordset(spCombo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCombo.RecordCount > 0 Then
                        
                        ZZInsumo = Trim(rstCombo!Insumo)
                        ZZCantidad = rstCombo!Cantidad * ZZCantiCombo
                        
                        ZZEntra = "S"
                        
                        For CicloIII = 1 To Renglon
                            
                            If Trim(ZZInsumo) <> "" Then
                                If UCase(Trim(WVector2.TextMatrix(CicloIII, 1))) = Trim(UCase(ZZInsumo)) Then
                                    WVector2.TextMatrix(CicloIII, 4) = Str$(Val(WVector2.TextMatrix(CicloIII, 4)) + ZZCantidad)
                                    ZZEntra = "N"
                                    Exit For
                                End If
                            End If
                                                
                        Next CicloIII
                        
                        If ZZEntra = "S" Then
                            
                            Renglon = Renglon + 1
                                
                            WVector2.Row = Renglon
                                
                            WVector2.Col = 1
                            WVector2.Text = Trim(rstCombo!Insumo)
                            Auxi1 = Trim(rstCombo!Insumo)
                            
                            WVector2.Col = 2
                            WVector2.Text = ""
                            
                            
                            WVector2.Col = 4
                            WVector2.Text = Pusing("###,###.##", Str$(rstCombo!Cantidad * ZZCantiCombo))
                            
                            rstCombo.Close
                                
                            If Trim(Auxi1) <> "" Then
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Insumo"
                                ZSql = ZSql + " Where Insumo.Codigo = " + "'" + Auxi1 + "'"
                                spInsumo = ZSql
                                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstInsumo.RecordCount > 0 Then
                                    WVector2.Col = 3
                                    WVector2.Text = rstInsumo!Descripcion
                                    rstInsumo.Close
                                End If
                            End If
                            
                        End If
                                            
                    End If
                
                Next WRenglon
            
            End If
            
        End If
    
        WVector2.Col = 1
        WVector2.Row = 1

    Next CicloII
End Sub

