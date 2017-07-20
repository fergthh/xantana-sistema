VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form prgcliente 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Clientes"
   ClientHeight    =   7605
   ClientLeft      =   930
   ClientTop       =   405
   ClientWidth     =   11415
   LinkTopic       =   "Form2"
   ScaleHeight     =   7605
   ScaleWidth      =   11415
   Begin VB.Frame Frame6 
      Caption         =   "Navegación por Registros"
      Height          =   1095
      Left            =   6240
      TabIndex        =   93
      Top             =   6360
      Width           =   4335
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
         Height          =   615
         Left            =   240
         MouseIcon       =   "cliente.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":030A
         TabIndex        =   97
         ToolTipText     =   "Primer Registro"
         Top             =   290
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
         Height          =   615
         Left            =   1200
         MouseIcon       =   "cliente.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":0A56
         TabIndex        =   96
         ToolTipText     =   "Registro Anterior"
         Top             =   290
         Width           =   855
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Siguiente F7"
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
         Left            =   2160
         MouseIcon       =   "cliente.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":11A2
         TabIndex        =   95
         ToolTipText     =   "Registro Siguiente"
         Top             =   290
         Width           =   975
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
         Height          =   615
         Left            =   3240
         MouseIcon       =   "cliente.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":18EE
         TabIndex        =   94
         ToolTipText     =   "Salida"
         Top             =   290
         Width           =   855
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10320
      TabIndex        =   92
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
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
      Height          =   615
      Left            =   5280
      MouseIcon       =   "cliente.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":203A
      TabIndex        =   91
      ToolTipText     =   "Salida"
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Lista 
      Caption         =   "Listar F9"
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
      Left            =   4320
      MouseIcon       =   "cliente.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":2B86
      TabIndex        =   90
      ToolTipText     =   "Impresion "
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consultar F4"
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
      Left            =   3240
      MouseIcon       =   "cliente.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":36D2
      TabIndex        =   89
      ToolTipText     =   "Consulta de Datos"
      Top             =   6600
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar F3"
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
      Left            =   2280
      MouseIcon       =   "cliente.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":421E
      TabIndex        =   88
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borrar  F2"
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
      Left            =   1320
      MouseIcon       =   "cliente.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":4D6A
      TabIndex        =   87
      ToolTipText     =   "Elimina el Registro"
      Top             =   6600
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Grabar F1"
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
      Left            =   360
      MouseIcon       =   "cliente.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "cliente.frx":58B6
      TabIndex        =   86
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6600
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10920
      Top             =   6600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10821
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales"
      TabPicture(0)   =   "cliente.frx":60F8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(1)=   "Frame4"
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(3)=   "Fantasia"
      Tab(0).Control(4)=   "DatosAdicinales"
      Tab(0).Control(5)=   "Razon"
      Tab(0).Control(6)=   "Cliente"
      Tab(0).Control(7)=   "lblLabels(1)"
      Tab(0).Control(8)=   "lblLabels(2)"
      Tab(0).Control(9)=   "lblLabels(0)"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Facturación"
      TabPicture(1)   =   "cliente.frx":6114
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label26"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label25"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label24"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabels(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DesCondicion"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label17"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DesTipoClie"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label12"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label15"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "DesListaPrecios"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "FechaAlta"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "LocalidadII"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "PostalII"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "ProvinciaII"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Historial"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "ClieLista"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Bonifica"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Vercontactos"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "DireccionII"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Condicion"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TipoClie"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "ObservacionesII"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "NroLista"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Expreso"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "cliente.frx":6130
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
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
         Left            =   10440
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   114
         Text            =   "cliente.frx":614C
         Top             =   6600
         Width           =   180
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos Impositivos"
         Height          =   1215
         Left            =   -74760
         TabIndex        =   74
         Top             =   4680
         Width           =   10455
         Begin MSMask.MaskEdBox Cuit 
            Height          =   285
            Left            =   960
            TabIndex        =   85
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   13
            Mask            =   "##-########-#"
            PromptChar      =   " "
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
            Height          =   975
            Left            =   4440
            TabIndex        =   78
            Top             =   120
            Width           =   5775
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
               Left            =   4080
               TabIndex        =   84
               Top             =   600
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
               Left            =   2280
               TabIndex        =   81
               Top             =   240
               Width           =   1455
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
               Left            =   4080
               TabIndex        =   83
               Top             =   240
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
               TabIndex        =   82
               Top             =   600
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
               Left            =   360
               TabIndex        =   80
               Top             =   600
               Width           =   1455
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
               Left            =   360
               TabIndex        =   79
               Top             =   240
               Width           =   1215
            End
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
            Left            =   3240
            MaxLength       =   6
            TabIndex        =   75
            Text            =   " "
            Top             =   480
            Width           =   855
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
            Left            =   240
            TabIndex        =   77
            Top             =   480
            Width           =   735
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
            Left            =   2640
            TabIndex        =   76
            Top             =   480
            Width           =   615
         End
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   72
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos de Contacto"
         Height          =   1815
         Left            =   -74760
         TabIndex        =   59
         Top             =   2760
         Width           =   10455
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
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   70
            Text            =   " "
            Top             =   660
            Width           =   3255
         End
         Begin VB.TextBox txtPaginaWeb 
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
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   68
            Text            =   " "
            Top             =   660
            Width           =   4215
         End
         Begin VB.TextBox txtNombreContacto 
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
            Left            =   7200
            MaxLength       =   50
            TabIndex        =   66
            Text            =   " "
            Top             =   300
            Width           =   3135
         End
         Begin VB.TextBox Telefono 
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
            Height          =   315
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   62
            Text            =   " "
            Top             =   300
            Width           =   3255
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
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   61
            Text            =   " "
            Top             =   1400
            Width           =   8775
         End
         Begin VB.TextBox fax 
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
            MaxLength       =   30
            TabIndex        =   60
            Text            =   " "
            Top             =   1050
            Width           =   3255
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
            Left            =   240
            TabIndex        =   71
            Top             =   660
            Width           =   1095
         End
         Begin VB.Label Label29 
            Caption         =   "Página Web:"
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
            Left            =   4920
            TabIndex        =   69
            Top             =   660
            Width           =   1215
         End
         Begin VB.Label Label19 
            Caption         =   "Responsable / Contacto"
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
            Left            =   4920
            TabIndex        =   67
            Top             =   300
            Width           =   2175
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
            Left            =   240
            TabIndex        =   65
            Top             =   300
            Width           =   975
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
            TabIndex        =   64
            Top             =   1400
            Width           =   1455
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
            Left            =   720
            TabIndex        =   63
            Top             =   1050
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Domiciliarios"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   50
         Top             =   1320
         Width           =   10455
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
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   54
            Text            =   " "
            Top             =   360
            Width           =   6135
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
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   53
            Text            =   " "
            Top             =   720
            Width           =   4215
         End
         Begin VB.TextBox Postal 
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
            Left            =   9000
            MaxLength       =   15
            TabIndex        =   52
            Text            =   " "
            Top             =   360
            Width           =   1335
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
            Left            =   7200
            TabIndex        =   51
            Text            =   " "
            Top             =   720
            Width           =   3135
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
            TabIndex        =   58
            Top             =   360
            Width           =   1215
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
            TabIndex        =   57
            Top             =   720
            Width           =   1215
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
            Left            =   7680
            TabIndex        =   56
            Top             =   360
            Width           =   1455
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
            Left            =   5880
            TabIndex        =   55
            Top             =   720
            Width           =   1095
         End
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
         Left            =   -72240
         MaxLength       =   50
         TabIndex        =   48
         Top             =   960
         Width           =   7815
      End
      Begin VB.TextBox ObservacionesII 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   2280
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "cliente.frx":614E
         Top             =   3120
         Width           =   7815
      End
      Begin VB.TextBox TipoClie 
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
         TabIndex        =   35
         Text            =   " "
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Condicion 
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
         TabIndex        =   34
         Text            =   " "
         Top             =   960
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
         TabIndex        =   33
         Top             =   2160
         Width           =   5175
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
         Left            =   6240
         TabIndex        =   32
         Top             =   4800
         Width           =   1455
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
         Left            =   9240
         TabIndex        =   31
         Top             =   6720
         Visible         =   0   'False
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
         Left            =   9120
         TabIndex        =   30
         Top             =   6480
         Width           =   1695
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
         Left            =   7920
         TabIndex        =   29
         Top             =   4800
         Width           =   1695
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
         Left            =   6960
         TabIndex        =   28
         Text            =   " "
         Top             =   2520
         Width           =   3135
      End
      Begin VB.TextBox PostalII 
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
         Left            =   9120
         MaxLength       =   15
         TabIndex        =   27
         Text            =   " "
         Top             =   2160
         Width           =   975
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
         TabIndex        =   26
         Text            =   " "
         Top             =   2520
         Width           =   3135
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
         Left            =   -63120
         MouseIcon       =   "cliente.frx":6150
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":645A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Pedidos de Clientes"
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
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
         Left            =   -70320
         MaxLength       =   50
         TabIndex        =   23
         Top             =   480
         Width           =   5895
      End
      Begin VB.TextBox Cliente 
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
         Left            =   -73440
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   1575
      End
      Begin MSMask.MaskEdBox FechaAlta 
         Height          =   285
         Left            =   8640
         TabIndex        =   37
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VB.Label DesListaPrecios 
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
         TabIndex        =   98
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label15 
         Caption         =   "Lista de Precios:"
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
         Left            =   600
         TabIndex        =   73
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombre de Fantasía:"
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
         Left            =   -74400
         TabIndex        =   49
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Observaciones:"
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
         Left            =   720
         TabIndex        =   47
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Categoria:"
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
         Left            =   1155
         TabIndex        =   46
         Top             =   1680
         Width           =   855
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
         Left            =   3240
         TabIndex        =   45
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label17 
         Caption         =   "Condicion de Pago"
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
         Left            =   360
         TabIndex        =   44
         Top             =   960
         Width           =   1815
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
         TabIndex        =   43
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Dirección de Entrega:"
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
         Left            =   160
         TabIndex        =   42
         Top             =   2160
         Width           =   1935
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
         Left            =   7560
         TabIndex        =   41
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label24 
         Caption         =   "Provincia:"
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
         TabIndex        =   40
         Top             =   2520
         Width           =   855
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
         Height          =   255
         Left            =   7800
         TabIndex        =   39
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "Localidad:"
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
         Left            =   1150
         TabIndex        =   38
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Razón Social:"
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
         Left            =   -71640
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Código"
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
         Left            =   -74400
         TabIndex        =   22
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame PantaContacto 
      Caption         =   "Contactos"
      Height          =   5415
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton CierraContactos 
         Caption         =   "Aceptar y Cerrar"
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
         MouseIcon       =   "cliente.frx":6D24
         MousePointer    =   99  'Custom
         Picture         =   "cliente.frx":702E
         TabIndex        =   20
         ToolTipText     =   "Salida"
         Top             =   3960
         Width           =   2175
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
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   18
         Text            =   " "
         Top             =   3120
         Width           =   3015
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
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   16
         Text            =   " "
         Top             =   3120
         Width           =   3495
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2760
         Width           =   7455
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
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   2040
         Width           =   3135
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
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   10
         Text            =   " "
         Top             =   2040
         Width           =   3495
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1680
         Width           =   7575
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
         Left            =   6480
         MaxLength       =   50
         TabIndex        =   6
         Text            =   " "
         Top             =   960
         Width           =   3135
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
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   4
         Text            =   " "
         Top             =   960
         Width           =   3495
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   600
         Width           =   7575
      End
      Begin VB.Line Line4 
         X1              =   1200
         X2              =   9120
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line3 
         X1              =   1200
         X2              =   9120
         Y1              =   1440
         Y2              =   1440
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
         Left            =   5760
         TabIndex        =   19
         Top             =   3120
         Width           =   735
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
         Left            =   840
         TabIndex        =   17
         Top             =   3120
         Width           =   855
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
         Left            =   840
         TabIndex        =   15
         Top             =   2760
         Width           =   855
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
         Left            =   5760
         TabIndex        =   13
         Top             =   2040
         Width           =   735
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
         Left            =   840
         TabIndex        =   11
         Top             =   2040
         Width           =   855
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
         Left            =   840
         TabIndex        =   9
         Top             =   1680
         Width           =   855
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
         Left            =   5760
         TabIndex        =   7
         Top             =   960
         Width           =   615
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
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   1095
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
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Extras 
      Height          =   4335
      Left            =   2640
      TabIndex        =   99
      Top             =   1320
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Frame FrameConsulta 
         Height          =   3015
         Left            =   480
         TabIndex        =   109
         Top             =   720
         Width           =   5655
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
            Height          =   1500
            ItemData        =   "cliente.frx":7870
            Left            =   360
            List            =   "cliente.frx":7877
            TabIndex        =   113
            Top             =   600
            Visible         =   0   'False
            Width           =   5055
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
            Left            =   1560
            TabIndex        =   112
            Top             =   840
            Visible         =   0   'False
            Width           =   2655
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
            Left            =   360
            TabIndex        =   111
            Top             =   240
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.CommandButton btnCerrarConsulta 
            Caption         =   "Cerrar"
            Height          =   495
            Left            =   2040
            TabIndex        =   110
            Top             =   2280
            Width           =   1695
         End
      End
      Begin VB.Frame FrameListado 
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
         Height          =   1815
         Left            =   480
         TabIndex        =   100
         Top             =   1200
         Visible         =   0   'False
         Width           =   5655
         Begin VB.CommandButton Acepta 
            Caption         =   "Confirmar (F11)"
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
            Left            =   3720
            MouseIcon       =   "cliente.frx":7885
            MousePointer    =   99  'Custom
            Picture         =   "cliente.frx":7B8F
            TabIndex        =   106
            ToolTipText     =   "Graba los Datos Ingresados"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Cancela 
            Caption         =   "Cancelar (F12)"
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
            Left            =   3720
            MouseIcon       =   "cliente.frx":7FD1
            MousePointer    =   99  'Custom
            Picture         =   "cliente.frx":82DB
            TabIndex        =   105
            ToolTipText     =   "Graba los Datos Ingresados"
            Top             =   960
            Width           =   1215
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
            TabIndex        =   104
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
            TabIndex        =   103
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
            Left            =   1920
            TabIndex        =   102
            Top             =   1200
            Value           =   -1  'True
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
            Left            =   600
            TabIndex        =   101
            Top             =   1200
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
            TabIndex        =   108
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
            Height          =   255
            Left            =   360
            TabIndex        =   107
            Top             =   360
            Width           =   1215
         End
      End
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
Private Valida As Boolean

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

Private Function DatosBasicosValidos()
    Dim validos As Boolean
    validos = True
    
    ' Validamos Datos de Identificacion
    If Trim(Cliente.Text) = "" Or Trim(Razon.Text) = "" Or Trim(Fantasia.Text) = "" Then
        validos = False
    End If
    
    ' Datos de contacto
    If Trim(Direccion.Text) = "" Or Trim(Localidad.Text) = "" Or Trim(Postal.Text) = "" Or Provincia.ListIndex < 0 Or Provincia.ListIndex > 25 Then
        validos = False
    End If
    
    ' Datos de Entrega.
    If Trim(DireccionII.Text) = "" Or Trim(LocalidadII.Text) = "" Or Trim(PostalII.Text) = "" Or ProvinciaII.ListIndex < 0 Or ProvinciaII.ListIndex > 25 Then
        validos = False
    End If
    
    ' Datos impositivos (Alguno tiene que se elegido)
    If Iva1.Value = False And Iva2.Value = False And Iva3.Value = False And Iva4.Value = False And Iva5.Value = False And Iva6.Value = False Then
        validos = False
    End If
    
    DatosBasicosValidos = validos
End Function

Sub Verifica_datos()

    ' Verificamos que los datos básicos hayan sido colocados.
    If Not DatosBasicosValidos() Then
        Valida = False
        MsgBox "La Razon social, el nombre de Fantasia, los datos de contacto y de imputacioón son obligatorios. Por favor, verifique y vuelva a intentarlo.", vbInformation
        Exit Sub
    End If
    
    ' Validar las fechas. ¿Son todas obligatorias?
    Auxi = "S"
    Valida_fecha FechaAlta, Auxi
    
    If Auxi = "N" Then
        Valida = False
        MsgBox "La Fecha de Alta no tiene un formato o valor valido", vbInformation
        Exit Sub
    End If
    
    ' Validar que Condicion de Pago exista.
    Auxi = "S"
    If Trim(Condicion.Text) = "" Then
        MsgBox "Debe indicarse una condición de Pago.", vbInformation
        Valida = False
        Exit Sub
    Else
        Valida_Condicion Condicion.Text, Auxi
    End If
    
    If Auxi = "N" Then
        Valida = False
        MsgBox "La Condición de Pago indicada no es válida.", vbInformation
        Exit Sub
    End If
    
    ' Validar que se coloque y que lista exista.
    Auxi = "S"
    If Trim(NroLista.Text) = "" Then
        Valida = False
        MsgBox "Debe indicarse una Lista de Precios..", vbInformation
        Exit Sub
    Else
        Valida_Lista NroLista.Text, Auxi
    End If
    
    If Auxi = "N" Then
        Valida = False
        MsgBox "La Lista de Precios indicada no es válida.", vbInformation
        Exit Sub
    End If
    
    ' Validar el Cuit en caso de colocar alguno. (¿Como en Administración?)
    If Trim(Cuit.Text) <> "" Then
    
        Auxi = "S"
        
        verifica_cuit Cuit.Text, Auxi
        
        If Auxi = "N" Then
            Valida = False
            MsgBox "La CUIT indicado no es válido.", vbInformation
            Exit Sub
        End If
    End If
    
End Sub

Private Sub Valida_Lista(Lista As String, Valida As String)

    ZSql = ""
    ZSql = ZSql + "Select Codigo"
    ZSql = ZSql + " FROM Lista"
    ZSql = ZSql + " Where Codigo = " + "'" + Trim(Lista) + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Valida = "S"
    Else
        Valida = "N"
    End If
    
    rstCliente.Close

End Sub

Private Sub Valida_Condicion(Condicion As String, Valida As String)

    ZSql = ""
    ZSql = ZSql + "Select Codigo"
    ZSql = ZSql + " FROM CondPago"
    ZSql = ZSql + " Where Codigo = " + "'" + Trim(Condicion) + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Valida = "S"
    Else
        Valida = "N"
    End If
    
    rstCliente.Close

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
        txtNombreContacto.Text = Trim(rstCliente!Responsable)
        txtPaginaWeb.Text = Trim(rstCliente!PaginaWeb)
        Observaciones.Text = Trim(rstCliente!Observaciones)
        ObservacionesII.Text = Trim(rstCliente!ObservacionesII)
        Cuit.Mask = ""
        Cuit.Text = Trim(rstCliente!Cuit)
        Cuit.Mask = "##-########-#" ' Volvemos a colocar la mascara al campo.
        EMail.Text = Trim(rstCliente!EMail)
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
    
    If Trim(Hasta.Text) = "" Or Val(Hasta.Text) = 0 Then
        Hasta.Text = "ZZZZ"
    End If
    
    Listado.ReportFileName = App.Path + "/cliente.rpt"
    
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
    FrameListado.Visible = False
    Extras.Visible = False
    
End Sub

Private Sub Bonifica_Click()
    ZZPasaCliente = Cliente
    ZZPasaProceso = 0
    PrgClienteBonifica.Show
End Sub

Private Sub btnCerrarConsulta_Click()
    Extras.Visible = False
    FrameConsulta.Visible = False
End Sub

Private Sub Cancela_Click()
    FrameListado.Visible = False
    Extras.Visible = False
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

    ' DEFINIR QUE DATOS SE VAN A VALIDAR.
    Call Verifica_datos
    
    If Not Valida Then
        Exit Sub
    End If
    
    Exit Sub
    
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
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Trim(Cliente.Text) + "'"
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
        ZSql = ZSql + " Responsable = " + "'" + txtNombreContacto.Text + "',"
        ZSql = ZSql + " PaginaWeb = " + "'" + txtPaginaWeb.Text + "',"
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
        ZSql = ZSql + " ObservacionesII = " + "'" + ObservacionesII.Text + "',"
        ZSql = ZSql + " Cuit = " + "'" + Cuit.Text + "',"
        ZSql = ZSql + " Email = " + "'" + EMail.Text + "',"
        ZSql = ZSql + " Fax = " + "'" + fax.Text + "',"
        ZSql = ZSql + " PorceIva = " + "'" + PorceIva.Text + "',"
        ZSql = ZSql + " Provincia = " + "'" + Mid$(Str$(Provincia.ListIndex), 2, 2) + "',"
        ZSql = ZSql + " ProvinciaII = " + "'" + Mid$(Str$(ProvinciaII.ListIndex), 2, 2) + "',"
        ZSql = ZSql + " Iva = " + "'" + WIva + "',"
        ZSql = ZSql + " Expreso = '',"
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
        ZSql = ZSql + "Responsable ,"
        ZSql = ZSql + "PaginaWeb ,"
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
        ZSql = ZSql + "ObservacionesII ,"
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
        ZSql = ZSql + "'" + txtNombreContacto.Text + "',"
        ZSql = ZSql + "'" + txtPaginaWeb.Text + "',"
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
        ZSql = ZSql + "'" + ObservacionesII.Text + "',"
        ZSql = ZSql + "'" + ZZFechaAlta + "',"
        ZSql = ZSql + "'" + Cuit.Text + "',"
        ZSql = ZSql + "'" + EMail.Text + "',"
        ZSql = ZSql + "'" + fax.Text + "',"
        ZSql = ZSql + "'" + PorceIva.Text + "',"
        ZSql = ZSql + "'" + Mid$(Str$(Provincia.ListIndex), 2, 2) + "',"
        ZSql = ZSql + "'" + Mid$(Str$(ProvinciaII.ListIndex), 2, 2) + "',"
        ZSql = ZSql + "'" + WIva + "',"
        ZSql = ZSql + "'',"
        ZSql = ZSql + "'" + TipoClie.Text + "',"
        ZSql = ZSql + "'" + NroLista.Text + "',"
        ZSql = ZSql + "'" + Condicion.Text + "')"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    
    Call CmdLimpiar_Click
    
    
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
'    Cuit.Text = ""
    EMail.Text = ""
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
    Extras.Visible = True
    Extras.ZOrder 0
    FrameConsulta.Visible = False
    FrameListado.Visible = True
    Desde.SetFocus
End Sub

Private Sub Razon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fantasia.SetFocus
    End If
    If KeyAscii = 27 Then
        Razon.Text = ""
    End If
End Sub

Private Sub Direccion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Postal.SetFocus
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
        Telefono.SetFocus
    End If
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NroLista_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = "SELECT Descripcion FROM Lista WHERE Codigo = '" & Trim(NroLista.Text) & "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesListaPrecios.Caption = rstCliente!Descripcion
            rstCliente.Close
            TipoClie.SetFocus
        Else
            NroLista.SetFocus
        End If
        
        Exit Sub
        
    End If
    If KeyAscii = 27 Then
        NroLista.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Postal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Localidad.SetFocus
    End If
    If KeyAscii = 27 Then
        Postal.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Replace(Cuit.Text, "-", "")) <> "" Then
        
            ' Validar el Cuit en caso de colocar alguno. (¿Como en Administración?)
            
            Auxi = "S"
                        
            If Len(Trim(Cuit.Text)) <> 13 Then
                Exit Sub
            End If
            
            verifica_cuit Cuit.Text, Auxi
            
            If Auxi = "N" Then
                Valida = False
                MsgBox "La CUIT indicado no es válido.", vbInformation
                Exit Sub
            End If
        
            Auxi = "S"
            If Auxi = "S" Then
                PorceIva.SetFocus
                    Else
                Cuit.SetFocus
            End If
                Else
            PorceIva.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cuit.Mask = ""
        Cuit.Text = ""
        Cuit.Mask = "##-########-#"
    End If
End Sub

Private Sub Telefono_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNombreContacto.SetFocus
    End If
    If KeyAscii = 27 Then
        Telefono.Text = ""
    End If
End Sub

Private Sub EMail_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPaginaWeb.SetFocus
    End If
    If KeyAscii = 27 Then
        EMail.Text = ""
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

Private Sub PorceIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva1.SetFocus
    End If
    If KeyAscii = 27 Then
        PorceIva.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuit.SetFocus
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
            DesCondicion.Caption = ""
            Condicion.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Condicion.Text = ""
        DesCondicion.Caption = ""
    End If
End Sub

Private Sub Fantasia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Direccion.SetFocus
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
    
        If Trim(Cliente.Text) = "" Then
            Exit Sub
        End If
    
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
    Opcion.AddItem "Lista de Precios"
    
    Extras.Visible = True
    Extras.ZOrder 0
    FrameConsulta.Visible = True
    Ayuda.Visible = False
    Pantalla.Visible = False
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
        
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lista"
            ZSql = ZSql + " Order by Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
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
                rstCondPago.Close
            End If
        
        Case Else
    End Select
    
    Call Mostrar_Resultados
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Mostrar_Resultados()
    Extras.Visible = True
    Extras.ZOrder 0
    FrameListado.Visible = False
    FrameConsulta.Visible = True
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
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
    
    FrameConsulta.Visible = False
    Extras.Visible = False
    
End Sub

Sub Form_Load()

    Valida = True

    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Observaciones.Text = ""
    EMail.Text = ""
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
    
    SSTab1.TabEnabled(0) = True
    
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
    Extras.ZOrder 0
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub NroLista_DblClick()

' FALTA VER BIEN QUE ONDA CON LOS CLICKS Y LAS AYUDAS, ME PARECE QE ESTAN COLGADOS.

    Opcion.Clear
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Expreso"
    Opcion.AddItem "TipoClie"
    Opcion.AddItem "Condicion"
    Opcion.AddItem "Listas"
    Rem Opcion.Visible = True
    Extras.ZOrder 0
    Opcion.ListIndex = 4
    
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

Private Sub txtNombreContacto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub txtNombreContacto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EMail.SetFocus
    End If
    If KeyAscii = 27 Then
        txtNombreContacto.Text = ""
    End If
End Sub



Private Sub txtPaginaWeb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub txtPaginaWeb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fax.SetFocus
    End If
    If KeyAscii = 27 Then
        txtPaginaWeb.Text = ""
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
    
    PantaContacto.Visible = True
    PantaContacto.ZOrder 0
    
    NombreI.SetFocus

End Sub
