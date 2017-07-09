VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgParametro 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de datos del la Empresa"
   ClientHeight    =   7380
   ClientLeft      =   645
   ClientTop       =   720
   ClientWidth     =   10605
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   10605
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
      Left            =   6120
      MouseIcon       =   "Parametro.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Parametro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Salida"
      Top             =   6240
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
      Left            =   3720
      MouseIcon       =   "Parametro.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "Parametro.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6240
      Width           =   855
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9763
      _Version        =   327680
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "Parametro.frx":1698
      Tab(0).ControlCount=   18
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(12)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(7)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabels(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabels(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabels(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabels(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabels(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabels(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "InicioAct"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Telefono"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "CondIva"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Nombre"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Direccion"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Localidad"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Cuit"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Actividad"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "IngBrutos"
      Tab(0).Control(17).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "Parametro.frx":16B4
      Tab(1).ControlCount=   40
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabels(14)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabels(25)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabels(28)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblLabels(27)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabels(26)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblLabels(24)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblLabels(21)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblLabels(20)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblLabels(19)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblLabels(18)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblLabels(17)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblLabels(16)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblLabels(13)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lblLabels(11)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblLabels(10)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lblLabels(9)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblLabels(8)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "lblLabels(15)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "lblLabels(22)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "lblLabels(23)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "CtaFondoFijo"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "CtaIva105"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "CtaGanancia"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "CtaIvaven"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "CtaVentas"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "CtaIb"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "CtaIva27"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "CtaIva5"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "CtaIva21"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "CtaProveedores"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "CtaDocumentos"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "CtaCheque"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "CtaEfectivo"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "CtaDeudores"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "CtaRetOtro"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "CtaRetIva"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "CtaRetGan"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "CtaImpCombustible"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "CtaImpInterno"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "CtaRetSuss"
      Tab(1).Control(39).Enabled=   0   'False
      Begin VB.TextBox CtaRetSuss 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   59
         Text            =   " "
         Top             =   3120
         Width           =   2200
      End
      Begin VB.TextBox CtaImpInterno 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   57
         Text            =   " "
         Top             =   2450
         Width           =   2200
      End
      Begin VB.TextBox CtaImpCombustible 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   55
         Text            =   " "
         Top             =   2800
         Width           =   2200
      End
      Begin VB.TextBox IngBrutos 
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
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   27
         Text            =   " "
         Top             =   2820
         Width           =   2295
      End
      Begin VB.TextBox Actividad 
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
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   26
         Text            =   " "
         Top             =   2460
         Width           =   6015
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
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   25
         Text            =   " "
         Top             =   2100
         Width           =   2295
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
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   24
         Text            =   " "
         Top             =   1380
         Width           =   6015
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
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   23
         Text            =   " "
         Top             =   1020
         Width           =   6015
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
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   22
         Text            =   " "
         Top             =   660
         Width           =   6015
      End
      Begin VB.TextBox CtaRetGan 
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   20
         TabIndex        =   21
         Text            =   " "
         Top             =   660
         Width           =   2200
      End
      Begin VB.TextBox CtaRetIva 
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   20
         TabIndex        =   20
         Text            =   " "
         Top             =   1020
         Width           =   2200
      End
      Begin VB.TextBox CtaRetOtro 
         BackColor       =   &H00FFFFFF&
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
         MaxLength       =   20
         TabIndex        =   19
         Text            =   " "
         Top             =   1380
         Width           =   2200
      End
      Begin VB.TextBox CtaDeudores 
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
         MaxLength       =   20
         TabIndex        =   18
         Text            =   " "
         Top             =   1740
         Width           =   2200
      End
      Begin VB.TextBox CondIva 
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
         Left            =   3120
         MaxLength       =   50
         TabIndex        =   17
         Text            =   " "
         Top             =   3540
         Width           =   6015
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
         Height          =   285
         Left            =   3120
         MaxLength       =   30
         TabIndex        =   16
         Text            =   " "
         Top             =   1740
         Width           =   3495
      End
      Begin VB.TextBox CtaEfectivo 
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
         MaxLength       =   20
         TabIndex        =   15
         Text            =   " "
         Top             =   2100
         Width           =   2200
      End
      Begin VB.TextBox CtaCheque 
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
         MaxLength       =   20
         TabIndex        =   14
         Text            =   " "
         Top             =   2460
         Width           =   2200
      End
      Begin VB.TextBox CtaDocumentos 
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
         MaxLength       =   20
         TabIndex        =   13
         Text            =   " "
         Top             =   2820
         Width           =   2200
      End
      Begin VB.TextBox CtaProveedores 
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
         MaxLength       =   20
         TabIndex        =   12
         Text            =   " "
         Top             =   3180
         Width           =   2200
      End
      Begin VB.TextBox CtaIva21 
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
         MaxLength       =   20
         TabIndex        =   11
         Text            =   " "
         Top             =   3540
         Width           =   2200
      End
      Begin VB.TextBox CtaIva5 
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
         MaxLength       =   20
         TabIndex        =   10
         Text            =   " "
         Top             =   3900
         Width           =   2200
      End
      Begin VB.TextBox CtaIva27 
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
         MaxLength       =   20
         TabIndex        =   9
         Text            =   " "
         Top             =   4260
         Width           =   2200
      End
      Begin VB.TextBox CtaIb 
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   8
         Text            =   " "
         Top             =   660
         Width           =   2200
      End
      Begin VB.TextBox CtaVentas 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   7
         Text            =   " "
         Top             =   1740
         Width           =   2200
      End
      Begin VB.TextBox CtaIvaven 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   6
         Text            =   " "
         Top             =   1380
         Width           =   2200
      End
      Begin VB.TextBox CtaGanancia 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   5
         Text            =   " "
         Top             =   1020
         Width           =   2200
      End
      Begin VB.TextBox CtaIva105 
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
         MaxLength       =   20
         TabIndex        =   4
         Text            =   " "
         Top             =   4620
         Width           =   2200
      End
      Begin VB.TextBox CtaFondoFijo 
         BackColor       =   &H00FFFFFF&
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
         Left            =   -67320
         MaxLength       =   20
         TabIndex        =   3
         Text            =   " "
         Top             =   2100
         Width           =   2200
      End
      Begin MSMask.MaskEdBox InicioAct 
         Height          =   285
         Left            =   3120
         TabIndex        =   28
         Top             =   3180
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label lblLabels 
         Caption         =   "Retencion de Suss"
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
         Index           =   23
         Left            =   -69840
         TabIndex        =   60
         Top             =   3150
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Otros Conceptos I"
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
         Index           =   22
         Left            =   -69840
         TabIndex        =   58
         Top             =   2450
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Otros Conceptos II"
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
         Index           =   15
         Left            =   -69840
         TabIndex        =   56
         Top             =   2760
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Inicio de Actividades"
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
         Left            =   600
         TabIndex        =   54
         Top             =   3180
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nro de Ingresos Brutos"
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
         Left            =   600
         TabIndex        =   53
         Top             =   2820
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Actividad"
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
         Left            =   600
         TabIndex        =   52
         Top             =   2460
         Width           =   2415
      End
      Begin VB.Label lblLabels 
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
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   51
         Top             =   2100
         Width           =   2415
      End
      Begin VB.Label lblLabels 
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
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   50
         Top             =   1380
         Width           =   2415
      End
      Begin VB.Label lblLabels 
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
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   49
         Top             =   1020
         Width           =   2415
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
         Index           =   2
         Left            =   600
         TabIndex        =   48
         Top             =   660
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ret. de Ganancias Recibidas"
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
         Index           =   8
         Left            =   -74880
         TabIndex        =   47
         Top             =   660
         Width           =   2655
      End
      Begin VB.Label lblLabels 
         Caption         =   "Retencion de Iva Recibidas"
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
         Index           =   9
         Left            =   -74880
         TabIndex        =   46
         Top             =   1020
         Width           =   3855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Retencion de I.Brutos"
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
         Index           =   10
         Left            =   -74880
         TabIndex        =   45
         Top             =   1380
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Deudores por Ventas"
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
         Index           =   11
         Left            =   -74880
         TabIndex        =   44
         Top             =   1740
         Width           =   2415
      End
      Begin VB.Label lblLabels 
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
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   43
         Top             =   3540
         Width           =   2415
      End
      Begin VB.Label lblLabels 
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
         Index           =   12
         Left            =   600
         TabIndex        =   42
         Top             =   1740
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Efectivo"
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
         Index           =   13
         Left            =   -74880
         TabIndex        =   41
         Top             =   2100
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Valores a Depositar"
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
         Index           =   16
         Left            =   -74880
         TabIndex        =   40
         Top             =   2460
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Documentos de Terceros"
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
         Index           =   17
         Left            =   -74880
         TabIndex        =   39
         Top             =   2820
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Proveedores"
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
         Index           =   18
         Left            =   -74880
         TabIndex        =   38
         Top             =   3180
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Iva Compras Insc."
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
         Index           =   19
         Left            =   -74880
         TabIndex        =   37
         Top             =   3540
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Percep. de Iva Recibida"
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
         Index           =   20
         Left            =   -74880
         TabIndex        =   36
         Top             =   3900
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Iva Compras Serv."
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
         Index           =   21
         Left            =   -74880
         TabIndex        =   35
         Top             =   4260
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Per. de Ing.Brutos Recibida"
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
         Index           =   24
         Left            =   -69840
         TabIndex        =   34
         Top             =   660
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ventas"
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
         Index           =   26
         Left            =   -69840
         TabIndex        =   33
         Top             =   1740
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Iva Ventas"
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
         Index           =   27
         Left            =   -69840
         TabIndex        =   32
         Top             =   1380
         Width           =   3855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Retencion de Ganancias"
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
         Index           =   28
         Left            =   -69840
         TabIndex        =   31
         Top             =   1020
         Width           =   3615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Iva Compras Esp."
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
         Index           =   25
         Left            =   -74880
         TabIndex        =   30
         Top             =   4620
         Width           =   3255
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fondo Fijo"
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
         Index           =   14
         Left            =   -69840
         TabIndex        =   29
         Top             =   2100
         Width           =   3735
      End
   End
End
Attribute VB_Name = "PrgParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Verifica_datos()
    Rem If Val(Cuenta.text) = 0 Then
    Rem     Cuenta.text = "0"
    Rem End If
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("#,###,###.##", Comision.text)
End Sub

Private Sub cmdAdd_Click()

    ZZNombreBase = WNombreBase

    txtOdbc = "Empresa"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    ZSql = ""
    ZSql = ZSql + "UPDATE empresa SET "
    ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "',"
    ZSql = ZSql + " Direccion = " + "'" + Direccion.Text + "',"
    ZSql = ZSql + " Localidad = " + "'" + Localidad.Text + "',"
    ZSql = ZSql + " Telefono = " + "'" + Telefono.Text + "',"
    ZSql = ZSql + " Cuit = " + "'" + Cuit.Text + "',"
    ZSql = ZSql + " Actividad = " + "'" + Actividad.Text + "',"
    ZSql = ZSql + " IngBrutos = " + "'" + IngBrutos.Text + "',"
    ZSql = ZSql + " InicioAct = " + "'" + InicioAct.Text + "',"
    ZSql = ZSql + " CondIva = " + "'" + CondIva.Text + "',"
    ZSql = ZSql + " CtaRetGan = " + "'" + CtaRetGan.Text + "',"
    ZSql = ZSql + " CtaRetIva = " + "'" + CtaRetIva.Text + "',"
    ZSql = ZSql + " CtaRetOtro = " + "'" + CtaRetOtro.Text + "',"
    ZSql = ZSql + " CtaDeudores = " + "'" + CtaDeudores.Text + "',"
    ZSql = ZSql + " CtaEfectivo = " + "'" + CtaEfectivo.Text + "',"
    ZSql = ZSql + " CtaCheque = " + "'" + CtaCheque.Text + "',"
    ZSql = ZSql + " CtaDocumentos = " + "'" + CtaDocumentos.Text + "',"
    ZSql = ZSql + " CtaProveedores = " + "'" + CtaProveedores.Text + "',"
    ZSql = ZSql + " CtaIva21 = " + "'" + CtaIva21.Text + "',"
    ZSql = ZSql + " CtaIva5 = " + "'" + CtaIva5.Text + "',"
    ZSql = ZSql + " CtaIva27 = " + "'" + CtaIva27.Text + "',"
    ZSql = ZSql + " CtaIva105 = " + "'" + CtaIva105.Text + "',"
    ZSql = ZSql + " CtaIb = " + "'" + CtaIb.Text + "',"
    ZSql = ZSql + " CtaGanancia = " + "'" + CtaGanancia.Text + "',"
    ZSql = ZSql + " CtaIvaVen = " + "'" + CtaIvaven.Text + "',"
    ZSql = ZSql + " CtaVentas = " + "'" + CtaVentas.Text + "',"
    ZSql = ZSql + " CtaFondoFijo = " + "'" + CtaFondoFijo.Text + "',"
    ZSql = ZSql + " CtaImpInterno = " + "'" + CtaImpInterno.Text + "',"
    ZSql = ZSql + " CtaImpCombustible = " + "'" + CtaImpCombustible.Text + "',"
    ZSql = ZSql + " CtaRetSuss = " + "'" + CtaRetSuss.Text + "'"
    ZSql = ZSql + " Where Empresa = " + "'" + WEmpresa + "'"
    
    spEmpresa = ZSql
    Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
        
    WNombreBase = ZZNombreBase
    
    txtOdbc = WNombreBase
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    

    ZSql = ""
    ZSql = ZSql + "UPDATE empresa SET "
    ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "',"
    ZSql = ZSql + " Direccion = " + "'" + Direccion.Text + "',"
    ZSql = ZSql + " Localidad = " + "'" + Localidad.Text + "',"
    ZSql = ZSql + " Telefono = " + "'" + Telefono.Text + "',"
    ZSql = ZSql + " Cuit = " + "'" + Cuit.Text + "',"
    ZSql = ZSql + " Actividad = " + "'" + Actividad.Text + "',"
    ZSql = ZSql + " IngBrutos = " + "'" + IngBrutos.Text + "',"
    ZSql = ZSql + " InicioAct = " + "'" + InicioAct.Text + "',"
    ZSql = ZSql + " CondIva = " + "'" + CondIva.Text + "',"
    ZSql = ZSql + " CtaRetGan = " + "'" + CtaRetGan.Text + "',"
    ZSql = ZSql + " CtaRetIva = " + "'" + CtaRetIva.Text + "',"
    ZSql = ZSql + " CtaRetOtro = " + "'" + CtaRetOtro.Text + "',"
    ZSql = ZSql + " CtaDeudores = " + "'" + CtaDeudores.Text + "',"
    ZSql = ZSql + " CtaEfectivo = " + "'" + CtaEfectivo.Text + "',"
    ZSql = ZSql + " CtaCheque = " + "'" + CtaCheque.Text + "',"
    ZSql = ZSql + " CtaDocumentos = " + "'" + CtaDocumentos.Text + "',"
    ZSql = ZSql + " CtaProveedores = " + "'" + CtaProveedores.Text + "',"
    ZSql = ZSql + " CtaIva21 = " + "'" + CtaIva21.Text + "',"
    ZSql = ZSql + " CtaIva5 = " + "'" + CtaIva5.Text + "',"
    ZSql = ZSql + " CtaIva27 = " + "'" + CtaIva27.Text + "',"
    ZSql = ZSql + " CtaIva105 = " + "'" + CtaIva105.Text + "',"
    ZSql = ZSql + " CtaIb = " + "'" + CtaIb.Text + "',"
    ZSql = ZSql + " CtaGanancia = " + "'" + CtaGanancia.Text + "',"
    ZSql = ZSql + " CtaIvaVen = " + "'" + CtaIvaven.Text + "',"
    ZSql = ZSql + " CtaVentas = " + "'" + CtaVentas.Text + "',"
    ZSql = ZSql + " CtaFondoFijo = " + "'" + CtaFondoFijo.Text + "',"
    ZSql = ZSql + " CtaImpInterno = " + "'" + CtaImpInterno.Text + "',"
    ZSql = ZSql + " CtaImpCombustible = " + "'" + CtaImpCombustible.Text + "',"
    ZSql = ZSql + " CtaRetSuss = " + "'" + CtaRetSuss.Text + "'"
    ZSql = ZSql + " Where Empresa = " + "'" + WEmpresa + "'"
    
    spEmpresa = ZSql
    Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    Call CmdLimpiar_Click
    Call CmdClose_Click
    
End Sub

Private Sub CmdLimpiar_Click()

    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Telefono.Text = ""
    Cuit.Text = ""
    Actividad.Text = ""
    IngBrutos.Text = ""
    InicioAct.Text = "  /  /    "
    CondIva.Text = ""
                
    CtaRetGan.Text = ""
    CtaRetIva.Text = ""
    CtaRetOtro.Text = ""
    CtaDeudores.Text = ""
    CtaEfectivo.Text = ""
    CtaCheque.Text = ""
    CtaDocumentos.Text = ""
    CtaProveedores.Text = ""
    CtaIva21.Text = ""
    CtaIva5.Text = ""
    CtaIva27.Text = ""
    CtaIva105.Text = ""
    CtaIb.Text = ""
    CtaGanancia.Text = ""
    CtaIvaven.Text = ""
    CtaVentas.Text = ""
    CtaFondoFijo.Text = ""
    CtaImpInterno.Text = ""
    CtaImpCombustible.Text = ""
    CtaRetSuss.Text = ""
    Tablas.Tab = 0
    Nombre.SetFocus
    
End Sub

Private Sub CmdClose_Click()
    Tablas.Tab = 0
    PrgParametro.Hide
    Unload Me
    MenuAdminis.Show
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
        Telefono.SetFocus
    End If
    If KeyAscii = 27 Then
        Localidad.Text = ""
    End If
End Sub

Private Sub Telefono_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuit.SetFocus
    End If
    If KeyAscii = 27 Then
        Telefono.Text = ""
    End If
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Actividad.SetFocus
    End If
    If KeyAscii = 27 Then
        Cuit.Text = ""
    End If
End Sub

Private Sub Actividad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IngBrutos.SetFocus
    End If
    If KeyAscii = 27 Then
        Actividad.Text = ""
    End If
End Sub

Private Sub IngBrutos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        InicioAct.SetFocus
    End If
    If KeyAscii = 27 Then
        IngBrutos.Text = ""
    End If
End Sub

Private Sub InicioAct_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CondIva.SetFocus
    End If
    If KeyAscii = 27 Then
        InicioAct.Text = "  /  /    "
    End If
End Sub

Private Sub CondIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        CondIva.Text = ""
    End If
End Sub

Private Sub CtaRetGan_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaRetIva.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaRetGan.Text = ""
    End If
End Sub

Private Sub CtaRetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaRetOtro.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaRetIva.Text = ""
    End If
End Sub

Private Sub CtaRetOtro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaDeudores.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaRetOtro.Text = ""
    End If
End Sub

Private Sub CtaDeudores_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaEfectivo.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaDeudores.Text = ""
    End If
End Sub

Private Sub CtaEfectivo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaCheque.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaEfectivo.Text = ""
    End If
End Sub

Private Sub CtaCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaDocumentos.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaCheque.Text = ""
    End If
End Sub

Private Sub CtaDocumentos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaProveedores.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaDocumentos.Text = ""
    End If
End Sub

Private Sub CtaProveedores_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaIva21.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaProveedores.Text = ""
    End If
End Sub

Private Sub CtaIva21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaIva5.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaIva21.Text = ""
    End If
End Sub

Private Sub CtaIva5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaIva27.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaIva5.Text = ""
    End If
End Sub

Private Sub CtaIva27_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaIva105.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaIva27.Text = ""
    End If
End Sub

Private Sub CtaIva105_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaIb.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaIva105.Text = ""
    End If
End Sub

Private Sub CtaIb_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaGanancia.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaIb.Text = ""
    End If
End Sub

Private Sub CtaGanancia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaIvaven.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaGanancia.Text = ""
    End If
End Sub

Private Sub CtaIvaven_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaVentas.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaIvaven.Text = ""
    End If
End Sub

Private Sub Ctaventas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaFondoFijo.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaVentas.Text = ""
    End If
End Sub

Private Sub CtaFondoFijo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaImpInterno.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaFondoFijo.Text = ""
    End If
End Sub

Private Sub CtaImpInterno_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaImpCombustible.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaImpInterno.Text = ""
    End If
End Sub

Private Sub CtaImpCombustible_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaRetSuss.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaImpCombustible.Text = ""
    End If
End Sub

Private Sub CtaRetSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CtaRetGan.SetFocus
    End If
    If KeyAscii = 27 Then
        CtaRetSuss.Text = ""
    End If
End Sub

Sub Form_Load()

    Tablas.TabCaption(0) = "Datos de la Empresa"
    Tablas.TabCaption(1) = "Imputacioines Contables"

    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Telefono.Text = ""
    Cuit.Text = ""
    Actividad.Text = ""
    IngBrutos.Text = ""
    InicioAct.Text = "  /  /    "
    CondIva.Text = ""
                
    CtaRetGan.Text = ""
    CtaRetIva.Text = ""
    CtaRetOtro.Text = ""
    CtaDeudores.Text = ""
    CtaEfectivo.Text = ""
    CtaCheque.Text = ""
    CtaDocumentos.Text = ""
    CtaProveedores.Text = ""
    CtaIva21.Text = ""
    CtaIva5.Text = ""
    CtaIva27.Text = ""
    CtaIva105.Text = ""
    CtaIb.Text = ""
    CtaGanancia.Text = ""
    CtaIvaven.Text = ""
    CtaVentas.Text = ""
    CtaFondoFijo.Text = ""
    CtaImpInterno.Text = ""
    CtaImpCombustible.Text = ""
    CtaRetSuss.Text = ""
    
    Tablas.Tab = 0
    
    ZZNombreBase = WNombreBase

    txtOdbc = "Empresa"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Empresa"
    ZSql = ZSql + " Where Empresa.Empresa = " + "'" + WEmpresa + "'"
    spEmpresa = ZSql
    Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
    If rstEmpresa.RecordCount > 0 Then
    
        Nombre.Text = rstEmpresa!Nombre
        Direccion.Text = rstEmpresa!Direccion
        Localidad.Text = rstEmpresa!Localidad
        Telefono.Text = rstEmpresa!Telefono
        Cuit.Text = rstEmpresa!Cuit
        Actividad.Text = rstEmpresa!Actividad
        IngBrutos.Text = rstEmpresa!IngBrutos
        InicioAct.Text = rstEmpresa!InicioAct
        CondIva.Text = rstEmpresa!CondIva
                
        CtaRetGan = rstEmpresa!CtaRetGan
        CtaRetIva = rstEmpresa!CtaRetIva
        CtaRetOtro = rstEmpresa!CtaRetOtro
        CtaDeudores = rstEmpresa!CtaDeudores
        CtaEfectivo = rstEmpresa!CtaEfectivo
        CtaCheque = rstEmpresa!CtaCheque
        CtaDocumentos = rstEmpresa!CtaDocumentos
        CtaProveedores = rstEmpresa!CtaProveedores
        CtaIva21 = rstEmpresa!CtaIva21
        CtaIva5 = rstEmpresa!CtaIva5
        CtaIva27 = rstEmpresa!CtaIva27
        CtaIva105 = rstEmpresa!CtaIva105
        CtaIb.Text = rstEmpresa!CtaIb
        CtaGanancia = rstEmpresa!CtaGanancia
        CtaIvaven = rstEmpresa!CtaIvaven
        CtaVentas = rstEmpresa!CtaVentas
        CtaFondoFijo = rstEmpresa!CtaFondoFijo
        CtaImpInterno = rstEmpresa!CtaImpInterno
        CtaImpCombustible = rstEmpresa!CtaImpCombustible
        CtaRetSuss = rstEmpresa!CtaRetSuss
        
        rstEmpresa.Close
    End If
    
    WNombreBase = ZZNombreBase
    
    txtOdbc = WNombreBase
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    On Error GoTo WError
    
    Select Case Tablas.Tab
        Case 0
            Nombre.SetFocus
        Case 1
            CtaRetGan.SetFocus
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub



Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Direccion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Localidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Telefono_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cuit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Actividad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub IngBrutos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub InicioAct_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CondIva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaRetGan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaRetIva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaRetOtro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaDeudores_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaEfectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaCheque_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaDocumentos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaProveedores_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaIva21_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaIva5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaIva27_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva105_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaIb_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaGanancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaIvaven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaVentas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaFondoFijo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaImpInterno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaImpCombustible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CtaRetSuss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub






