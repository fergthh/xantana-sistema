VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgEscalas 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Configuracion del Sistema"
   ClientHeight    =   7380
   ClientLeft      =   645
   ClientTop       =   720
   ClientWidth     =   10605
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   10605
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
      Left            =   3840
      MouseIcon       =   "escalas.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "escalas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6120
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
      Left            =   6240
      MouseIcon       =   "escalas.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "escalas.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Salida"
      Top             =   6120
      Width           =   855
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5535
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   9763
      _Version        =   327680
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
      TabPicture(0)   =   "escalas.frx":1698
      Tab(0).ControlCount=   42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(7)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabels(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabels(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabels(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabels(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabels(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabels(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLabels(13)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblLabels(16)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblLabels(17)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblLabels(18)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblLabels(19)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblLabels(20)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblLabels(21)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Line2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line3"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lblLabels(27)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblLabels(29)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblLabels(30)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "RetMinima"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Escala4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Escala3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Escala2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Escala1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Minimo3"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Minimo2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Minimo1"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Escala5"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Tasa1"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Tasa2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Tasa3"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Tasa4"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Tasa5"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TasaGen"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TasaBienes"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Numero1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TasaNoInscripto"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Minimo4"
      Tab(0).Control(41).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "escalas.frx":16B4
      Tab(1).ControlCount=   10
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabels(8)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabels(9)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabels(10)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblLabels(11)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblLabels(28)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "PorceBienes"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "PorceServicios"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "PorceTranspo"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "MinimoIva"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Numero2"
      Tab(1).Control(9).Enabled=   -1  'True
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "escalas.frx":16D0
      Tab(2).ControlCount=   16
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLabels(14)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblLabels(15)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblLabels(22)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblLabels(23)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblLabels(24)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblLabels(25)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "lblLabels(26)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblLabels(31)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "IvaNoInscripto"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "IvaInscripto"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "Percepcion"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "Punto"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "CantiFac"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "CantiRem"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "CantiArti"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "IvaServicio"
      Tab(2).Control(15).Enabled=   -1  'True
      Begin VB.TextBox IvaServicio 
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
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   66
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Minimo4 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   64
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox TasaNoInscripto 
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
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   60
         Text            =   " "
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox Numero2 
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
         Left            =   -69600
         MaxLength       =   10
         TabIndex        =   58
         Text            =   " "
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Numero1 
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
         Left            =   8280
         MaxLength       =   10
         TabIndex        =   56
         Text            =   " "
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox CantiArti 
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
         Left            =   -70200
         MaxLength       =   4
         TabIndex        =   54
         Text            =   " "
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox CantiRem 
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
         Left            =   -70200
         MaxLength       =   4
         TabIndex        =   52
         Text            =   " "
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox CantiFac 
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
         Left            =   -70200
         MaxLength       =   4
         TabIndex        =   50
         Text            =   " "
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox Punto 
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
         Left            =   -70200
         MaxLength       =   4
         TabIndex        =   48
         Text            =   " "
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox Percepcion 
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
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   46
         Text            =   " "
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox TasaBienes 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   38
         Text            =   " "
         Top             =   4440
         Width           =   1215
      End
      Begin VB.TextBox TasaGen 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   36
         Text            =   " "
         Top             =   4080
         Width           =   1215
      End
      Begin VB.TextBox Tasa5 
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
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   35
         Text            =   " "
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox Tasa4 
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
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   34
         Text            =   " "
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Tasa3 
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
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   33
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Tasa2 
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
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   32
         Text            =   " "
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Tasa1 
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
         Left            =   6480
         MaxLength       =   10
         TabIndex        =   31
         Text            =   " "
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Escala5 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   29
         Text            =   " "
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox IvaInscripto 
         Alignment       =   1  'Right Justify
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
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox IvaNoInscripto 
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
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   26
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox MinimoIva 
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
         Left            =   -69600
         MaxLength       =   10
         TabIndex        =   24
         Text            =   " "
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox PorceTranspo 
         Alignment       =   1  'Right Justify
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
         Left            =   -69600
         MaxLength       =   10
         TabIndex        =   22
         Text            =   " "
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox PorceServicios 
         Alignment       =   1  'Right Justify
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
         Left            =   -69600
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox PorceBienes 
         Alignment       =   1  'Right Justify
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
         Left            =   -69600
         MaxLength       =   10
         TabIndex        =   18
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Minimo1 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Minimo2 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   9
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Minimo3 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   8
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Escala1 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   7
         Text            =   " "
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Escala2 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   6
         Text            =   " "
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox Escala3 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   5
         Text            =   " "
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Escala4 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   4
         Text            =   " "
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox RetMinima 
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
         Left            =   3120
         MaxLength       =   10
         TabIndex        =   3
         Text            =   " "
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcentaje Iva Servicios Publicos"
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
         Index           =   31
         Left            =   -74400
         TabIndex        =   67
         Top             =   1680
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo Pago Servicios"
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
         Index           =   30
         Left            =   600
         TabIndex        =   65
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tasa No Inscripto"
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
         Index           =   29
         Left            =   4560
         TabIndex        =   61
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ultimo Nro. de Retencion de Iva"
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
         Left            =   -74400
         TabIndex        =   59
         Top             =   2520
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ultimo Nro. de Retencion de Ganancias"
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
         Left            =   4560
         TabIndex        =   57
         Top             =   4920
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cantidad de Articulos en cada Comprobante"
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
         Left            =   -74400
         TabIndex        =   55
         Top             =   3480
         Width           =   3975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cantidad de Remitos"
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
         Left            =   -74400
         TabIndex        =   53
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cantidad de Facturas"
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
         Left            =   -74400
         TabIndex        =   51
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Punto de Venta"
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
         Left            =   -74400
         TabIndex        =   49
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcentaje de Percepcion de Iva"
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
         Left            =   -74400
         TabIndex        =   47
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "     Escala        Honorarios"
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
         Left            =   8280
         TabIndex        =   45
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Line Line3 
         X1              =   7920
         X2              =   8160
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line2 
         X1              =   8160
         X2              =   8160
         Y1              =   3840
         Y2              =   2160
      End
      Begin VB.Line Line1 
         X1              =   7920
         X2              =   8160
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Tasa 4"
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
         Left            =   4560
         TabIndex        =   44
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Tasa 3"
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
         Left            =   4560
         TabIndex        =   43
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Tasa 2"
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
         Left            =   4560
         TabIndex        =   42
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Tasa 1"
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
         Left            =   4560
         TabIndex        =   41
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Tasa 5"
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
         Left            =   4560
         TabIndex        =   40
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tasa Bienes y Servicios"
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
         Left            =   600
         TabIndex        =   39
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tasa General"
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
         Left            =   600
         TabIndex        =   37
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 5"
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
         TabIndex        =   30
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcentaje Iva Inscripto"
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
         Left            =   -74400
         TabIndex        =   28
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcentaje Iva No Insc."
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
         Left            =   -74400
         TabIndex        =   27
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Minimo IVA"
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
         Left            =   -74400
         TabIndex        =   25
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcentaje de trabajos s/inmueble ajeno"
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
         Left            =   -74400
         TabIndex        =   23
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcentaje locaciones - prestaciones de servicio"
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
         Left            =   -74400
         TabIndex        =   21
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcentaje Compra-Venta de cosas muebles"
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
         Left            =   -74400
         TabIndex        =   19
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo Pago Otros"
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
         TabIndex        =   17
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo Pago Honorarios"
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
         TabIndex        =   16
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo Pago Alquileres"
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
         TabIndex        =   15
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 1"
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
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 2"
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
         TabIndex        =   13
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 3"
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
         TabIndex        =   12
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 4"
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
         TabIndex        =   11
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Retencion Minima"
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
         TabIndex        =   10
         Top             =   4920
         Width           =   2415
      End
   End
End
Attribute VB_Name = "PrgEscalas"
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

    Call Verifica_datos
                
    ZSql = ""
    ZSql = ZSql + "UPDATE Parametro SET "
    ZSql = ZSql + " Minimo1 = " + "'" + Minimo1.Text + "',"
    ZSql = ZSql + " Minimo2 = " + "'" + Minimo2.Text + "',"
    ZSql = ZSql + " Minimo3 = " + "'" + Minimo3.Text + "',"
    ZSql = ZSql + " Minimo4 = " + "'" + Minimo4.Text + "',"
    ZSql = ZSql + " Escala1 = " + "'" + Escala1.Text + "',"
    ZSql = ZSql + " Escala2 = " + "'" + Escala2.Text + "',"
    ZSql = ZSql + " Escala3 = " + "'" + Escala3.Text + "',"
    ZSql = ZSql + " Escala4 = " + "'" + Escala4.Text + "',"
    ZSql = ZSql + " Escala5 = " + "'" + Escala5.Text + "',"
    ZSql = ZSql + " Tasa1 = " + "'" + Tasa1.Text + "',"
    ZSql = ZSql + " Tasa2 = " + "'" + Tasa2.Text + "',"
    ZSql = ZSql + " Tasa3 = " + "'" + Tasa3.Text + "',"
    ZSql = ZSql + " Tasa4 = " + "'" + Tasa4.Text + "',"
    ZSql = ZSql + " Tasa5 = " + "'" + Tasa5.Text + "',"
    ZSql = ZSql + " TasaGen = " + "'" + TasaGen.Text + "',"
    ZSql = ZSql + " TasaBienes = " + "'" + TasaBienes.Text + "',"
    ZSql = ZSql + " TasaNoInscripto = " + "'" + TasaNoInscripto.Text + "',"
    ZSql = ZSql + " RetMinima = " + "'" + RetMinima.Text + "',"
    ZSql = ZSql + " PorceBienes = " + "'" + PorceBienes.Text + "',"
    ZSql = ZSql + " PorceServicios = " + "'" + PorceServicios.Text + "',"
    ZSql = ZSql + " PorceTranspo = " + "'" + PorceTranspo.Text + "',"
    ZSql = ZSql + " MinimoIva = " + "'" + MinimoIva.Text + "',"
    ZSql = ZSql + " IvaInscripto = " + "'" + IvaInscripto.Text + "',"
    ZSql = ZSql + " IvaNoInscripto = " + "'" + IvaNoInscripto.Text + "',"
    ZSql = ZSql + " IvaServicio = " + "'" + IvaServicio.Text + "'"
    ZSql = ZSql + " Where Clave = 1"
    
    spParametro = ZSql
    Set rstParametro = db.OpenRecordset(spParametro, dbOpenSnapshot, dbSQLPassThrough)




                
    ZSql = ""
    ZSql = ZSql + "UPDATE Configuracion SET "
    ZSql = ZSql + " Iva1 = " + "'" + IvaInscripto.Text + "',"
    ZSql = ZSql + " Iva2 = " + "'" + IvaNoInscripto.Text + "',"
    ZSql = ZSql + " IvaServicio = " + "'" + IvaServicio.Text + "',"
    ZSql = ZSql + " Percepcion = " + "'" + Percepcion.Text + "',"
    ZSql = ZSql + " Punto = " + "'" + Punto.Text + "',"
    ZSql = ZSql + " CantiFac = " + "'" + CantiFac.Text + "',"
    ZSql = ZSql + " CantiRem = " + "'" + CantiRem.Text + "',"
    ZSql = ZSql + " CantiArti = " + "'" + CantiArti.Text + "'"
    ZSql = ZSql + " Where Clave = 1"
    
    spConfiguracion = ZSql
    Set rstConfiguracion = db.OpenRecordset(spConfiguracion, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE NroRet SET "
    ZSql = ZSql + " Numero = " + "'" + Numero1.Text + "'"
    ZSql = ZSql + " Where Clave = 1"
    spNroRet = ZSql
    Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE NroRet SET "
    ZSql = ZSql + " Numero = " + "'" + Numero2.Text + "'"
    ZSql = ZSql + " Where Clave = 2"
    spNroRet = ZSql
    Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)

    Call CmdLimpiar_Click
    Call CmdClose_Click
    
End Sub

Private Sub CmdLimpiar_Click()
    Minimo1.Text = ""
    Minimo2.Text = ""
    Minimo3.Text = ""
    Minimo4.Text = ""
    Escala1.Text = ""
    Escala2.Text = ""
    Escala3.Text = ""
    Escala4.Text = ""
    Escala5.Text = ""
    Tasa1.Text = ""
    Tasa2.Text = ""
    Tasa3.Text = ""
    Tasa4.Text = ""
    Tasa5.Text = ""
    TasaGen.Text = ""
    TasaBienes.Text = ""
    TasaNoInscripto.Text = ""
    RetMinima.Text = ""
    PorceBienes.Text = ""
    PorceServicios.Text = ""
    PorceTranspo.Text = ""
    MinimoIva.Text = ""
    IvaInscripto.Text = ""
    IvaNoInscripto.Text = ""
    IvaServicio.Text = ""
    Percepcion.Text = ""
    Punto.Text = ""
    CantiFac.Text = ""
    CantiRem.Text = ""
    CantiArti.Text = "2"
    
    Rem Minimo1.SetFocus
    Tablas.Tab = 0
    Minimo1.SetFocus
    
End Sub

Private Sub CmdClose_Click()
    PrgEscalas.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Minimo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Minimo1.Text = Pusing("#,###,###.##", Minimo1.Text)
        Minimo2.SetFocus
    End If
    If KeyAscii = 27 Then
        Minimo1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Minimo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Minimo2.Text = Pusing("#,###,###.##", Minimo2.Text)
        Minimo3.SetFocus
    End If
    If KeyAscii = 27 Then
        Minimo2.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Minimo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Minimo3.Text = Pusing("#,###,###.##", Minimo3.Text)
        Minimo4.SetFocus
    End If
    If KeyAscii = 27 Then
        Minimo3.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Minimo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Minimo4.Text = Pusing("#,###,###.##", Minimo4.Text)
        Escala1.SetFocus
    End If
    If KeyAscii = 27 Then
        Minimo4.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Escala1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Escala1.Text = Pusing("#,###,###.##", Escala1.Text)
        Tasa1.SetFocus
    End If
    If KeyAscii = 27 Then
        Escala1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Tasa1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tasa1.Text = Pusing("###,###.##", Tasa1.Text)
        Escala2.SetFocus
    End If
    If KeyAscii = 27 Then
        Tasa1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Escala2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Escala2.Text = Pusing("#,###,###.##", Escala2.Text)
        Tasa2.SetFocus
    End If
    If KeyAscii = 27 Then
        Escala2.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Tasa2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tasa2.Text = Pusing("###,###.##", Tasa2.Text)
        Escala3.SetFocus
    End If
    If KeyAscii = 27 Then
        Tasa2.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Escala3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Escala3.Text = Pusing("#,###,###.##", Escala3.Text)
        Tasa3.SetFocus
    End If
    If KeyAscii = 27 Then
        Escala3.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Tasa3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tasa3.Text = Pusing("###,###.##", Tasa3.Text)
        Escala4.SetFocus
    End If
    If KeyAscii = 27 Then
        Tasa3.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Escala4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Escala4.Text = Pusing("#,###,###.##", Escala4.Text)
        Tasa4.SetFocus
    End If
    If KeyAscii = 27 Then
        Escala4.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Tasa4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tasa4.Text = Pusing("###,###.##", Tasa4.Text)
        Escala5.SetFocus
    End If
    If KeyAscii = 27 Then
        Tasa4.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Escala5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Escala5.Text = Pusing("#,###,###.##", Escala5.Text)
        Tasa5.SetFocus
    End If
    If KeyAscii = 27 Then
        Escala5.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Tasa5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tasa5.Text = Pusing("###,###.##", Tasa5.Text)
        TasaGen.SetFocus
    End If
    If KeyAscii = 27 Then
        Tasa5.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub TasaGen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TasaGen.Text = Pusing("###,###.##", TasaGen.Text)
        TasaBienes.SetFocus
    End If
    If KeyAscii = 27 Then
        TasaGen.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub TasaBienes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TasaBienes.Text = Pusing("###,###.##", TasaBienes.Text)
        TasaNoInscripto.SetFocus
    End If
    If KeyAscii = 27 Then
        TasaBienes.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub TasaNoInscripto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TasaNoInscripto.Text = Pusing("###,###.##", TasaNoInscripto.Text)
        RetMinima.SetFocus
    End If
    If KeyAscii = 27 Then
        TasaNoInscripto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetMinima_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetMinima.Text = Pusing("#,###,###.##", RetMinima.Text)
        Numero1.SetFocus
    End If
    If KeyAscii = 27 Then
        RetMinima.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Minimo1.SetFocus
    End If
    If KeyAscii = 27 Then
        Numero1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub PorceBienes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorceBienes.Text = Pusing("#,###,###.##", PorceBienes.Text)
        PorceServicios.SetFocus
    End If
    If KeyAscii = 27 Then
        PorceBienes.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub PorceServicios_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorceServicios.Text = Pusing("#,###,###.##", PorceServicios.Text)
        PorceTranspo.SetFocus
    End If
    If KeyAscii = 27 Then
        PorceServicios.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub PorceTranspo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorceTranspo.Text = Pusing("#,###,###.##", PorceTranspo.Text)
        MinimoIva.SetFocus
    End If
    If KeyAscii = 27 Then
        PorceTranspo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub MinimoIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MinimoIva.Text = Pusing("#,###,###.##", MinimoIva.Text)
        Numero2.SetFocus
    End If
    If KeyAscii = 27 Then
        MinimoIva.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PorceBienes.SetFocus
    End If
    If KeyAscii = 27 Then
        Numero2.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub IvaInscripto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IvaInscripto.Text = Pusing("#,###,###.##", IvaInscripto.Text)
        IvaNoInscripto.SetFocus
    End If
    If KeyAscii = 27 Then
        IvaInscripto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub IvaNoInscripto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IvaNoInscripto.Text = Pusing("#,###,###.##", IvaNoInscripto.Text)
        IvaServicio.SetFocus
    End If
    If KeyAscii = 27 Then
        IvaNoInscripto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub IvaServicio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IvaServicio.Text = Pusing("#,###,###.##", IvaServicio.Text)
        Percepcion.SetFocus
    End If
    If KeyAscii = 27 Then
        IvaServicio.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Percepcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Percepcion.Text = Pusing("#,###,###.##", Percepcion.Text)
        Punto.SetFocus
    End If
    If KeyAscii = 27 Then
        Percepcion.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CantiFac.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CantiFac_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CantiRem.SetFocus
    End If
    If KeyAscii = 27 Then
        CantiFac.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CantiRem_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CantiArti.SetFocus
    End If
    If KeyAscii = 27 Then
        CantiRem.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CantiArti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        IvaInscripto.SetFocus
    End If
    If KeyAscii = 27 Then
        CantiArti.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Tablas.TabCaption(0) = "Retencion de Ganancias (2784)"
    Tablas.TabCaption(1) = "Retencion de Iva (3125)"
    Tablas.TabCaption(2) = "Parametros Generales"

    Minimo1.Text = ""
    Minimo2.Text = ""
    Minimo3.Text = ""
    Minimo4.Text = ""
    Escala1.Text = ""
    Escala2.Text = ""
    Escala3.Text = ""
    Escala4.Text = ""
    Escala5.Text = ""
    Tasa1.Text = ""
    Tasa2.Text = ""
    Tasa3.Text = ""
    Tasa4.Text = ""
    Tasa5.Text = ""
    TasaGen.Text = ""
    TasaBienes.Text = ""
    TasaNoInscripto.Text = ""
    RetMinima.Text = ""
    PorceBienes.Text = ""
    PorceServicios.Text = ""
    PorceTranspo.Text = ""
    MinimoIva.Text = ""
    IvaInscripto.Text = ""
    IvaNoInscripto.Text = ""
    IvaServicio.Text = ""
    Percepcion.Text = ""
    Punto.Text = ""
    CantiFac.Text = ""
    CantiRem.Text = ""
    CantiArti.Text = "2"
    
    Tablas.Tab = 0
    Rem Minimo1.SetFocus
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Parametro"
    ZSql = ZSql + " Where Parametro.Clave = 1"
    spParametro = ZSql
    Set rstParametro = db.OpenRecordset(spParametro, dbOpenSnapshot, dbSQLPassThrough)
    If rstParametro.RecordCount > 0 Then
        Minimo1.Text = rstParametro!Minimo1
        Minimo2.Text = rstParametro!Minimo2
        Minimo3.Text = rstParametro!Minimo3
        Minimo4.Text = rstParametro!Minimo4
        Escala1.Text = rstParametro!Escala1
        Escala2.Text = rstParametro!Escala2
        Escala3.Text = rstParametro!Escala3
        Escala4.Text = rstParametro!Escala4
        Escala5.Text = rstParametro!Escala5
        Tasa1.Text = rstParametro!Tasa1
        Tasa2.Text = rstParametro!Tasa2
        Tasa3.Text = rstParametro!Tasa3
        Tasa4.Text = rstParametro!Tasa4
        Tasa5.Text = rstParametro!Tasa5
        TasaGen.Text = rstParametro!TasaGen
        TasaBienes.Text = rstParametro!TasaBienes
        TasaNoInscripto.Text = rstParametro!TasaNoInscripto
        RetMinima.Text = rstParametro!RetMinima
        PorceBienes.Text = rstParametro!PorceBienes
        PorceServicios.Text = rstParametro!PorceServicios
        PorceTranspo.Text = rstParametro!PorceTranspo
        MinimoIva.Text = rstParametro!MinimoIva
        IvaInscripto.Text = rstParametro!IvaInscripto
        IvaNoInscripto.Text = rstParametro!IvaNoInscripto
        IvaServicio.Text = rstParametro!IvaServicio
        Minimo1.Text = Pusing("#,###,###.##", Minimo1.Text)
        Minimo2.Text = Pusing("#,###,###.##", Minimo2.Text)
        Minimo3.Text = Pusing("#,###,###.##", Minimo3.Text)
        Minimo4.Text = Pusing("#,###,###.##", Minimo4.Text)
        Escala1.Text = Pusing("#,###,###.##", Escala1.Text)
        Escala2.Text = Pusing("#,###,###.##", Escala2.Text)
        Escala3.Text = Pusing("#,###,###.##", Escala3.Text)
        Escala4.Text = Pusing("#,###,###.##", Escala4.Text)
        Escala5.Text = Pusing("#,###,###.##", Escala5.Text)
        Tasa1.Text = Pusing("###,###.##", Tasa1.Text)
        Tasa2.Text = Pusing("###,###.##", Tasa2.Text)
        Tasa3.Text = Pusing("###,###.##", Tasa3.Text)
        Tasa4.Text = Pusing("###,###.##", Tasa4.Text)
        Tasa5.Text = Pusing("###,###.##", Tasa5.Text)
        TasaGen.Text = Pusing("###,###.##", TasaGen.Text)
        TasaBienes.Text = Pusing("###,###.##", TasaBienes.Text)
        TasaNoInscripto.Text = Pusing("###,###.##", TasaNoInscripto.Text)
        RetMinima.Text = Pusing("#,###,###.##", RetMinima.Text)
        PorceBienes.Text = Pusing("#,###,###.##", PorceBienes.Text)
        PorceServicios.Text = Pusing("#,###,###.##", PorceServicios.Text)
        PorceTranspo.Text = Pusing("#,###,###.##", PorceTranspo.Text)
        MinimoIva.Text = Pusing("#,###,###.##", MinimoIva.Text)
        IvaInscripto.Text = Pusing("#,###,###.##", IvaInscripto.Text)
        IvaNoInscripto.Text = Pusing("#,###,###.##", IvaNoInscripto.Text)
        IvaServicio.Text = Pusing("#,###,###.##", IvaServicio.Text)
        rstParametro.Close
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Configuracion"
    ZSql = ZSql + " Where Configuracion.Clave = 1"
    spConfiguracion = ZSql
    Set rstConfiguracion = db.OpenRecordset(spConfiguracion, dbOpenSnapshot, dbSQLPassThrough)
    If rstConfiguracion.RecordCount > 0 Then
        IvaInscripto.Text = rstConfiguracion!Iva1
        IvaNoInscripto.Text = rstConfiguracion!Iva2
        IvaServicio.Text = rstConfiguracion!IvaServicio
        Percepcion.Text = rstConfiguracion!Percepcion
        Punto.Text = rstConfiguracion!Punto
        Percepcion.Text = Pusing("#,###,###.##", Percepcion.Text)
        IvaInscripto.Text = Pusing("#,###,###.##", IvaInscripto.Text)
        IvaNoInscripto.Text = Pusing("#,###,###.##", IvaNoInscripto.Text)
        IvaServicio.Text = Pusing("#,###,###.##", IvaServicio.Text)
        CantiFac.Text = rstConfiguracion!CantiFac
        CantiRem.Text = rstConfiguracion!CantiRem
        CantiArti.Text = rstConfiguracion!CantiArti
        rstConfiguracion.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM NroRet"
    ZSql = ZSql + " Where NroRet.Clave = 1"
    spNroRet = ZSql
    Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
    If rstNroRet.RecordCount > 0 Then
        Numero1.Text = rstNroRet!Numero
        rstNroRet.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM NroRet"
    ZSql = ZSql + " Where NroRet.Clave = 2"
    spNroRet = ZSql
    Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
    If rstNroRet.RecordCount > 0 Then
        Numero2.Text = rstNroRet!Numero
        rstNroRet.Close
    End If
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    On Error GoTo WError
    
    Select Case Tablas.Tab
        Case 0
            Minimo1.SetFocus
        Case 1
            PorceBienes.SetFocus
        Case 2
            IvaInscripto.SetFocus
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Minimo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Minimo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Minimo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Minimo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Minimo5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Escala1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Escala2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Escala3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Escala4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Escala5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tasa1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tasa2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tasa3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tasa4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tasa5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tasagen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TasaBienes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TasaNoInscripto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetMinima_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Numero1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PorceBienes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PorceTranspo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PorceServicios_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub MinimoIva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Numero2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub IvaInscripto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub IvaNoInscripto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub IvaServicio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Percepcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Punto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CantiFac_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CantiRem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub CantiArti_KeyDown(KeyCode As Integer, Shift As Integer)
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





