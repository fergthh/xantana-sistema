VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompras 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Comprobantes de Proveedores"
   ClientHeight    =   8490
   ClientLeft      =   435
   ClientTop       =   375
   ClientWidth     =   10935
   LinkTopic       =   "Form2"
   ScaleHeight     =   8490
   ScaleWidth      =   10935
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
      Left            =   9840
      MouseIcon       =   "cOMPRAS.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "cOMPRAS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Salida"
      Top             =   5640
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
      Left            =   9840
      MouseIcon       =   "cOMPRAS.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "cOMPRAS.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Consulta de Datos"
      Top             =   4560
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
      Left            =   9840
      MouseIcon       =   "cOMPRAS.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "cOMPRAS.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3480
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
      Left            =   9840
      MouseIcon       =   "cOMPRAS.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "cOMPRAS.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Elimina el Registro"
      Top             =   2400
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
      Left            =   9840
      MouseIcon       =   "cOMPRAS.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "cOMPRAS.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1320
      Width           =   855
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   4575
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8070
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "cOMPRAS.frx":387C
      Tab(0).ControlCount=   42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DesConcepto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Total"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label19"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label18"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label16"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label10"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblLabels(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label20"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DesProveedorIva"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label22"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label23"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label24"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label25"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DesCentro"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label21"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "DesBanco"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label27"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Fecha"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Vencimiento"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Periodo"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Observaciones"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Concepto"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Exento"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Iva27"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Iva21"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Ib"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Iva5"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Neto"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Iva105"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "ProveedorIva"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "ImpInterno"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "ImpCombustible"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Centro"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Banco"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Paridad"
      Tab(0).Control(41).Enabled=   0   'False
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "cOMPRAS.frx":3898
      Tab(1).ControlCount=   13
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Debito"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "Credito"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "WTitulo(4)"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "WTitulo(3)"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "WTitulo(2)"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "WTitulo(1)"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "WTexto2"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "WCombo1"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "WTexto1"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "WTexto3"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "WVector1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "WImpo1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label1"
      Tab(1).Control(12).Enabled=   0   'False
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "cOMPRAS.frx":38B4
      Tab(2).ControlCount=   11
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label15"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Wimpo2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "WVector11"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "WTexto31"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "WTitulo1(3)"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "WTitulo1(2)"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "WTitulo1(1)"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "WTexto21"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "WCombo11"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "WTexto11"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "SumaProyecto"
      Tab(2).Control(10).Enabled=   -1  'True
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "cOMPRAS.frx":38D0
      Tab(3).ControlCount=   15
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label26"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "WVector22"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "WTexto32"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "WTexto12"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "WCombo12"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "WTexto22"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "WTitulo2(1)"
      Tab(3).Control(6).Enabled=   -1  'True
      Tab(3).Control(7)=   "WTitulo2(2)"
      Tab(3).Control(7).Enabled=   -1  'True
      Tab(3).Control(8)=   "WTitulo2(3)"
      Tab(3).Control(8).Enabled=   -1  'True
      Tab(3).Control(9)=   "WTitulo2(4)"
      Tab(3).Control(9).Enabled=   -1  'True
      Tab(3).Control(10)=   "LeeRemito"
      Tab(3).Control(10).Enabled=   -1  'True
      Tab(3).Control(11)=   "Remito"
      Tab(3).Control(11).Enabled=   -1  'True
      Tab(3).Control(12)=   "WTitulo2(5)"
      Tab(3).Control(12).Enabled=   -1  'True
      Tab(3).Control(13)=   "WTitulo2(6)"
      Tab(3).Control(13).Enabled=   -1  'True
      Tab(3).Control(14)=   "IMpoRemito"
      Tab(3).Control(14).Enabled=   -1  'True
      Begin VB.TextBox Paridad 
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
         MaxLength       =   15
         TabIndex        =   108
         Text            =   " "
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox IMpoRemito 
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
         Height          =   375
         Left            =   -67680
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   3600
         Width           =   1215
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
         Left            =   -70920
         Locked          =   -1  'True
         TabIndex        =   105
         Top             =   2040
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
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   104
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox Remito 
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
         Height          =   375
         Left            =   -70200
         MaxLength       =   10
         TabIndex        =   102
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton LeeRemito 
         Caption         =   "Lee Remito"
         Height          =   495
         Left            =   -68880
         TabIndex        =   101
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox Banco 
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
         Left            =   6960
         TabIndex        =   97
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Centro 
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
         Left            =   2520
         TabIndex        =   94
         Top             =   3360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox ImpCombustible 
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
         MaxLength       =   15
         TabIndex        =   91
         Text            =   " "
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox ImpInterno 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   90
         Text            =   " "
         Top             =   2280
         Width           =   1335
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
         Left            =   -73200
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   2040
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
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   2040
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
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   2160
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
         Index           =   1
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   2040
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
         Left            =   -73800
         TabIndex        =   83
         Top             =   1560
         Width           =   375
      End
      Begin VB.ComboBox WCombo12 
         Height          =   315
         Left            =   -72720
         TabIndex        =   82
         Top             =   1560
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
         Left            =   -74400
         TabIndex        =   81
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox ProveedorIva 
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
         MaxLength       =   11
         TabIndex        =   78
         Text            =   " "
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox Iva105 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   70
         Text            =   " "
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox SumaProyecto 
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
         Left            =   -68280
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Debito 
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
         Left            =   -68880
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Credito 
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
         Left            =   -67200
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox WTexto11 
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
         Left            =   -73920
         TabIndex        =   60
         Top             =   1320
         Width           =   375
      End
      Begin VB.ComboBox WCombo11 
         Height          =   315
         Left            =   -72240
         TabIndex        =   59
         Top             =   1320
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto21 
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
         Left            =   -73320
         TabIndex        =   58
         Top             =   1320
         Width           =   375
      End
      Begin VB.TextBox WTitulo1 
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
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   57
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox WTitulo1 
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
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox WTitulo1 
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
         Left            =   -72720
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1800
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
         Left            =   -73320
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1920
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
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1920
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
         Left            =   -74280
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   1920
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
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1920
         Width           =   375
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
         Left            =   -73800
         TabIndex        =   48
         Top             =   1440
         Width           =   375
      End
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   -73320
         TabIndex        =   47
         Top             =   1440
         Visible         =   0   'False
         Width           =   390
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
         Left            =   -74280
         TabIndex        =   46
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Neto 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   31
         Text            =   " "
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox Iva5 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   27
         Text            =   " "
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox Ib 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   26
         Text            =   " "
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox Iva21 
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
         MaxLength       =   15
         TabIndex        =   25
         Text            =   " "
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Iva27 
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
         MaxLength       =   15
         TabIndex        =   24
         Text            =   " "
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Exento 
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
         MaxLength       =   15
         TabIndex        =   23
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Concepto 
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
         Left            =   2520
         TabIndex        =   22
         Top             =   3000
         Width           =   855
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
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   21
         Top             =   3720
         Width           =   6015
      End
      Begin MSMask.MaskEdBox Periodo 
         Height          =   285
         Left            =   6480
         TabIndex        =   28
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Vencimiento 
         Height          =   285
         Left            =   2520
         TabIndex        =   29
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   2520
         TabIndex        =   30
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   285
         Left            =   -74760
         TabIndex        =   52
         Top             =   1440
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   53
         Top             =   480
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto31 
         Height          =   285
         Left            =   -72840
         TabIndex        =   61
         Top             =   1320
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
      Begin MSFlexGridLib.MSFlexGrid WVector11 
         Height          =   2655
         Left            =   -74160
         TabIndex        =   62
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   4683
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox WTexto32 
         Height          =   285
         Left            =   -73320
         TabIndex        =   84
         Top             =   1560
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
      Begin MSFlexGridLib.MSFlexGrid WVector22 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   85
         Top             =   720
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.Label Label27 
         Caption         =   "Paridad"
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
         TabIndex        =   107
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label26 
         Caption         =   "Remito"
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
         Left            =   -71400
         TabIndex        =   103
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label DesBanco 
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
         Height          =   285
         Left            =   7920
         TabIndex        =   100
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label21 
         Caption         =   "Banco"
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
         Left            =   6240
         TabIndex        =   98
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label DesCentro 
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
         Left            =   3480
         TabIndex        =   96
         Top             =   3360
         Visible         =   0   'False
         Width           =   5055
      End
      Begin VB.Label Label25 
         Caption         =   "Centro de Costo"
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
         TabIndex        =   95
         Top             =   3360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label24 
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
         Height          =   375
         Left            =   4080
         TabIndex        =   93
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label23 
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
         Left            =   240
         TabIndex        =   92
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label22 
         Caption         =   "Proveedor Iva"
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
         TabIndex        =   80
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label DesProveedorIva 
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
         Height          =   285
         Left            =   4080
         TabIndex        =   79
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "Importe 10,5%"
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
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Wimpo2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -71400
         TabIndex        =   69
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe Neto de Comprobante"
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
         Height          =   255
         Left            =   -74160
         TabIndex        =   68
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label WImpo1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -72000
         TabIndex        =   67
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Importe Total de Comprobante"
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   66
         Top             =   3240
         Width           =   2775
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Emision"
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
         Left            =   240
         TabIndex        =   45
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Importe Neto"
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
         TabIndex        =   44
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha de vencimiento"
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
         TabIndex        =   43
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha de Iva"
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
         TabIndex        =   42
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Iva R.G. 3337"
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
         TabIndex        =   41
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Importe Perc. I.B."
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
         TabIndex        =   40
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "Importe Iva 21%"
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
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label17 
         Caption         =   "Importe Iva 27%"
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
         TabIndex        =   38
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "Importe No Gravado"
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
         Left            =   4080
         TabIndex        =   37
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label19 
         Caption         =   "Importe Total"
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
         TabIndex        =   36
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         TabIndex        =   35
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
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
         TabIndex        =   34
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label DesConcepto 
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
         Left            =   3480
         TabIndex        =   33
         Top             =   3000
         Width           =   5055
      End
      Begin VB.Label Label9 
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
         Left            =   240
         TabIndex        =   32
         Top             =   3720
         Width           =   1695
      End
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
      Height          =   330
      Left            =   120
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   9135
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
      Height          =   1980
      Left            =   3000
      TabIndex        =   8
      Top             =   6360
      Visible         =   0   'False
      Width           =   3135
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
      Left            =   600
      TabIndex        =   17
      Text            =   " "
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox TipoComp 
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
      Left            =   840
      TabIndex        =   1
      Text            =   " "
      Top             =   600
      Width           =   1335
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
      ItemData        =   "cOMPRAS.frx":38EC
      Left            =   120
      List            =   "cOMPRAS.frx":38F3
      TabIndex        =   5
      Top             =   6120
      Visible         =   0   'False
      Width           =   9135
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
      Left            =   4320
      MaxLength       =   4
      TabIndex        =   3
      Text            =   " "
      Top             =   600
      Width           =   855
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
      Left            =   6240
      MaxLength       =   8
      TabIndex        =   4
      Text            =   " "
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Letra 
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
      MaxLength       =   1
      TabIndex        =   2
      Text            =   " "
      Top             =   600
      Width           =   495
   End
   Begin VB.Frame Frame4 
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
      Height          =   975
      Left            =   7680
      TabIndex        =   10
      Top             =   0
      Width           =   3135
      Begin VB.OptionButton Contado4 
         Caption         =   "Banco"
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
         Left            =   1560
         TabIndex        =   99
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Contado3 
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
         Left            =   1560
         TabIndex        =   77
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Contado2 
         Caption         =   "En Cta.Cte."
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
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Contado1 
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
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
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   15
      Left            =   360
      TabIndex        =   19
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label8 
      Caption         =   "Punto"
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
      Left            =   3600
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Numero"
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
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label13 
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
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label11 
      Caption         =   "Letra"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   600
      Width           =   495
   End
   Begin VB.Label DesProveedor 
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
      Height          =   285
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "PrgCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Dato As String
Private Auxi As String
Private WImpo As Double
Private WProveedor As String
Private WTipo As String
Private WPunto As String
Private WNumero As String
Private Uno As Double
Private Dos As Double
Private WTipoProveedor As Integer
Private WRenglon As String
Private WDias As Integer
Private WFecha As String
Private WVencimiento As String
Private WSuma1 As Double
Private WSuma2 As Double
Private ZSalida As String
Private ZImpo As String
Dim movstk As Integer
Dim WNroStk As String
Dim ZMes As String
Dim ZAno As String
Dim WTrabajo(100, 10) As String
Dim ZRenglon As Integer
Dim ZLugar As Integer
Dim ZUltimo As Integer

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Dim WParametros1(10, 10) As Double
Dim WFormato1(10) As String
Dim WControl1 As String

Dim WParametros2(10, 10) As Double
Dim WFormato2(10) As String
Dim WControl2 As String

Sub Calcula_total()

    WImpo = 0
    Call Format_datos
    
    Dato = Neto.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva21.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva5.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva27.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Iva105.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Ib.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = Exento.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = ImpInterno.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    Dato = ImpCombustible.Text
    Rem Call Convierte_datos(Dato, Auxi)
    If Val(Dato) <> 0 Then
        WImpo = WImpo + Val(Dato)
    End If
    
    
    Total.Caption = WImpo
    Total.Caption = Pusing("###,###.##", Total.Caption)
    
End Sub

Sub Alinea_Datos()
    WProveedor = Trim(Proveedor.Text)
    Tipo.Text = Str$(TipoComp.ListIndex + 1)
    WTipo = Tipo.Text
    Call Ceros(WTipo, 2)
    Tipo.Text = WTipo
    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
    Punto.Text = WPunto
    WNumero = Numero.Text
    Call Ceros(WNumero, 8)
    Numero.Text = WNumero
    Letra.Text = Left$(Letra.Text, 1)
End Sub

Sub Imprime_Descripcion()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = rstProveedor!Nombre
        WDias = rstProveedor!Dias + 1
        rstProveedor.Close
            Else
        DesProveedor.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        DesConcepto.Caption = rstConceptos!Nombre
        rstConceptos.Close
            Else
        DesConcepto.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proyecto"
    ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + Centro.Text + "'"
    spProyecto = ZSql
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        DesCentro.Caption = rstProyecto!Descripcion
        rstProyecto.Close
            Else
        DesCentro.Caption = ""
    End If

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ProveedorIva.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedorIva.Caption = rstProveedor!Nombre
        rstProveedor.Close
            Else
        DesProveedorIva.Caption = ""
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Banco"
    ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
    spBanco = ZSql
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        DesBanco.Caption = rstBanco!Nombre
        rstBanco.Close
            Else
        DesBanco.Caption = ""
    End If
    
End Sub

Sub Verifica_datos()
    If Val(Neto.Text) = 0 Then
        Neto.Text = "0"
    End If
    If Val(Iva21.Text) = 0 Then
        Iva21.Text = "0"
    End If
    If Val(Iva5.Text) = 0 Then
        Iva5.Text = "0"
    End If
    If Val(Iva27.Text) = 0 Then
        Iva27.Text = "0"
    End If
    If Val(Iva105.Text) = 0 Then
        Iva105.Text = "0"
    End If
    If Val(Ib.Text) = 0 Then
        Ib.Text = "0"
    End If
    If Val(Exento.Text) = 0 Then
        Exento.Text = "0"
    End If
    If Val(ImpInterno.Text) = 0 Then
        ImpInterno.Text = "0"
    End If
    If Val(ImpCombustible.Text) = 0 Then
        ImpCombustible.Text = "0"
    End If
    If Val(Total.Caption) = 0 Then
        Total.Caption = "0"
    End If
    If Val(Concepto.Text) = 0 Then
        Concepto.Text = "0"
    End If
    If Val(Centro.Text) = 0 Then
        Centro.Text = "0"
    End If
    If Val(Banco.Text) = 0 Then
        Banco.Text = "0"
    End If
    If Val(Paridad.Text) = 0 Then
        Paridad.Text = "0"
    End If
End Sub

Sub Format_datos()
    If Val(Neto.Text) <> 0 Then
        Neto.Text = Pusing("###,###.##", Neto.Text)
            Else
        Neto.Text = ""
    End If
    If Val(Iva21.Text) <> 0 Then
        Iva21.Text = Pusing("###,###.##", Iva21.Text)
            Else
        Iva21.Text = ""
    End If
    If Val(Iva5.Text) <> 0 Then
        Iva5.Text = Pusing("###,###.##", Iva5.Text)
            Else
        Iva5.Text = ""
    End If
    If Val(Iva27.Text) <> 0 Then
        Iva27.Text = Pusing("###,###.##", Iva27.Text)
            Else
        Iva27.Text = ""
    End If
    If Val(Iva105.Text) <> 0 Then
        Iva105.Text = Pusing("###,###.##", Iva105.Text)
            Else
        Iva105.Text = ""
    End If
    If Val(Ib.Text) <> 0 Then
        Ib.Text = Pusing("###,###.##", Ib.Text)
            Else
        Ib.Text = ""
    End If
    If Val(ImpInterno.Text) <> 0 Then
        ImpInterno.Text = Pusing("###,###.##", ImpInterno.Text)
            Else
        ImpInterno.Text = ""
    End If
    If Val(ImpCombustible.Text) <> 0 Then
        ImpCombustible.Text = Pusing("###,###.##", ImpCombustible.Text)
            Else
        ImpCombustible.Text = ""
    End If
    If Val(Exento.Text) <> 0 Then
        Exento.Text = Pusing("###,###.##", Exento.Text)
            Else
        Exento.Text = ""
    End If
    If Val(Paridad.Text) <> 0 Then
        Paridad.Text = Pusing("###,###.##", Paridad.Text)
            Else
        Paridad.Text = ""
    End If
    Total.Caption = Pusing("###,###.##", Total.Caption)
End Sub

Sub Imprime_Datos()

    Call Alinea_Datos
    
    WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        Proveedor.Text = rstIvaComp!Proveedor
        TipoComp.ListIndex = rstIvaComp!Tipo - 1
        Letra.Text = rstIvaComp!Letra
        Punto.Text = rstIvaComp!Punto
        Numero.Text = rstIvaComp!Numero
        Call Alinea_Datos
        Fecha.Text = rstIvaComp!Fecha
        Vencimiento.Text = rstIvaComp!Vencimiento
        Periodo.Text = rstIvaComp!Periodo
        Neto.Text = Abs(rstIvaComp!Neto)
        Iva21.Text = Abs(rstIvaComp!Iva21)
        Iva5.Text = Abs(rstIvaComp!Iva5)
        Iva27.Text = Abs(rstIvaComp!Iva27)
        Iva105.Text = Abs(rstIvaComp!Iva105)
        Ib.Text = Abs(rstIvaComp!Ib)
        ImpInterno.Text = Abs(rstIvaComp!ImpInterno)
        ImpCombustible.Text = Abs(rstIvaComp!ImpCombustible)
        Exento.Text = Abs(rstIvaComp!Exento)
        Concepto.Text = rstIvaComp!Concepto
        Centro.Text = rstIvaComp!Centro
        Observaciones.Text = Trim(rstIvaComp!Observaciones)
        Call Calcula_total
        Contado1.Value = False
        Contado2.Value = False
        Contado3.Value = False
        Contado4.Value = False
        Select Case Val(rstIvaComp!Contado)
            Case 1
                Contado1.Value = True
            Case 2
                Contado2.Value = True
            Case 3
                Contado3.Value = True
            Case 4
                Contado4.Value = True
            Case Else
        End Select
        Remito.Text = IIf(IsNull(rstIvaComp!Remito), "", rstIvaComp!Remito)
        Remito.Text = Trim(Remito.Text)
        Paridad.Text = IIf(IsNull(rstIvaComp!Paridad), "0", rstIvaComp!Paridad)
        ProveedorIva.Text = IIf(IsNull(rstIvaComp!ProveedorIva), "", rstIvaComp!ProveedorIva)
        Banco.Text = IIf(IsNull(rstIvaComp!Banco), "", rstIvaComp!Banco)
        movstk = IIf(IsNull(rstIvaComp!movstk), "0", rstIvaComp!movstk)
        rstIvaComp.Close
        Call Format_datos
        Call Imprime_Descripcion
    End If
End Sub

Private Sub cmdAdd_Click()

    Tipo.Text = TipoComp.ListIndex + 1
    Tablas.Tab = 1
    Tablas.Tab = 0
            
    WPasa = "S"
    Call Verifica_datos
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        rstProveedor.Close
            Else
        WPasa = "N"
        m$ = "Codigo de Proveedor Incorrecto"
        a% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        rstConceptos.Close
            Else
        WPasa = "N"
        m$ = "Codigo de Concepto Incorrecto"
        a% = MsgBox(m$, 0, "Archivo de Ingresos de Conceptos de Compras")
    End If
        
    Call Valida_fecha(Fecha.Text, Auxi)
    If Auxi <> "S" Then
        WPasa = "N"
        m$ = "Formato de Fecha de emision, formato valido : dd/mm/aaaa"
        a% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    End If
        
    Call Valida_fecha(Vencimiento.Text, Auxi)
    If Auxi <> "S" Then
        WPasa = "N"
        m$ = "Formato de Fecha de vencimiento (1), formato valido : dd/mm/aaaa"
        a% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    End If

    Call Valida_fecha(Periodo.Text, Auxi)
    If Auxi <> "S" Then
        WPasa = "N"
        m$ = "Formato de Fecha de Iva, formato valido : dd/mm/aaaa"
        a% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    End If
        
    If (Val(Tipo.Text) < 1 Or Val(Tipo.Text) > 3) Then
        If Val(Tipo.Text) <> 7 And Val(Tipo.Text) <> 8 Then
            WPasa = "N"
            m$ = "Tipo de Comprobante Invalido"
            a% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
        End If
    End If
    
    If Left$(Letra.Text, 1) <> "A" And Left$(Letra.Text, 1) <> "B" And Left$(Letra.Text, 1) <> "C" And Left$(Letra.Text, 1) <> "X" And Left$(Letra.Text, 1) <> "Z" And Left$(Letra.Text, 1) <> "M" Then
        WPasa = "N"
        m$ = "Letra del Comprobante Invalido"
        a% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    End If
    
    WSuma1 = Val(Debito.Text)
    WSuma2 = Val(Credito.Text)
    Call Redondeo(WSuma1)
    Call Redondeo(WSuma2)
    
    Rem If WSuma1 <> WSuma2 Then
    Rem     WPasa = "N"
    Rem     m$ = "Asiento Contable Desbalanceado"
    Rem     A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    Rem         Else
    Rem     WSuma1 = Val(Debito.Text)
   Rem      WSuma2 = Val(Total.Caption)
    Rem     Call Redondeo(WSuma1)
    Rem     Call Redondeo(WSuma2)
    Rem     If WSuma1 <> WSuma2 Then
    Rem         WPasa = "N"
    Rem         m$ = "Los valores del comprobante no conciden con el asiento contable"
    Rem         A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    Rem     End If
    Rem End If
    
    Rem If Val(SumaProyecto.Text) <> Val(Neto.Text) + Val(Exento.Text) Then
    Rem     WPasa = "N"
    Rem     m$ = "Los valores de la Discriminacion por Proyecto no conicide con el importe de comprobante"
    Rem     A% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
    Rem End If
    
    WNumero = Numero.Text
    Call Ceros(WNumero, 8)
    Numero.Text = WNumero
    Call Alinea_Datos
    XProveedor = Proveedor.Text
    XTipo = TipoComp.ListIndex
    XLetra = Letra.Text
    XPunto = Punto.Text
    XNumero = Numero.Text
    
    WClave = WProveedor + Letra.Text + WTipo + WPunto + WNumero
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCtePrv"
    ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + WClave + "'"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCtePrv.RecordCount > 0 Then
        If rstCtaCtePrv!Saldo <> rstCtaCtePrv!Total Then
            m$ = "El comprobante ya se encuentra cancelado total o parcialmente"
            a% = MsgBox(m$, 0, "Archivo de Ingresos de Comprobantes")
            WPasa = "N"
        End If
        rstCtaCtePrv.Close
    End If
    
    
    If WPasa = "S" Then
    
        Call Alinea_Datos
        Call Verifica_datos
    
        ZSql = ""
        ZSql = ZSql + "DELETE ImpCyb"
        ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
        ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Letra = " + "'" + Letra.Text + "'"
        ZSql = ZSql + " and Punto = " + "'" + WPunto + "'"
        ZSql = ZSql + " and Numero = " + "'" + WNumero + "'"
        spImpCyb = ZSql
        Set rstImpCyb = db.OpenRecordset(spImpCyb, dbOpenSnapshot, dbSQLPassThrough)
    
        ZSql = ""
        ZSql = ZSql + "DELETE ImpProy"
        ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
        ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Letra = " + "'" + Letra.Text + "'"
        ZSql = ZSql + " and Punto = " + "'" + WPunto + "'"
        ZSql = ZSql + " and Numero = " + "'" + WNumero + "'"
        spImpProy = ZSql
        Set rstImpProy = db.OpenRecordset(spImpProy, dbOpenSnapshot, dbSQLPassThrough)
        
        If movstk <> 0 Then
        
            Erase WTrabajo
            ZLugar = 0
        
            For ZRenglon = 1 To 100
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Compras"
                ZSql = ZSql + " Where Compras.Numero = " + "'" + Str$(movstk) + "'"
                ZSql = ZSql + " and Compras.Renglon = " + "'" + Str$(ZRenglon) + "'"
                spCompras = ZSql
                Set rstCompras = db.OpenRecordset(spCompras, dbOpenSnapshot, dbSQLPassThrough)
                If rstCompras.RecordCount > 0 Then
                    ZLugar = ZLugar + 1
                    WTrabajo(ZLugar, 1) = rstCompras!Articulo
                    WTrabajo(ZLugar, 2) = rstCompras!Cantidad
                    rstCompras.Close
                End If
                
            Next ZRenglon
            
            For ZRenglon = 1 To ZLugar
            
                WArticulo = WTrabajo(ZRenglon, 1)
                WCantidad = WTrabajo(ZRenglon, 2)
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Stock = Stock - " + WCantidad
                ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            Next ZRenglon
            
            ZSql = ""
            ZSql = ZSql + "DELETE Compras"
            ZSql = ZSql + " Where Numero = " + "'" + Str$(movstk) + "'"
            spCompras = ZSql
            Set rstCompras = db.OpenRecordset(spCompras, dbOpenSnapshot, dbSQLPassThrough)
        
        End If

        Renglon = 0
        WRenglon = 0
        
        For IRow = 1 To 100
            
            WVector22.Row = IRow
            
            WVector22.Col = 1
            Articulo = WVector22.Text
                    
            WVector22.Col = 3
            Cantidad = Val(WVector22.Text)
            
            WVector22.Col = 4
            Costo = Val(WVector22.Text)
                    
            If Cantidad <> 0 Then
                    
                If movstk = 0 Then
                
                    movstk = 1
                    
                    ZSql = ""
                    ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
                    ZSql = ZSql + " FROM Compras"
                    spCompras = ZSql
                    Set rstCompras = db.OpenRecordset(spCompras, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCompras.RecordCount > 0 Then
                        rstCompras.MoveLast
                        ZUltimo = IIf(IsNull(rstCompras!NUMEROMayor), "0", rstCompras!NUMEROMayor)
                        movstk = ZUltimo + 1
                        rstCompras.Close
                    End If
                    
                End If
                WNroStk = Str$(movstk)
        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = WNroStk
                Call Ceros(Auxi1, 6)
                
                WClave = Auxi1 + Auxi
                WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Compras ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Costo ,"
                ZSql = ZSql + "Observaciones )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WNroStk + "',"
                ZSql = ZSql + "'" + Str$(Renglon) + "',"
                ZSql = ZSql + "'" + Fecha.Text + "',"
                ZSql = ZSql + "'" + WOrdFecha + "',"
                ZSql = ZSql + "'" + Articulo + "',"
                ZSql = ZSql + "'" + Str$(Cantidad) + "',"
                ZSql = ZSql + "'" + Str$(Costo) + "',"
                ZSql = ZSql + "'" + Observaciones.Text + "')"
                spCompras = ZSql
                Set rstCompras = db.OpenRecordset(spCompras, dbOpenSnapshot, dbSQLPassThrough)
                
                Rem If Costo <> 0 Then
                Rem     ZSql = ""
                Rem     ZSql = ZSql + "UPDATE Articulo SET "
                Rem     ZSql = ZSql + " Costo = " + Str$(Costo)
                Rem     ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                Rem     spArticulo = ZSql
                Rem     Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                Rem End If
                
                Rem ZSql = ""
                Rem ZSql = ZSql + "UPDATE Articulo SET "
                Rem ZSql = ZSql + " Stock = Stock + " + Str$(Cantidad)
                Rem ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                Rem spArticulo = ZSql
                Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                                        
        Next IRow
        
        Call Alinea_Datos
        Call Verifica_datos
        
        Rem Graba el iva ventas
        
        WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero
        
        WNeto = Val(Neto.Text)
        WIva21 = Val(Iva21.Text)
        WIva5 = Val(Iva5.Text)
        WIva27 = Val(Iva27.Text)
        WIva105 = Val(Iva105.Text)
        WIb = Val(Ib.Text)
        WImpInterno = Val(ImpInterno.Text)
        WImpCombustible = Val(ImpCombustible.Text)
        WExento = Val(Exento.Text)
        
        Select Case Val(Tipo.Text)
            Case 1
                WImpre = "FC"
            Case 7
                WImpre = "TK"
            Case 8
                WImpre = "RC"
            Case 2
                WImpre = "ND"
            Case 3
                WImpre = "NC"
                WNeto = Val(Neto.Text) * -1
                WIva21 = Val(Iva21.Text) * -1
                WIva5 = Val(Iva5.Text) * -1
                WIva27 = Val(Iva27.Text) * -1
                WIva105 = Val(Iva105.Text) * -1
                WIb = Val(Ib.Text) * -1
                WImpInterno = Val(ImpInterno.Text) * -1
                WImpCombustible = Val(ImpCombustible.Text) * -1
                WExento = Val(Exento.Text) * -1
            Case Else
                WImpre = "  "
        End Select
        
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WCodigoEmpresa = WEmpresa
        WOrdPeriodo = Right$(Periodo.Text, 4) + Mid$(Periodo.Text, 4, 2) + Left$(Periodo.Text, 2)
        WImpreNumero = WNumero
        If Contado1.Value = True Then
            WContado = "1"
        End If
        If Contado2.Value = True Then
            WContado = "2"
        End If
        If Contado3.Value = True Then
            WContado = "3"
        End If
        If Contado4.Value = True Then
            WContado = "4"
        End If
                
        WProveedorIva = ProveedorIva.Text
        If Trim(ProveedorIva.Text) = "" Then
            WProveedorIva = Proveedor.Text
        End If
        
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM IvaComp"
        ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
        spIvaComp = ZSql
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
            rstIvaComp.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE IvaComp SET "
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " Vencimiento = " + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + " Periodo = " + "'" + Periodo.Text + "',"
            ZSql = ZSql + " Concepto = " + "'" + Concepto.Text + "',"
            ZSql = ZSql + " Centro = " + "'" + Centro.Text + "',"
            ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " Remito = " + "'" + Remito.Text + "',"
            ZSql = ZSql + " Paridad = " + "'" + Paridad.Text + "',"
            ZSql = ZSql + " Neto = " + "'" + Str$(WNeto) + "',"
            ZSql = ZSql + " Iva21 = " + "'" + Str$(WIva21) + "',"
            ZSql = ZSql + " Iva5 = " + "'" + Str$(WIva5) + "',"
            ZSql = ZSql + " Iva27 = " + "'" + Str$(WIva27) + "',"
            ZSql = ZSql + " Iva105 = " + "'" + Str$(WIva105) + "',"
            ZSql = ZSql + " Ib = " + "'" + Str$(WIb) + "',"
            ZSql = ZSql + " ImpInterno = " + "'" + Str$(WImpInterno) + "',"
            ZSql = ZSql + " ImpCombustible = " + "'" + Str$(WImpCombustible) + "',"
            ZSql = ZSql + " Exento = " + "'" + Str$(WExento) + "',"
            ZSql = ZSql + " Impre = " + "'" + WImpre + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
            ZSql = ZSql + " Contado = " + "'" + WContado + "',"
            ZSql = ZSql + " ProveedorIva = " + "'" + WProveedorIva + "',"
            ZSql = ZSql + " Banco= " + "'" + Banco.Text + "',"
            ZSql = ZSql + " MovStk = " + "'" + Str$(movstk) + "',"
            ZSql = ZSql + " CodigoEmpresa = " + "'" + WCodigoEmpresa + "',"
            ZSql = ZSql + " ImpreNumero = " + "'" + WImpreNumero + "',"
            ZSql = ZSql + " OrdPeriodo = " + "'" + WOrdPeriodo + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO IvaComp ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Vencimiento ,"
            ZSql = ZSql + "Periodo ,"
            ZSql = ZSql + "Concepto ,"
            ZSql = ZSql + "Centro ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "Paridad ,"
            ZSql = ZSql + "Neto ,"
            ZSql = ZSql + "Iva21 ,"
            ZSql = ZSql + "Iva5 ,"
            ZSql = ZSql + "Iva27 ,"
            ZSql = ZSql + "Iva105 ,"
            ZSql = ZSql + "Ib ,"
            ZSql = ZSql + "ImpInterno ,"
            ZSql = ZSql + "ImpCombustible ,"
            ZSql = ZSql + "Exento ,"
            ZSql = ZSql + "Impre ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Contado ,"
            ZSql = ZSql + "ProveedorIva ,"
            ZSql = ZSql + "Banco ,"
            ZSql = ZSql + "MovStk ,"
            ZSql = ZSql + "CodigoEmpresa ,"
            ZSql = ZSql + "ImpreNumero ,"
            ZSql = ZSql + "OrdPeriodo )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + Letra.Text + "',"
            ZSql = ZSql + "'" + WPunto + "',"
            ZSql = ZSql + "'" + WNumero + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + Vencimiento.Text + "',"
            ZSql = ZSql + "'" + Periodo.Text + "',"
            ZSql = ZSql + "'" + Concepto.Text + "',"
            ZSql = ZSql + "'" + Centro.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + Remito.Text + "',"
            ZSql = ZSql + "'" + Paridad.Text + "',"
            ZSql = ZSql + "'" + Str$(WNeto) + "',"
            ZSql = ZSql + "'" + Str$(WIva21) + "',"
            ZSql = ZSql + "'" + Str$(WIva5) + "',"
            ZSql = ZSql + "'" + Str$(WIva27) + "',"
            ZSql = ZSql + "'" + Str$(WIva105) + "',"
            ZSql = ZSql + "'" + Str$(WIb) + "',"
            ZSql = ZSql + "'" + Str$(WImpInterno) + "',"
            ZSql = ZSql + "'" + Str$(WImpCombustible) + "',"
            ZSql = ZSql + "'" + Str$(WExento) + "',"
            ZSql = ZSql + "'" + WImpre + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + WContado + "',"
            ZSql = ZSql + "'" + WProveedorIva + "',"
            ZSql = ZSql + "'" + Banco.Text + "',"
            ZSql = ZSql + "'" + Str$(movstk) + "',"
            ZSql = ZSql + "'" + WCodigoEmpresa + "',"
            ZSql = ZSql + "'" + WImpreNumero + "',"
            ZSql = ZSql + "'" + WOrdPeriodo + "')"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                
        End If
                
        WProveedor = Trim(Proveedor.Text)
        WTipo = Tipo.Text
        WLetra = Letra.Text
        WPunto = Punto.Text
        WNumero = Numero.Text
        WFecha = Fecha.Text
        WVencimiento = Vencimiento.Text
        
        Rem Graba la cta.cte
        
        Call Alinea_Datos
        WClave = WProveedor + Letra.Text + WTipo + WPunto + WNumero
        
        ZSql = ""
        ZSql = ZSql + "DELETE CtaCtePrv"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
        spCtaCtePrv = ZSql
        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
        
        If Val(WContado) = 2 Then
        
            Call Alinea_Datos
            
            WTotal = Val(Total.Caption)
            WSaldo = Val(Total.Caption)
            
            Select Case Val(WTipo)
                Case 1
                    WImpre = "FC"
                Case 7
                    WImpre = "TK"
                Case 8
                    WImpre = "RC"
                Case 2
                    WImpre = "ND"
                Case 3
                    WImpre = "NC"
                    WTotal = WTotal * -1
                    WSaldo = WSaldo * -1
                Case Else
                    WImpre = ""
            End Select
            
            WOrdFecha = Right$(Fecha, 4) + Mid$(Fecha, 4, 2) + Left$(Fecha, 2)
            WOrdvencimiento = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
            
                Else
        
            Call Alinea_Datos
            
            WTotal = Val(Total.Caption)
            WSaldo = 0
            
            Select Case Val(WTipo)
                Case 1
                    WImpre = "CO"
                Case 7
                    WImpre = "TK"
                Case 8
                    WImpre = "RC"
                Case 2
                    WImpre = "ND"
                Case 3
                    WImpre = "NC"
                    WTotal = WTotal * -1
                    WSaldo = WSaldo * -1
                Case Else
                    WImpre = ""
            End Select
            
            WOrdFecha = Right$(Fecha, 4) + Mid$(Fecha, 4, 2) + Left$(Fecha, 2)
            WOrdvencimiento = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
                
        End If
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CtaCtePrv ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "OrdVencimiento ,"
        ZSql = ZSql + "Impre )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WClave + "',"
        ZSql = ZSql + "'" + WProveedor + "',"
        ZSql = ZSql + "'" + WLetra + "',"
        ZSql = ZSql + "'" + WTipo + "',"
        ZSql = ZSql + "'" + WPunto + "',"
        ZSql = ZSql + "'" + WNumero + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + "1" + "',"
        ZSql = ZSql + "'" + Vencimiento.Text + "',"
        ZSql = ZSql + "'" + Str$(WTotal) + "',"
        ZSql = ZSql + "'" + Str$(WSaldo) + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "',"
        ZSql = ZSql + "'" + WOrdFecha + "',"
        ZSql = ZSql + "'" + WOrdvencimiento + "',"
        ZSql = ZSql + "'" + WImpre + "')"
        spCtaCtePrv = ZSql
        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        GrabaCentro = "S"
        
        For Ciclo = 1 To 100
        
            WRenglon = Str$(Ciclo)
            Call Ceros(WRenglon, 2)
            
            If Val(WVector1.TextMatrix(Ciclo, 3)) <> 0 Or Val(WVector1.TextMatrix(Ciclo, 4)) <> 0 Then
            
                WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero + WRenglon
                        
                ZSql = ""
                ZSql = ZSql + "INSERT INTO ImpCyb ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Proveedor ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Letra ,"
                ZSql = ZSql + "Punto ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Cuenta ,"
                ZSql = ZSql + "Debito ,"
                ZSql = ZSql + "Credito ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Observaciones )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WProveedor + "',"
                ZSql = ZSql + "'" + WTipo + "',"
                ZSql = ZSql + "'" + Letra.Text + "',"
                ZSql = ZSql + "'" + WPunto + "',"
                ZSql = ZSql + "'" + WNumero + "',"
                ZSql = ZSql + "'" + WRenglon + "',"
                ZSql = ZSql + "'" + WVector1.TextMatrix(Ciclo, 1) + "',"
                ZSql = ZSql + "'" + WVector1.TextMatrix(Ciclo, 3) + "',"
                ZSql = ZSql + "'" + WVector1.TextMatrix(Ciclo, 4) + "',"
                ZSql = ZSql + "'" + Fecha.Text + "',"
                ZSql = ZSql + "'" + WOrdFecha + "',"
                ZSql = ZSql + "'" + Observaciones.Text + "')"
                spImpCyb = ZSql
                Set rstImpCyb = db.OpenRecordset(spImpCyb, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            If Val(WVector11.TextMatrix(Ciclo, 3)) <> 0 Then
            
                GrabaCentro = "N"
                WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero + WRenglon
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO ImpProy ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Proveedor ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Letra ,"
                ZSql = ZSql + "Punto ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Proyecto ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Concepto )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + WProveedor + "',"
                ZSql = ZSql + "'" + WTipo + "',"
                ZSql = ZSql + "'" + Letra.Text + "',"
                ZSql = ZSql + "'" + WPunto + "',"
                ZSql = ZSql + "'" + WNumero + "',"
                ZSql = ZSql + "'" + WRenglon + "',"
                ZSql = ZSql + "'" + WVector11.TextMatrix(Ciclo, 1) + "',"
                ZSql = ZSql + "'" + WVector11.TextMatrix(Ciclo, 3) + "',"
                ZSql = ZSql + "'" + Fecha.Text + "',"
                ZSql = ZSql + "'" + WOrdFecha + "',"
                ZSql = ZSql + "'" + Concepto.Text + "')"
                spImpProy = ZSql
                Set rstImpProy = db.OpenRecordset(spImpProy, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        Next Ciclo
        
        If GrabaCentro = "S" And Val(Centro.Text) <> 0 Then
        
            WRenglon = "1"
            Call Ceros(WRenglon, 2)
            
            ZImpo = Val(Neto.Text) + Val(Exento.Text)
            ZImpo = Pusing("###,###.##", ZImpo)
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpProy ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Proyecto ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Concepto )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + Letra.Text + "',"
            ZSql = ZSql + "'" + WPunto + "',"
            ZSql = ZSql + "'" + WNumero + "',"
            ZSql = ZSql + "'" + WRenglon + "',"
            ZSql = ZSql + "'" + Centro.Text + "',"
            ZSql = ZSql + "'" + ZImpo + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + Concepto.Text + "')"
            spImpProy = ZSql
            Set rstImpProy = db.OpenRecordset(spImpProy, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        Call CmdLimpiar_Click
        
    End If
    
    Proveedor.SetFocus
    
End Sub

Private Sub cmdDelete_Click()

    WNumero = Numero.Text
    Call Ceros(WNumero, 8)
    Numero.Text = WNumero
    Call Alinea_Datos
    XProveedor = Proveedor.Text
    XTipo = TipoComp.ListIndex
    XLetra = Letra.Text
    XPunto = Punto.Text
    XNumero = Numero.Text
    
    WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero
    WClaveCtaCte = WProveedor + Letra.Text + WTipo + WPunto + WNumero

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        rstIvaComp.Close
        T$ = "Comprobantes del Proveedor"
        m$ = "Desea Borrar el Comprobante "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            ZSql = ""
            ZSql = ZSql + "DELETE IvaComp"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
            spIvaComp = ZSql
            Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        
            ZSql = ""
            ZSql = ZSql + "DELETE CtaCtePrv"
            ZSql = ZSql + " Where Clave = " + "'" + WClaveCtaCte + "'"
            spCtaCtePrv = ZSql
            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "DELETE ImpCyb"
            ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
            ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Punto = " + "'" + WPunto + "'"
            ZSql = ZSql + " and Numero = " + "'" + WNumero + "'"
            spImpCyb = ZSql
            Set rstImpCyb = db.OpenRecordset(spImpCyb, dbOpenSnapshot, dbSQLPassThrough)
    
            ZSql = ""
            ZSql = ZSql + "DELETE ImpProy"
            ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
            ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Punto = " + "'" + WPunto + "'"
            ZSql = ZSql + " and Numero = " + "'" + WNumero + "'"
            spImpProy = ZSql
            Set rstImpProy = db.OpenRecordset(spImpProy, dbOpenSnapshot, dbSQLPassThrough)
        
            If movstk <> 0 Then
        
                Rem Erase WTrabajo
                Rem ZLugar = 0
        
                Rem For ZRenglon = 1 To 100
                    
                Rem     ZSql = ""
                Rem     ZSql = ZSql + "Select *"
                Rem     ZSql = ZSql + " FROM Compras"
                Rem     ZSql = ZSql + " Where Compras.Numero = " + "'" + Str$(movstk) + "'"
                Rem     ZSql = ZSql + " and Compras.Renglon = " + "'" + Str$(ZRenglon) + "'"
                Rem     spCompras = ZSql
                Rem     Set rstCompras = db.OpenRecordset(spCompras, dbOpenSnapshot, dbSQLPassThrough)
                Rem     If rstCompras.RecordCount > 0 Then
                Rem         ZLugar = ZLugar + 1
                Rem         WTrabajo(ZLugar, 1) = rstCompras!Articulo
                Rem         WTrabajo(ZLugar, 2) = rstCompras!Cantidad
                Rem         rstCompras.Close
                Rem     End If
                
                Rem Next ZRenglon
            
                Rem For ZRenglon = 1 To ZLugar
            
                Rem     WArticulo = WTrabajo(ZRenglon, 1)
                Rem     WCantidad = WTrabajo(ZRenglon, 2)
                
                Rem     ZSql = ""
                Rem     ZSql = ZSql + "UPDATE Articulo SET "
                Rem     ZSql = ZSql + " Stock = Stock - " + WCantidad
                Rem     ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                Rem     spArticulo = ZSql
                Rem     Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                Rem Next ZRenglon
            
                ZSql = ""
                ZSql = ZSql + "DELETE Compras"
                ZSql = ZSql + " Where Numero = " + "'" + Str$(movstk) + "'"
                spCompras = ZSql
                Set rstCompras = db.OpenRecordset(spCompras, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            Call CmdLimpiar_Click
            
        End If
    End If
    
    Proveedor.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Proveedor.Text = ""
    Tipo.Text = ""
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Periodo.Text = "  /  /    "
    Remito.Text = ""
    Paridad.Text = ""
    Neto.Text = ""
    Iva21.Text = ""
    Iva5.Text = ""
    Iva27.Text = ""
    Iva105.Text = ""
    Ib.Text = ""
    ImpInterno.Text = ""
    ImpCombustible.Text = ""
    Exento.Text = ""
    Concepto.Text = ""
    Centro.Text = ""
    Total.Caption = ""
    Contado1.Value = False
    Contado2.Value = True
    Contado3.Value = False
    Contado4.Value = False
    DesProveedor.Caption = ""
    DesConcepto.Caption = ""
    DesCentro.Caption = ""
    Observaciones.Text = ""
    ProveedorIva.Text = ""
    DesProveedorIva.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Remito.Text = ""
    IMpoRemito.Text = ""
    
    
    movstk = 0
    
    
    TipoComp.ListIndex = 0
    Tablas.Tab = 0
    
    Call Limpia_Vector
    Call Limpia_Vector1
    Call Limpia_Vector2
    
    Proveedor.SetFocus
    
End Sub

Private Sub CmdClose_Click()
    PrgCompras.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub LeeRemito_Click()

    If Val(Remito.Text) = 0 Then
        Exit Sub
    End If

    Call Limpia_Vector2
    Renglon = 0

    
    If ZZNivel = 1 Then
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Remito"
    ZSql = ZSql + " Where Remito.Proveedor = " + "'" + Proveedor.Text + "'"
    ZSql = ZSql + " and Remito.Remito = " + "'" + Remito.Text + "'"
    ZSql = ZSql + " Order by Remito.Clave"
        
    spRemito = ZSql
    Set rstRemito = db.OpenRecordset(spRemito, dbOpenSnapshot, dbSQLPassThrough)
    If rstRemito.RecordCount > 0 Then
    
        With rstRemito
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
                        
                    WVector22.Row = Renglon
                        
                    WVector22.Col = 1
                    WVector22.Text = rstRemito!Insumo
                        
                    WVector22.Col = 3
                    WVector22.Text = Pusing("###,###", Str$(rstRemito!Cantidad))
                    
                    WVector22.Col = 5
                    WVector22.Text = Str$(rstRemito!Orden)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstRemito.Close
    End If
    
    ZZSuma = 0
    
    For Ciclo = 1 To Renglon
    
        ZZInsumo = WVector22.TextMatrix(Ciclo, 1)
        ZZOrden = WVector22.TextMatrix(Ciclo, 5)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Insumo"
        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZZInsumo + "'"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        If rstInsumo.RecordCount > 0 Then
            WVector22.TextMatrix(Ciclo, 2) = rstInsumo!Descripcion
            rstInsumo.Close
        End If

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Orden"
        ZSql = ZSql + " Where Orden.Numero = " + "'" + ZZOrden + "'"
        ZSql = ZSql + " and Orden.Insumo = " + "'" + ZZInsumo + "'"
        spOrden = ZSql
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
        
            ZZPrecio = rstOrden!Precio
            ZZMoneda = rstOrden!MONEDA
            If ZZMoneda = 2 And Val(Paridad.Text) <> 0 Then
                ZZPrecio = ZZPrecio * Val(Paridad.Text)
            End If
            WVector22.TextMatrix(Ciclo, 4) = Str$(ZZPrecio)
            WVector22.TextMatrix(Ciclo, 4) = Pusing("###,###.###", WVector22.TextMatrix(Ciclo, 4))
            WVector22.TextMatrix(Ciclo, 5) = (Val(WVector22.TextMatrix(Ciclo, 3)) * Val(WVector22.TextMatrix(Ciclo, 4)))
            WVector22.TextMatrix(Ciclo, 5) = Pusing("###,###.##", WVector22.TextMatrix(Ciclo, 5))
            ZZSuma = ZZSuma + Val(WVector22.TextMatrix(Ciclo, 5))
            rstOrden.Close
        End If

    Next Ciclo
    
    IMpoRemito.Text = Str$(ZZSuma)
    IMpoRemito.Text = Pusing("###,###.##", IMpoRemito.Text)
    
    
    WVector22.Col = 1
    WVector22.Row = 1
    
    If ZZNivel = 1 Then
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        txtOdbc = "FraganciasII"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If


End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WProveedor = Trim(UCase(Proveedor.Text))
        Proveedor.Text = WProveedor
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = rstProveedor!Nombre
            WDias = rstProveedor!Dias + 1
            rstProveedor.Close
            Letra.SetFocus
                Else
            Proveedor.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
    End If
End Sub

Private Sub Letra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Letra.Text = UCase(Letra.Text)
        If Left$(Letra.Text, 1) = "A" Or Left$(Letra.Text, 1) = "B" Or Left$(Letra.Text, 1) = "C" Or Left$(Letra.Text, 1) = "X" Or Left$(Letra.Text, 1) = "Z" Or Left$(Letra.Text, 1) = "M" Then
            Punto.SetFocus
                Else
            Letra.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Letra.Text = ""
    End If
End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
        Punto.Text = WPunto
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WNumero = Numero.Text
        Call Ceros(WNumero, 8)
        Numero.Text = WNumero
        Call Alinea_Datos
        
        XProveedor = Proveedor.Text
        XTipo = TipoComp.ListIndex
        XLetra = Letra.Text
        XPunto = Punto.Text
        XNumero = Numero.Text
        
        WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM IvaComp"
        ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
        spIvaComp = ZSql
        Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvaComp.RecordCount > 0 Then
            rstIvaComp.Close
            Proveedor.Text = XProveedor
            TipoComp.ListIndex = XTipo
            Letra.Text = XLetra
            Punto.Text = XPunto
            Numero.Text = XNumero
            Call Imprime_Datos
            Call Proceso
            Existe = "S"
                Else
            CmdLimpiar_Click
            Proveedor.Text = XProveedor
            TipoComp.ListIndex = XTipo
            Letra.Text = XLetra
            Punto.Text = XPunto
            Numero.Text = XNumero
            Existe = "N"
            Call Imprime_Descripcion
        End If
        Tipo.Text = Str$(TipoComp.ListIndex + 1)
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            If Periodo.Text = "  /  /    " Then
                Periodo.Text = Fecha.Text
            End If
            If Vencimiento.Text = "  /  /    " Then
                WFecha = Fecha.Text
                Call Calcula_vencimiento(WFecha, WDias, WVencimiento)
                Vencimiento.Text = WVencimiento
            End If
            Vencimiento.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub



Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            WFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            WVencimiento = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
            If WVencimiento >= WFecha Then
                Periodo.SetFocus
                    Else
                Vencimiento.SetFocus
            End If
                Else
            Vencimiento.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vencimiento.Text = "  /  /    "
    End If
End Sub

Private Sub Periodo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Periodo.Text, Auxi)
        If Auxi = "S" Then
            Paridad.SetFocus
                Else
            Periodo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Periodo.Text = "  /  /    "
    End If
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Neto.SetFocus
    End If
    If KeyAscii = 27 Then
        Paridad.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Neto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Neto.Text = Pusing("###,###.##", Neto.Text)
        If (Letra.Text = "A" Or Letra.Text = "M") And Val(Iva21.Text) = 0 Then
            Iva21.Text = Val(Neto.Text) * (ConfigIva1 / 100)
            Iva21.Text = Pusing("###,###.##", Iva21.Text)
        End If
        Call Calcula_total
        If Letra.Text = "A" Or Letra.Text = "M" Then
            Iva21.SetFocus
                Else
            Exento.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Neto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva21.Text = Pusing("###,###.##", Iva21.Text)
        Call Calcula_total
        Iva5.SetFocus
    End If
    If KeyAscii = 27 Then
        Iva21.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva5.Text = Pusing("###,###.##", Iva5.Text)
        Call Calcula_total
        Iva27.SetFocus
    End If
    If KeyAscii = 27 Then
        Iva5.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva27_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva27.Text = Pusing("###,###.##", Iva27.Text)
        Call Calcula_total
        Ib.SetFocus
    End If
    If KeyAscii = 27 Then
        Iva27.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ib_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ib.Text = Pusing("###,###.##", Ib.Text)
        Call Calcula_total
        Exento.SetFocus
    End If
    If KeyAscii = 27 Then
        Ib.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Exento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Exento.Text = Pusing("###,###.##", Exento.Text)
        Call Calcula_total
        ImpInterno.SetFocus
    End If
    If KeyAscii = 27 Then
        Exento.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ImpInterno_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ImpInterno.Text = Pusing("###,###.##", ImpInterno.Text)
        Call Calcula_total
        ImpCombustible.SetFocus
    End If
    If KeyAscii = 27 Then
        ImpInterno.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ImpCombustible_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ImpCombustible.Text = Pusing("###,###.##", ImpCombustible.Text)
        Call Calcula_total
        Iva105.SetFocus
    End If
    If KeyAscii = 27 Then
        ImpCombustible.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Iva105_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva105.Text = Pusing("###,###.##", Iva105.Text)
        Call Calcula_total
        Concepto.SetFocus
    End If
    If KeyAscii = 27 Then
        Iva105.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Concepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Conceptos"
        ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
        spConceptos = ZSql
        Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        If rstConceptos.RecordCount > 0 Then
            DesConcepto.Caption = rstConceptos!Nombre
            Observaciones.SetFocus
            rstConceptos.Close
                Else
            Concepto.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Concepto.Text = ""
        DesConcepto.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Centro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Centro.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + Centro.Text + "'"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                DesCentro.Caption = rstProyecto!Descripcion
                Observaciones.SetFocus
                rstProyecto.Close
                    Else
                Centro.SetFocus
            End If
                Else
            DesCentro.Caption = ""
            Observaciones.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Centro.Text = ""
        DesCentro.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Fecha.SetFocus
        Rem Tablas.Tab = 1
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub ProveedorIva_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ProveedorIva.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            DesProveedorIva.Caption = rstProveedor!Nombre
            rstProveedor.Close
            Call Format_datos
        End If
    End If
    If KeyAscii = 27 Then
        ProveedorIva.Text = ""
        DesProveedor.Caption = "2"
    End If
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            DesBanco.Caption = rstBanco!Nombre
            rstBanco.Close
                Else
            DesBanco.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        Banco.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call LeeRemito_Click
    End If
    If KeyAscii = 27 Then
        Remito.Text = ""
        Call LeeRemito_Click
    End If
End Sub


Private Sub Consulta_Click()

     Pantalla.Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Conceptos de Compra"
     Opcion.AddItem "Cuentas Contables"
     Opcion.AddItem "Proyectos"
     Opcion.AddItem "Proveedores Iva"
     Opcion.AddItem "Articulos"
     Opcion.AddItem "Centro de Costos"

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
        Case 0, 4
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
                            IngresaItem = rstProveedor!Proveedor + " " + rstProveedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Order by Conceptos.Concepto"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                With rstConceptos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstConceptos!Concepto) + " " + rstConceptos!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstConceptos!Concepto
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstConceptos.Close
            End If
        
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Order by Cuenta.Cuenta"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                With rstCuenta
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCuenta!Cuenta + " " + rstCuenta!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCuenta!Cuenta
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCuenta.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Order by Proyecto.Codigo"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstProyecto!Codigo) + " " + rstProyecto!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstProyecto!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
                
            End If
            
        Case 5
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
            
        Case 6
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Order by Banco.Banco"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                With rstBanco
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Banco) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Banco
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstBanco.Close
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
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Concepto.Text = WIndice.List(Indice)
            Call Concepto_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            WTexto1.Text = WIndice.List(Indice)
            Call WTexto1_KeyDown(13, 0)
            
        Case 3
            If ZSalida = 0 Then
                Indice = Pantalla.ListIndex
                Centro.Text = WIndice.List(Indice)
                Call Centro_KeyPress(13)
                    Else
                Indice = Pantalla.ListIndex
                WTexto21.Text = WIndice.List(Indice)
                Call WTexto21_KeyDown(13, 0)
            End If
            
        Case 4
            Indice = Pantalla.ListIndex
            ProveedorIva.Text = WIndice.List(Indice)
            Call ProveedorIva_KeyPress(13)
            
        Case 5
            Indice = Pantalla.ListIndex
            WTexto12.Text = WIndice.List(Indice)
            Call WTexto12_KeyDown(13, 0)
            
        Case 6
            Indice = Pantalla.ListIndex
            Banco.Text = WIndice.List(Indice)
            Call Banco_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Call Limpia_Vector1
    Call Limpia_Vector2
    
    Tablas.TabCaption(0) = "Datos del Comprobante"
    Tablas.TabCaption(1) = "Asiento Contable"
    Tablas.TabCaption(2) = "Centro de Costo"
    Tablas.TabCaption(3) = "Articulos"
    
    Tablas.Tab = 0

    Proveedor.Text = ""
    Tipo.Text = ""
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Vencimiento.Text = "  /  /    "
    Periodo.Text = "  /  /    "
    Remito.Text = ""
    Paridad.Text = ""
    Neto.Text = ""
    Iva21.Text = ""
    Iva5.Text = ""
    Iva27.Text = ""
    Iva105.Text = ""
    Ib.Text = ""
    ImpInterno.Text = ""
    ImpCombustible.Text = ""
    Exento.Text = ""
    Total.Caption = ""
    Contado1.Value = False
    Contado2.Value = True
    Contado3.Value = False
    Contado4.Value = False
    Concepto.Text = ""
    Centro.Text = ""
    Observaciones.Text = ""
    ProveedorIva.Text = ""
    DesProveedor.Caption = ""
    DesConcepto.Caption = ""
    DesCentro.Caption = ""
    DesProveedor.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Remito.Text = ""
    IMpoRemito.Text = ""
    
    movstk = 0
    
    TipoComp.Clear
    
    TipoComp.AddItem "Factura"
    TipoComp.AddItem "N.Debito"
    TipoComp.AddItem "N.Credito"
    TipoComp.AddItem ""
    TipoComp.AddItem ""
    TipoComp.AddItem ""
    TipoComp.AddItem "Ticket"
    TipoComp.AddItem "Recibo"
    
    TipoComp.ListIndex = 0
    
    Tablas.Tab = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Configuracion"
    ZSql = ZSql + " Where Configuracion.Clave = 1"
    spConfiguracion = ZSql
    Set rstConfiguracion = db.OpenRecordset(spConfiguracion, dbOpenSnapshot, dbSQLPassThrough)
    If rstConfiguracion.RecordCount > 0 Then
        ConfigIva1 = rstConfiguracion!Iva1
        ConfigIva2 = rstConfiguracion!Iva2
        ConfigPercepcion = rstConfiguracion!Percepcion
        ConfigPunto = rstConfiguracion!Punto
        rstConfiguracion.Close
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
        Case 0, 4
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
                            IngresaItem = rstProveedor!Proveedor + " " + rstProveedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Where Conceptos.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Conceptos.Concepto"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                With rstConceptos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstConceptos!Concepto) + " " + rstConceptos!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstConceptos!Concepto
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstConceptos.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cuenta.Cuenta"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                With rstCuenta
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCuenta!Cuenta + " " + rstCuenta!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCuenta!Cuenta
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCuenta.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Where Proyecto.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Proyecto.Codigo"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstProyecto!Codigo) + " " + rstProyecto!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstProyecto!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
            
        Case 6
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Banco.Banco"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                With rstBanco
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Banco) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Banco
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstBanco.Close
            End If
            
            
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Concepto de Compra"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Concepto_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Concepto de Compra"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Centro_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Conceptos"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Proyectos"
    Rem Opcion.Visible = True
    ZSalida = 0
    Opcion.ListIndex = 3

End Sub

Private Sub ProveedorIva_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Concepto de Compra"
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem "Proveedores Iva"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 4
    
    Call Opcion_Click

End Sub

Private Sub Banco_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Concepto de Compra"
    Opcion.AddItem ""
    Opcion.AddItem ""
    Opcion.AddItem "Proveedores Iva"
    Opcion.AddItem "Articulo"
    Opcion.AddItem "Banco"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 6
    
    Call Opcion_Click

End Sub

Private Sub Proceso()

    Call Limpia_Vector
    Call Limpia_Vector1
    Call Limpia_Vector2
    
    Lugar1 = 0
    Lugar2 = 0
    
    For Ciclo = 1 To 100
    
        WRenglon = Str$(Ciclo)
        Call Ceros(WRenglon, 2)
        
        WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero + WRenglon
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ImpCyb"
        ZSql = ZSql + " Where ImpCyb.Clave = " + "'" + WClave + "'"
        spImpCyb = ZSql
        Set rstImpCyb = db.OpenRecordset(spImpCyb, dbOpenSnapshot, dbSQLPassThrough)
        If rstImpCyb.RecordCount > 0 Then
        
            Lugar1 = Lugar1 + 1
            WVector1.Row = Lugar1
            WVector1.Col = 1
            WVector1.Text = rstImpCyb!Cuenta
            WVector1.Col = 3
            WVector1.Text = Str$(rstImpCyb!Debito)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            WVector1.Col = 4
            WVector1.Text = Str$(rstImpCyb!Credito)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            rstImpCyb.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.TextMatrix(WVector1.Row, 1) + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstCuenta!Descripcion
                rstCuenta.Close
            End If
            
        End If
        
        WClave = WProveedor + WTipo + Letra.Text + WPunto + WNumero + WRenglon
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ImpProy"
        ZSql = ZSql + " Where ImpProy.Clave = " + "'" + WClave + "'"
        spImpProy = ZSql
        Set rstImpProy = db.OpenRecordset(spImpProy, dbOpenSnapshot, dbSQLPassThrough)
        If rstImpProy.RecordCount > 0 Then
        
            Lugar2 = Lugar2 + 1
            WVector11.Row = Lugar2
            WVector11.Col = 1
            WVector11.Text = rstImpProy!Proyecto
            ZProyecto = rstImpProy!Proyecto
            WVector11.Col = 3
            WVector11.Text = Str$(rstImpProy!Importe)
            WVector11.Text = Pusing("###,###.##", WVector11.Text)
            rstImpProy.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + ZProyecto + "'"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                WVector11.Col = 2
                WVector11.Text = rstProyecto!Descripcion
                rstProyecto.Close
            End If
        End If
        
    Next Ciclo
    
    Debito.Text = ""
    Credito.Text = ""
    WDebito = 0
    WCredito = 0
    For Ciclo = 1 To 100
        WDebito = WDebito + Val(WVector1.TextMatrix(Ciclo, 3))
        WCredito = WCredito + Val(WVector1.TextMatrix(Ciclo, 4))
    Next Ciclo
    Debito.Text = Str$(WDebito)
    Credito.Text = Str$(WCredito)
    Debito.Text = Pusing(WFormato(3), Debito.Text)
    Credito.Text = Pusing(WFormato(4), Credito.Text)
    
    SumaProyecto.Text = ""
    WSumaProyecto = 0
    For Ciclo = 1 To 100
        WSumaProyecto = WSumaProyecto + Val(WVector11.TextMatrix(Ciclo, 3))
    Next Ciclo
    SumaProyecto.Text = Str$(WSumaProyecto)
    SumaProyecto.Text = Pusing(WFormato1(3), SumaProyecto.Text)
        
    For ZRenglon = 1 To 100
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Compras"
        ZSql = ZSql + " Where Compras.Numero = " + "'" + Str$(movstk) + "'"
        ZSql = ZSql + " and Compras.Renglon = " + "'" + Str$(ZRenglon) + "'"
        spCompras = ZSql
        Set rstCompras = db.OpenRecordset(spCompras, dbOpenSnapshot, dbSQLPassThrough)
        If rstCompras.RecordCount > 0 Then
            WVector22.TextMatrix(ZRenglon, 1) = rstCompras!Articulo
            WVector22.TextMatrix(ZRenglon, 3) = Pusing("###,###", Str$(rstCompras!Cantidad))
            WVector22.TextMatrix(ZRenglon, 4) = Pusing("###,###.###", Str$(rstCompras!Costo))
            WVector22.TextMatrix(ZRenglon, 5) = Pusing("###,###.##", Str$(rstCompras!Cantidad * rstCompras!Costo))
            Auxi1 = rstCompras!Articulo
            rstCompras.Close
        End If
    
    Next ZRenglon
    
    For ZRenglon = 1 To 100
    
        If Trim(WVector22.TextMatrix(ZRenglon, 1)) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Insumo"
            ZSql = ZSql + " Where Insumo.Codigo = " + "'" + WVector22.TextMatrix(ZRenglon, 1) + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstInsumo.RecordCount > 0 Then
                WVector22.TextMatrix(ZRenglon, 2) = rstInsumo!Descripcion
                rstInsumo.Close
            End If
        End If
        
    Next ZRenglon
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    On Error GoTo WError
    
    Select Case Tablas.Tab
        Case 0
            Fecha.SetFocus
        Case 1
            If Val(Debito.Text) = 0 And Val(Credito.Text) = 0 Then
            
                Tipo.Text = Str$(TipoComp.ListIndex + 1)
                WTipo = Tipo.Text
                Call Ceros(WTipo, 2)
                Tipo.Text = WTipo
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Empresa"
                ZSql = ZSql + " Where Empresa.Empresa = " + "'" + WEmpresa + "'"
                spEmpresa = ZSql
                Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
                If rstEmpresa.RecordCount > 0 Then
                    WCtaEfectivo = rstEmpresa!CtaEfectivo
                    WCtaFondoFijo = rstEmpresa!CtaFondoFijo
                    WCtaIva21 = rstEmpresa!CtaIva21
                    WCtaIva5 = rstEmpresa!CtaIva5
                    WCtaIva27 = rstEmpresa!CtaIva27
                    WCtaIb = rstEmpresa!CtaIb
                    WCtaImpInterno = rstEmpresa!CtaImpInterno
                    WCtaImpCombustible = rstEmpresa!CtaImpCombustible
                    WCtaIva105 = rstEmpresa!CtaIva105
                    WCtaProveedores = rstEmpresa!CtaProveedores
                    rstEmpresa.Close
                End If
            
                WFila = 0
                
                If Val(Total.Caption) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    If Contado1.Value = True Then
                        WVector1.Text = WCtaEfectivo
                            Else
                        If Contado3.Value = True Then
                            WVector1.Text = WCtaFondoFijo
                                Else
                            If Contado4.Value = True Then
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM Banco"
                                ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
                                spBanco = ZSql
                                Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                                If rstBanco.RecordCount > 0 Then
                                    WVector1.Text = rstBanco!Cuenta
                                    rstBanco.Close
                                End If
                                    Else
                                WVector1.Text = WCtaProveedores
                            End If
                        End If
                    End If
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = ""
                            WVector1.Col = 4
                            WVector1.Text = Total.Caption
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = ""
                            WVector1.Col = 3
                            WVector1.Text = Total.Caption
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    End Select
                            
                End If
                
                If Val(Neto.Text) <> 0 Or Val(Exento.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Conceptos"
                    ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
                    spConceptos = ZSql
                    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstConceptos.RecordCount > 0 Then
                        WVector1.Text = rstConceptos!Cuenta
                        rstConceptos.Close
                    End If
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = Str$(Val(Neto.Text) + Val(Exento.Text))
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = Str$(Val(Neto.Text) + Val(Exento.Text))
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                If Val(Iva21.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    WVector1.Text = WCtaIva21
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = Iva21.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = Iva21.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                If Val(Iva5.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    WVector1.Text = WCtaIva5
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = Iva5.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = Iva5.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                If Val(Iva27.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    WVector1.Text = WCtaIva27
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = Iva27.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = Iva27.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                If Val(Ib.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    WVector1.Text = WCtaIb
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = Ib.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = Ib.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                If Val(ImpInterno.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    WVector1.Text = WCtaImpInterno
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = ImpInterno.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = ImpInterno.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                If Val(ImpCombustible.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    WVector1.Text = WCtaImpCombustible
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = ImpCombustible.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = ImpCombustible.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                If Val(Iva105.Text) <> 0 Then
                    WFila = WFila + 1
                    WVector1.Row = WFila
                    WVector1.Col = 1
                    WVector1.Text = WCtaIva105
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cuenta"
                    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                    spCuenta = ZSql
                    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCuenta.RecordCount > 0 Then
                        WVector1.Col = 2
                        WVector1.Text = rstCuenta!Descripcion
                        rstCuenta.Close
                    End If
                    
                    Select Case Val(WTipo)
                        Case 1, 2, 7
                            WVector1.Col = 3
                            WVector1.Text = Iva105.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 4
                            WVector1.Text = ""
                        Case Else
                            WVector1.Col = 4
                            WVector1.Text = Iva105.Text
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            WVector1.Col = 3
                            WVector1.Text = ""
                    End Select
                End If
                
                Debito.Text = ""
                Credito.Text = ""
                WDebito = 0
                WCredito = 0
                For Ciclo = 1 To 100
                    WDebito = WDebito + Val(WVector1.TextMatrix(Ciclo, 3))
                    WCredito = WCredito + Val(WVector1.TextMatrix(Ciclo, 4))
                Next Ciclo
                Debito.Text = Str$(WDebito)
                Credito.Text = Str$(WCredito)
                Debito.Text = Pusing(WFormato(3), Debito.Text)
                Credito.Text = Pusing(WFormato(4), Credito.Text)
                
            End If
            
            WImpo1.Caption = Total.Caption
            WImpo1.Caption = Pusing("###,###.##", WImpo1.Caption)
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        Case 2
            Wimpo2.Caption = Str$(Val(Neto.Text) + Val(Exento.Text))
            Wimpo2.Caption = Pusing("###,###.##", Wimpo2.Caption)
            WVector11.Col = 1
            WVector11.Row = 1
            Call StartEdit1
        Case 3
            WVector12.Col = 1
            WVector12.Row = 1
            Call StartEdit2
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub


Rem
Rem Controles de la grilla
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
        Debito.Text = ""
        Credito.Text = ""
        WDebito = 0
        WCredito = 0
        For Ciclo = 1 To 100
            WDebito = WDebito + Val(WVector1.TextMatrix(Ciclo, 3))
            WCredito = WCredito + Val(WVector1.TextMatrix(Ciclo, 4))
        Next Ciclo
        Debito.Text = Str$(WDebito)
        Credito.Text = Str$(WCredito)
        Debito.Text = Pusing(WFormato(3), Debito.Text)
        Credito.Text = Pusing(WFormato(4), Credito.Text)
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
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
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
                Call Control_Grilla
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
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
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
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

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 1
            WVector1.Col = WVector1.Col + 2
        Case 4
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
            If Val(WVector1.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + WVector1.Text + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstCuenta!Descripcion
                    WVector1.Col = XColumna
                    rstCuenta.Close
                        Else
                    WControl = "N"
                End If
                
            End If
        Case 3, 4
            WVector1.Col = XColumna
        Case Else
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Then
        Exit Sub
    End If

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
        WVector1.Col = 4
        WAuxi3 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Or WAuxi3 <> "" Then
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
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
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

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 5
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
                WVector1.Text = "Cuenta"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Debito"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 4
                WVector1.Text = "Credito"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
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
        
        Rem despliega los totales
        
        Select Case Ciclo
            Case 3
                Debito.Text = ""
                Debito.Left = WTitulo(Ciclo).Left
                Debito.Top = 3240
                Debito.Width = WTitulo(Ciclo).Width
                Debito.Height = WTitulo(Ciclo).Height
                Debito.Visible = True
            Case 4
                Credito.Text = ""
                Credito.Left = WTitulo(Ciclo).Left
                Credito.Top = 3240
                Credito.Width = WTitulo(Ciclo).Width
                Credito.Height = WTitulo(Ciclo).Height
                Credito.Visible = True
            Case Else
        End Select
        
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
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
    Select Case WVector1.Col
        Case 1
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Conceptos"
            Opcion.AddItem "Cuentas Contables"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 2
    
            Call Opcion_Click
        Case Else
    End Select
End Sub

Private Sub WTexto2_DblClick()
    Select Case WVector1.Col
        Case 1
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Conceptos"
            Opcion.AddItem "Cuentas Contables"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 2
    
            Call Opcion_Click
        Case Else
    End Select
End Sub



Rem
Rem
Rem
Rem aca empieza los comandos del vector2
Rem
Rem
Rem
Rem


Private Sub GridEditText1(ByVal KeyAscii As Integer)

    XColumna = WVector11.Col
    XTipoDato = WParametros1(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto11.Left = WVector11.CellLeft + WVector11.Left
            WTexto11.Top = WVector11.CellTop + WVector11.Top
            WTexto11.Width = WVector11.CellWidth
            WTexto11.Height = WVector11.CellHeight
            WTexto11.MaxLength = WParametros1(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto11.Text = WVector11.Text
                    WTexto11.SelStart = Len(WTexto11.Text)
                Case Else
                    WTexto11.Text = Chr$(KeyAscii)
                    WTexto11.SelStart = 1
            End Select
            WTexto11.Visible = True
            WTexto11.SetFocus
        Case 1
            WTexto21.Left = WVector11.CellLeft + WVector11.Left
            WTexto21.Top = WVector11.CellTop + WVector11.Top
            WTexto21.Width = WVector11.CellWidth
            WTexto21.Height = WVector11.CellHeight
            WTexto21.MaxLength = WParametros1(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto21.Text = WVector11.Text
                    Rem WTexto21.SelStart = Len(WTexto21.Text)
                    WTexto21.SelStart = 0
                Case Else
                    WTexto21.Text = Chr$(KeyAscii)
                    WTexto21.SelStart = 1
            End Select
            WTexto21.Visible = True
            WTexto21.SetFocus
        Case 2
            WTexto31.Left = WVector11.CellLeft + WVector11.Left
            WTexto31.Top = WVector11.CellTop + WVector11.Top
            WTexto31.Width = WVector11.CellWidth
            WTexto31.Height = WVector11.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector11.Text) = 10 Then
                        WTexto31.Text = WVector11.Text
                            Else
                        WTexto31.Text = "  /  /    "
                    End If
                    WTexto31.SelStart = 0
                Case Else
                    WTexto31.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto31.SelStart = 1
            End Select
            WTexto31.Visible = True
            WTexto31.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit1()
    Pasa = 0
    If WCombo11.Visible Then
        Pasa = 0
        WVector11.Text = WCombo11.Text
        WCombo11.Visible = False
            Else
        If WTexto11.Visible Then
            Pasa = 1
            WVector11.Text = WTexto11.Text
            WTexto11.Visible = False
                Else
            If WTexto21.Visible Then
                Pasa = 1
                WVector11.Text = WTexto21.Text
                WTexto21.Visible = False
                    Else
                If WTexto31.Visible Then
                    Pasa = 1
                    WVector11.Text = WTexto31.Text
                    WTexto31.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato1(WVector11.Col) <> "" Then
            WVector11.Text = Pusing(WFormato1(WVector11.Col), WVector11.Text)
        End If
        SumaProyecto.Text = ""
        WSumaProyecto = 0
        For Ciclo = 1 To 100
            WSumaProyecto = WSumaProyecto + Val(WVector11.TextMatrix(Ciclo, 3))
        Next Ciclo
        SumaProyecto.Text = Str$(WSumaProyecto)
        SumaProyecto.Text = Pusing(WFormato(3), SumaProyecto.Text)
    End If
End Sub

Private Sub GridEditCombo1()
    ' Position the ComboBox over the cell.
    WCombo11.Left = WVector11.CellLeft + WVector11.Left
    WCombo11.Top = WVector11.CellTop + WVector11.Top
    WCombo11.Width = WVector11.CellWidth
    WCombo11.Visible = True
    WCombo11.SetFocus
End Sub

Private Sub WTexto11_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto11.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto11.Text = WVector11.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                Call Control_Grilla1
            End If
            Call StartEdit1

        Case vbKeyDown
            ' Move down 1 row.
            WVector11.SetFocus
            DoEvents
            If WVector11.Row < WVector11.Rows - 1 Then
                Call Control_Campo1
                If WControl1 = "S" Then
                    WVector11.Row = WVector11.Row + 1
                End If
            End If
            Call StartEdit1

        Case vbKeyUp
            ' Move up 1 row.
            WVector11.SetFocus
            DoEvents
            If WVector11.Row > WVector11.FixedRows Then
                Call Control_Campo1
                If WControl1 = "S" Then
                    WVector11.Row = WVector11.Row - 1
                End If
            End If
            Call StartEdit1

    End Select
End Sub

Private Sub WTexto21_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto21.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto21.Text = WVector11.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                Call Control_Grilla1
            End If
            Call StartEdit1
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector11.SetFocus
            DoEvents
            If WVector11.Row < WVector11.Rows - 1 Then
                Call Control_Campo1
                If WControl1 = "S" Then
                    WVector11.Row = WVector11.Row + 1
                End If
            End If
            Call StartEdit1

        Case vbKeyUp
            ' Move up 1 row.
            WVector11.SetFocus
            DoEvents
            If WVector11.Row > WVector11.FixedRows Then
                Call Control_Campo1
                If WControl1 = "S" Then
                    WVector11.Row = WVector11.Row - 1
                End If
            End If
            Call StartEdit1

    End Select
End Sub

Private Sub WTexto31_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto31.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto31.Text = WVector11.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector11.SetFocus
            Call Control_Campo1
            If WControl1 = "S" Then
                Call Control_Grilla1
            End If
            Call StartEdit1

        Case vbKeyDown
            ' Move down 1 row.
            WVector11.SetFocus
            DoEvents
            If WVector11.Row < WVector11.Rows - 1 Then
                Call Control_Campo1
                If WControl1 = "S" Then
                    WVector11.Row = WVector11.Row + 1
                End If
            End If
            Call StartEdit1

        Case vbKeyUp
            ' Move up 1 row.
            WVector11.SetFocus
            DoEvents
            If WVector11.Row > WVector11.FixedRows Then
                Call Control_Campo1
                If WControl1 = "S" Then
                    WVector11.Row = WVector11.Row - 1
                End If
            End If
            Call StartEdit1

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto11_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto21_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto31_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo11_Click()
    WVector11.SetFocus
End Sub

Private Sub WVector11_Click()
    StartEdit1
End Sub

Private Sub WVector11_LeaveCell()
    EndEdit1
End Sub

Private Sub WVector11_GotFocus()
    EndEdit1
End Sub

Rem Desde aca empieza las rutinas a cambiar

Private Sub StartEdit1()
    Select Case WParametros1(4, WVector11.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo11.Clear
            WCombo11.AddItem "Campo1"
            WCombo11.AddItem "Campo2"
            On Error Resume Next
            WCombo11.Text = WVector11.Text
            On Error GoTo 0
            GridEditCombo1
        Case Else
            If WParametros1(2, WVector11.Col) = 0 Then
                GridEditText1 Asc(" ")
            End If
    End Select
End Sub

Private Sub WVector11_KeyPress(KeyAscii As Integer)
    XColumna = WVector11.Col
    Select Case WParametros1(4, WVector11.Col)
        Case 1
        Case Else
            If WParametros1(2, XColumna) = 0 Then
                GridEditText1 KeyAscii
            End If
    End Select
End Sub

Private Sub Control_Grilla1()
    Select Case WVector11.Col
        Case 1
            WVector11.Col = WVector11.Col + 2
        Case 3
            If WVector11.Row < WVector11.Rows - 1 Then
                WVector11.Row = WVector11.Row + 1
            End If
            WVector11.Col = 1
        Case Else
            If WVector11.Col < WVector11.Cols - 1 Then
                WVector11.Col = WVector11.Col + 1
            End If
    End Select
    WVector11.SetFocus
    GridEditText1 KeyAscii
End Sub

Private Sub Control_Campo1()
    XColumna = WVector11.Col
    XFila = WVector11.Row
    WControl1 = "S"
    Select Case XColumna
        Case 1
            If Val(WVector11.Text) <> 0 Then
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proyecto"
                ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + WVector11.Text + "'"
                spProyecto = ZSql
                Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
                If rstProyecto.RecordCount > 0 Then
                    WVector11.Col = 2
                    WVector11.Text = rstProyecto!Descripcion
                    WVector11.Col = XColumna
                    rstProyecto.Close
                        Else
                    WControl1 = "N"
                End If
                
            End If
        Case 3, 4, 5
            WVector11.Col = XColumna
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector1()

    WVector11.Clear
    
    Rem ponga la grilla en negritas
    WVector11.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto11.FontName = WVector11.FontName
    WTexto11.FontSize = WVector11.FontSize
    WTexto11.Visible = False
    WTexto21.FontName = WVector11.FontName
    WTexto21.FontSize = WVector11.FontSize
    WTexto21.Visible = False
    WTexto31.FontName = WVector11.FontName
    WTexto31.FontSize = WVector11.FontSize
    WTexto31.Visible = False
    WCombo11.FontName = WVector11.FontName
    WCombo11.FontSize = WVector11.FontSize
    WCombo11.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector11.FixedCols = 1
    WVector11.Cols = 4
    WVector11.FixedRows = 1
    WVector11.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector11.Text = "Articulo"
    
    Rem Longitud
    Rem WVector11.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros1(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros1(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros1(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros1(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector11.ColWidth(0) = 200
    WVector11.Row = 0
    For Ciclo = 1 To WVector11.Cols - 1
        WVector11.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector11.Text = "Concepto"
                WVector11.ColWidth(Ciclo) = 1500
                WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros1(1, Ciclo) = 10
                WParametros1(2, Ciclo) = 0
                WParametros1(3, Ciclo) = 1
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = ""
            Case 2
                WVector11.Text = "Descripcion"
                WVector11.ColWidth(Ciclo) = 4000
                WVector11.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros1(1, Ciclo) = 50
                WParametros1(2, Ciclo) = 1
                WParametros1(3, Ciclo) = 0
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = ""
            Case 3
                WVector11.Text = "Importe"
                WVector11.ColWidth(Ciclo) = 1500
                WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros1(1, Ciclo) = 10
                WParametros1(2, Ciclo) = 0
                WParametros1(3, Ciclo) = 1
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = "###,###.##"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector11.Row = 0
    For Ciclo = 1 To WVector11.Cols - 1
        WVector11.Col = Ciclo
        WTitulo1(Ciclo).Text = WVector11.Text
        WTitulo1(Ciclo).Left = WVector11.CellLeft + WVector11.Left
        WTitulo1(Ciclo).Top = WVector11.CellTop + WVector11.Top
        WTitulo1(Ciclo).Width = WVector11.CellWidth
        WTitulo1(Ciclo).Height = WVector11.CellHeight
        WTitulo1(Ciclo).Visible = True
        
        Rem despliega los totales
        
        Select Case Ciclo
            Case 3
                SumaProyecto.Text = ""
                SumaProyecto.Left = WTitulo1(Ciclo).Left
                SumaProyecto.Top = 3240
                SumaProyecto.Width = WTitulo1(Ciclo).Width
                SumaProyecto.Height = WTitulo1(Ciclo).Height
                SumaProyecto.Visible = True
            Case Else
        End Select
        
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector11.Cols - 1
        WAncho = WAncho + WVector11.ColWidth(Ciclo)
    Next Ciclo
    WVector11.Width = WAncho

    ' Size the columns.
    Font.Name = WVector11.Font.Name
    Font.Size = WVector11.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector11.AllowUserResizing = flexResizeBoth
    
    WVector11.Col = 1
    WVector11.Row = 1
    
End Sub

Private Sub WVector11_Scroll()
    WTexto11.Visible = False
    WTexto21.Visible = False
    WTexto31.Visible = False
End Sub

Private Sub WVector11_DblClick()

    If WVector11.Col = 0 Then
        Exit Sub
    End If

    WTexto11.Visible = False
    WTexto21.Visible = False
    WTexto31.Visible = False

    For Ciclo = 1 To WVector11.Cols - 1
        WVector11.Col = Ciclo
        WVector11.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector11.Rows - 1
        WVector11.Row = Ciclo
        WVector11.Col = 1
        WAuxi1 = WVector11.Text
        WVector11.Col = 3
        WAuxi2 = WVector11.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector11.Cols - 1
                WVector11.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector11.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector1
    
    For Ciclo = 1 To EntraVector
        WVector11.Row = Ciclo
        For da = 1 To WVector11.Cols - 1
            WVector11.Col = da
            WVector11.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
End Sub

Private Sub WTexto21_DblClick()
    Select Case WVector11.Col
        Case 1
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Conceptos"
            Opcion.AddItem "Cuentas Contables"
            Opcion.AddItem "Proyectos"
            Rem Opcion.Visible = True
            ZSalida = 1
            Opcion.ListIndex = 3
    
            Call Opcion_Click
        Case Else
    End Select
End Sub

Private Sub WTexto11_DblClick()
    Select Case WVector11.Col
        Case 1
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Conceptos"
            Opcion.AddItem "Cuentas Contables"
            Opcion.AddItem "Proyectos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 3
    
            Call Opcion_Click
        Case Else
    End Select
End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TipoComp_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Letra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Punto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Contado1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Contado2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Contado3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Contado4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Periodo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Neto_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Ib_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ImpInterno_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ImpCombustible_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva105_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva21_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva27_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Exento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Concepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Centro_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub ProveedorIva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Banco_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub




Rem
Rem
Rem
Rem aca empieza los comandos del vector3
Rem
Rem
Rem
Rem


Private Sub GridEditText2(ByVal KeyAscii As Integer)

    XColumna = WVector22.Col
    XTipoDato = WParametros2(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector22.CellLeft + WVector22.Left
            WTexto12.Top = WVector22.CellTop + WVector22.Top
            WTexto12.Width = WVector22.CellWidth
            WTexto12.Height = WVector22.CellHeight
            WTexto12.MaxLength = WParametros2(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector22.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector22.CellLeft + WVector22.Left
            WTexto22.Top = WVector22.CellTop + WVector22.Top
            WTexto22.Width = WVector22.CellWidth
            WTexto22.Height = WVector22.CellHeight
            WTexto22.MaxLength = WParametros2(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector22.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector22.CellLeft + WVector22.Left
            WTexto32.Top = WVector22.CellTop + WVector22.Top
            WTexto32.Width = WVector22.CellWidth
            WTexto32.Height = WVector22.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector22.Text) = 10 Then
                        WTexto32.Text = WVector22.Text
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

Private Sub EndEdit2()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector22.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector22.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector22.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector22.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato2(WVector22.Col) <> "" Then
            WVector22.Text = Pusing(WFormato2(WVector22.Col), WVector22.Text)
        End If
    End If
End Sub

Private Sub GridEditCombo2()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector22.CellLeft + WVector22.Left
    WCombo12.Top = WVector22.CellTop + WVector22.Top
    WCombo12.Width = WVector22.CellWidth
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
            WTexto12.Text = WVector22.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector22.SetFocus
            DoEvents
            Call Control_Campo2
            If WControl2 = "S" Then
                Call Control_Grilla2
            End If
            Call StartEdit2

        Case vbKeyDown
            ' Move down 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row < WVector22.Rows - 1 Then
                Rem Call Control_Campo2
                WControl2 = "S"
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row > WVector22.FixedRows Then
                Rem Call Control_Campo2
                WControl2 = "S"
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row - 1
                End If
            End If
            Call StartEdit2

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto22.Text = WVector22.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector22.SetFocus
            DoEvents
            Call Control_Campo2
            If WControl2 = "S" Then
                Call Control_Grilla2
            End If
            Call StartEdit2
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row < WVector22.Rows - 1 Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row > WVector22.FixedRows Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row - 1
                End If
            End If
            Call StartEdit2

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
            WTexto32.Text = WVector22.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector22.SetFocus
            Call Control_Campo2
            If WControl2 = "S" Then
                Call Control_Grilla2
            End If
            Call StartEdit2

        Case vbKeyDown
            ' Move down 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row < WVector22.Rows - 1 Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row > WVector22.FixedRows Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row - 1
                End If
            End If
            Call StartEdit2

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
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector22.SetFocus
End Sub

Private Sub WVector22_Click()
    StartEdit2
End Sub

Private Sub WVector22_LeaveCell()
    EndEdit2
End Sub

Private Sub WVector22_GotFocus()
    EndEdit2
End Sub

Rem Desde aca empieza las rutinas a cambiar

Private Sub StartEdit2()
    Select Case WParametros2(4, WVector22.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector22.Text
            On Error GoTo 0
            GridEditCombo2
        Case Else
            If WParametros2(2, WVector22.Col) = 0 Then
                GridEditText2 Asc(" ")
            End If
    End Select
End Sub

Private Sub WVector22_KeyPress(KeyAscii As Integer)
    XColumna = WVector22.Col
    Select Case WParametros2(4, WVector22.Col)
        Case 1
        Case Else
            If WParametros2(2, XColumna) = 0 Then
                GridEditText2 KeyAscii
            End If
    End Select
End Sub

Private Sub Control_Grilla2()
    Select Case WVector22.Col
        Case 4
            If WVector22.Row < WVector22.Rows - 1 Then
                WVector22.Row = WVector22.Row + 1
            End If
            WVector22.Col = 1
        Case Else
            If WVector22.Col < WVector22.Cols - 1 Then
                WVector22.Col = WVector22.Col + 1
            End If
    End Select
    WVector22.SetFocus
    GridEditText2 KeyAscii
End Sub

Private Sub Control_Campo2()
    XColumna = WVector22.Col
    XFila = WVector22.Row
    WControl2 = "S"
    Select Case XColumna
        Case 1
            If WVector22.Text <> "" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector22.Text + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector22.Col = 2
                    WVector22.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                        Else
                    WControl2 = "N"
                End If
                
            End If
        Case 3, 4, 5
            WVector22.Col = XColumna
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector2()

    WVector22.Clear
    
    Rem ponga la grilla en negritas
    WVector22.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector22.FontName
    WTexto12.FontSize = WVector22.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector22.FontName
    WTexto22.FontSize = WVector22.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector22.FontName
    WTexto32.FontSize = WVector22.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector22.FontName
    WCombo12.FontSize = WVector22.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector22.FixedCols = 1
    WVector22.Cols = 7
    WVector22.FixedRows = 1
    WVector22.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector22.Text = "Articulo"
    
    Rem Longitud
    Rem WVector22.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros2(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros2(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros2(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros2(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector22.ColWidth(0) = 200
    WVector22.Row = 0
    For Ciclo = 1 To WVector22.Cols - 1
        WVector22.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector22.Text = "Articulo"
                WVector22.ColWidth(Ciclo) = 2000
                WVector22.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros2(1, Ciclo) = 20
                WParametros2(2, Ciclo) = 0
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 2
                WVector22.Text = "Descripcion"
                WVector22.ColWidth(Ciclo) = 2500
                WVector22.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros2(1, Ciclo) = 50
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 3
                WVector22.Text = "Cantidad"
                WVector22.ColWidth(Ciclo) = 1200
                WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 10
                WParametros2(2, Ciclo) = 0
                WParametros2(3, Ciclo) = 1
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
            Case 4
                WVector22.Text = "Costo"
                WVector22.ColWidth(Ciclo) = 1200
                WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 10
                WParametros2(2, Ciclo) = 0
                WParametros2(3, Ciclo) = 1
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = "###,###.##"
            Case 5
                WVector22.Text = "Importe"
                WVector22.ColWidth(Ciclo) = 1200
                WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 10
                WParametros2(2, Ciclo) = 0
                WParametros2(3, Ciclo) = 1
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = "###,###.##"
            Case 6
                WVector22.Text = ""
                WVector22.ColWidth(Ciclo) = 10
                WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 10
                WParametros2(2, Ciclo) = 0
                WParametros2(3, Ciclo) = 1
                WParametros2(4, Ciclo) = 0
                WFormato2(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector22.Row = 0
    For Ciclo = 1 To WVector22.Cols - 1
        WVector22.Col = Ciclo
        WTitulo2(Ciclo).Text = WVector22.Text
        WTitulo2(Ciclo).Left = WVector22.CellLeft + WVector22.Left
        WTitulo2(Ciclo).Top = WVector22.CellTop + WVector22.Top
        WTitulo2(Ciclo).Width = WVector22.CellWidth
        WTitulo2(Ciclo).Height = WVector22.CellHeight
        WTitulo2(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector22.Cols - 1
        WAncho = WAncho + WVector22.ColWidth(Ciclo)
    Next Ciclo
    WVector22.Width = WAncho

    ' Size the columns.
    Font.Name = WVector22.Font.Name
    Font.Size = WVector22.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector22.AllowUserResizing = flexResizeBoth
    
    WVector22.Col = 1
    WVector22.Row = 1
    
End Sub

Private Sub WVector22_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub

Private Sub WVector22_DblClick()

    If WVector22.Col = 0 Then
        Exit Sub
    End If

    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False

    For Ciclo = 1 To WVector22.Cols - 1
        WVector22.Col = Ciclo
        WVector22.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector22.Rows - 1
        WVector22.Row = Ciclo
        WVector22.Col = 1
        WAuxi1 = WVector22.Text
        WVector22.Col = 3
        WAuxi2 = WVector22.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector22.Cols - 1
                WVector22.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector22.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector2
    
    For Ciclo = 1 To EntraVector
        WVector22.Row = Ciclo
        For da = 1 To WVector22.Cols - 1
            WVector22.Col = da
            WVector22.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
End Sub

Private Sub WTexto12_DblClick()
    Select Case WVector22.Col
        Case 1
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Conceptos"
            Opcion.AddItem "Cuentas Contables"
            Opcion.AddItem "Cuentas Contables"
            Opcion.AddItem "Cuentas Contables"
            Opcion.AddItem "Cuentas Contables"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 5
    
            Call Opcion_Click
        Case Else
    End Select
End Sub







