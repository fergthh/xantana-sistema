VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDevol 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Notas de Credito por Devolucion"
   ClientHeight    =   7995
   ClientLeft      =   510
   ClientTop       =   375
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7995
   ScaleWidth      =   10890
   Visible         =   0   'False
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   8400
      TabIndex        =   55
      Text            =   "Text2"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   9720
      TabIndex        =   54
      Text            =   "Text1"
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Vendedor 
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
      Left            =   7440
      MaxLength       =   6
      TabIndex        =   51
      Text            =   " "
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Descuento1 
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
      Left            =   4920
      MaxLength       =   6
      TabIndex        =   44
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Descuento3 
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
      MaxLength       =   6
      TabIndex        =   43
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Descuento2 
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
      Left            =   6120
      MaxLength       =   6
      TabIndex        =   42
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox Partida 
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
      MaxLength       =   6
      TabIndex        =   41
      Text            =   " "
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Campana 
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
      MaxLength       =   10
      TabIndex        =   40
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox Lista 
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
      Left            =   2880
      MaxLength       =   6
      TabIndex        =   39
      Text            =   " "
      Top             =   840
      Width           =   495
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
      Left            =   9000
      MouseIcon       =   "Devol.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Devol.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Menu Principal"
      Top             =   3000
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
      Left            =   9000
      MouseIcon       =   "Devol.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "Devol.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Consulta de Datos"
      Top             =   1920
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
      Left            =   8040
      MouseIcon       =   "Devol.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "Devol.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Limpia la pantalla"
      Top             =   4080
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
      Left            =   8040
      MouseIcon       =   "Devol.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "Devol.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Elimina el Registro"
      Top             =   3000
      Width           =   855
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
      Left            =   8040
      MouseIcon       =   "Devol.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "Devol.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1920
      Width           =   855
   End
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
      Left            =   9120
      MouseIcon       =   "Devol.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "Devol.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Impresion"
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   4560
      TabIndex        =   21
      Top             =   5880
      Width           =   2895
      Begin VB.Label SubTotal 
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
         Left            =   1440
         TabIndex        =   50
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Neto"
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
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "SubTotal"
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
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Iva 21%"
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
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Iva 10.5%"
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
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Neto 
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
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Iva1 
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
         Left            =   1440
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Iva2 
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
         Left            =   1440
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
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
         Left            =   1440
         TabIndex        =   24
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Dto 
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
         Left            =   1440
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Descuento"
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
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
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
      Left            =   9480
      MaxLength       =   8
      TabIndex        =   19
      Text            =   " "
      Top             =   120
      Width           =   1215
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
      Left            =   8640
      MaxLength       =   4
      TabIndex        =   18
      Top             =   120
      Width           =   735
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
      Left            =   8160
      MaxLength       =   1
      TabIndex        =   17
      Top             =   120
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
      Left            =   3360
      TabIndex        =   15
      Top             =   3240
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   360
      Left            =   3480
      TabIndex        =   14
      Top             =   3600
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
      Left            =   4080
      TabIndex        =   13
      Top             =   3240
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3720
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
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3720
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
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3720
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
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   375
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
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   4335
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9000
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "FactuPro.rpt"
      CopiesToPrinter =   2
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   3615
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
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.ListBox WIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1740
      ItemData        =   "Devol.frx":43C8
      Left            =   120
      List            =   "Devol.frx":43CF
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   4335
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   16
      Top             =   3240
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
      Height          =   4335
      Left            =   120
      TabIndex        =   32
      Top             =   1320
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7646
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label18 
      Caption         =   "Vendedor"
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
      Left            =   6120
      TabIndex        =   53
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label DesVendedor 
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
      Left            =   8280
      TabIndex        =   52
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Descuento"
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
      Left            =   3720
      TabIndex        =   48
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Partida"
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
      TabIndex        =   47
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Campaña"
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
      Left            =   3720
      TabIndex        =   46
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label16 
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
      Height          =   285
      Left            =   2040
      TabIndex        =   45
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Nro de Comprobante"
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
      Left            =   6120
      TabIndex        =   20
      Top             =   120
      Width           =   1815
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
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   2775
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
      TabIndex        =   5
      Top             =   120
      Width           =   1575
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
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "PrgDevol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoDto1 As Double
Private WImpoDto2 As Double
Private WImpoDto3 As Double
Private WDescuento As Double
Private WDescuento1 As Double
Private WDescuento2 As Double
Private WDescuento3 As Double
Private WCodIva As String
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private WTipo As String
Private WProvincia As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private Mes(0 To 30) As String
Private XIndice As Single
Private WTipopro As Integer
Private XTalle As String
Private XColor As String
Private XArticulo As String
Private XTexto1 As String
Private XTexto2 As String
Dim CantiFac As Integer
Dim CantiRem As Integer
Dim CantiArti As Integer
Dim ZMes As String
Dim ZAno As String
Dim ZZCambia As String



Dim ZZClave As String
Dim ZZLetra As String
Dim ZZTipo As String
Dim ZZPunto As String
Dim ZZNumero As String
Dim ZZRenglon As String
Dim ZZCliente As String
Dim ZZfecha As String
Dim ZZEstado As String
Dim ZZVencimiento As String
Dim ZZTotal As String
Dim ZZSaldo As String
Dim ZZOrdFecha As String
Dim ZZOrdVencimiento As String
Dim ZZImpre As String
Dim ZZNeto As String
Dim ZZIva1 As String
Dim ZZIva2 As String
Dim ZZPedido As String
Dim ZZRemito As String
Dim ZZOrden As String
Dim ZZProvincia As String
Dim ZZVendedor As String
Dim ZZCosto As String
Dim ZZImporte1 As String
Dim ZZImporte2 As String
Dim ZZImporte3 As String
Dim ZZImporte4 As String
Dim ZZImporte5 As String
Dim ZZImporte6 As String
Dim ZZImporte7 As String
Dim ZZTipoventa As String
Dim ZZProyecto As String
Dim ZZParidad As String
Dim ZZTotalUs As String
Dim ZZSaldoUs As String
Dim ZZRemito1 As String
Dim ZZRemito2 As String
Dim ZZBusqueda As String
Dim ZZBase1 As String
Dim ZZBase2 As String

Dim ZZDescuento1 As String
Dim ZZDescuento2 As String
Dim ZZDescuento3 As String
Dim ZZPartida As String
Dim ZZLista As String
Dim ZZCampana As String

Dim ZZCantidad As String
Dim ZZCantidadII As String

Dim WWArticulo As String
Dim WWDescripcion As String
Dim WWCantidad As String
Dim WWPrecio As String
Dim WWImpre As Double
Dim WWDto(10) As Double
Dim WWIva(10) As Double


Dim WVector(100, 10) As String


Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Calcula_FechaVto()

    Rem With rstPago
    Rem    .Index = "Pago"
    Rem    .Seek "=", WPago1
    Rem    If .NoMatch = False Then
    Rem        WPlazo1 = !Plazo
    Rem        WTasa = !Tasa
    Rem        WDescuento = !Descuento
    Rem        WPago = !Nombre
    Rem    End If
    Rem End With
    
    Rem WFecha = Fecha.Text
    Rem Call Calcula_vencimiento(WFecha, WPlazo1, Wvencimiento)
    
    Rem With rstPago
    Rem     .Index = "Pago"
    Rem     .Seek "=", WPago2
    Rem     If .NoMatch = False Then
    Rem         WPlazo2 = !Plazo
    Rem     End If
    Rem End With
    
    Rem Call Calcula_vencimiento(WFecha, WPlazo2, WVencimiento1)

End Sub

Private Sub Consulta_Click()

    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulos"

    Opcion.Visible = True
     
 End Sub

Private Sub Impresion_Click()

    Rem Call Impresion
    
    WVector1.Col = 1
    WVector1.Row = 1
        
    Numero.SetFocus
    
End Sub


Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
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
                            IngresaItem = !Cliente + " " + !Razon
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
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
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
                rstArticulo.Close
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

Private Sub Calcula_Click()

    WNeto = 0

    For a = 1 To 50
        WCantidad = Val(WVector1.TextMatrix(a, 3))
        WPrecio = Val(WVector1.TextMatrix(a, 4))
        Text1.Text = Letra.Text
        If Trim(UCase(Letra.Text)) = "B" Then
            WWImpre = WPrecio * 1.21
            Call Redondeo(WWImpre)
            WPrecio = WWImpre
            If WPrecio <> 0 Then
                Text2.Text = Str$(WPrecio)
            End If
        End If
        WNeto = WNeto + (WPrecio * WCantidad)
    Next a
    
    Call Calcula_Importe

End Sub

Private Sub Calcula_Importe()

    WImpoDto1 = 0
    WImpoDto2 = 0
    WImpoDto3 = 0
    WDescuento1 = Val(Descuento1.Text)
    WDescuento2 = Val(Descuento2.Text)
    WDescuento3 = Val(Descuento3.Text)
    
    WDescuento = WDescuento1
    If WDescuento <> 0 Then
        WImpoDto1 = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto1)
        WNeto = WNeto - WImpoDto1
    End If
    
    WDescuento = WDescuento2
    If WDescuento <> 0 Then
        WImpoDto2 = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto2)
        WNeto = WNeto - WImpoDto2
    End If
    
    WDescuento = WDescuento3
    If WDescuento <> 0 Then
        WImpoDto3 = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto3)
        WNeto = WNeto - WImpoDto3
    End If
    
    WImpoDto = WImpoDto1 + WImpoDto2 + WImpoDto3
    
    WIva1 = 0
    WIva2 = 0
    
    If Val(WEmpresa) = 1 Then
        Select Case Val(WCodIva)
            Case 2
                WIva1 = WNeto * ((ConfigIva1) / 100)
                WIva2 = WNeto * ((ConfigIva2) / 100)
                Call Redondeo(WIva1)
                Call Redondeo(WIva2)
            Case 1
                WIva1 = WNeto * ((ConfigIva1) / 100)
                Call Redondeo(WIva1)
            Case Else
        End Select
    End If
    
    WWIva(1) = WIva1
    WWIva(2) = WIva2
    
    WTotal = WNeto + WIva1 + WIva2
    
    SubTotal.Caption = Str$(WNeto + WImpoDto)
    Dto.Caption = Str$(WImpoDto)
    Neto.Caption = Str$(WNeto)
    Iva1.Caption = Str$(WIva1)
    Iva2.Caption = Str$(WIva2)
    Total.Caption = Str$(WTotal)
    
    SubTotal.Caption = Pusing("###,###.##", SubTotal.Caption)
    Dto.Caption = Pusing("###,###.##", Dto.Caption)
    Neto.Caption = Pusing("###,###.##", Neto.Caption)
    Iva1.Caption = Pusing("###,###.##", Iva1.Caption)
    Iva2.Caption = Pusing("###,###.##", Iva2.Caption)
    Total.Caption = Pusing("###,###.##", Total.Caption)

End Sub

Private Sub cmdClose_Click()
    PrgDevol.Hide
    Unload Me
    Menu4.Show
End Sub

Private Sub Graba_Click()

    Call Calcula_Click
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
            
    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
    
    ZZTipo = "02"
    ZZImpre = "DV"
            
    ZZPunto = WPunto
    ZZLetra = Letra.Text
    ZZNumero = Auxi
    ZZRenglon = "01"
    ZZCliente = Cliente.Text
    ZZfecha = Fecha.Text
    ZZEstado = "0"
    ZZVencimiento = Fecha.Text
    ZZTotal = Str$(WTotal * -1)
    ZZSaldo = Str$(WTotal * -1)
    ZZTotalUs = Str$(WTotal * -1)
    ZZSaldoUs = Str$(WTotal * -1)
    If Trim(UCase(Letra.Text)) = "B" Then
        WNeto = WTotal / (1 + ((ConfigIva1) / 100))
        Call Redondeo(WNeto)
        WIva1 = WTotal - WNeto
        WIva2 = 0
        ZZNeto = Str$(WNeto * -1)
        ZZNetoTotal = Str$(WNeto * -1)
        ZZIva1 = Str$(WIva1 * -1)
        ZZIva2 = Str$(WIva2 * -1)
            Else
        ZZNeto = Str$(WNeto * -1)
        ZZNetoTotal = Str$(WNeto * -1)
        ZZIva1 = Str$(WIva1 * -1)
        ZZIva2 = Str$(WIva2 * -1)
    End If
    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZZOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZZPedido = ""
    ZZRemito = ""
    ZZOrden = ""
    ZZProvincia = WProvincia
    ZZVendedor = Vendedor.Text
    ZZCosto = "0"
    ZZImporte1 = "0"
    ZZImporte2 = "0"
    ZZImporte3 = "0"
    ZZImporte4 = "0"
    ZZImporte5 = "0"
    ZZImporte6 = "0"
    ZZImporte7 = "0"
    ZZTipoventa = "0"
    ZZProyecto = ""
    ZZParidad = "0"
    ZZRemito1 = ""
    ZZRemito2 = ""
    ZZBusqueda = ZZLetra + WPunto + Auxi
    ZZDescuento = ""
    
    ZZDescuento1 = Descuento1.Text
    ZZDescuento2 = Descuento2.Text
    ZZDescuento3 = Descuento3.Text
    ZZPartida = Partida.Text
    ZZPago = ""
    ZZLista = Lista.Text
    ZZOCompra = ""
    ZZCampana = Campana.Text
    ZZDespacho1 = ""
    ZZDespacho2 = ""
    ZZBase1 = DateDiff("d", "01/01/2000", Fecha.Text)
    ZZBase2 = "0"
    ZZLinea = ""
    
    If Val(WEmpresa) = 1 Then
        If Partida.Text = "V" Then
            ZZNetoTotal = Str$(WNeto * 2)
        End If
        If Partida.Text = "W" Then
            ZZNetoTotal = Str$(WNeto * 3)
        End If
            
        If Partida.Text = "M" Then
            ZZNetoTotal = Str$(WNeto * 2)
        End If
        If Partida.Text = "Z" Then
            ZZNetoTotal = Str$(WNeto * 3)
        End If
    End If
    
    ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
    
    ZSql = ""
    ZSql = ZSql + "INSERT INTO CtaCte ("
    ZSql = ZSql + "Clave ,"
    ZSql = ZSql + "Letra ,"
    ZSql = ZSql + "Tipo ,"
    ZSql = ZSql + "Punto ,"
    ZSql = ZSql + "Numero ,"
    ZSql = ZSql + "Renglon ,"
    ZSql = ZSql + "Cliente ,"
    ZSql = ZSql + "fecha ,"
    ZSql = ZSql + "Estado ,"
    ZSql = ZSql + "Vencimiento ,"
    ZSql = ZSql + "Total ,"
    ZSql = ZSql + "Saldo ,"
    ZSql = ZSql + "OrdFecha  ,"
    ZSql = ZSql + "OrdVencimiento ,"
    ZSql = ZSql + "Impre ,"
    ZSql = ZSql + "Neto ,"
    ZSql = ZSql + "NetoTotal ,"
    ZSql = ZSql + "Iva1 ,"
    ZSql = ZSql + "Iva2 ,"
    ZSql = ZSql + "Pedido ,"
    ZSql = ZSql + "Remito ,"
    ZSql = ZSql + "Orden ,"
    ZSql = ZSql + "Provincia ,"
    ZSql = ZSql + "Vendedor ,"
    ZSql = ZSql + "Costo ,"
    ZSql = ZSql + "Importe1 ,"
    ZSql = ZSql + "Importe2 ,"
    ZSql = ZSql + "Importe3 ,"
    ZSql = ZSql + "Importe4 ,"
    ZSql = ZSql + "Importe5 ,"
    ZSql = ZSql + "Importe6 ,"
    ZSql = ZSql + "Importe7 ,"
    ZSql = ZSql + "Tipoventa ,"
    ZSql = ZSql + "Proyecto ,"
    ZSql = ZSql + "Paridad ,"
    ZSql = ZSql + "TotalUs ,"
    ZSql = ZSql + "SaldoUs ,"
    ZSql = ZSql + "Remito1 ,"
    ZSql = ZSql + "Remito2 ,"
    ZSql = ZSql + "Descuento1 ,"
    ZSql = ZSql + "Descuento2 ,"
    ZSql = ZSql + "Descuento3 ,"
    ZSql = ZSql + "Partida ,"
    ZSql = ZSql + "Pago ,"
    ZSql = ZSql + "Lista ,"
    ZSql = ZSql + "Linea ,"
    ZSql = ZSql + "OCompra ,"
    ZSql = ZSql + "Campana ,"
    ZSql = ZSql + "Despacho1 ,"
    ZSql = ZSql + "Despacho2 ,"
    ZSql = ZSql + "Base1 ,"
    ZSql = ZSql + "Base2 ,"
    ZSql = ZSql + "Busqueda )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + ZZClave + "',"
    ZSql = ZSql + "'" + ZZLetra + "',"
    ZSql = ZSql + "'" + ZZTipo + "',"
    ZSql = ZSql + "'" + ZZPunto + "',"
    ZSql = ZSql + "'" + ZZNumero + "',"
    ZSql = ZSql + "'" + ZZRenglon + "',"
    ZSql = ZSql + "'" + ZZCliente + "',"
    ZSql = ZSql + "'" + ZZfecha + "',"
    ZSql = ZSql + "'" + ZZEstado + "',"
    ZSql = ZSql + "'" + ZZVencimiento + "',"
    ZSql = ZSql + "'" + ZZTotal + "',"
    ZSql = ZSql + "'" + ZZSaldo + "',"
    ZSql = ZSql + "'" + ZZOrdFecha + "',"
    ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
    ZSql = ZSql + "'" + ZZImpre + "',"
    ZSql = ZSql + "'" + ZZNeto + "',"
    ZSql = ZSql + "'" + ZZNetoTotal + "',"
    ZSql = ZSql + "'" + ZZIva1 + "',"
    ZSql = ZSql + "'" + ZZIva2 + "',"
    ZSql = ZSql + "'" + ZZPedido + "',"
    ZSql = ZSql + "'" + ZZRemito + "',"
    ZSql = ZSql + "'" + ZZOrden + "',"
    ZSql = ZSql + "'" + ZZProvincia + "',"
    ZSql = ZSql + "'" + ZZVendedor + "',"
    ZSql = ZSql + "'" + ZZCosto + "',"
    ZSql = ZSql + "'" + ZZImporte1 + "',"
    ZSql = ZSql + "'" + ZZImporte2 + "',"
    ZSql = ZSql + "'" + ZZImporte3 + "',"
    ZSql = ZSql + "'" + ZZImporte4 + "',"
    ZSql = ZSql + "'" + ZZImporte5 + "',"
    ZSql = ZSql + "'" + ZZImporte6 + "',"
    ZSql = ZSql + "'" + ZZImporte7 + "',"
    ZSql = ZSql + "'" + ZZTipoventa + "',"
    ZSql = ZSql + "'" + ZZProyecto + "',"
    ZSql = ZSql + "'" + ZZParidad + "',"
    ZSql = ZSql + "'" + ZZTotalUs + "',"
    ZSql = ZSql + "'" + ZZSaldoUs + "',"
    ZSql = ZSql + "'" + ZZRemito1 + "',"
    ZSql = ZSql + "'" + ZZRemito2 + "',"
    ZSql = ZSql + "'" + ZZDescuento1 + "',"
    ZSql = ZSql + "'" + ZZDescuento2 + "',"
    ZSql = ZSql + "'" + ZZDescuento3 + "',"
    ZSql = ZSql + "'" + ZZPartida + "',"
    ZSql = ZSql + "'" + ZZPago + "',"
    ZSql = ZSql + "'" + ZZLista + "',"
    ZSql = ZSql + "'" + ZZLinea + "',"
    ZSql = ZSql + "'" + ZZOCompra + "',"
    ZSql = ZSql + "'" + ZZCampana + "',"
    ZSql = ZSql + "'" + ZZDespacho1 + "',"
    ZSql = ZSql + "'" + ZZDespacho2 + "',"
    ZSql = ZSql + "'" + ZZBase1 + "',"
    ZSql = ZSql + "'" + ZZBase2 + "',"
    ZSql = ZSql + "'" + ZZBusqueda + "')"
                            
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    Renglon = 0
    WRenglon = 0
        
    For a = 1 To 50
        
        WRenglon = WRenglon + 1
            
        WVector1.Row = WRenglon
            
        WVector1.Col = 1
        Articulo = UCase(WVector1.Text)
        
        WVector1.Col = 2
        DesArticulo = WVector1.Text
                    
        WVector1.Col = 3
        Cantidad = Val(WVector1.Text) * -1
                    
        WVector1.Col = 4
        Precio = Val(WVector1.Text)
            
        If Cantidad <> 0 Then
                    
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Numero.Text)
            Call Ceros(Auxi1, 8)
                    
            ZZTipo = "02"
            ZZNumero = Numero.Text
            ZZRenglon = Renglon
            ZZArticulo = Articulo
            ZZCantidad = Str$(Cantidad)
            ZZDescripcion = DesArticulo
            ZZPrecio = Str$(Precio)
            ZZPrecioUs = Str$(Precio)
            ZZImporte = Str$(Precio * Cantidad)
            ZZImporteUs = Str$(Precio * Cantidad)
            ZZCliente = Cliente.Text
            ZZParidad = "0"
            ZZVendedor = "0"
            ZZRubro = "0"
            ZZLinea = "0"
            ZZCosto1 = "0"
            ZZCosto2 = "0"
            ZZCoeficiente = "0"
            ZZPedido = "0"
            ZZfecha = Fecha.Text
            ZZImporte1 = "0"
            ZZImporte2 = "0"
            ZZImporte3 = "0"
            ZZImporte4 = "0"
            ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZZWArticulo = ""
            ZZRemito = ""
            ZZClave = "02" + Auxi1 + Auxi
            ZZWDate = Date$
            ZZClaveCtacte = Left$(ZZClave, 10) + "01"
            ZZImprefactura = "NOTA DE CREDITO"
            ZZNroFactura = Auxi1
            ZZTalle = Talle
            ZZColor = XXColor
            ZZCuenta = WCuenta
            ZZDescuento1 = Descuento1.Text
            ZZDescuento2 = Descuento2.Text
            ZZDescuento3 = Descuento3.Text
            ZZPartida = Partida.Text
            ZZCampana = Campana.Text
            ZZCantidadII = ZZCantidad
            ZZPrecioII = ZZPrecio
            
            If Val(WEmpresa) = 1 Then
                If Partida.Text = "V" Then
                    ZZCantidad = Str$(Val(ZZCantidad) * 2)
                End If
                If Partida.Text = "W" Then
                    ZZCantidad = Str$(Val(ZZCantidad) * 3)
                End If
            End If
            
            If Val(WEmpresa) = 1 Then
                If Partida.Text = "M" Then
                    ZZPrecio = Str$(Val(ZZPrecio) * 2)
                End If
                If Partida.Text = "Z" Then
                    ZZPrecio = Str$(Val(ZZPrecio) * 3)
                End If
            End If
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Estadistica ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "CantidadII ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "PrecioII ,"
            ZSql = ZSql + "PrecioUs ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "ImporteUs ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Paridad ,"
            ZSql = ZSql + "Vendedor ,"
            ZSql = ZSql + "Rubro ,"
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Costo1 ,"
            ZSql = ZSql + "Costo2 ,"
            ZSql = ZSql + "Coeficiente ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Importe1 ,"
            ZSql = ZSql + "Importe2 ,"
            ZSql = ZSql + "Importe3 ,"
            ZSql = ZSql + "Importe4 ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "WArticulo ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "WDate ,"
            ZSql = ZSql + "Marca ,"
            ZSql = ZSql + "ClaveCtacte ,"
            ZSql = ZSql + "Imprefactura ,"
            ZSql = ZSql + "NroFactura ,"
            ZSql = ZSql + "Descuento1 ,"
            ZSql = ZSql + "Descuento2 ,"
            ZSql = ZSql + "Descuento3 ,"
            ZSql = ZSql + "Campana ,"
            ZSql = ZSql + "Talle ,"
            ZSql = ZSql + "Color ,"
            ZSql = ZSql + "Partida )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZPunto + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZCantidadII + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZPrecioII + "',"
            ZSql = ZSql + "'" + ZZPrecioUs + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZImporteUs + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZParidad + "',"
            ZSql = ZSql + "'" + ZZVendedor + "',"
            ZSql = ZSql + "'" + ZZRubro + "',"
            ZSql = ZSql + "'" + ZZLinea + "',"
            ZSql = ZSql + "'" + ZZCosto1 + "',"
            ZSql = ZSql + "'" + ZZCosto2 + "',"
            ZSql = ZSql + "'" + ZZCoeficiente + "',"
            ZSql = ZSql + "'" + ZZPedido + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZImporte1 + "',"
            ZSql = ZSql + "'" + ZZImporte2 + "',"
            ZSql = ZSql + "'" + ZZImporte3 + "',"
            ZSql = ZSql + "'" + ZZImporte4 + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + ZZWArticulo + "',"
            ZSql = ZSql + "'" + ZZRemito + "',"
            ZSql = ZSql + "'" + ZZWDate + "',"
            ZSql = ZSql + "'" + ZZMarca + "',"
            ZSql = ZSql + "'" + ZZClaveCtacte + "',"
            ZSql = ZSql + "'" + ZZImprefactura + "',"
            ZSql = ZSql + "'" + ZZNroFactura + "',"
            ZSql = ZSql + "'" + ZZDescuento1 + "',"
            ZSql = ZSql + "'" + ZZDescuento2 + "',"
            ZSql = ZSql + "'" + ZZDescuento3 + "',"
            ZSql = ZSql + "'" + ZZCampana + "',"
            ZSql = ZSql + "'" + ZZTalle + "',"
            ZSql = ZSql + "'" + ZZColor + "',"
            ZSql = ZSql + "'" + ZZPartida + "')"
                            
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Salidas = Salidas - " + "'" + Str$(Abs(Val(ZZCantidad))) + "',"
            ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Abs(Val(ZZCantidad))) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                            
        End If
                                        
    Next a
        
    If Val(WEmpresa) = 1 Then
        Call Impresion_Factura
            Else
        Call Impresion_Remito
    End If
    
        
    Call Limpia_Click
        
    Cliente.SetFocus
        
End Sub

Private Sub CmdDelete_Click()

    T$ = "Baja de Comprobantes"
    m$ = "Desea Borrar el Comprobante "
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
            
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
            
        WTipo = "02"
            
        ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
           
        Rem ZSql = ""
        Rem ZSql = ZSql + "Select *"
        Rem ZSql = ZSql + " FROM Ctacte"
        Rem ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
        Rem spCtaCte = ZSql
        Rem Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstCtaCte.RecordCount > 0 Then
        Rem     ZSaldo = rstCtaCte!Saldo
        Rem     ZTotal = rstCtaCte!Total
        Rem     rstCtaCte.Close
       Rem      If ZSaldo <> ZTotal Then
        Rem         m$ = "El comprobante se encuentra total o parcialmente cancelado"
        Rem         a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
        Rem         Exit Sub
        Rem     End If
        Rem End If
    
    
        Erase WVector
    
        For WRenglon = 1 To 50
        
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            WClave = "02" + Auxi + Auxi1
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Estadistica"
            ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
            
                Articulo = rstEstadistica!Articulo
                Cantidad = Abs(rstEstadistica!Cantidad)
                CantidadII = Abs(rstEstadistica!CantidadII)
                    
                rstEstadistica.Close
                    
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Salidas = Salidas + " + "'" + Str$(Cantidad) + "',"
                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(Cantidad) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
            End If
            
        Next WRenglon
        
        ZSql = ""
        ZSql = ZSql + "DELETE Estadistica"
        ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "02" + "'"
        ZSql = ZSql + " and Estadistica.Numero = " + "'" + Numero.Text + "'"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        
        ZSql = ""
        ZSql = ZSql + "DELETE CtaCte"
        ZSql = ZSql + " Where Letra = " + "'" + Letra.Text + "'"
        ZSql = ZSql + " and Tipo = " + "'" + "02" + "'"
        ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        
        Call Limpia_Click
        
        Cliente.SetFocus
        
    End If

End Sub

Private Sub Limpia_Click()

    Call Limpia_Vector

    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Descuento1.Text = ""
    Descuento2.Text = ""
    Descuento3.Text = ""
    Partida.Text = ""
    Lista.Text = ""
    Campana.Text = ""
    
    Renglon = 0
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""

    Graba.Enabled = True
    Cliente.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case 1
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Indice = Pantalla.ListIndex
            ClaveVen$ = WIndice.List(Indice)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ClaveVen$ + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstArticulo!Codigo
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                WVector1.Col = 4
                Select Case Val(Lista.Text)
                    Case 1
                        WVector1.Text = Str$(rstArticulo!Precio1)
                    Case 2
                        WVector1.Text = Str$(rstArticulo!Precio2)
                    Case 3
                        WVector1.Text = Str$(rstArticulo!Precio3)
                    Case 4
                        WVector1.Text = Str$(rstArticulo!Precio4)
                    Case 5
                        WVector1.Text = Str$(rstArticulo!Precio5)
                    Case 6
                        WVector1.Text = Str$(rstArticulo!Precio6)
                    Case Else
                        WVector1.Text = ""
                End Select
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "C.Final"
    Iva(4) = "Exento"
    Iva(5) = "MOnotributo"
    Iva(6) = "Exterior"
    
    Mes(1) = "Enero"
    Mes(2) = "Febrero"
    Mes(3) = "Marzo"
    Mes(4) = "Abril"
    Mes(5) = "Mayo"
    Mes(6) = "Junio"
    Mes(7) = "Julio"
    Mes(8) = "Agosto"
    Mes(9) = "Septiembre"
    Mes(10) = "Octubre"
    Mes(11) = "Noviembre"
    Mes(12) = "Diciembre"
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"
    
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Descuento1.Text = ""
    Descuento2.Text = ""
    Descuento3.Text = ""
    Partida.Text = ""
    Lista.Text = ""
    Campana.Text = ""
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
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
        CantiFac = rstConfiguracion!CantiFac
        CantiRem = rstConfiguracion!CantiRem
        CantiArti = rstConfiguracion!CantiArti
        rstConfiguracion.Close
    End If
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 50
    
        If Val(Punto.Text) = 1 Then
    
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "02" + Auxi + Auxi1
            
                Else
            
            ZZZLetra = Letra.Text
            
            ZZZPunto = Punto.Text
            Call Ceros(ZZZPunto, 1)
            
            ZZZNumeroFac = Numero.Text
            Call Ceros(ZZZNumeroFac, 6)
                
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "02" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi1
                
        End If
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Estadistica"
        ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
        
            Canti = rstEstadistica!Cantidad
            
            Renglon = Renglon + 1
                    
            WVector1.Row = Renglon
                    
            WVector1.Col = 1
            WVector1.Text = rstEstadistica!Articulo
            Auxi1 = rstEstadistica!Articulo
                
            WVector1.Col = 2
            WVector1.Text = IIf(IsNull(rstEstadistica!Descripcion), "", rstEstadistica!Descripcion)
                
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###", Str$(rstEstadistica!CantidadII * -1))
                
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!PrecioII))
            
            rstEstadistica.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                rstArticulo.Close
            End If
                
        End If
    
    Next WRenglon

    Call Calcula_Click
    
    Graba.Enabled = True

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
            DesCliente.Caption = rstCliente!Razon
            Partida.Text = rstCliente!Partida
            Descuento1.Text = Str$(rstCliente!Descuento1)
            Descuento2.Text = Str$(rstCliente!Descuento2)
            Descuento3.Text = Str$(rstCliente!Descuento3)
            Descuento1.Text = Pusing("###,###.##", Descuento1.Text)
            Descuento2.Text = Pusing("###,###.##", Descuento2.Text)
            Descuento3.Text = Pusing("###,###.##", Descuento3.Text)
            Vendedor.Text = rstCliente!Vendedor
            WProvincia = rstCliente!Provincia
            WCodIva = rstCliente!Iva
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            If Val(WEmpresa) = 1 Then
                Select Case Val(WCodIva)
                    Case 1, 2
                        Letra.Text = "A"
                    Case Else
                        Letra.Text = "B"
                End Select
                    Else
                Letra.Text = "A"
            End If
            
            If Letra.Text = "B" Then
                m$ = "COLOQUE EL FORMULARIO B"
                a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            End If
            
            rstCliente.Close
            
            Rem ZSql = ""
            Rem ZSql = ZSql + "Select *"
            Rem ZSql = ZSql + " FROM ClienteAdicional"
            Rem ZSql = ZSql + " Where ClienteAdicional.Cliente = " + "'" + Cliente.Text + "'"
            Rem ZSql = ZSql + " and ClienteAdicional.Linea = " + "'" + Linea.Text + "'"
            Rem spClienteAdicional = ZSql
            Rem Set rstClienteAdicional = db.OpenRecordset(spClienteAdicional, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstClienteAdicional.RecordCount > 0 Then
            Rem     Descuento1.Text = Str$(rstClienteAdicional!Descuento1)
            Rem     Descuento2.Text = Str$(rstClienteAdicional!Descuento2)
            Rem     Descuento3.Text = Str$(rstClienteAdicional!Descuento3)
            Rem     Descuento1.Text = Pusing("###,###.##", Descuento1.Text)
            Rem     Descuento2.Text = Pusing("###,###.##", Descuento2.Text)
            Rem     Descuento3.Text = Pusing("###,###.##", Descuento3.Text)
            Rem     Vendedor.Text = rstClienteAdicional!Vendedor
            Rem     rstClienteAdicional.Close
            Rem End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Codigo = " + "'" + Vendedor.Text + "'"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                DesVendedor.Caption = rstVendedor!Nombre
                rstVendedor.Close
            End If
                
            WPunto = Str(ConfigPunto)
            Call Ceros(WPunto, 4)
            Punto.Text = WPunto
                
            Numero.Text = "1"
            WTipo = "02"
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Ctacte.Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
            ZSql = ZSql + " Order by Ctacte.Numero"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                With rstCtaCte
                    .MoveLast
                    Do
                        If .BOF = False Then
                    
                            If Letra.Text = rstCtaCte!Letra And Punto.Text = rstCtaCte!Punto Then
                                If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                    Numero.Text = Str$(Val(rstCtaCte!Numero) + 1)
                                    Exit Do
                                End If
                            End If
                                
                            .MovePrevious
                            
                            If .BOF = True Then
                                Exit Do
                            End If
                                
                                Else
                            
                            Exit Do
                    
                        End If
                    Loop
                End With
                rstCtaCte.Close
            End If
                
            Lista.SetFocus
                Else
            Cliente.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Numero.Text
        Call Ceros(WPunto, 4)
        
        Numero.Text = "1"
        WTipo = "02"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Letra = " + "'" + Letra.Text + "'"
        ZSql = ZSql + " and Ctacte.Punto = " + "'" + Punto.Text + "'"
        ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
        ZSql = ZSql + " Order by Ctacte.Numero"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            With rstCtaCte
                .MoveLast
                Do
                    If .BOF = False Then
                    
                        If Letra.Text = rstCtaCte!Letra And Punto.Text = rstCtaCte!Punto Then
                            If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                Numero.Text = Str$(Val(rstCtaCte!Numero) + 1)
                                Exit Do
                            End If
                        End If
                                
                        .MovePrevious
                            
                        If .BOF = True Then
                            Exit Do
                        End If
                                
                            Else
                            
                        Exit Do
                    
                    End If
                Loop
            End With
            rstCtaCte.Close
        End If
        
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
            
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
            
        WTipo = "02"
        ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            
            Fecha.Text = rstCtaCte!Fecha
            Cliente.Text = rstCtaCte!Cliente
            Descuento1.Text = Str$(rstCtaCte!Descuento1)
            Descuento1.Text = Pusing("###,###.##", Descuento1.Text)
            Descuento2.Text = Str$(rstCtaCte!Descuento2)
            Descuento2.Text = Pusing("###,###.##", Descuento2.Text)
            Descuento3.Text = Str$(rstCtaCte!Descuento3)
            Descuento3.Text = Pusing("###,###.##", Descuento3.Text)
            Partida.Text = rstCtaCte!Partida
            Campana.Text = rstCtaCte!Campana
            Lista.Text = rstCtaCte!Lista
                
            rstCtaCte.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                WProvincia = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                rstCliente.Close
            End If
                
            Call Proceso_Click
                
                Else
                    
            Graba.Enabled = True
            WNumero = Numero.Text
            Numero.Text = WNumero
            Fecha.SetFocus
                
        End If
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
            Lista.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Descuento1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento1.Text = Pusing("###,###.##", Descuento1.Text)
        Descuento2.SetFocus
    End If
    If KeyAscii = 27 Then
        Descuento1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento2.Text = Pusing("###,###.##", Descuento2.Text)
        Descuento3.SetFocus
    End If
    If KeyAscii = 27 Then
        Descuento2.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento3.Text = Pusing("###,###.##", Descuento3.Text)
        Partida.SetFocus
    End If
    If KeyAscii = 27 Then
        Descuento3.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Partida_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Partida.Text
            Case "H", "V", "W", "M", "Z"
                Lista.SetFocus
            Case Else
                Partida.SetFocus
        End Select
    End If
    If KeyAscii = 27 Then
        Partida.Text = ""
    End If
End Sub

Private Sub Lista_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Val(Lista.Text)
            Case 1, 2, 3, 4, 5, 6
                Campana.SetFocus
            Case Else
                Lista.SetFocus
        End Select
    End If
    If KeyAscii = 27 Then
        Lista.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Campana_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Campana.Text = ""
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
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Cliente + " " + !Razon
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
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
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
                rstArticulo.Close
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
            
        Rem F1,F2,F3,F4,f9,F10
        Case 112, 113, 114, 115, 120, 121
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
            
        Rem F1,F2,F3,F4,f9,F10
        Case 112, 113, 114, 115, 120, 121
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
            
        Rem F1,F2,F3,F4,f9,F10
        Case 112, 113, 114, 115, 120, 121
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
            If WVector1.Text <> "999999" Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    WVector1.Col = 4
                    Select Case Val(Lista.Text)
                        Case 1
                            WVector1.Text = Str$(rstArticulo!Precio1)
                        Case 2
                            WVector1.Text = Str$(rstArticulo!Precio2)
                        Case 3
                            WVector1.Text = Str$(rstArticulo!Precio3)
                        Case 4
                            WVector1.Text = Str$(rstArticulo!Precio4)
                        Case 5
                            WVector1.Text = Str$(rstArticulo!Precio5)
                        Case 6
                            WVector1.Text = Str$(rstArticulo!Precio6)
                        Case Else
                            WVector1.Text = ""
                    End Select
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    WVector1.Col = 2
                    rstArticulo.Close
                        Else
                    WControl = "N"
                End If
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
    Call Calcula_Click
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
        WAuxi3 = WVector1.Text
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

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Call Opcion_Click
    
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
    WVector1.Cols = 5
    WVector1.FixedRows = 1
    WVector1.Rows = 51
    
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
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3100
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 1300
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

Private Sub Cliente_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulo"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = ""
    sCentimos = "centavos"
    
    Numero = CStr(Val(Total.Caption))
    
    XTexto1 = Numero2Letra(Numero, , sMoneda & " ", sCentimos & " ")
    XTexto1 = XTexto1 + Space$(100)
    
    Pasa = 0
    
    For da = 60 To 1 Step -1
        If Mid$(XTexto1, da, 1) = Space$(1) Then
            Pasa = 1
        End If
        If Pasa = 1 Then
            If Mid$(XTexto1, da, 1) <> Space$(1) Then
                Exit For
            End If
        End If
    Next da
    
    XTexto2 = Mid$(XTexto1, da + 2, 100)
    XTexto1 = Left$(XTexto1, da)
    
End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descuento1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descuento2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descuento3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Partida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Campana_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call Graba_Click
        Case 113
            Call CmdDelete_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub




















Sub Impresion_Factura()

    Open "Lpt1" For Output As #99 Len = 255
    Rem Open "Dada.txt" For Output As #99 Len = 255
    
    WLetra = Letra.Text

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = Trim(rstCliente!Razon)
        WDireccion = Trim(rstCliente!Direccion)
        WLocalidad = Trim(rstCliente!Localidad)
        WPostal = Trim(rstCliente!Postal)
        WTelefono = Trim(rstCliente!Telefono)
        WObservaciones = Trim(rstCliente!Observaciones)
        WCuit = Trim(rstCliente!Cuit)
        WEmail = Trim(rstCliente!EMail)
        WFax = Trim(rstCliente!fax)
        WProvincia = Val(rstCliente!Provincia)
        WIva = Val(rstCliente!Iva)
        WExpreso = Str$(rstCliente!Expreso)
        WPartida = rstCliente!Partida
        WVendedor = Str$(rstCliente!Vendedor)
        WDescuento1 = Str$(rstCliente!Descuento1)
        WDescuento2 = Str$(rstCliente!Descuento2)
        WDescuento3 = Str$(rstCliente!Descuento3)
        rstCliente.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Str$(WExpreso) + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        WFleteroNombre = Trim(rstExpreso!Nombre)
        WFleteroDireccion = Trim(rstExpreso!Direccion)
        WFleteroCuit = Trim(rstExpreso!Cuit)
        rstExpreso.Close
    End If
    

    If WLetra = "A" Then
        
        Print #99, Tab(50); "NOTA DE CREDITO"
        Print #99, ""
        Print #99, ""
        Print #99, Tab(4); "";
        Print #99, Tab(44); Fecha.Text;
        Print #99, Tab(74); "";
        Print #99, Tab(114); Fecha.Text
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        
        Print #99, ""

        Print #99, ""
        Print #99, ""

        Print #99, Tab(34); WDireccion;
        Print #99, Tab(104); WDireccion
        Print #99, Tab(3); WRazon;
        Print #99, Tab(34); WLocalidad;
        Print #99, Tab(72); WRazon;
        Print #99, Tab(104); WLocalidad
        Print #99, Tab(34); Provincia(Val(WProvincia)); " "; WPostal;
        Print #99, Tab(104); Provincia(Val(WProvincia)); " "; WPostal

        Print #99, ""
        Print #99, ""

            Else

        Print #99, Tab(50); "NOTA DE CREDITO"
        Print #99, ""
        Print #99, ""
        Print #99, Tab(4); "";
        Print #99, Tab(44); Fecha.Text
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""

        Print #99, ""
        Print #99, ""

        Print #99, Tab(34); WDireccion
        Print #99, Tab(3); WRazon;
        Print #99, Tab(34); WLocalidad
        Print #99, Tab(34); Provincia(Val(WProvincia)); " "; WPostal

        Print #99, ""
        Print #99, ""
        
    End If

    If WLetra = "A" Then

        Print #99, Tab(11); Iva(WIva);
        Print #99, Tab(41); WCuit;
        Print #99, Tab(61); Cliente.Text;
        Print #99, Tab(76); Iva(WIva);
        Print #99, Tab(108); WCuit;
        Print #99, Tab(128); Cliente.Text;

            Else

        Print #99, Tab(11); Iva(WIva);
        Print #99, Tab(41); WCuit;
        Print #99, Tab(61); Cliente.Text;

    End If

    If WLetra = "A" Then
    
        Print #99, ""
        Print #99, ""

        Print #99, Tab(59); Numero.Text

        Print #99, ""
        Print #99, ""
        Print #99, ""
        
            Else
            
        Print #99, ""
        Print #99, ""

        Print #99, Tab(59); Numero.Text

        Print #99, ""
        Print #99, ""
        Print #99, ""
        
    End If


    ZLineas = 0

    For Ciclo = 1 To 19
    
        WWArticulo = WVector1.TextMatrix(Ciclo, 1)
        WWDescripcion = WVector1.TextMatrix(Ciclo, 2)
        WWCantidad = WVector1.TextMatrix(Ciclo, 3)
        WWPrecio = WVector1.TextMatrix(Ciclo, 4)
    
        If Trim(WWArticulo) <> "" Then

            WWImpre1 = Right$("     " + Pusing("#####", WWCantidad), 5)
            
            If WWArticulo <> "999999" Then
                Print #99, Tab(3); Left$(WWArticulo, 10);
            End If

            Print #99, Tab(13); Left$(WWDescripcion, 23);
            Print #99, Tab(39); WWImpre1;
            
            If Val(WIva) <> 1 Then
                WWImpre = Val(WWPrecio) * 1.21
                Call Redondeo(WWImpre)
                WWImpre2 = Right$("        " + Pusing("##,###.##", Str$(WWImpre)), 8)
                WWImpre3 = Right$("          " + Pusing("###,###.##", Str$(WWImpre * Val(WWCantidad))), 10)
                Print #99, Tab(45); WWImpre2;
                Print #99, Tab(56); WWImpre3;
                    Else
                WWImpre = Val(WWPrecio)
                Call Redondeo(WWImpre)
                WWImpre2 = Right$("        " + Pusing("##,###.##", Str$(WWImpre)), 8)
                WWImpre3 = Right$("          " + Pusing("###,###.##", Str$(WWImpre * Val(WWCantidad))), 10)
                Print #99, Tab(45); WWImpre2;
                Print #99, Tab(56); WWImpre3;
            End If
            
            If Letra.Text = "A" Then
                WWImpre2 = Right$("        " + Pusing("##,###.##", Str$(WWImpre)), 8)
                If WWArticulo <> "999999" Then
                    Print #99, Tab(70); WWArticulo;
                End If
                Print #99, Tab(81); Left$(WWDescripcion, 23);
                Print #99, Tab(107); WWImpre1;
                Print #99, Tab(112); WWImpre2
            End If
            
            ZLineas = ZLineas + 1
            
        End If
    Next Ciclo

    For ZImprelinea = ZLineas To 19
        Print #99, ""
    Next ZImprelinea

    If Letra.Text = "A" Then

        Print #99, Tab(13); ""

        If Val(Descuento1.Text) <> 0 Then
            WWImpre1 = Right$("     " + Pusing("#####.##", Descuento1.Text), 5)
            WWImpre2 = Right$("          " + Pusing("###,###.##", Str$(WImpoDto1)), 10)
            Print #99, Tab(39); "Bonif. "; WWImpre1;
            Print #99, Tab(56); WWImpre2
                Else
            Print #99, ""
        End If

        If Val(Descuento2.Text) <> 0 Then
            WWImpre1 = Right$("     " + Pusing("#####.##", Descuento2.Text), 5)
            WWImpre2 = Right$("          " + Pusing("###,###.##", Str$(WImpoDto2)), 10)
            Print #99, Tab(39); "Bonif. "; WWImpre1;
            Print #99, Tab(56); WWImpre2
                Else
            Print #99, ""
        End If

        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        
        Print #99, Tab(72); Left$(WFleteroNombre, 20);
        Print #99, Tab(93); Left$(WFleteroDireccion, 17);
        Print #99, Tab(111); Left$(WFleteroCuit, 15)

        Print #99, ""

            Else

        Print #99, ""

        If Val(Descuento1.Text) <> 0 Then
            Print #99, Tab(39); "Bonif. "; Pusing("##.##", Descuento1.Text);
            Print #99, Tab(56); Pusing("###,###.##", Str$(WImpoDto1))
                Else
            Print #99, ""
        End If

        If Val(Descuento2.Text) <> 0 Then
            Print #99, Tab(39); "Bonif. "; Pusing("##.##", Descuento2.Text);
            Print #99, Tab(56); Pusing("###,###.##", Str$(WImpoDto2))
                Else
            Print #99, ""
        End If

        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""

        Print #99, Tab(8); Left$(WFleteroNombre, 20);
        Print #99, Tab(30); Left$(WFleteroDireccion, 17);
        Print #99, Tab(50); Left$(WFleteroCuit, 15)
        
        Print #99, ""

    End If
    
    WWImpre1 = Right$("          " + Pusing("###,###.##", Neto.Caption), 10)
    WWImpre2 = Right$("          " + Pusing("###,###.##", Str$(WWIva(1))), 10)
    WWImpre3 = Right$("          " + Pusing("###,###.##", Str$(WWIva(2))), 10)
    WWImpre4 = Right$("          " + Pusing("###,###.##", Total.Caption), 10)

    If Letra.Text = "A" Then
            
        Print #99, Tab(56); WWImpre1
        Print #99, ""
        Print #99, Tab(56); WWImpre1

        If WWIva(1) <> 0 Then
            Print #99, Tab(49); "21 %";
            Print #99, Tab(56); WWImpre2
                Else
            Print #99, ""
        End If


        If WWIva(2) <> 0 Then
            Print #99, Tab(49); "10.5%";
            Print #99, Tab(56); WWImpre3
                Else
            Print #99, ""
        End If

        Print #99, ""

        Print #99, Tab(5); Partida.Text;
        Print #99, Tab(10); Right$(Vendedor.Text, 2);
        Print #99, Tab(56); WWImpre4;
        Print #99, Tab(72); Partida.Text;
        Print #99, Tab(77); Right$(Vendedor.Text, 2)

        Print #99, ""
        Print #99, ""
        Print #99, ""

        If Val(WDespacho$) <> 0 Then
            Print #99, Tab(5); "Despacho : "; Numero0$; " / "; Numero1$;
            Print #99, Tab(72); "Despacho : "; Numero0$; " / "; Numero1$
                Else
            Print #99, ""
        End If

            Else

        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""

        Print #99, Tab(5); Partida.Text;
        Print #99, Tab(10); Right$(Vendedor.Text, 2);
        Print #99, Tab(56); WWImpre4
        
        Print #99, ""
        Print #99, ""
        Print #99, ""

        If Val(WDespacho$) <> 0 Then
            Print #99, Tab(5); "Despacho : "; Numero0$; " / "; Numero1$
                Else
            Print #99, ""
        End If

    End If

    Print #99, ""
    Print #99, ""
    Print #99, ""
    Print #99, ""
    Print #99, ""
    Print #99, ""
    Print #99, ""
    Print #99, ""
    
    Close #99

End Sub

Sub Impresion_Remito()

    Open "Lpt1" For Output As #99 Len = 255
    Rem Open "Dada.txt" For Output As #99 Len = 255

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = Trim(rstCliente!Razon)
        WDireccion = Trim(rstCliente!Direccion)
        WLocalidad = Trim(rstCliente!Localidad)
        WPostal = Trim(rstCliente!Postal)
        WTelefono = Trim(rstCliente!Telefono)
        WObservaciones = Trim(rstCliente!Observaciones)
        WCuit = Trim(rstCliente!Cuit)
        WEmail = Trim(rstCliente!EMail)
        WFax = Trim(rstCliente!fax)
        WProvincia = Val(rstCliente!Provincia)
        WIva = Val(rstCliente!Iva)
        WExpreso = Str$(rstCliente!Expreso)
        WPartida = rstCliente!Partida
        WVendedor = Str$(rstCliente!Vendedor)
        WDescuento1 = Str$(rstCliente!Descuento1)
        WDescuento2 = Str$(rstCliente!Descuento2)
        WDescuento3 = Str$(rstCliente!Descuento3)
        rstCliente.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Str$(WExpreso) + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        WFleteroNombre = Trim(rstExpreso!Nombre)
        WFleteroDireccion = Trim(rstExpreso!Direccion)
        WFleteroCuit = Trim(rstExpreso!Cuit)
        rstExpreso.Close
    End If
    
    ZLineas = 0
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)

    Print #99, Chr$(18)

    Print #99, "                                      +---+"
    Print #99, "+-------------------------------------| A |-----------------------------------+"
    Print #99, "|                                     +---+                                   |"
    Print #99, "| Y E N A D I   S. A.                   | MERCADERIA EN CONSIGNACION          |"
    Print #99, "| Av. R.S.Ortiz 436 (Ex.Canning)        |Dev. N§            :0004-0000" + Right$(Auxi, 4) + "    |"
    Print #99, "| Capital Federal   C.P.:1414           |Fecha              :" + Fecha.Text + "       |"
    Print #99, "| TE : 4514-9777 (ROT.)                 |Cuit N§            :30-62995041-5    |"
    Print #99, "| IVA : RESPONSABLE INSCRIPTO           |Ing.Brutos N§      :901-943310-7     |"
    Print #99, "|                                       |Caja Previsional N§:1.514.757        |"
    Print #99, "|                                       |                                     |"
    Print #99, "|-----------------------------------------------------------------------------|"
    Print #99, "|                                                                             |"
    Print #99, "|Señor    :" + WRazon;
    Print #99, Tab(65); "(" + Cliente.Text + ")";
    Print #99, Tab(79); "|"
    Print #99, "|Domicilio:" + WDireccion;
    Print #99, Tab(79); "|"
    Print #99, "|Provincia:" + Provincia(Val(WProvincia));
    Print #99, Tab(65); "C.P.:" + WPostal;
    Print #99, Tab(79); "|"
    Print #99, "|Localidad:" + WLocalidad;
    Print #99, Tab(79); "|"
    Print #99, "|C.U.I.T. :" + WCuit;
    Print #99, Tab(79); "|"
    Print #99, "|Iva.     :" + Iva(WIva);
    Print #99, Tab(79); "|"
    If Val(WDespacho$) <> 0 Then
        Print #99, Using; "| Despacho : " + Despacho1.Text + "  " + Despacho2.Text;
        Print #99, Tab(79); "|"
            Else
        Print #99, "|                                                                             |"
    End If
    Print #99, "|-----------------------------------------------------------------------------|"
    Print #99, "| PLAZO DE LIQUIDACION DE LA OPERACION : 30 DIAS                              |"
    Print #99, "|-----------------------------------------------------------------------------|"
    Print #99, "| Codigo   |           Descripcion                | Cantidad  |  P.Unit. U$S  |"
    Print #99, "|-----------------------------------------------------------------------------|"
    Print #99, "|          |                                      |           |               |"

    For Ciclo = 1 To 19
    
        WWArticulo = WVector1.TextMatrix(Ciclo, 1)
        WWDescripcion = WVector1.TextMatrix(Ciclo, 2)
        WWCantidad = WVector1.TextMatrix(Ciclo, 3)
        WWPrecio = WVector1.TextMatrix(Ciclo, 4)

        If Trim(WWArticulo) <> "" Then
        
            WWImpre1 = Right$("       " + Pusing("###,###", WWCantidad), 7)
            WWImpre2 = Right$("           " + Pusing("###,###.##", WWPrecio), 10)

            Print #99, "|" + Left$(WWArticulo, 10);
            Print #99, Tab(12); "|" + Left$(WWDescripcion, 35);
            Print #99, Tab(51); "|" + WWImpre1;
            Print #99, Tab(63); "|" + WWImpre2;
            Print #99, Tab(79); "|"

            Lineas% = Lineas% + 1
        End If
    Next Ciclo
    
    For Imprelinea = Lineas% To 20
            Print #99, "|          |                                      |           |               |"
    Next Imprelinea

    Print #99, "|-----------------------------------------------------------------------------|"
    Print #99, "|   Vendedor : " + Vendedor.Text;
    Print #99, Tab(79); "|"
    Print #99, "|   ";
    Print #99, Tab(79); "|"
    Print #99, "|   ";
    Print #99, Tab(79); "|"
    Print #99, "|   ";
    Print #99, Tab(79); "|"
    Print #99, "|   Expreso :" + Left$(WFleteroNombre, 20);
    Print #99, Tab(42); Left$(WFleteroDireccion, 18);
    Print #99, Tab(65); WFleteroCuit;
    Print #99, Tab(79); "|"
    
    Print #99, "|-----------------------------------------------------------------------------|"
    Print #99, "| reintegrada, en caso de no ser vendida, a simple requerimiento de la        |"
    Print #99, "| comitente                                                                   |"
    Print #99, "|-----------------------------------------------------------------------------|"

    Print #99, ""
    Print #99, ""
    Print #99, ""

    WWImpre1 = Right$("           " + Pusing("###,###.##", SubTotal.Caption), 10)
    WWImpre2 = Right$("           " + Pusing("###,###.##", Dto.Caption), 10)
    WWImpre3 = Right$("           " + Pusing("###,###.##", Total.Caption), 10)


    Print #99, "|-----------------------------------------------------------------------------|"
    Print #99, "|                                                Sub-Total    |   " + WWImpre1;
    Print #99, Tab(79); "|"
    Print #99, "|                                                Bonificacion |   " + WWImpre2;
    Print #99, Tab(79); "|"
    Print #99, "|                                                Total        |   " + WWImpre3;
    Print #99, Tab(79); "|"
    Print #99, "|-----------------------------------------------------------------------------|"


    Print #99, ""
    Print #99, ""
    Print #99, ""
    Print #99, ""
    Print #99, ""
        
    Close #99

End Sub



