VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCustodia 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Cheques en Custodia"
   ClientHeight    =   7995
   ClientLeft      =   30
   ClientTop       =   615
   ClientWidth     =   11880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   Begin VB.Frame PantaVaria 
      Height          =   1935
      Left            =   480
      TabIndex        =   36
      Top             =   1200
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox Cuenta 
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
         TabIndex        =   43
         Text            =   " "
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Titular 
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
         TabIndex        =   37
         Text            =   " "
         Top             =   360
         Width           =   4695
      End
      Begin MSMask.MaskEdBox FechaEmision 
         Height          =   285
         Left            =   2040
         TabIndex        =   39
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
      Begin VB.Label Label7 
         Caption         =   "Cuenta"
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
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Fec.Emision"
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
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Titular"
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
         TabIndex        =   38
         Top             =   360
         Width           =   1455
      End
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
      Index           =   9
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   4440
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
      Index           =   8
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   5280
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
      Index           =   7
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   5040
      Width           =   375
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
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   32
      Text            =   " "
      Top             =   480
      Width           =   735
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
      Index           =   6
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton Cheque 
      Caption         =   "Cheques F6"
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
      Left            =   3960
      MouseIcon       =   "Custodia.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Custodia.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cartera de Cheques"
      Top             =   6840
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
      Left            =   4920
      MouseIcon       =   "Custodia.frx":0799
      MousePointer    =   99  'Custom
      Picture         =   "Custodia.frx":0AA3
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Impresion de Orden de Pago"
      Top             =   6840
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
      Left            =   120
      MouseIcon       =   "Custodia.frx":12E5
      MousePointer    =   99  'Custom
      Picture         =   "Custodia.frx":15EF
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   6840
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
      Left            =   1080
      MouseIcon       =   "Custodia.frx":1E31
      MousePointer    =   99  'Custom
      Picture         =   "Custodia.frx":213B
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Elimina el Registro"
      Top             =   6840
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
      Left            =   2040
      MouseIcon       =   "Custodia.frx":297D
      MousePointer    =   99  'Custom
      Picture         =   "Custodia.frx":2C87
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Limpia la pantalla"
      Top             =   6840
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
      Left            =   3000
      MouseIcon       =   "Custodia.frx":34C9
      MousePointer    =   99  'Custom
      Picture         =   "Custodia.frx":37D3
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Consulta de Datos"
      Top             =   6840
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
      Left            =   5880
      MouseIcon       =   "Custodia.frx":4015
      MousePointer    =   99  'Custom
      Picture         =   "Custodia.frx":431F
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Menu Principal"
      Top             =   6840
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9600
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      PrintFileName   =   "impreCUstodia.rpt"
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
      Left            =   6480
      TabIndex        =   22
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3960
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3960
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3360
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3360
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3360
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
      Left            =   2040
      TabIndex        =   14
      Top             =   2760
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   2040
      TabIndex        =   13
      Top             =   3360
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
      Left            =   2640
      TabIndex        =   12
      Top             =   2760
      Width           =   375
   End
   Begin VB.ListBox WVector 
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   975
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Importe 
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
      MaxLength       =   15
      TabIndex        =   2
      Text            =   " "
      Top             =   1200
      Width           =   1335
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
      Height          =   1260
      Left            =   8160
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   120
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5040
      TabIndex        =   5
      Top             =   120
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
      Height          =   6060
      ItemData        =   "Custodia.frx":4B61
      Left            =   6480
      List            =   "Custodia.frx":4B68
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   5295
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3240
      TabIndex        =   18
      Top             =   2760
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
      Height          =   4575
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   8070
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin MSMask.MaskEdBox Acredita 
      Height          =   285
      Left            =   1800
      TabIndex        =   31
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
   Begin VB.Label lblLabels 
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
      Index           =   1
      Left            =   120
      TabIndex        =   35
      Top             =   480
      Width           =   1455
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2760
      TabIndex        =   34
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Fec.Acreditacion"
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
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Importe"
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
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Creditos 
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
      Left            =   4200
      TabIndex        =   9
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. 1) Ef.    2) Cheques  "
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
      Left            =   840
      TabIndex        =   8
      Top             =   6360
      Width           =   3375
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1575
   End
End
Attribute VB_Name = "PrgCustodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Private dada As String
Private Vector(10, 6) As String
Private Numero As String
Private Imprelin As Single
Dim BajaCheque(100) As String
Dim DatosCheque(100, 10) As String
Dim ZMes As String
Dim ZAno As String
Dim ZZLugar As Integer


Dim ZZCodigo As String
Dim ZZRenglon As String
Dim ZZBanco As String
Dim ZZImporte As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZAcredita As String
Dim ZZAcreditaOrd As String
Dim ZZTipo2 As String
Dim ZZNumero2 As String
Dim ZZFecha2 As String
Dim ZZObservaciones2 As String
Dim ZZImporte2 As String
Dim ZZEmpresa As String
Dim ZZClave As String
Dim ZZClaveCheque As String


Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Suma_Datos()
    Creditos.Caption = ""
    For IRow = 1 To 100
        Auxi = WVector1.TextMatrix(IRow, 5)
        Call Conver(Auxi, dada)
        If Val(Auxi) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Auxi))
        End If
    Next IRow
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
End Sub

Private Sub Lee_Datos()

    Call Limpia_Vector

    Renglon = 0
    Debito = 0
    Credito = 0
    
    Do
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        WClave = Codigo.Text + Auxi1
             
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Custodia"
        ZSql = ZSql + " Where Custodia.Clave = " + "'" + WClave + "'"
        spCustodia = ZSql
        Set rstCustodia = db.OpenRecordset(spCustodia, dbOpenSnapshot, dbSQLPassThrough)
        If rstCustodia.RecordCount > 0 Then
             
            Credito = Credito + 1
            WVector1.Row = Credito
            WVector1.Col = 1
            WVector1.Text = rstCustodia!Tipo2
            WVector1.Col = 2
            WVector1.Text = rstCustodia!Numero2
            WVector1.Col = 3
            WVector1.Text = rstCustodia!Fecha2
            WVector1.Col = 4
            If rstCustodia!Observaciones2 <> "" Then
                WVector1.Text = rstCustodia!Observaciones2
            End If
            WVector1.Col = 5
            WVector1.Text = rstCustodia!Importe2
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            
            rstCustodia.Close
            
                Else
                
            Exit Do
            
        End If
        
    Loop
End Sub

Sub Verifica_datos()
    If Importe.Text = 0 Then
        Importe.Text = "0"
    End If
End Sub

Sub Format_datos()
    Importe.Text = Pusing("###,###.##", Importe.Text)
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Banco"
    ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
    spBanco = ZSql
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        Banco.Text = rstBanco!Banco
        DesBanco.Caption = rstBanco!Nombre
        Call Format_datos
        rstBanco.Close
    End If
End Sub

Private Sub cmdAdd_Click()

    Rem If WLicencia <> "1234-5678-ABCD-EFGH" And Val(Codigo.text) > 10 Then
    Rem     WMsg$ = "La version del sistema es para un uso limitado de movimientos." + Chr$(13) + _
    REM          "El objetivo es el de verificar las opciones y el funcionamiento del mismo." + Chr$(13) + _
    REM          "Para poder utilizar el sistema sin limite de movimientos se debe adquirir la version definitiva."
    Rem     A% = MsgBox(WMsg$, 0, "Sistema de Control de Gestion")
    Rem     Exit Sub
    Rem End If
    
    Existe = ""

    If Codigo.Text <> "" And Fecha.Text <> "" And Banco.Text <> "" Then
    
        If Existe <> "S" Then
    
            Call Suma_Datos
        
            Debito = 0
            Credito = 0
        
            If Val(Importe.Text) <> 0 Then
                Debito = Val(Importe.Text)
            End If
        
            If Val(Creditos.Caption) <> 0 Then
                Credito = Val(Creditos.Caption)
            End If
        
            If Debito = Credito Then
    
                Renglon = 0
                
                For IRow = 1 To 100
                
                    WRow = IRow
                    WVector1.Col = 5
                    WVector1.Row = IRow
                    Auxi = WVector1.Text
                    Call Conver(Auxi, dada)
                    If Val(Auxi) <> 0 Then
                    
                        Renglon = Renglon + 1
                        
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        Auxi2 = Str$(Val(Codigo.Text))
                        Call Ceros(Auxi2, 6)
                        
                        ZZCodigo = Auxi2
                        ZZRenglon = Auxi1
                        ZZBanco = Banco.Text
                        ZZImporte = Importe.Text
                        ZZfecha = Fecha.Text
                        ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        ZZAcredita = Acredita.Text
                        ZZAcreditaOrd = Right$(Acredita.Text, 4) + Mid$(Acredita.Text, 4, 2) + Left$(Acredita.Text, 2)
                        WVector1.Col = 1
                        ZZTipo2 = WVector1.Text
                        WVector1.Col = 2
                        ZZNumero2 = WVector1.Text
                        WVector1.Col = 3
                        ZZFecha2 = WVector1.Text
                        WVector1.Col = 4
                        ZZObservaciones2 = WVector1.Text
                        WVector1.Col = 5
                        ZZImporte2 = Auxi
                        ZZEmpresa = WEmpresa
                        ZZClave = ZZCodigo + ZZRenglon
                        ZZClaveCheque = BajaCheque(IRow)
                        WVector1.Col = 6
                        ZZClaveLectora = WVector1.Text
                        WVector1.Col = 7
                        ZZTitular = WVector1.Text
                        WVector1.Col = 8
                        ZZFechaEmision = WVector1.Text
                        ZZSucursal = DatosCheque(IRow, 1)
                        WVector1.Col = 9
                        ZZCuenta = WVector1.Text
                        
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Custodia ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "Codigo ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Banco ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Importe ,"
                        ZSql = ZSql + "Acredita ,"
                        ZSql = ZSql + "AcreditaOrd ,"
                        ZSql = ZSql + "Tipo2 ,"
                        ZSql = ZSql + "Numero2 ,"
                        ZSql = ZSql + "Fecha2 ,"
                        ZSql = ZSql + "Importe2  ,"
                        ZSql = ZSql + "Observaciones2 ,"
                        ZSql = ZSql + "Empresa ,"
                        ZSql = ZSql + "Impolista ,"
                        ZSql = ZSql + "Titular ,"
                        ZSql = ZSql + "FechaEmision ,"
                        ZSql = ZSql + "Sucursal ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "ClaveLectora ,"
                        ZSql = ZSql + "ClaveCheque )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + ZZClave + "',"
                        ZSql = ZSql + "'" + ZZCodigo + "',"
                        ZSql = ZSql + "'" + ZZRenglon + "',"
                        ZSql = ZSql + "'" + ZZBanco + "',"
                        ZSql = ZSql + "'" + ZZfecha + "',"
                        ZSql = ZSql + "'" + ZZFechaOrd + "',"
                        ZSql = ZSql + "'" + ZZImporte + "',"
                        ZSql = ZSql + "'" + ZZAcredita + "',"
                        ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
                        ZSql = ZSql + "'" + ZZTipo2 + "',"
                        ZSql = ZSql + "'" + ZZNumero2 + "',"
                        ZSql = ZSql + "'" + ZZFecha2 + "',"
                        ZSql = ZSql + "'" + ZZImporte2 + "',"
                        ZSql = ZSql + "'" + ZZObservaciones2 + "',"
                        ZSql = ZSql + "'" + ZZEmpresa + "',"
                        ZSql = ZSql + "'" + ZZImpolista + "',"
                        ZSql = ZSql + "'" + ZZTitular + "',"
                        ZSql = ZSql + "'" + ZZFechaEmision + "',"
                        ZSql = ZSql + "'" + ZZSucursal + "',"
                        ZSql = ZSql + "'" + ZZCuenta + "',"
                        ZSql = ZSql + "'" + ZZClaveLectora + "',"
                        ZSql = ZSql + "'" + ZZClaveCheque + "')"
                            
                        spCustodia = ZSql
                        Set rstCustodia = db.OpenRecordset(spCustodia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        Select Case Val(ZZTipo2)
                            Case 2
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Recibos SET "
                                ZSql = ZSql + " Estado2 = " + "'" + "X" + "',"
                                ZSql = ZSql + " Orden = " + "'" + "0" + "',"
                                ZSql = ZSql + " Deposito = " + "'" + Codigo.Text + "',"
                                ZSql = ZSql + " Destino = " + "'" + DesBanco.Caption + "',"
                                ZSql = ZSql + " BancoSalida = " + "'" + Banco.Text + "',"
                                ZSql = ZSql + " ProveedorSalida = " + "'" + "0" + "'"
                                ZSql = ZSql + " Where Clave = " + "'" + BajaCheque(IRow) + "'"
                                spRecibos = ZSql
                                Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                                
                            Case Else
                        End Select
                    
                    End If
                
                Next IRow
        
                Call ImpreCustodia

                Call CmdLimpiar_Click
                Codigo.SetFocus
        
            End If
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
    
        T$ = "Custodia"
        m$ = "Desea Borrar el Custodia "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            For IRow = 1 To 100
            
                Auxi1 = Str$(IRow)
                Call Ceros(Auxi1, 2)
                WClave = Codigo.Text + Auxi1
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Custodia"
                ZSql = ZSql + " Where Custodia.Clave = " + "'" + WClave + "'"
                spCustodia = ZSql
                Set rstCustodia = db.OpenRecordset(spCustodia, dbOpenSnapshot, dbSQLPassThrough)
                If rstCustodia.RecordCount > 0 Then
                
                    ZZClaveCheque = rstCustodia!ClaveCheque
                    rstCustodia.Close
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Recibos SET "
                    ZSql = ZSql + " Estado2 = " + "'" + "P" + "',"
                    ZSql = ZSql + " Orden = " + "'" + "0" + "',"
                    ZSql = ZSql + " Deposito = " + "'" + "0" + "',"
                    ZSql = ZSql + " Destino = " + "'" + "" + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + ZZClaveCheque + "'"
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)

                End If
                
            Next IRow
            
        End If
        
        ZSql = ""
        ZSql = ZSql + "DELETE Custodia"
        ZSql = ZSql + " Where Custodia.Codigo = " + "'" + Codigo.Text + "'"
        spCustodia = ZSql
        Set rstCustodia = db.OpenRecordset(spCustodia, dbOpenSnapshot, dbSQLPassThrough)
        
        Call CmdLimpiar_Click
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector
    Pantalla.Visible = False

    Codigo.Text = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Importe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Acredita.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Creditos.Caption = ""
    Codigo.SetFocus

    Codigo.Text = "1"
    ZSql = ""
    ZSql = ZSql + "Select *"
    Rem ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Custodia"
    ZSql = ZSql + " Order by fechaOrd, Codigo"
    spCustodia = ZSql
    Set rstCustodia = db.OpenRecordset(spCustodia, dbOpenSnapshot, dbSQLPassThrough)
    If rstCustodia.RecordCount > 0 Then
        rstCustodia.MoveLast
        Rem ZUltimo = IIf(IsNull(rstCustodia!CodigoMayor), "0", rstCustodia!CodigoMayor)
        ZUltimo = IIf(IsNull(rstCustodia!Codigo), "0", rstCustodia!Codigo)
        Codigo.Text = ZUltimo + 1
        rstCustodia.Close
    End If

End Sub

Private Sub CmdClose_Click()
    PrgCustodia.Hide
    Unload Me
    MenuA3.Show
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Existe = "N"
            
        Auxi1 = Codigo.Text
        Call Ceros(Auxi1, 6)
        Codigo.Text = Auxi1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Custodia"
        ZSql = ZSql + " Where Custodia.Codigo = " + "'" + Codigo.Text + "'"
        spCustodia = ZSql
        Set rstCustodia = db.OpenRecordset(spCustodia, dbOpenSnapshot, dbSQLPassThrough)
        If rstCustodia.RecordCount > 0 Then
        
            Existe = "S"
            If rstCustodia!Banco <> "" Then
                Banco.Text = rstCustodia!Banco
            End If
            If rstCustodia!Importe <> "" Then
                Importe.Text = rstCustodia!Importe
            End If
            Fecha.Text = rstCustodia!Fecha
            Acredita.Text = rstCustodia!Acredita
            
            rstCustodia.Close
            
        End If
        
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Fecha.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Banco.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Banco.Text) <> 0 Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                DesBanco.Caption = rstBanco!Nombre
                Acredita.SetFocus
                rstBanco.Close
                    Else
                Banco.SetFocus
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Banco.Text = ""
        DesBanco.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Acredita_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Acredita.Text, Auxi)
        If Auxi = "S" Then
            Importe.SetFocus
                Else
            Acredita.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Acredita.Text = "  /  /    "
    End If
End Sub

Private Sub Importe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Importe.Text = Pusing("###,###.##", Importe.Text)
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Importe.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     Ayuda.Visible = False
     Opcion.Clear
     Opcion.AddItem "Bancos"
     Opcion.AddItem "Cheques en Cartera"
     Opcion.Visible = True
     
End Sub

Private Sub Impresion_Click()
    Call ImpreCustodia
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    WVector.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
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
                            IngresaItem = Str$(rstBanco!Banco) + " " + rstBanco!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstBanco!Banco
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstBanco.Close
            End If
            Ayuda.Visible = True
            Ayuda.Text = ""
            Ayuda.SetFocus
                
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.TipoReg = '2'"
            ZSql = ZSql + " and Recibos.Tipo2 = '02'"
            ZSql = ZSql + " and Recibos.Estado2 <> 'X'"
            ZSql = ZSql + " Order by Recibos.FechaOrd2"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                With rstRecibos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Auxi$ = Str$(rstRecibos!Importe2)
                            Auxi$ = Mascara("###,###.##", Auxi$)
                            Numero = Str$(Val(rstRecibos!Numero2))
                            Call Ceros(Numero, 6)
                            IngresaItem = Numero + "    " + rstRecibos!Fecha2 + "      " + Auxi$ + "      " + rstRecibos!Banco2
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstRecibos!Clave
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRecibos.Close
            End If
     
        Case Else
    End Select
    
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()

    Select Case XIndice
        Case 0
            Ayuda.Visible = False
            Opcion.Visible = False
            Pantalla.Visible = False
            Indice = Pantalla.ListIndex
            Banco.Text = WIndice.List(Indice)
            Call Banco_KeyPress(13)
                
        Case Else
    End Select
    
End Sub

Private Sub Pantalla_DblClick()

    Select Case XIndice
        Case 1
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Auxi = WVector.List(Indice)
            If Auxi <> "X" Then
            
            For IRow = 1 To 100
                WVector1.Col = 5
                WVector1.Row = IRow
                Auxi = WVector1.Text
                Call Conver(Auxi, dada)
                If Val(Auxi) = 0 Then
                    Exit For
                End If
            Next IRow
            
            Indice = Pantalla.ListIndex
            WClave = WIndice.List(Indice)
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.Clave = " + "'" + WClave + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                
                WVector1.Col = 1
                WVector1.Text = 2
                    
                WVector1.Col = 2
                WVector1.Text = rstRecibos!Numero2
                
                WVector1.Col = 3
                WVector1.Text = rstRecibos!Fecha2
                
                WVector1.Col = 4
                WVector1.Text = rstRecibos!Banco2
                
                WVector1.Col = 5
                WVector1.Text = rstRecibos!Importe2
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    
                BajaCheque(WVector1.Row) = WIndice.List(Indice)
                DatosCheque(WVector1.Row, 1) = rstRecibos!SucursalCheque
                
                ZZLugar = WVector1.Row
                
                rstRecibos.Close
                    
                Call Suma_Datos
                    
                WVector1.Row = XRow
                WVector1.Col = 1
                Call StartEdit
                    
                Pantalla.List(Indice) = ""
                WIndice.List(Indice) = ""
                    
                
                   
                PantaVaria.Visible = True
                Titular.Text = ""
                FechaEmision.Text = "  /  /    "
                Titular.SetFocus
                    
            End If
            
            End If
                
        Case Else
    End Select
    
End Sub


Private Sub Form_Load()

    Call Limpia_Vector
    Pantalla.Visible = False
    
    Codigo.Text = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Importe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Acredita.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Creditos.Caption = ""
    
    Codigo.Text = "1"
    ZSql = ""
    ZSql = ZSql + "Select *"
    Rem ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Custodia"
    ZSql = ZSql + " Order by fechaOrd, Codigo"
    spCustodia = ZSql
    Set rstCustodia = db.OpenRecordset(spCustodia, dbOpenSnapshot, dbSQLPassThrough)
    If rstCustodia.RecordCount > 0 Then
        rstCustodia.MoveLast
        Rem ZUltimo = IIf(IsNull(rstCustodia!CodigoMayor), "0", rstCustodia!CodigoMayor)
        ZUltimo = IIf(IsNull(rstCustodia!Codigo), "0", rstCustodia!Codigo)
        Codigo.Text = ZUltimo + 1
        rstCustodia.Close
    End If
    

    Rem ZZDeposito = 1
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select *"
    Rem Rem ZSql = ZSql + "Select Max(Deposito) as [DepositoMayor]"
    Rem ZSql = ZSql + " FROM Depositos"
    Rem ZSql = ZSql + " Order by fechaOrd, Deposito"
    Rem spDepositos = ZSql
    Rem Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstDepositos.RecordCount > 0 Then
    Rem     rstDepositos.MoveLast
    Rem     Rem ZUltimo = IIf(IsNull(rstDepositos!DepositoMayor), "0", rstDepositos!DepositoMayor)
    Rem     ZUltimo = IIf(IsNull(rstDepositos!Deposito), "0", rstDepositos!Deposito)
    Rem     ZZDeposito = ZUltimo + 1
    Rem     rstDepositos.Close
    Rem End If
    Rem
    Rem If ZZDeposito > ZZCustodia Then
    Rem     Codigo.Text = ZZDeposito
    Rem         Else
    Rem     Codigo.Text = ZZCustodia
    Rem End If
     
End Sub

Private Sub ImpreCustodia()

    T$ = "Impresion de Comprobante del Custodia"
    m$ = "Desea realizar la impresion del comprobante"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        Auxi2 = Str$(Val(Codigo.Text))
        Call Ceros(Auxi2, 6)
    
        Listado.ReportFileName = "ImpreCustodia.rpt"
        
        Listado.WindowTitle = "Comprobante de Custodia"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height

        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        
        Listado.SQLQuery = "SELECT Custodia.Codigo, Custodia.Fecha, Custodia.Numero2, Custodia.Fecha2, Custodia.Importe2, Custodia.Observaciones2, Custodia.Sucursal, Custodia.Titular, Custodia.fechaemision, Custodia.Cuenta " _
                + "From " _
                + DSQ + ".dbo.Custodia Custodia " _
                + "Where " _
                + "Custodia.Codigo >= " + Codigo.Text + " AND " _
                + "Custodia.Codigo <= " + Codigo.Text
    
        Listado.Connect = Connect()
    
        Uno = "{Custodia.Codigo} in " + Codigo.Text + " to " + Codigo.Text
    
        Listado.GroupSelectionFormula = Uno
        Listado.SelectionFormula = Uno
    
        Listado.Destination = 1
        Rem Listado.Destination = 0
        Listado.Action = 1

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
                            IngresaItem = Str$(rstBanco!Banco) + " " + rstBanco!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstBanco!Banco
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
        Call Suma_Datos
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
            
        Rem F1,F2,F3,F4,f6,f9,F10
        Case 112, 113, 114, 115, 117, 120, 121
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
            
        Rem F1,F2,F3,F4,f6,f9,F10
        Case 112, 113, 114, 115, 117, 120, 121
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
            
        Rem F1,F2,F3,F4,f6,f9,F10
        Case 112, 113, 114, 115, 117, 120, 121
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
        Case 3
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
            If WVector1.Text <> "" Then
                dada = Len(Trim(WVector1.Text))
                If Len(Trim(WVector1.Text)) = 29 Then
                
                    For IRow = 1 To 50
                        If WVector1.TextMatrix(IRow, 1) = "" Then
                            XRow = WVector1.Row
                            Exit For
                        End If
                    Next IRow
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Recibos"
                    ZSql = ZSql + " Where Recibos.ClaveLectora = " + "'" + WVector1.Text + "'"
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibos.RecordCount > 0 Then
                        
                        WVector1.Row = XRow
                        
                        WVector1.Col = 1
                        WVector1.Text = "2"
                        
                        WVector1.Col = 2
                        WVector1.Text = rstRecibos!Numero2
                        
                        WVector1.Col = 3
                        WVector1.Text = rstRecibos!Fecha2
                        
                        WVector1.Col = 4
                        WVector1.Text = rstRecibos!Banco2
                        
                        WVector1.Col = 5
                        WVector1.Text = rstRecibos!Importe2
                        WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        
                        WVector1.Col = 6
                        WVector1.Text = rstRecibos!ClaveLectora
                            
                        BajaCheque(WVector1.Row) = rstRecibos!Clave
                        DatosCheque(WVector1.Row, 1) = rstRecibos!SucursalCheque
                        
                        rstRecibos.Close
                                    
                        Call Suma_Datos
                                
                        WVector1.Col = 1
                        WVector1.Row = WVector1.Row + 1
                        WControl = "N"
                        Call StartEdit
                        
                            Else
                        
                        WVector1.Text = ""
                            
                    End If
                    
                        Else
            
            
                    If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 3 Then
                        Auxi$ = Str$(Val(WVector1.Text))
                        Call Ceros(Auxi$, 2)
                        WVector1.Text = Auxi$
                            
                        Select Case Val(WVector1.Text)
                            Case 1
                                WVector1.Col = 2
                                WVector1.Text = ""
                                WVector1.Col = 3
                                WVector1.Text = ""
                                WVector1.Col = 4
                                WVector1.Text = ""
                                WVector1.Col = 5
                                WVector1.Text = Importe.Text
                                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                                Call Suma_Datos
                                WVector1.Col = 0
                                WVector1.Row = WVector1.Row + 1
                            Case Else
                        End Select
                            
                            Else
                                
                        WControl = "N"
                            
                    End If
                End If
            End If
            
        Case 3
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
    WVector1.Cols = 10
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
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1150
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 6
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
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

Private Sub Banco_DblClick()

    Opcion.Clear
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cartera de Cheques"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Cheque_Click()

    Opcion.Clear
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cartera de Cheques"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Banco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Acredita_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Importe_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 117
            Call Cheque_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub

















Private Sub Titular_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaEmision.SetFocus
    End If
    If KeyAscii = 27 Then
        Titular.Text = ""
    End If
End Sub

Private Sub FechaEmision_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(FechaEmision.Text, Auxi)
        If Auxi = "S" Then
            Cuenta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaEmision.Text = "  /  /    "
    End If
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.TextMatrix(ZZLugar, 7) = Titular.Text
        WVector1.TextMatrix(ZZLugar, 8) = FechaEmision.Text
        WVector1.TextMatrix(ZZLugar, 9) = Cuenta.Text
        PantaVaria.Visible = False
    End If
    If KeyAscii = 27 Then
        Cuenta.Text = ""
    End If
End Sub
