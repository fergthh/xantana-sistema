VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form prgArticulo 
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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   9000
      TabIndex        =   91
      Top             =   4080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox StockVI 
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
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   89
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StockIII 
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
      Left            =   6480
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   87
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox StockV 
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
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   85
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StockII 
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
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   83
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox StockIV 
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
      TabIndex        =   81
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox StockI 
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
      TabIndex        =   79
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   12840
      TabIndex        =   75
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame PantaArticulo 
      Height          =   4335
      Left            =   240
      TabIndex        =   60
      Top             =   5640
      Width           =   9255
      Begin VB.CheckBox TipoBusqueda 
         Caption         =   "Todos"
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
         Left            =   7800
         TabIndex        =   78
         Top             =   3720
         Width           =   975
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   63
         Top             =   1080
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   1320
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   3975
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   7011
         _Version        =   393216
         BackColor       =   16777152
      End
   End
   Begin VB.Frame PantaPrecios 
      Height          =   4095
      Left            =   240
      TabIndex        =   35
      Top             =   5400
      Visible         =   0   'False
      Width           =   11295
      Begin VB.ComboBox Moneda 
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
         Left            =   9480
         TabIndex        =   74
         Top             =   240
         Width           =   1695
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   73
         Top             =   1920
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   1920
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   71
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
         Index           =   4
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   70
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   69
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
         Index           =   6
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   68
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
         Index           =   7
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   3000
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   2760
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
         Left            =   3240
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2520
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
         Index           =   10
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   2280
         Width           =   375
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
         Left            =   120
         MaxLength       =   6
         TabIndex        =   52
         Top             =   1080
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
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   51
         Top             =   1080
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
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   50
         Top             =   1080
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
         Left            =   3840
         MaxLength       =   10
         TabIndex        =   49
         Top             =   1080
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
         MaxLength       =   6
         TabIndex        =   48
         Top             =   1080
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
         MaxLength       =   10
         TabIndex        =   47
         Top             =   1080
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
         MaxLength       =   6
         TabIndex        =   46
         Top             =   1080
         Width           =   1095
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
         MaxLength       =   10
         TabIndex        =   45
         Top             =   1080
         Width           =   1095
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
         Left            =   960
         MaxLength       =   8
         TabIndex        =   38
         Top             =   240
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   2415
         Left            =   120
         TabIndex        =   36
         Top             =   1560
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   4260
         _Version        =   393216
         BackColor       =   16777152
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   5520
         TabIndex        =   41
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   7800
         TabIndex        =   42
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label14 
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
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label13 
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
         Left            =   2760
         TabIndex        =   59
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label12 
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
         TabIndex        =   58
         Top             =   720
         Width           =   855
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
         TabIndex        =   57
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Left            =   1320
         TabIndex        =   56
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
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
         Left            =   3960
         TabIndex        =   55
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
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
         TabIndex        =   54
         Top             =   720
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
         Left            =   9000
         TabIndex        =   53
         Top             =   720
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
         Left            =   4800
         TabIndex        =   44
         Top             =   240
         Width           =   855
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
         Left            =   7080
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
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
         Left            =   360
         TabIndex        =   40
         Top             =   240
         Width           =   1215
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
         Left            =   1920
         TabIndex        =   39
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.TextBox Sector 
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
      Left            =   1080
      MaxLength       =   8
      TabIndex        =   32
      Top             =   2880
      Width           =   855
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
      Left            =   6600
      MaxLength       =   10
      TabIndex        =   30
      Text            =   " "
      Top             =   240
      Visible         =   0   'False
      Width           =   255
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
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   29
      Text            =   " "
      Top             =   240
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
      Left            =   5955
      MaxLength       =   10
      TabIndex        =   28
      Text            =   " "
      Top             =   240
      Visible         =   0   'False
      Width           =   180
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
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   27
      Text            =   " "
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Etiqueta 
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
      Left            =   6000
      TabIndex        =   25
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ComboBox Facturable 
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
      Left            =   6000
      TabIndex        =   23
      Top             =   3240
      Width           =   1455
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
      Left            =   1560
      MaxLength       =   25
      TabIndex        =   22
      Top             =   960
      Width           =   3975
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
      Left            =   1080
      TabIndex        =   20
      Top             =   3240
      Width           =   1455
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
      Left            =   7200
      MouseIcon       =   "articulo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Pedidos de Clientes"
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Stock 
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
      TabIndex        =   17
      Top             =   1560
      Width           =   1095
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
      MouseIcon       =   "articulo.frx":0BD4
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   4320
      Width           =   735
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
      MouseIcon       =   "articulo.frx":1720
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Elimina el Registro"
      Top             =   4320
      Width           =   735
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
      Left            =   1920
      MouseIcon       =   "articulo.frx":226C
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":2576
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Limpia la pantalla"
      Top             =   4320
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
      Left            =   2760
      MouseIcon       =   "articulo.frx":2DB8
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":30C2
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Consulta de Datos"
      Top             =   4320
      Width           =   735
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
      Left            =   6960
      MouseIcon       =   "articulo.frx":3904
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":3C0E
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salida"
      Top             =   4320
      Width           =   735
   End
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
      Height          =   975
      Left            =   3600
      MouseIcon       =   "articulo.frx":4450
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":475A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Primer Registro"
      Top             =   4320
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
      Left            =   4440
      MouseIcon       =   "articulo.frx":4B9C
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":4EA6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Registro Anterior"
      Top             =   4320
      Width           =   735
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
      Left            =   5280
      MouseIcon       =   "articulo.frx":52E8
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":55F2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Registro Siguiente"
      Top             =   4320
      Width           =   735
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
      Height          =   975
      Left            =   6120
      MouseIcon       =   "articulo.frx":5A34
      MousePointer    =   99  'Custom
      Picture         =   "articulo.frx":5D3E
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salida"
      Top             =   4320
      Width           =   735
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
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
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
      Left            =   8640
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   960
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
      TabIndex        =   4
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   2
      Top             =   600
      Width           =   5535
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
      Height          =   3180
      ItemData        =   "articulo.frx":6180
      Left            =   8040
      List            =   "articulo.frx":6187
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   5775
   End
   Begin MSMask.MaskEdBox FechaInactivo 
      Height          =   285
      Left            =   3000
      TabIndex        =   31
      Top             =   3240
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
   Begin VB.Label Label22 
      Caption         =   "En Tercero"
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
      TabIndex        =   90
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "Deposito III"
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
      TabIndex        =   88
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label20 
      Caption         =   "Mk"
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
      Left            =   2760
      TabIndex        =   86
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label18 
      Caption         =   "Produccion"
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
      Left            =   2760
      TabIndex        =   84
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label17 
      Caption         =   "P/Facturar"
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
      TabIndex        =   82
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "General"
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
      TabIndex        =   80
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Reducida"
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
      Left            =   360
      TabIndex        =   77
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Extendida"
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
      TabIndex        =   76
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Sector"
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
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label DesSector 
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
      Left            =   3120
      TabIndex        =   33
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label4 
      Caption         =   "Etiqueta"
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
      Left            =   4800
      TabIndex        =   26
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Facturable"
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
      Left            =   4800
      TabIndex        =   24
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   240
      X2              =   7680
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   7680
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7680
      Y1              =   1440
      Y2              =   1440
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
      Left            =   240
      TabIndex        =   21
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label27 
      Caption         =   "Stock"
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
      TabIndex        =   18
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   6
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "prgArticulo"
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
    prgArticulo.Hide
    Unload Me
    MenuVen.Show
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

Private Sub Descripcion_Keypress(KeyAscii As Integer)
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
            Tipo.SetFocus
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
            Fragancia.SetFocus
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
            Calidad.SetFocus
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
            Tamano.SetFocus
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

    Linea.Text = ""
    Tipo.Text = ""
    Fragancia.Text = ""
    Calidad.Text = ""
    Tamano.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    Sector.Text = ""
    DesSector.Caption = ""
    Stock.Text = ""
    StockI.Text = ""
    StockII.Text = ""
    StockII.Text = ""
    StockIV.Text = ""
    StockV.Text = ""
    StockVI.Text = ""
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
    
    Moneda.Clear
    
    Moneda.AddItem "Pesos"
    Moneda.AddItem "Dolares"
    
    Moneda.ListIndex = 1
    
    TipoBusqueda.Value = 0
    
    Activo.Clear
    
    Activo.AddItem "Si"
    Activo.AddItem "No"
    
    Activo.ListIndex = 0

    Facturable.Clear
     
    Facturable.AddItem "Si"
    Facturable.AddItem "No"
    
    Facturable.ListIndex = 0
     
    
    Etiqueta.Clear
    
    Etiqueta.AddItem ""
    Etiqueta.AddItem "Si"
    Etiqueta.AddItem "No"
    
    Etiqueta.ListIndex = 0
    
    Call Limpia_Vector
    Call Limpia_Vector2
    Call LInea_DblClick
    
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




Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 3
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
                WVector1.Text = "Codigo"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 5000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
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

