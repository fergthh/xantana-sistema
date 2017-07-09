VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgpago 
   AutoRedraw      =   -1  'True
   Caption         =   "Orden de Pago"
   ClientHeight    =   8055
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8055
   ScaleWidth      =   11880
   Begin VB.Frame BusquedaCheque 
      Height          =   3255
      Left            =   5400
      TabIndex        =   67
      Top             =   3240
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela"
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
         Left            =   3000
         TabIndex        =   80
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         Left            =   960
         TabIndex        =   78
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Maximo 
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
         TabIndex        =   72
         Text            =   " "
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Compensacion 
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
         TabIndex        =   70
         Text            =   " "
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox Efectivo 
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
         TabIndex        =   68
         Text            =   " "
         Top             =   240
         Width           =   1335
      End
      Begin MSMask.MaskEdBox PlazoMaximo 
         Height          =   285
         Left            =   2520
         TabIndex        =   74
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox PlazoInicial 
         Height          =   285
         Left            =   2520
         TabIndex        =   76
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label17 
         Caption         =   "Plazo Inicial"
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
         TabIndex        =   77
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Plazo Maximo"
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
         TabIndex        =   75
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Maximo Cheque"
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
         Left            =   600
         TabIndex        =   73
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Compensacion"
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
         Left            =   600
         TabIndex        =   71
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Left            =   600
         TabIndex        =   69
         Top             =   240
         Width           =   1575
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
      Index           =   13
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Busqueda 
      Caption         =   "Busqueda"
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
      MouseIcon       =   "Pago.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   79
      ToolTipText     =   "Cartera de Cheques"
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox Solicitud 
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
      Left            =   5400
      TabIndex        =   64
      Text            =   " "
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
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
      Left            =   11040
      MouseIcon       =   "Pago.frx":0799
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":0AA3
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Menu Principal"
      Top             =   2400
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
      Left            =   6240
      MouseIcon       =   "Pago.frx":12E5
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":15EF
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Consulta de Datos"
      Top             =   2400
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
      Left            =   5280
      MouseIcon       =   "Pago.frx":1E31
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":213B
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2400
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
      Left            =   4320
      MouseIcon       =   "Pago.frx":297D
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":2C87
      Style           =   1  'Graphical
      TabIndex        =   60
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
      Left            =   3360
      MouseIcon       =   "Pago.frx":34C9
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":37D3
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2400
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
      Left            =   10080
      MouseIcon       =   "Pago.frx":4015
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":431F
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Impresion de Orden de Pago"
      Top             =   2400
      Width           =   855
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
      Left            =   8160
      MouseIcon       =   "Pago.frx":4B61
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":4E6B
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Cartera de Cheques"
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton CtaCte 
      Caption         =   "Cta.Cte. F5"
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
      Left            =   7200
      MouseIcon       =   "Pago.frx":52FA
      MousePointer    =   99  'Custom
      Picture         =   "Pago.frx":5604
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Cuenta Corriente de Proveedores"
      Top             =   2400
      Width           =   855
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
      Left            =   8040
      TabIndex        =   54
      Text            =   " "
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox ReteIva 
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
      Left            =   5760
      TabIndex        =   52
      Text            =   " "
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame IngreCuenta1 
      Caption         =   "Cuenta Contable"
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
      Height          =   855
      Left            =   3480
      TabIndex        =   47
      Top             =   3360
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox Cuenta1 
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
         TabIndex        =   48
         Text            =   " "
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label8 
         Caption         =   "Codigo"
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
         TabIndex        =   49
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame IngreCuenta 
      Caption         =   "Cuenta Contable"
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
      Height          =   855
      Left            =   3480
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   4935
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
         Left            =   1560
         TabIndex        =   19
         Text            =   " "
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Codigo"
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
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
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
      Left            =   960
      TabIndex        =   42
      Top             =   4320
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   41
      Top             =   4320
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
      Left            =   2160
      TabIndex        =   40
      Top             =   4320
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
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4920
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4920
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4920
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   4920
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   4920
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
      Index           =   6
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   5400
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   5400
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   5400
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
      Index           =   9
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   5400
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
      Index           =   10
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   5400
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
      Index           =   11
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   4920
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
      Index           =   12
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   5400
      Width           =   375
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
      Left            =   6840
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.TextBox Retencion 
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
      Left            =   5760
      TabIndex        =   24
      Text            =   " "
      Top             =   1560
      Width           =   975
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
      Left            =   1680
      TabIndex        =   22
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
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
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   15
      Text            =   " "
      Top             =   840
      Width           =   5055
   End
   Begin Crystal.CrystalReport LISTADO 
      Left            =   7680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Impreord.rpt"
      WindowTitle     =   "Orden de Pago"
      CopiesToPrinter =   2
      WindowState     =   2
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Orden de Pago"
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
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2655
      Begin VB.OptionButton Tipo4 
         Caption         =   "Transferencias"
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
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton Tipo3 
         Caption         =   "Pagos Varios"
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
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Pagos de Cta.Cte."
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
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Tipo2 
         Caption         =   "Anticipos"
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
         Top             =   960
         Width           =   1335
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
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   2
      Text            =   " "
      Top             =   480
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
      Height          =   1020
      Left            =   7560
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      Top             =   120
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
   Begin VB.TextBox Orden 
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
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7320
      TabIndex        =   6
      Top             =   0
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
      Height          =   1260
      ItemData        =   "Pago.frx":5ECE
      Left            =   6840
      List            =   "Pago.frx":5ED5
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Top             =   4320
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
      Height          =   3375
      Left            =   0
      TabIndex        =   43
      Top             =   3840
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   5953
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Diferencia 
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
      Left            =   10320
      TabIndex        =   66
      Top             =   7560
      Width           =   1095
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
      Left            =   5880
      TabIndex        =   65
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label12 
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
      Height          =   255
      Left            =   6840
      TabIndex        =   55
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "Ret. Iva"
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
      Left            =   4560
      TabIndex        =   53
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "VALORES A ENTREGAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   51
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "COMPROBANTES A CANCELAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   50
      Top             =   3480
      Width           =   5535
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Forma de Pago : 1) Efectivo  2) Banco  4) Comp  5) Caja"
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
      Left            =   5160
      TabIndex        =   46
      Top             =   7320
      Width           =   5055
   End
   Begin VB.Label Debitos 
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
      Left            =   2880
      TabIndex        =   45
      Top             =   7320
      Width           =   1095
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
      Left            =   10320
      TabIndex        =   44
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "Ganancias"
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
      Left            =   4560
      TabIndex        =   26
      Top             =   1560
      Width           =   1575
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
      Height          =   255
      Left            =   2520
      TabIndex        =   23
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label LabelBanco 
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
      Left            =   120
      TabIndex        =   21
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
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
      TabIndex        =   14
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1455
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
      Left            =   3120
      TabIndex        =   9
      Top             =   480
      Width           =   3615
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
      Left            =   2880
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   " "
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Orden de Pago"
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
      Width           =   1815
   End
End
Attribute VB_Name = "Prgpago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Debito As Double
Private Credito As Double
Private WImpresion(200, 10) As String
Private WImpre2(100, 10) As String
Private WDebito(100, 2) As String
Private WCredito(100, 4) As String
Private WCuenta(100) As String
Private WCuenta1(100) As String
Private WCuentaBco As String
Private WVectorIva(100, 4) As String

Private Numero As String
Private WNumero As String
Private WSaldo As Double
Private WRetencion As Double
Private WReteIva As Double
Private WCuatri  As String
Private WEmpNombre As String
Private WEmpDirecion As String
Private WEmpLocalidad As String
Private WEmpCuit As String
Private WPrvDireccion As String
Private WPrvCuit As String
Private WLeyenda(10) As String
Private WTipo As String
Private WTipoprv As Single
Private WTipoiva As Single
Private WTipoReteiva As Single
Private WExepcion As Double
Private WNeto As Double
Private WAnticipo As Double
Private WBruto As Double
Private WIva As Double
Private WRetenido As Double
Private WFecha As String
Private WNroRet As Integer
Private WNroRet1 As Integer
Private XNeto As Double
Private XBruto As Double
Private XIva As Double
Private XTBase As Double
Private XImpor As Double
Private XPara(0 To 10) As Double
Private WTasa1(10) As Double
Private WAuxi As Double
Private Total As Double
Private Auxi As String
Private Auxi11 As String
Private XSaldo As Double
Private Tipocuenta As String
Private AuxiFecha As String
Private WProveedor As String
Private WTipocta As Integer
Dim BajaCheque(100) As String
Private WCtaChequeRecha As String
Dim WMinimo1 As Double
Dim WMinimo2 As Double
Dim WMinimo3 As Double
Dim WMinimo4 As Double
Dim WRetMinima As Double
Dim Existe As String
Dim PorceRIva As Double
Private XNetoIva As Double
Private XBrutoIva As Double
Private XIvaIva As Double
Private XReteIva As Double
Private WPorceFactura As Double
Private NetoParcial As Double
Private IvaParcial As Double
Private NetoTotal As Double
Private IvaTotal As Double
Private FacturaTotal As Double
Private ZSaldo As Double
Private ZLetra As String
Private ZTipo As String
Private ZPunto As String
Private ZNumero As String
Private ZProveedor As String
Dim ZMes As String
Dim ZAno As String

Dim ZZClave As String
Dim ZZOrden As String
Dim ZZRenglon As String
Dim ZZProveedor As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZTipoOrd As String
Dim ZZRetGanancias As String
Dim ZZRetIva As String
Dim ZZRetOtra As String
Dim ZZRetencion As String
Dim ZZTipoReg As String
Dim ZZTipo1 As String
Dim ZZLetra1 As String
Dim ZZPunto1 As String
Dim ZZNumero1 As String
Dim ZZImporte1 As String
Dim ZZTipo2 As String
Dim ZZNumero2 As String
Dim ZZFecha2 As String
Dim ZZBanco2 As String
Dim ZZImporte2 As String
Dim ZZObservaciones2 As String
Dim ZZConcepto As String
Dim ZZObservaciones As String
Dim ZZImporte As String
Dim ZZFechaOrd2 As String
Dim ZZImpoList As String
Dim ZZCuenta As String
Dim ZZSolicitud As String
Dim ZZClaveCheque As String
Dim ZZNroRet As String
Dim ZZNroRet1 As String
Dim ZZPorceIva As String
Dim ZZPorceRIva As String
Dim ZZExepcion As String
Dim ZZImpo1 As String
Dim ZZImpo2 As String
Dim ZZImpo3 As String
Dim ZZImpo4 As String
Dim ZXBrutoIva As String
Dim ZXNetoIva As String
Dim ZXIvaIva As String

Dim ZZLetra  As String
Dim ZZTipo  As String
Dim ZZPunto  As String
Dim ZZNumero  As String
Dim ZZEstado As String
Dim ZZVencimiento As String
Dim ZZVencimiento1 As String
Dim ZZTotal As String
Dim ZZSaldo As String
Dim ZZOrdFecha As String
Dim ZZOrdVencimiento As String
Dim ZZImpre As String
Dim ZZSaldoList As String
Dim ZZNroInterno As String
Dim ZZLista As String
Dim ZZAcumulado As String
Dim ZZEmpresa As String
Dim ZZCodigoEmpresa As String

Dim ZZFecha1  As String
Dim ZZDescripcion  As String
Dim ZZDia  As String
Dim ZZMes  As String
Dim ZZAno  As String
Dim ZZNombre  As String
Dim ZConcepto  As String
Dim ZDesCuenta  As String
Dim ZZRetib  As String
Dim ZZTasa   As String

Dim WPlazo1 As Integer
Dim WVencimiento As String
Dim WAno As String
Dim ZFecha As String
Dim ZCheque(100, 5) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub Suma_Datos()

    Debitos.Caption = ""
    Creditos.Caption = ""
    
    For IRow = 1 To 50
        Debitos.Caption = Str$(Val(Debitos.Caption) + Val(WVector1.TextMatrix(IRow, 5)))
        Creditos.Caption = Str$(Val(Creditos.Caption) + Val(WVector1.TextMatrix(IRow, 12)))
    Next IRow
    
    If Existe = "N" Then
        Call calcret_Click
    End If
    
    Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Retencion.Text) + Val(Reteiva.Text))
    
    Debitos.Caption = Pusing("###,###.##", Debitos.Caption)
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
    
    Diferencia.Caption = Str$(Val(Debitos.Caption) - Val(Creditos.Caption))
    Diferencia.Caption = Pusing("###,###.##", Diferencia.Caption)
    
End Sub

Private Sub Lee_Datos()

    Call Limpia_Vector
    
    Renglon = 0
    Debito = 0
    Credito = 0
    
    Erase BajaCheque
    Erase WCuenta
    Erase WCuenta1
    
    
    
    
    Do
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
    
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pagos"
        ZSql = ZSql + " Where Pagos.Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Pagos.Renglon = " + "'" + Auxi1 + "'"
        spPagos = ZSql
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
            
            Select Case Val(rstPagos!Tiporeg)
                Case 1
                    Debito = Debito + 1
                    WVector1.Row = Debito
                    WVector1.Col = 1
                    WVector1.Text = rstPagos!Tipo1
                    WVector1.Col = 2
                    WVector1.Text = rstPagos!Letra1
                    WVector1.Col = 3
                    WVector1.Text = rstPagos!Punto1
                    WVector1.Col = 4
                    WVector1.Text = rstPagos!Numero1
                    WVector1.Col = 5
                    WVector1.Text = rstPagos!Importe1
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    WVector1.Col = 6
                    WVector1.Text = rstPagos!Observaciones2
                    WCuenta(WVector1.Row) = rstPagos!Cuenta
                    If rstPagos!Banco2 <> 0 Then
                        Banco.Text = rstPagos!Banco2
                    End If
                    
                Case 2
                    Credito = Credito + 1
                    WVector1.Row = Credito
                    WVector1.Col = 7
                    WVector1.Text = rstPagos!Tipo2
                    WVector1.Col = 8
                    WVector1.Text = rstPagos!Numero2
                    WVector1.Col = 9
                    WVector1.Text = rstPagos!Fecha2
                    WVector1.Col = 10
                    WVector1.Text = rstPagos!Banco2
                    WVector1.Col = 11
                    If rstPagos!Observaciones2 <> "" Then
                        WVector1.Text = rstPagos!Observaciones2
                    End If
                    WVector1.Col = 12
                    WVector1.Text = rstPagos!Importe2
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    BajaCheque(WVector1.Row) = rstPagos!ClaveCheque
                    WCuenta1(WVector1.Row) = rstPagos!Cuenta
                Case Else
            End Select
            
            rstPagos.Close
            
                Else
            Exit Do
        End If
    Loop
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = rstProveedor!Nombre
        WPrvDireccion = rstProveedor!Direccion
        WPrvCuit = rstProveedor!Cuit
        WTipoprv = rstProveedor!Ganancia
        WTipoiva = rstProveedor!Iva
        rstProveedor.Close
            Else
        DesProveedor.Caption = ""
        WPrvDireccion = ""
        WPrvCuit = ""
        WTipoprv = 0
        WTipoiva = 0
        WTipoReteiva = 0
        WExepcion = 0
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

Private Sub Acepta_Click()

    OrdInicial = Right$(PlazoInicial.Text, 4) + Mid$(PlazoInicial.Text, 4, 2) + Left$(PlazoInicial.Text, 2)
    OrdMaximo = Right$(PlazoMaximo.Text, 4) + Mid$(PlazoMaximo.Text, 4, 2) + Left$(PlazoMaximo.Text, 2)
    
    ZSuma = Val(Compensacion.Text)
    ZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.TipoReg = '2'"
    ZSql = ZSql + " and Recibos.Tipo2 = '02'"
    ZSql = ZSql + " and Recibos.Estado2 <> 'X'"
    ZSql = ZSql + " Order by Recibos.FechaOrd2 desc"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstRecibos!FechaOrd2 >= OrdInicial And rstRecibos!FechaOrd2 <= OrdMaximo Then
                        If rstRecibos!Importe2 <> 0 Then
                            If Val(Maximo.Text) = 0 Or rstRecibos!Importe2 <= Val(Maximo.Text) Then
                                ZClaseCheque = IIf(IsNull(rstRecibos!ClaseCheque), "0", rstRecibos!ClaseCheque)
                                If Val(ZClaseCheque) <> 2 Then
                                    ZPrueba = ZSuma + rstRecibos!Importe2
                                    If ZPrueba < Val(Debitos.Caption) Then
                                        ZSuma = ZPrueba
                                        ZLugar = ZLugar + 1
                                        ZCheque(ZLugar, 1) = rstRecibos!Numero2
                                        ZCheque(ZLugar, 2) = rstRecibos!Fecha2
                                        ZCheque(ZLugar, 3) = rstRecibos!Banco2
                                        ZCheque(ZLugar, 4) = Str$(rstRecibos!Importe2)
                                        ZCheque(ZLugar, 5) = rstRecibos!Clave
                                    End If
                                End If
                            End If
                        End If
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    If ZSuma + Val(Efectivo.Text) >= Val(Debitos.Caption) Then
    
        For Ciclo = 1 To ZLugar
            
            WVector1.Row = Ciclo
            ZZLugar = Ciclo
                
            WVector1.Col = 8
            WVector1.Text = ZCheque(Ciclo, 1)
            
            WVector1.Col = 9
            WVector1.Text = ZCheque(Ciclo, 2)
            
            WVector1.Col = 10
            WVector1.Text = ""
                
            WVector1.Col = 11
            WVector1.Text = ZCheque(Ciclo, 3)
            
            WVector1.Col = 12
            WVector1.Text = ZCheque(Ciclo, 4)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        
            WVector1.Col = 7
            WVector1.Text = "03"
                
            BajaCheque(WVector1.Row) = ZCheque(Ciclo, 5)
        
        Next Ciclo
        
        ZDife = Val(Debitos.Caption) - ZSuma
        
        If Val(Efectivo.Text) <> 0 Then
        
            If Val(Efectivo.Text) > ZDife Then
            
                ZZLugar = ZZLugar + 1
            
                WVector1.Row = ZZLugar
                    
                WVector1.Col = 8
                WVector1.Text = ""
                
                WVector1.Col = 9
                WVector1.Text = ""
                
                WVector1.Col = 10
                WVector1.Text = ""
                    
                WVector1.Col = 11
                WVector1.Text = ""
                
                WVector1.Col = 12
                WVector1.Text = Str$(ZDife)
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            
                WVector1.Col = 7
                WVector1.Text = "01"
                    
                BajaCheque(ZZCiclo) = ""
                
                ZDife = 0
                    
                    Else
            
                ZZLugar = ZZLugar + 1
            
                WVector1.Row = ZZLugar
                    
                WVector1.Col = 8
                WVector1.Text = ""
                
                WVector1.Col = 9
                WVector1.Text = ""
                
                WVector1.Col = 10
                WVector1.Text = ""
                    
                WVector1.Col = 11
                WVector1.Text = ""
                
                WVector1.Col = 12
                WVector1.Text = Efectivo.Text
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            
                WVector1.Col = 7
                WVector1.Text = "01"
                    
                BajaCheque(ZZCiclo) = ""
                
                ZDife = ZDife - Val(Efectivo.Text)
                
            End If
            
        End If
        
        If Val(Compensacion.Text) <> 0 Then
        
            ZZLugar = ZZLugar + 1
        
            WVector1.Row = ZZLugar
                
            WVector1.Col = 8
            WVector1.Text = ""
            
            WVector1.Col = 9
            WVector1.Text = ""
            
            WVector1.Col = 10
            WVector1.Text = ""
                
            WVector1.Col = 11
            WVector1.Text = ""
            
            WVector1.Col = 12
            WVector1.Text = Compensacion.Text
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        
            WVector1.Col = 7
            WVector1.Text = "04"
                
            BajaCheque(ZZCiclo) = ""
                
            ZDife = 0
            
        End If

        Call Suma_Datos
        
            Else
            
        m$ = "No hay valores que cumplan con las condiciones requeridas"
        A% = MsgBox(m$, 0, "Busqueda de Cheques en cartera")
        
    End If
    
    BusquedaCheque.Visible = False

End Sub

Private Sub Busqueda_Click()

    WFecha = Fecha.Text
    
    WAno = Mid$(WFecha, 7, 4)
    Ano = Val(WAno)
    WMes = Mid$(WFecha, 4, 2)
    Mes = Val(WMes)
    WDia = Mid$(WFecha, 1, 2)
    Dia = Val(WDia)
    
    WAno = Str$(Val(WAno) - 1)
    Call Ceros(WAno, 4)

    ZFecha = Mid$(WFecha, 1, 6) + WAno
    
    Dife = DateDiff("d", ZFecha, WFecha)

    WPlazo1 = Dife - 29
    Call Calcula_vencimiento(ZFecha, WPlazo1, WVencimiento)
    
    Efectivo.Text = ""
    Compensacion.Text = ""
    Maximo.Text = ""
    PlazoMaximo.Text = "  /  /    "
    PlazoInicial.Text = WVencimiento

    BusquedaCheque.Visible = True
    Efectivo.SetFocus
    
End Sub

Private Sub Cancela_Click()
    BusquedaCheque.Visible = False
End Sub

Private Sub Efectivo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compensacion.SetFocus
    End If
    If KeyAscii = 27 Then
        Efectivo.Text = ""
    End If
End Sub

Private Sub Compensacion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Maximo.SetFocus
    End If
    If KeyAscii = 27 Then
        Compensacion.Text = ""
    End If
End Sub

Private Sub Maximo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PlazoMaximo.SetFocus
    End If
    If KeyAscii = 27 Then
        Maximo.Text = ""
    End If
End Sub

Private Sub PlazoMaximo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(PlazoMaximo.Text)) = 8 Then
            PlazoMaximo.Text = Left$(PlazoMaximo.Text, 6) + "20" + Right$(Trim(PlazoMaximo.Text), 2)
        End If
        Call Valida_fecha1(PlazoMaximo.Text, Auxi)
        If Auxi = "S" Then
            PlazoInicial.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        PlazoMaximo.Text = "  /  /    "
    End If
End Sub

Private Sub PlazoInicial_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(PlazoInicial.Text)) = 8 Then
            PlazoInicial.Text = Left$(PlazoInicial.Text, 6) + "20" + Right$(Trim(PlazoInicial.Text), 2)
        End If
        Call Valida_fecha1(PlazoInicial.Text, Auxi)
        If Auxi = "S" Then
            Rem Efectivo.SetFocus
            Call Acepta_Click
        End If
    End If
    If KeyAscii = 27 Then
        PlazoInicial.Text = "  /  /    "
    End If
End Sub

Private Sub cmdAdd_Click()
                        
    Rem VER GABACION

    Rem If WLicencia <> "1234-5678-ABCD-EFGH" And Val(Orden.Text) > 10 Then
    Rem     WMsg$ = "La version del sistema es para un uso limitado de movimientos." + Chr$(13) + _
    rem          "El objetivo es el de verificar las opciones y el funcionamiento del mismo." + Chr$(13) + _
    rem          "Para poder utilizar el sistema sin limite de movimientos se debe adquirir la version definitiva."
    Rem     A% = MsgBox(WMsg$, 0, "Sistema de Control de Gestion")
    Rem     Exit Sub
    Rem End If

    If Orden.Text <> "" And Fecha.Text <> "" Then
    
        For IRow = 1 To 50
        
            WRow = IRow
            If Val(WVector1.TextMatrix(IRow, 7)) = 2 Then
                WNumeroCheque = Val(WVector1.TextMatrix(IRow, 8))
                WBancoCheque = WVector1.TextMatrix(IRow, 10)
                
                Entra = "N"
    
                
            End If
            
        Next IRow
    
        If Proveedor.Text <> "" Or Tipo3.Value = True Or Tipo4.Value = True Then
    
            Auxi1 = Orden.Text
            Call Ceros(Auxi1, 6)
            Orden.Text = Auxi1
            
            Existe = "N"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pagos"
            ZSql = ZSql + " Where Pagos.Orden = " + "'" + Orden.Text + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            If rstPagos.RecordCount > 0 Then
                m$ = "Orden de Pago ya existente"
                A% = MsgBox(m$, 0, "Ingreso de Ordenes de Pago")
                Existe = "S"
                rstPagos.Close
            End If
            
            If Existe <> "S" Then
    
                Call Suma_Datos
        
                Debito = 0
                Credito = 0
                If Val(Debitos.Caption) <> 0 Then
                    Debito = Val(Debitos.Caption)
                End If
        
                If Val(Creditos.Caption) <> 0 Then
                    Credito = Val(Creditos.Caption)
                End If
        
                If Debito <> Credito Then
                    m$ = "Los valores de la orden de pago no balancean"
                    A% = MsgBox(m$, 0, "Ingreso de Ordenes de Pago")
                End If
        
                If Debito = Credito Then
                
                    WNroRet = 0
                    WNroRet1 = 0
                
                    If Val(Retencion.Text) <> 0 Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM NroRet"
                        ZSql = ZSql + " Where NroRet.Clave = 1"
                        spNroRet = ZSql
                        Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
                        If rstNroRet.RecordCount > 0 Then
                        
                            WNroRet = rstNroRet!Numero + 1
                            rstNroRet.Close
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE NroRet SET "
                            ZSql = ZSql + " Numero = Numero + 1"
                            ZSql = ZSql + " Where Clave = 1"
                            spNroRet = ZSql
                            Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            WNroRet = 1
                            
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO NroRet ("
                            ZSql = ZSql + "Clave ,"
                            ZSql = ZSql + "Numero )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + "1" + "',"
                            ZSql = ZSql + "'" + "1" + "')"
                            spNroRet = ZSql
                            Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                    
                    End If
                    
                    If Val(Reteiva.Text) <> 0 Then
                    
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM NroRet"
                        ZSql = ZSql + " Where NroRet.Clave = 2"
                        spNroRet = ZSql
                        Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
                        If rstNroRet.RecordCount > 0 Then
                        
                            WNroRet1 = rstNroRet!Numero + 1
                            rstNroRet.Close
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE NroRet SET "
                            ZSql = ZSql + " Numero = Numero + 1"
                            ZSql = ZSql + " Where Clave = 2"
                            spNroRet = ZSql
                            Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
                            
                                Else
                                
                            WNroRet1 = 1
                            
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO NroRet ("
                            ZSql = ZSql + "Clave ,"
                            ZSql = ZSql + "Numero )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + "2" + "',"
                            ZSql = ZSql + "'" + "1" + "')"
                            spNroRet = ZSql
                            Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                        
                    End If
        
                    If Val(Banco.Text) = 0 Then
                        Banco.Text = "0"
                    End If
                    
                    WProveedor = Trim(Proveedor.Text)
                    Renglon = 0
                    
                    For IRow = 1 To 50
                    
                        WRow = IRow
                        
                        WVector1.Col = 5
                        WVector1.Row = IRow
                        If Val(WVector1.Text) <> 0 Then
                        
                            Renglon = Renglon + 1
                            Auxi1 = Str$(Renglon)
                            Call Ceros(Auxi1, 2)
                            
                            ZZOrden = Orden.Text
                            ZZRenglon = Auxi1
                            ZZProveedor = Trim(Proveedor.Text)
                            ZZfecha = Fecha.Text
                            ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            ZZImporte = Str$(Debito)
                            ZZRetencion = Retencion.Text
                            ZZRetIva = Reteiva.Text
                            ZZObservaciones = Observaciones.Text
                            ZZCuenta = ""
                            
                            If Tipo1.Value = True Then
                                ZZTipoOrd = "1"
                            End If
                            If Tipo2.Value = True Then
                                ZZTipoOrd = "2"
                            End If
                            If Tipo3.Value = True Then
                                ZZTipoOrd = "3"
                                ZZCuenta = WCuenta(IRow)
                            End If
                            If Tipo4.Value = True Then
                                ZZTipoOrd = "4"
                                ZZCuenta = WCuenta(IRow)
                            End If
                        
                            ZZTipoReg = "1"
                            WVector1.Col = 1
                            ZZTipo1 = Left$(WVector1.Text, 2)
                            If Val(ZZTipo1) <> 0 Then
                                Call Ceros(ZZTipo1, 2)
                            End If
                            WVector1.Col = 2
                            ZZLetra1 = Left$(WVector1.Text, 1)
                            WVector1.Col = 3
                            ZZPunto1 = Left$(WVector1.Text, 4)
                            WVector1.Col = 4
                            ZZNumero1 = Left$(WVector1.Text, 8)
                            WVector1.Col = 5
                            ZZImporte1 = WVector1.Text
                            WVector1.Col = 6
                            
                            ZZObservaciones2 = Left$(WVector1.Text, 30)
                            ZZTipo2 = ""
                            ZZNumero2 = ""
                            ZZFecha2 = ""
                            ZZFechaOrd2 = ""
                            If Tipo4.Value = True Then
                                ZZBanco2 = Banco.Text
                                    Else
                                ZZBanco2 = "0"
                            End If
                            ZZImporte2 = "0"
                            
                            ZZSolicitud = Str$(Solicitud.ListIndex)
                            ZZClaveCheque = ""
                            ZZClaveLectora = ""
                            
                            ZZNroRet = WNroRet
                            ZZNroRet1 = WNroRet1
                            ZZExepcion = WExepcion
                            ZZPorceIva = PorceIva.Text
                            ZZPorceRIva = PorceRIva
                            
                            ZZRetGanancias = ""
                            ZZRetOtra = ""
                            ZZConcepto = ""
                            ZZImpoList = ""
                            ZZImpo1 = ""
                            ZZImpo2 = ""
                            ZZImpo3 = ""
                            ZZImpo4 = ""
                            
                            ZXBrutoIva = ""
                            ZXNetoIva = ""
                            ZXIvaIva = ""
                            
                            ZZClave = ZZOrden + ZZRenglon
                            
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Pagos ("
                            ZSql = ZSql + "Clave ,"
                            ZSql = ZSql + "Orden ,"
                            ZSql = ZSql + "Renglon ,"
                            ZSql = ZSql + "Proveedor ,"
                            ZSql = ZSql + "Fecha ,"
                            ZSql = ZSql + "FechaOrd ,"
                            ZSql = ZSql + "TipoOrd ,"
                            ZSql = ZSql + "RetGanancias ,"
                            ZSql = ZSql + "RetIva ,"
                            ZSql = ZSql + "RetOtra ,"
                            ZSql = ZSql + "Retencion ,"
                            ZSql = ZSql + "TipoReg ,"
                            ZSql = ZSql + "Tipo1 ,"
                            ZSql = ZSql + "Letra1 ,"
                            ZSql = ZSql + "Punto1 ,"
                            ZSql = ZSql + "Numero1 ,"
                            ZSql = ZSql + "Importe1 ,"
                            ZSql = ZSql + "Tipo2 ,"
                            ZSql = ZSql + "Numero2 ,"
                            ZSql = ZSql + "Fecha2 ,"
                            ZSql = ZSql + "Banco2 ,"
                            ZSql = ZSql + "Importe2 ,"
                            ZSql = ZSql + "Observaciones2 ,"
                            ZSql = ZSql + "Concepto ,"
                            ZSql = ZSql + "Observaciones ,"
                            ZSql = ZSql + "Importe ,"
                            ZSql = ZSql + "FechaOrd2 ,"
                            ZSql = ZSql + "ImpoList ,"
                            ZSql = ZSql + "Cuenta ,"
                            ZSql = ZSql + "Solicitud ,"
                            ZSql = ZSql + "ClaveCheque ,"
                            ZSql = ZSql + "ClaveLectora ,"
                            ZSql = ZSql + "NroRet ,"
                            ZSql = ZSql + "NroRet1 ,"
                            ZSql = ZSql + "PorceIva ,"
                            ZSql = ZSql + "PorceRIva ,"
                            ZSql = ZSql + "Exepcion ,"
                            ZSql = ZSql + "Impo1 ,"
                            ZSql = ZSql + "Impo2 ,"
                            ZSql = ZSql + "Impo3 ,"
                            ZSql = ZSql + "Impo4 ,"
                            ZSql = ZSql + "XBrutoIva ,"
                            ZSql = ZSql + "XNetoIva ,"
                            ZSql = ZSql + "XIvaIva )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + ZZClave + "',"
                            ZSql = ZSql + "'" + ZZOrden + "',"
                            ZSql = ZSql + "'" + ZZRenglon + "',"
                            ZSql = ZSql + "'" + ZZProveedor + "',"
                            ZSql = ZSql + "'" + ZZfecha + "',"
                            ZSql = ZSql + "'" + ZZFechaOrd + "',"
                            ZSql = ZSql + "'" + ZZTipoOrd + "',"
                            ZSql = ZSql + "'" + ZZRetGanancias + "',"
                            ZSql = ZSql + "'" + ZZRetIva + "',"
                            ZSql = ZSql + "'" + ZZRetOtra + "',"
                            ZSql = ZSql + "'" + ZZRetencion + "',"
                            ZSql = ZSql + "'" + ZZTipoReg + "',"
                            ZSql = ZSql + "'" + ZZTipo1 + "',"
                            ZSql = ZSql + "'" + ZZLetra1 + "',"
                            ZSql = ZSql + "'" + ZZPunto1 + "',"
                            ZSql = ZSql + "'" + ZZNumero1 + "',"
                            ZSql = ZSql + "'" + ZZImporte1 + "',"
                            ZSql = ZSql + "'" + ZZTipo2 + "',"
                            ZSql = ZSql + "'" + ZZNumero2 + "',"
                            ZSql = ZSql + "'" + ZZFecha2 + "',"
                            ZSql = ZSql + "'" + ZZBanco2 + "',"
                            ZSql = ZSql + "'" + ZZImporte2 + "',"
                            ZSql = ZSql + "'" + ZZObservaciones2 + "',"
                            ZSql = ZSql + "'" + ZZConcepto + "',"
                            ZSql = ZSql + "'" + ZZObservaciones + "',"
                            ZSql = ZSql + "'" + ZZImporte + "',"
                            ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                            ZSql = ZSql + "'" + ZZImpoList + "',"
                            ZSql = ZSql + "'" + ZZCuenta + "',"
                            ZSql = ZSql + "'" + ZZSolicitud + "',"
                            ZSql = ZSql + "'" + ZZClaveCheque + "',"
                            ZSql = ZSql + "'" + ZZClaveLectora + "',"
                            ZSql = ZSql + "'" + ZZNroRet + "',"
                            ZSql = ZSql + "'" + ZZNroRet1 + "',"
                            ZSql = ZSql + "'" + ZZPorceIva + "',"
                            ZSql = ZSql + "'" + ZZPorceRIva + "',"
                            ZSql = ZSql + "'" + ZZExepcion + "',"
                            ZSql = ZSql + "'" + ZZImpo1 + "',"
                            ZSql = ZSql + "'" + ZZImpo2 + "',"
                            ZSql = ZSql + "'" + ZZImpo3 + "',"
                            ZSql = ZSql + "'" + ZZImpo4 + "',"
                            ZSql = ZSql + "'" + ZXBrutoIva + "',"
                            ZSql = ZSql + "'" + ZXNetoIva + "',"
                            ZSql = ZSql + "'" + ZXIvaIva + "')"
                            
                            spPagos = ZSql
                            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                            If Tipo1.Value = True Then
                            
                                WLetra = ZZLetra1
                                WTipo = ZZTipo1
                                WPunto = ZZPunto1
                                WNumero = ZZNumero1
                                WImporte = ZZImporte1
                                Claveven$ = WProveedor + WLetra + WTipo + WPunto + WNumero
                                
                                ZSql = ""
                                ZSql = ZSql + "UPDATE CtaCtePrv SET "
                                ZSql = ZSql + " Saldo = Saldo - " + "'" + WImporte + "'"
                                ZSql = ZSql + " Where Clave = " + "'" + Claveven$ + "'"
                                spCtaCtePrv = ZSql
                                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                                
                            End If
                            
                        End If
                
                        WVector1.Col = 12
                        WVector1.Row = IRow
                        If Val(WVector1.Text) <> 0 Then
                        
                            Renglon = Renglon + 1
                            Auxi1 = Str$(Renglon)
                            Call Ceros(Auxi1, 2)
                            
                            ZZOrden = Orden.Text
                            ZZRenglon = Auxi1
                            ZZProveedor = Trim(Proveedor.Text)
                            ZZfecha = Fecha.Text
                            ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            ZZImporte = Str$(Debito)
                            ZZRetencion = Retencion.Text
                            ZZRetIva = Reteiva.Text
                            ZZObservaciones = Observaciones.Text
                            If Tipo1.Value = True Then
                                ZZTipoOrd = "1"
                            End If
                            If Tipo2.Value = True Then
                                ZZTipoOrd = "2"
                            End If
                            If Tipo3.Value = True Then
                                ZZTipoOrd = "3"
                            End If
                            If Tipo4.Value = True Then
                                ZZTipoOrd = "4"
                            End If
                            
                            ZZTipoReg = "2"
                            ZZTipo1 = ""
                            ZZLetra1 = ""
                            ZZPunto1 = ""
                            ZZNumero1 = ""
                            ZZImporte1 = 0
                            WVector1.Col = 7
                            ZZTipo2 = Left$(WVector1.Text, 2)
                            If Val(ZZTipo2) <> 0 Then
                                Call Ceros(ZZTipo2, 2)
                            End If
                            WVector1.Col = 8
                            ZZNumero2 = Left$(WVector1.Text, 8)
                            
                            WVector1.Col = 9
                            ZZFecha2 = Left$(WVector1.Text, 10)
                            ZZFechaOrd2 = Right$(ZZFecha2, 4) + Mid$(ZZFecha2, 4, 2) + Left$(ZZFecha2, 2)
                            WVector1.Col = 10
                            ZZBanco2 = WVector1.Text
                            WVector1.Col = 11
                            ZZObservaciones2 = Left$(WVector1.Text, 20)
                            WVector1.Col = 12
                            ZZImporte2 = WVector1.Text
                            WVector1.Col = 13
                            ZZClaveLectora = WVector1.Text
                            If Val(ZZTipo2) = 4 Then
                                ZZCuenta = WCuenta1(IRow)
                                    Else
                                ZZCuenta = ""
                            End If
                            
                            ZZSolicitud = Solicitud.ListIndex
                            If Val(ZZTipo2) = 3 Then
                                ZZClaveCheque = BajaCheque(WVector1.Row)
                                    Else
                                ZZClaveCheque = ""
                            End If
                            If Val(ZZTipo2) = 4 Then
                                ZZCuenta = WCuenta1(IRow)
                                    Else
                                ZZCuenta = ""
                            End If
                            
                            ZZClave = ZZOrden + ZZRenglon
                            ZZNroRet = WNroRet
                            ZZNroRet1 = WNroRet1
                            ZZExepcion = WExepcion
                            ZZPorceIva = PorceIva.Text
                            ZZPorceRIva = PorceRIva
                            
                            ZXBrutoIva = WVectorIva(IRow, 1)
                            ZXNetoIva = WVectorIva(IRow, 2)
                            ZXIvaIva = WVectorIva(IRow, 3)
                            
                            
                            ZZClave = ZZOrden + ZZRenglon
                            
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Pagos ("
                            ZSql = ZSql + "Clave ,"
                            ZSql = ZSql + "Orden ,"
                            ZSql = ZSql + "Renglon ,"
                            ZSql = ZSql + "Proveedor ,"
                            ZSql = ZSql + "Fecha ,"
                            ZSql = ZSql + "FechaOrd ,"
                            ZSql = ZSql + "TipoOrd ,"
                            ZSql = ZSql + "RetGanancias ,"
                            ZSql = ZSql + "RetIva ,"
                            ZSql = ZSql + "RetOtra ,"
                            ZSql = ZSql + "Retencion ,"
                            ZSql = ZSql + "TipoReg ,"
                            ZSql = ZSql + "Tipo1 ,"
                            ZSql = ZSql + "Letra1 ,"
                            ZSql = ZSql + "Punto1 ,"
                            ZSql = ZSql + "Numero1 ,"
                            ZSql = ZSql + "Importe1 ,"
                            ZSql = ZSql + "Tipo2 ,"
                            ZSql = ZSql + "Numero2 ,"
                            ZSql = ZSql + "Fecha2 ,"
                            ZSql = ZSql + "Banco2 ,"
                            ZSql = ZSql + "Importe2 ,"
                            ZSql = ZSql + "Observaciones2 ,"
                            ZSql = ZSql + "Concepto ,"
                            ZSql = ZSql + "Observaciones ,"
                            ZSql = ZSql + "Importe ,"
                            ZSql = ZSql + "FechaOrd2 ,"
                            ZSql = ZSql + "ImpoList ,"
                            ZSql = ZSql + "Cuenta ,"
                            ZSql = ZSql + "Solicitud ,"
                            ZSql = ZSql + "ClaveCheque ,"
                            ZSql = ZSql + "ClaveLectora ,"
                            ZSql = ZSql + "NroRet ,"
                            ZSql = ZSql + "NroRet1 ,"
                            ZSql = ZSql + "PorceIva ,"
                            ZSql = ZSql + "PorceRIva ,"
                            ZSql = ZSql + "Exepcion ,"
                            ZSql = ZSql + "Impo1 ,"
                            ZSql = ZSql + "Impo2 ,"
                            ZSql = ZSql + "Impo3 ,"
                            ZSql = ZSql + "Impo4 ,"
                            ZSql = ZSql + "XBrutoIva ,"
                            ZSql = ZSql + "XNetoIva ,"
                            ZSql = ZSql + "XIvaIva )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + ZZClave + "',"
                            ZSql = ZSql + "'" + ZZOrden + "',"
                            ZSql = ZSql + "'" + ZZRenglon + "',"
                            ZSql = ZSql + "'" + ZZProveedor + "',"
                            ZSql = ZSql + "'" + ZZfecha + "',"
                            ZSql = ZSql + "'" + ZZFechaOrd + "',"
                            ZSql = ZSql + "'" + ZZTipoOrd + "',"
                            ZSql = ZSql + "'" + ZZRetGanancias + "',"
                            ZSql = ZSql + "'" + ZZRetIva + "',"
                            ZSql = ZSql + "'" + ZZRetOtra + "',"
                            ZSql = ZSql + "'" + ZZRetencion + "',"
                            ZSql = ZSql + "'" + ZZTipoReg + "',"
                            ZSql = ZSql + "'" + ZZTipo1 + "',"
                            ZSql = ZSql + "'" + ZZLetra1 + "',"
                            ZSql = ZSql + "'" + ZZPunto1 + "',"
                            ZSql = ZSql + "'" + ZZNumero1 + "',"
                            ZSql = ZSql + "'" + ZZImporte1 + "',"
                            ZSql = ZSql + "'" + ZZTipo2 + "',"
                            ZSql = ZSql + "'" + ZZNumero2 + "',"
                            ZSql = ZSql + "'" + ZZFecha2 + "',"
                            ZSql = ZSql + "'" + ZZBanco2 + "',"
                            ZSql = ZSql + "'" + ZZImporte2 + "',"
                            ZSql = ZSql + "'" + ZZObservaciones2 + "',"
                            ZSql = ZSql + "'" + ZZConcepto + "',"
                            ZSql = ZSql + "'" + ZZObservaciones + "',"
                            ZSql = ZSql + "'" + ZZImporte + "',"
                            ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                            ZSql = ZSql + "'" + ZZImpoList + "',"
                            ZSql = ZSql + "'" + ZZCuenta + "',"
                            ZSql = ZSql + "'" + ZZSolicitud + "',"
                            ZSql = ZSql + "'" + ZZClaveCheque + "',"
                            ZSql = ZSql + "'" + ZZClaveLectora + "',"
                            ZSql = ZSql + "'" + ZZNroRet + "',"
                            ZSql = ZSql + "'" + ZZNroRet1 + "',"
                            ZSql = ZSql + "'" + ZZPorceIva + "',"
                            ZSql = ZSql + "'" + ZZPorceRIva + "',"
                            ZSql = ZSql + "'" + ZZExepcion + "',"
                            ZSql = ZSql + "'" + ZZImpo1 + "',"
                            ZSql = ZSql + "'" + ZZImpo2 + "',"
                            ZSql = ZSql + "'" + ZZImpo3 + "',"
                            ZSql = ZSql + "'" + ZZImpo4 + "',"
                            ZSql = ZSql + "'" + ZXBrutoIva + "',"
                            ZSql = ZSql + "'" + ZXNetoIva + "',"
                            ZSql = ZSql + "'" + ZXIvaIva + "')"
                            
                            spPagos = ZSql
                            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                            
                            If Val(ZZTipo2) = 3 Then
                            
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Recibos SET "
                                ZSql = ZSql + " Estado2 = " + "'" + "X" + "',"
                                ZSql = ZSql + " Orden = " + "'" + Orden.Text + "',"
                                ZSql = ZSql + " Deposito = " + "'" + "0" + "',"
                                ZSql = ZSql + " Destino = " + "'" + DesProveedor.Caption + "',"
                                ZSql = ZSql + " ProveedorSalida = " + "'" + Proveedor.Text + "',"
                                ZSql = ZSql + " BancoSalida = " + "'" + "0" + "'"
                                ZSql = ZSql + " Where Clave = " + "'" + BajaCheque(WVector1.Row) + "'"
                                spRecibos = ZSql
                                Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                            
                            End If
                    
                        End If
                
                    Next IRow
                    
                    Renglon = 0
                    
                    If Tipo1.Value = True Then
        
                        WLetra = "A"
                        WTipo = "04"
                        WPunto = "0000"
                        WNumero = Orden.Text
                        WProveedor = Trim(Proveedor.Text)
        
                        Call Ceros(WNumero, 8)
                        
                        ZZProveedor = Trim(Proveedor.Text)
                        ZZLetra = WLetra
                        ZZTipo = WTipo
                        ZZPunto = WPunto
                        ZZNumero = WNumero
                        ZZfecha = Fecha.Text
                        ZZEstado = "1"
                        ZZVencimiento = "  /  /    "
                        ZZTotal = Str$(Debito * -1)
                        ZZSaldo = "0"
                        ZZClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                        ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        ZZOrdVencimiento = "00000000"
                        ZZImpre = "OP"
                        ZZObservaciones = Observaciones.Text
                        
                        
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
                        ZSql = ZSql + "Vencimiento1 ,"
                        ZSql = ZSql + "Total ,"
                        ZSql = ZSql + "Saldo ,"
                        ZSql = ZSql + "OrdFecha ,"
                        ZSql = ZSql + "OrdVencimiento ,"
                        ZSql = ZSql + "Impre ,"
                        ZSql = ZSql + "SaldoList ,"
                        ZSql = ZSql + "NroInterno ,"
                        ZSql = ZSql + "Lista ,"
                        ZSql = ZSql + "Acumulado ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Empresa ,"
                        ZSql = ZSql + "CodigoEmpresa )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + ZZClave + "',"
                        ZSql = ZSql + "'" + ZZProveedor + "',"
                        ZSql = ZSql + "'" + ZZLetra + "',"
                        ZSql = ZSql + "'" + ZZTipo + "',"
                        ZSql = ZSql + "'" + ZZPunto + "',"
                        ZSql = ZSql + "'" + ZZNumero + "',"
                        ZSql = ZSql + "'" + ZZfecha + "',"
                        ZSql = ZSql + "'" + ZZEstado + "',"
                        ZSql = ZSql + "'" + ZZVencimiento + "',"
                        ZSql = ZSql + "'" + ZZVencimiento1 + "',"
                        ZSql = ZSql + "'" + ZZTotal + "',"
                        ZSql = ZSql + "'" + ZZSaldo + "',"
                        ZSql = ZSql + "'" + ZZOrdFecha + "',"
                        ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
                        ZSql = ZSql + "'" + ZZImpre + "',"
                        ZSql = ZSql + "'" + ZZSaldoList + "',"
                        ZSql = ZSql + "'" + ZZNroInterno + "',"
                        ZSql = ZSql + "'" + ZZLista + "',"
                        ZSql = ZSql + "'" + ZZAcumulado + "',"
                        ZSql = ZSql + "'" + ZZObservaciones + "',"
                        ZSql = ZSql + "'" + ZZEmpresa + "',"
                        ZSql = ZSql + "'" + ZZCodigoEmpresa + "')"
                            
                        spCtaCtePrv = ZSql
                        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    If Tipo2.Value = True Then
        
                        WLetra = "A"
                        WTipo = "05"
                        WPunto = "0000"
                        WNumero = Orden.Text
                        WProveedor = Trim(Proveedor.Text)
        
                        Call Ceros(WNumero, 8)
        
                        ZZProveedor = Trim(Proveedor.Text)
                        ZZLetra = WLetra
                        ZZTipo = WTipo
                        ZZPunto = WPunto
                        ZZNumero = WNumero
                        ZZfecha = Fecha.Text
                        ZZEstado = "1"
                        ZZVencimiento = "  /  /    "
                        ZZTotal = Str$(Debito * -1)
                        ZZSaldo = Str$(Debito * -1)
                        ZZClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                        ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        ZZOrdVencimiento = "00000000"
                        ZZImpre = "AN"
                        ZZObservaciones = Observaciones.Text
                        
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
                        ZSql = ZSql + "Vencimiento1 ,"
                        ZSql = ZSql + "Total ,"
                        ZSql = ZSql + "Saldo ,"
                        ZSql = ZSql + "OrdFecha ,"
                        ZSql = ZSql + "OrdVencimiento ,"
                        ZSql = ZSql + "Impre ,"
                        ZSql = ZSql + "SaldoList ,"
                        ZSql = ZSql + "NroInterno ,"
                        ZSql = ZSql + "Lista ,"
                        ZSql = ZSql + "Acumulado ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Empresa ,"
                        ZSql = ZSql + "CodigoEmpresa )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + ZZClave + "',"
                        ZSql = ZSql + "'" + ZZProveedor + "',"
                        ZSql = ZSql + "'" + ZZLetra + "',"
                        ZSql = ZSql + "'" + ZZTipo + "',"
                        ZSql = ZSql + "'" + ZZPunto + "',"
                        ZSql = ZSql + "'" + ZZNumero + "',"
                        ZSql = ZSql + "'" + ZZfecha + "',"
                        ZSql = ZSql + "'" + ZZEstado + "',"
                        ZSql = ZSql + "'" + ZZVencimiento + "',"
                        ZSql = ZSql + "'" + ZZVencimiento1 + "',"
                        ZSql = ZSql + "'" + ZZTotal + "',"
                        ZSql = ZSql + "'" + ZZSaldo + "',"
                        ZSql = ZSql + "'" + ZZOrdFecha + "',"
                        ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
                        ZSql = ZSql + "'" + ZZImpre + "',"
                        ZSql = ZSql + "'" + ZZSaldoList + "',"
                        ZSql = ZSql + "'" + ZZNroInterno + "',"
                        ZSql = ZSql + "'" + ZZLista + "',"
                        ZSql = ZSql + "'" + ZZAcumulado + "',"
                        ZSql = ZSql + "'" + ZZObservaciones + "',"
                        ZSql = ZSql + "'" + ZZEmpresa + "',"
                        ZSql = ZSql + "'" + ZZCodigoEmpresa + "')"
                            
                        spCtaCtePrv = ZSql
                        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
        
                    WFecha = Right$(Fecha.Text, 2) + Mid$(Fecha.Text, 4, 2)
                    Auxi = Trim(Proveedor.Text)
                    
                    Claveven$ = WFecha + Auxi
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Retencion SET "
                    ZSql = ZSql + " Neto = Neto + " + "'" + Str$(XNeto) + "',"
                    ZSql = ZSql + " Retenido = Retenido + " + "'" + Retencion.Text + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + Claveven$ + "'"
                    spRetencion = ZSql
                    Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
                    
                    T$ = "Impresion de Orden de Pago"
                    m$ = "Desea realizar la impresion del comprobante"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 6 Then
                        Call IMPREORDEN
                    End If
                    If Val(Retencion.Text) <> 0 Then
                        T$ = "Impresion de Orden de Pago"
                        m$ = "Desea realizar la impresion de la retencion"
                        Respuesta% = MsgBox(m$, 32 + 4, T$)
                        If Respuesta% = 6 Then
                            Call Impreret
                        End If
                    End If

                    Orden.SetFocus
                    Call CmdLimpiar_Click
        
                End If
        
            End If
        
        End If
        
    End If
End Sub

Private Sub CmdDelete_Click()
    If Orden.Text <> "" Then

        If Tipo2.Value = True Then
            WLetra = "A"
            WTipo = "05"
            WPunto = "0000"
            WNumero = Orden.Text
            WProveedor = Trim(Proveedor.Text)
        
            Call Ceros(WNumero, 8)
        
            ZZClave = WProveedor + WLetra + WTipo + WPunto + WNumero
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCtePrv"
            ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClave + "'"
            spCtaCtePrv = ZSql
            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCtePrv.RecordCount > 0 Then
                WSaldo = rstCtaCtePrv!Saldo
                WTotal = rstCtaCtePrv!Total
                rstCtaCtePrv.Close
                If WSaldo <> WTotal Then
                    m$ = "El anticipo no se puede eliminar debido a que ya a sido aplicado"
                    A% = MsgBox(m$, 0, "Baja de Ordenes de Pago")
                    Exit Sub
                End If
            End If
            
        End If
    
        T$ = "Orden de Pagos"
        m$ = "Desea Borrar la Orden de pago "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
                   
            For da = 1 To 99
            
                Auxi1 = Str$(da)
                Call Ceros(Auxi1, 2)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Pagos"
                ZSql = ZSql + " Where Pagos.Orden = " + "'" + Orden.Text + "'"
                ZSql = ZSql + " and Pagos.Renglon = " + "'" + Auxi1 + "'"
                spPagos = ZSql
                Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                If rstPagos.RecordCount > 0 Then
                
                    If rstPagos!Tiporeg = "1" Then
                    
                        ZZLetra = rstPagos!Letra1
                        ZZTipo = rstPagos!Tipo1
                        ZZPunto = rstPagos!Punto1
                        ZZNumero = rstPagos!Numero1
                        ZZImporte = Str$(rstPagos!Importe1)
                        ZZClaveCheque = rstPagos!ClaveCheque
                        ZZProveedor = Trim(rstPagos!Proveedor)
                        ZZTipo2 = rstPagos!Tipo2
                                            
                        If Tipo1.Value = True Then
                        
                            Claveven$ = ZZProveedor + ZZLetra + ZZTipo + ZZPunto + ZZNumero
                                    
                            ZSql = ""
                            ZSql = ZSql + "UPDATE CtaCtePrv SET "
                            ZSql = ZSql + " Saldo = Saldo + " + "'" + ZZImporte + "'"
                            ZSql = ZSql + " Where Clave = " + "'" + Claveven$ + "'"
                            spCtaCtePrv = ZSql
                            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                            
                        End If
                            
                                Else
                               
                        ZZTipo2 = rstPagos!Tipo2
                        ZZClaveCheque = rstPagos!ClaveCheque
                        If Val(ZZTipo2) = 3 Then
                        
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
                        
                    End If
                    
                End If
                    
            Next da
            
            ZSql = ""
            ZSql = ZSql + "DELETE Pagos"
            ZSql = ZSql + " Where Pagos.Orden = " + "'" + Orden.Text + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        
            If Tipo1.Value = True Then
        
                WLetra = "A"
                WTipo = "04"
                WPunto = "0000"
                WNumero = Orden.Text
                WProveedor = Trim(Proveedor.Text)
        
                Call Ceros(WNumero, 8)
                
                WClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                ZSql = ""
                ZSql = ZSql + "DELETE CtaCtePrv"
                ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + WClave + "'"
                spCtaCtePrv = ZSql
                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        
            If Tipo2.Value = True Then
        
                WLetra = "A"
                WTipo = "05"
                WPunto = "0000"
                WNumero = Orden.Text
                WProveedor = Trim(Proveedor.Text)
        
                Call Ceros(WNumero, 8)
        
                WClave = WProveedor + WLetra + WTipo + WPunto + WNumero
                ZSql = ""
                ZSql = ZSql + "DELETE CtaCtePrv"
                ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + WClave + "'"
                spCtaCtePrv = ZSql
                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                WTipoiva = Val(rstProveedor!Iva)
                rstProveedor.Close
            End If
        
            Call Calcula_Base_Retenciones
            
            Rem XBruto = Val(Debitos.Caption)
            Rem If WTipoiva = 2 Then
            Rem     XNeto = (XBruto / (1 + (ConfigIva1 / 100)))
            Rem         Else
            Rem     XNeto = XBruto
            Rem End If
            
            XBruto = Val(Debitos.Caption)
            XNeto = NetoTotal
        
            WFecha = Right$(Fecha.Text, 2) + Mid$(Fecha.Text, 4, 2)
            Auxi = Trim(Proveedor.Text)
            
            Claveven$ = WFecha + Auxi
            ZSql = ""
            ZSql = ZSql + "UPDATE Retencion SET "
            ZSql = ZSql + " Neto = Neto - " + "'" + Str$(XNeto) + "',"
            ZSql = ZSql + " Retenido = Retenido - " + "'" + Retencion.Text + "'"
            ZSql = ZSql + " Where Clave = " + "'" + Claveven$ + "'"
            spRetencion = ZSql
            Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
            
            Call CmdLimpiar_Click
            Orden.SetFocus
        
        End If
        
    End If
    
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector
    
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    WNroRet = 0
    WNroRet1 = 0
    
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    
    Debitos.Caption = ""
    Creditos.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    Reteiva.Text = ""
    PorceIva.Text = ""
    
    Solicitud.ListIndex = 2
    
    Orden.SetFocus
    
    Orden.Text = "1"
    ZSql = ""
    ZSql = ZSql + "Select Pagos.Orden"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Orden < " + "'100000'"
    ZSql = ZSql + " Order by Pagos.Orden"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        rstPagos.MoveLast
        ZUltimo = IIf(IsNull(rstPagos!Orden), "0", rstPagos!Orden)
        Orden.Text = ZUltimo + 1
        rstPagos.Close
    End If
    
    
    Pantalla.Visible = False
    Opcion.Visible = False
    
    Ingrecuenta.Visible = False
    IngreCuenta1.Visible = False
    Erase WCuenta
    Erase WCuenta1
    
End Sub

Private Sub CmdClose_Click()
    Prgpago.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Impresion_Click()

    T$ = "Impresion de Orden de Pago"
    m$ = "Desea realizar la impresion del comprobante"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call IMPREORDEN
    End If
    
    
    If Val(Retencion.Text) <> 0 Then
        T$ = "Impresion de Orden de Pago"
        m$ = "Desea realizar la impresion de la retencion"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            Call Impreret
        End If
    End If
    
End Sub

Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        Existe = "N"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pagos"
        ZSql = ZSql + " Where Pagos.Orden = " + "'" + Orden.Text + "'"
        spPagos = ZSql
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
            Existe = "S"
            WNroRet = rstPagos!Nroret
            WNroRet1 = rstPagos!Nroret1
            Proveedor.Text = Trim(rstPagos!Proveedor)
            Fecha.Text = rstPagos!Fecha
            Retencion.Text = Str$(rstPagos!Retencion)
            Retencion.Text = Pusing("###,###.##", Retencion.Text)
            Reteiva.Text = Str$(rstPagos!RetIva)
            Reteiva.Text = Pusing("###,###.##", Reteiva.Text)
            WNroRet = rstPagos!Nroret
            WNroRet1 = rstPagos!Nroret1
            Tipo1.Value = False
            Tipo2.Value = False
            Tipo3.Value = False
            Tipo4.Value = False
            Select Case Val(rstPagos!TipoOrd)
                Case 1
                    Tipo1.Value = True
                Case 2
                    Tipo2.Value = True
                Case 3
                    Tipo3.Value = True
                Case 4
                    Tipo4.Value = True
                Case Else
            End Select
            Observaciones.Text = rstPagos!Observaciones
            Solicitud.ListIndex = rstPagos!Solicitud
            rstPagos.Close
        End If
        
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Imprime_Datos
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Proveedor.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Orden.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(Fecha.Text)) = 8 Then
            Fecha.Text = Left$(Fecha.Text, 6) + "20" + Right$(Trim(Fecha.Text), 2)
        End If
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Solicitud_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Proveedor.SetFocus
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Proveedor.Text) <> "" Then
        
            Proveedor.Text = Trim(UCase(Proveedor.Text))
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                aa = rstProveedor!Ganancia
                Proveedor.Text = Trim(rstProveedor!Proveedor)
                DesProveedor.Caption = rstProveedor!Nombre
                WPrvDireccion = rstProveedor!Direccion
                WPrvCuit = rstProveedor!Cuit
                WTipoprv = rstProveedor!Ganancia
                WTipoiva = rstProveedor!Iva
                rstProveedor.Close
                Observaciones.SetFocus
            End If
            
                Else
                
            Observaciones.SetFocus
            
        End If
        
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Tipo4.Value = True Then
            Banco.SetFocus
                Else
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
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
                Banco.Text = rstBanco!Banco
                DesBanco.Caption = rstBanco!Nombre
                WCtabanco = rstBanco!Cuenta
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
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

Private Sub Consulta_Click()

     Pantalla.Visible = False
     Ayuda.Visible = False
    
     XRow = WVector1.Row
     XCol = WVector1.Col

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Banco"
     Opcion.AddItem "Cuenta Contables"

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
            Ayuda.Visible = True
            Ayuda.Text = ""
            
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
                            IngresaItem = !Proveedor + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
            Ayuda.SetFocus
            
        Case 1
            Ayuda.Visible = True
            Ayuda.Text = ""
            
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
            
            Ayuda.SetFocus
            
        Case 2
            Ayuda.Visible = True
            Ayuda.Text = ""
            
            If Cuenta.Visible = True Or Cuenta1.Visible = True Then
            
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
                                IngresaItem = !Cuenta + " " + !Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Cuenta
                                WIndice.AddItem IngresaItem
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCuenta.Close
                End If
            
            End If
            Ayuda.SetFocus
            
        Case 3
            If Trim(Proveedor.Text) <> "" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCtePrv"
                ZSql = ZSql + " Where CtaCtePrv.Proveedor = " + "'" + Proveedor.Text + "'"
                ZSql = ZSql + " Order by CtaCtePrv.OrdFecha"
                spCtaCtePrv = ZSql
                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCtePrv.RecordCount > 0 Then
                    With rstCtaCtePrv
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                ZSaldo = rstCtaCtePrv!Saldo
                                Call Redondeo(ZSaldo)
                                If ZSaldo <> 0 Then
                                    Auxi$ = Str$(ZSaldo)
                                    Auxi$ = Mascara("#,###,###.##", Auxi$)
                                    IngresaItem = rstCtaCtePrv!Impre + " " + rstCtaCtePrv!Letra + " " + rstCtaCtePrv!Punto + " " + rstCtaCtePrv!Numero + " " + rstCtaCtePrv!Fecha + " " + Auxi$
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstCtaCtePrv!Clave
                                    WIndice.AddItem IngresaItem
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCtaCtePrv.Close
                End If
            End If
     
        Case 4
            Ayuda.Visible = True
            Ayuda.Text = ""
            
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
                        
                            ZClaseCheque = IIf(IsNull(rstRecibos!ClaseCheque), "0", rstRecibos!ClaseCheque)
                            If Val(ZClaseCheque) <> 2 Then
                                Auxi$ = Str$(rstRecibos!Importe2)
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                Numero = Str$(Val(rstRecibos!Numero2))
                                Call Ceros(Numero, 6)
                                IngresaItem = Numero + "    " + rstRecibos!Fecha2 + "      " + Auxi$ + "      " + rstRecibos!Banco2
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstRecibos!Clave
                                WIndice.AddItem IngresaItem
                            End If
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
            Indice = Pantalla.ListIndex
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Banco.Text = WIndice.List(Indice)
            Call Banco_KeyPress(13)
            
        Case 2
            If Cuenta.Visible = True Then
                Indice = Pantalla.ListIndex
                Cuenta.Text = WIndice.List(Indice)
                Ayuda.Visible = False
                Pantalla.Visible = False
                Cuenta.SetFocus
            End If
            If Cuenta1.Visible = True Then
                Indice = Pantalla.ListIndex
                Cuenta1.Text = WIndice.List(Indice)
                Ayuda.Visible = False
                Pantalla.Visible = False
                Cuenta1.SetFocus
            End If
            
        Case 3
            If Tipo1.Value = True Then
                Entra = "S"
                Indice = Pantalla.ListIndex
                Compara1 = WIndice.List(Indice)
        
                For IRow = 1 To 50
                    WProveedor = Trim(Proveedor.Text)
                    Compara2 = WProveedor + WVector1.TextMatrix(IRow, 2)
                    Compara2 = Compara2 + WVector1.TextMatrix(IRow, 1)
                    Compara2 = Compara2 + WVector1.TextMatrix(IRow, 3)
                    Compara2 = Compara2 + WVector1.TextMatrix(IRow, 4)
                    If Compara1 = Compara2 Then
                        Entra = "N"
                        Exit For
                    End If
                Next IRow
            
                If Entra = "S" Then
            
                    For IRow = 1 To 50
                        If WVector1.TextMatrix(IRow, 1) = "" Then
                            XRow = WVector1.Row
                            Exit For
                        End If
                    Next IRow
                    
                    Indice = Pantalla.ListIndex
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM CtaCtePrv"
                    ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + WIndice.List(Indice) + "'"
                    spCtaCtePrv = ZSql
                    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCtaCtePrv.RecordCount > 0 Then
                
                        WVector1.Row = XRow
                            
                        WVector1.Col = 1
                        WVector1.Text = rstCtaCtePrv!Tipo
                    
                        WVector1.Col = 2
                        WVector1.Text = rstCtaCtePrv!Letra
                
                        WVector1.Col = 3
                        WVector1.Text = rstCtaCtePrv!Punto
                
                        WVector1.Col = 4
                        WVector1.Text = rstCtaCtePrv!Numero
                
                        WVector1.Col = 5
                        WVector1.Text = rstCtaCtePrv!Saldo
                        WVector1.Text = Pusing("#,###,###.##", WVector1.Text)
                    
                        WVector1.Col = 6
                        WVector1.Text = ""
                            
                        WVector1.Col = 1
                        WVector1.Text = rstCtaCtePrv!Tipo
                            
                        If rstCtaCtePrv!Letra = "X" Then
                            Solicitud.ListIndex = 1
                                Else
                            Solicitud.ListIndex = 2
                        End If
                        
                        rstCtaCtePrv.Close
                            
                        Call Suma_Datos
                            
                        WVector1.Col = 1
                        WVector1.Row = WVector1.Row + 1
                        Call StartEdit
                    
                    End If
            
                End If
            
            End If
                
        Case 4
            For IRow = 1 To 50
                If WVector1.TextMatrix(IRow, 7) = "" Then
                    XRow = WVector1.Row
                    Exit For
                End If
            Next IRow
            
            Indice = Pantalla.ListIndex
                        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.Clave = " + "'" + WIndice.List(Indice) + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                
                WVector1.Row = XRow
                        
                WVector1.Col = 7
                WVector1.Text = "03"
                    
                WVector1.Col = 8
                WVector1.Text = rstRecibos!Numero2
                
                WVector1.Col = 9
                WVector1.Text = rstRecibos!Fecha2
                
                WVector1.Col = 10
                WVector1.Text = ""
                    
                WVector1.Col = 11
                WVector1.Text = rstRecibos!Banco2
                
                WVector1.Col = 12
                WVector1.Text = rstRecibos!Importe2
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            
                WVector1.Col = 7
                WVector1.Text = "03"
                
                rstRecibos.Close
                    
                BajaCheque(WVector1.Row) = WIndice.List(Indice)
                            
                Call Suma_Datos
                        
                WVector1.Col = 7
                WVector1.Row = WVector1.Row + 1
                Call StartEdit
                
                Pantalla.List(Indice) = ""
                WIndice.List(Indice) = ""
                    
            End If
                
        Case Else
    End Select
    
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                rstCuenta.Close
                WCuenta(WVector1.Row) = Cuenta.Text
                Ingrecuenta.Visible = False
                WVector1.Col = WVector1.Col + 1
                Call StartEdit
                    Else
                Cuenta.SetFocus
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Cuenta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cuenta1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta1.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta1.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                rstCuenta.Close
                WCuenta1(WVector1.Row) = Cuenta1.Text
                IngreCuenta1.Visible = False
                If WVector1.Row < WVector1.Rows - 1 Then
                    WVector1.Row = WVector1.Row + 1
                End If
                WVector1.Col = 7
                Call StartEdit
                    Else
                Cuenta1.SetFocus
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Cuenta1.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Erase BajaCheque
    
    Solicitud.Clear
    
    Solicitud.AddItem ""
    Solicitud.AddItem "X"
    Solicitud.AddItem "Normal"
    
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    
    Debitos.Caption = ""
    Creditos.Caption = ""
    Diferencia.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    Reteiva.Text = ""
    PorceIva.Text = ""
    
    Solicitud.ListIndex = 2
    
    WNroRet = 0
    WNroRet1 = 0
    
    WLeyenda(1) = "Compra de Bienes"
    WLeyenda(2) = "Ejericio Prof. Lib. c/Aj.Inf."
    WLeyenda(3) = "Alquileres y Arrendamientos"
    
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
    
    PorceIva.Text = ConfigIva1
    PorceIva.Text = Pusing("###,###.##", PorceIva.Text)
    
    Orden.Text = "1"
    
    ZSql = ""
    ZSql = ZSql + "Select Pagos.Orden"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Orden < " + "'090000'"
    ZSql = ZSql + " Order by Pagos.Orden"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        rstPagos.MoveLast
        ZUltimo = IIf(IsNull(rstPagos!Orden), "0", rstPagos!Orden)
        Orden.Text = ZUltimo + 1
        rstPagos.Close
    End If
    
End Sub

Private Sub IMPREORDEN()

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
    
    PorceIva.Text = ConfigIva1
    PorceIva.Text = Pusing("###,###.##", PorceIva.Text)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Parametro"
    ZSql = ZSql + " Where Parametro.Clave = 1"
    spParametro = ZSql
    Set rstParametro = db.OpenRecordset(spParametro, dbOpenSnapshot, dbSQLPassThrough)
    If rstParametro.RecordCount > 0 Then
        WMinimo1 = rstParametro!Minimo1
        WMinimo2 = rstParametro!Minimo2
        WMinimo3 = rstParametro!Minimo3
        WMinimo4 = rstParametro!Minimo4
        WEscala1 = rstParametro!Escala1
        WEscala2 = rstParametro!Escala2
        WEscala3 = rstParametro!Escala3
        WEscala4 = rstParametro!Escala4
        WEscala5 = rstParametro!Escala5
        XTasa1 = rstParametro!Tasa1
        XTasa2 = rstParametro!Tasa2
        XTasa3 = rstParametro!Tasa3
        XTasa4 = rstParametro!Tasa4
        XTasa5 = rstParametro!Tasa5
        WRetMinima = rstParametro!RetMinima
        WPorceBienes = rstParametro!PorceBienes / 100
        WPorceServicios = rstParametro!PorceServicios / 100
        WPorceTranspo = rstParametro!PorceTranspo / 100
        WMinimoIva = rstParametro!MinimoIva
        WIvaInscripto = rstParametro!IvaInscripto
        WIvaNoInscripto = rstParametro!IvaNoInscripto
        WTasaGen = rstParametro!TasaGen / 100
        WTasaBienes = rstParametro!TasaBienes / 100
        WTasaNoInscripto = rstParametro!TasaNoInscripto / 100
        rstParametro.Close
    End If
    
    XPara(0) = 0
    XPara(1) = WEscala1
    XPara(2) = WEscala2
    XPara(3) = WEscala3
    XPara(4) = WEscala4
    XPara(5) = WEscala5
    
    WTasa1(1) = XTasa1 / 100
    WTasa1(2) = XTasa2 / 100
    WTasa1(3) = XTasa3 / 100
    WTasa1(4) = XTasa4 / 100
    WTasa1(5) = XTasa5 / 100
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpreOrd"
    spImpreOrd = ZSql
    Set rstImpreOrd = db.OpenRecordset(spImpreOrd, dbOpenSnapshot, dbSQLPassThrough)
    
    Impretit = WNombreEmpresa

    Cantidad = 0
    Total = 0
    SubTotaL = 0
        
    Erase WImpresion, WDebito, WCredito, WImpre2
        
    For IRow = 1 To 50
        WRow = IRow
        WVector1.Col = 5
        WVector1.Row = IRow
        If Val(WVector1.Text) <> 0 Then
            Cantidad = Cantidad + 1
            WVector1.Col = 1
            Select Case Val(Left$(WVector1.Text, 2))
                Case 1
                    WImpresion(Cantidad, 2) = "Factura"
                Case 2
                    WImpresion(Cantidad, 2) = "N.Debito"
                Case 3
                    WImpresion(Cantidad, 2) = "N.Credito"
                Case 99
                    WImpresion(Cantidad, 2) = "Varios"
                Case Else
                    WImpresion(Cantidad, 2) = ""
            End Select
                            
            WVector1.Col = 4
            WImpresion(Cantidad, 3) = Left$(WVector1.Text, 8)
            WVector1.Col = 6
            WImpresion(Cantidad, 4) = WVector1.Text
            WVector1.Col = 5
            WImpresion(Cantidad, 5) = WVector1.Text
            If Val(WImpresion(Cantidad, 2)) = 3 Or Val(WImpresion(Cantidad, 2)) = 5 Then
                Total = Total - Val(WImpresion(Cantidad, 5))
                    Else
                Total = Total + Val(WImpresion(Cantidad, 5))
            End If
                    
            WVector1.Col = 1
            WTipo = WVector1.Text
            WVector1.Col = 2
            WLetra = WVector1.Text
            WVector1.Col = 3
            WPunto = WVector1.Text
            WVector1.Col = 4
            WNumero = WVector1.Text
            
            
            
            WProveedor = Trim(Proveedor.Text)
                
            WVector1.Col = 1
            XTipo = WTipo
                    
            WVector1.Col = 2
            XLetra = WLetra
                
            WVector1.Col = 3
            XPunto = WPunto
                
            WVector1.Col = 4
            XNumero = WNumero

            WClave = WProveedor + XLetra + XTipo + XPunto + XNumero
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCtePrv"
            ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClave + "'"
            spCtaCtePrv = ZSql
            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCtePrv.RecordCount > 0 Then
                WImpresion(Cantidad, 1) = rstCtaCtePrv!Fecha
                rstCtaCtePrv.Close
            End If
                    
        End If
    Next IRow
        
    If Tipo1.Value = True Or Tipo2.Value = True Then
        WDebito(1, 1) = WCtaProveedor
        WDebito(1, 2) = Total
            Else
        For IRow = 0 To 9
            WRow = IRow
            WVector1.Col = 5
            WVector1.Row = IRow
            If Val(WVector1.Text) <> 0 Then
                WDebito(IRow + 1, 1) = WCuenta(IRow)
                WDebito(IRow + 1, 2) = Val(WVector1.Text)
            End If
        Next IRow
                    
    End If

    WCredito(1, 1) = WCtaProveedor
    If Retenido <> 0 Then
        WCredito(1, 2) = Retenido
    End If
        
    Lugar = 1
    Impre2 = 0
        
    For IRow = 0 To 9
        WVector1.Col = 12
        WVector1.Row = IRow
        If Val(WVector1.Text) <> 0 Then
            Lugar = Lugar + 1
            WCredito(Lugar, 4) = WVector1.Text
            WVector1.Col = 7
            Select Case Val(WVector1.Text)
                Case 2
                    WCredito(Lugar, 1) = "999999"
                    WVector1.Col = 10
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Banco"
                    ZSql = ZSql + " Where Banco.Banco = " + "'" + WVector1.Text + "'"
                    spBanco = ZSql
                    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                    If rstBanco.RecordCount > 0 Then
                        WCredito(Lugar, 1) = rstBanco!Cuenta
                        rstBanco.Close
                    End If
                Case 3, 4
                    WCredito(Lugar, 1) = WCtaCheques
                Case Else
                    WCredito(Lugar, 1) = WCtaEfectivo
            End Select
                    
            Impre2 = Impre2 + 1
            WVector1.Col = 8
            WImpre2(Impre2, 1) = WVector1.Text
            WVector1.Col = 11
            WImpre2(Impre2, 2) = WVector1.Text
            WVector1.Col = 12
            WImpre2(Impre2, 3) = WVector1.Text
                    
            WVector1.Col = 11
            WCredito(Lugar, 2) = WVector1.Text
            WVector1.Col = 8
            WCredito(Lugar, 3) = WVector1.Text
            WVector1.Col = 12
            WCredito(Lugar, 4) = WVector1.Text
        End If
    Next IRow
        
    SubTotaL = Total - Retenido
    TotalDebito = Total
    TotalCredito = Total
    
    WNombre = Observaciones.Text
    If Proveedor.Text <> "" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            WNombre = rstProveedor!Nombre
            rstProveedor.Close
        End If
    End If
    
    Renglon = 0
    
    
    For Ciclo = 1 To 50
        
        WVector1.Row = Ciclo
            
        If Tipo2.Value = True Then
            WTipo = "05"
                Else
            WVector1.Col = 1
            WTipo = WVector1.Text
        End If
        
        WVector1.Col = 2
        WLetra = WVector1.Text
            
        WVector1.Col = 3
        WPunto = WVector1.Text
            
        WVector1.Col = 4
        WNumero = WVector1.Text
                    
        WVector1.Col = 5
        WImporte = WVector1.Text
            
        WVector1.Col = 6
        WDescripcion = WVector1.Text
            
        WFecha = "  /  /    "
            
        Call Ceros(WTipo, 2)
            
        If Val(WImporte) <> 0 Then
        
            WProveedor = Trim(Proveedor.Text)
            XTipo = WTipo
            XLetra = WLetra
            XPunto = WPunto
            XNumero = WNumero
            WClave = WProveedor + XLetra + XTipo + XPunto + XNumero
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCtePrv"
            ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + WClave + "'"
            spCtaCtePrv = ZSql
            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCtePrv.RecordCount > 0 Then
                WFecha = rstCtaCtePrv!Fecha
                rstCtaCtePrv.Close
            End If
                    
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Orden.Text
            Call Ceros(Auxi1, 6)
                    
            ZZOrden = Orden.Text
            ZZRenglon = Str$(Renglon)
            ZZProveedor = Trim(Proveedor.Text)
            ZZfecha = Fecha.Text
            ZZTipoReg = "1"
            ZZTipo = WTipo
            ZZNumero = WNumero
            ZZFecha1 = WFecha
            ZZImporte = WImporte
            ZZDescripcion = WDescripcion
            ZZTotal = Creditos.Caption
            ZZRetencion = Retencion.Text
            ZZObservaciones = Observaciones.Text
            ZZDia = Left$(Fecha, 2)
            ZZMes = Mid$(Fecha, 4, 2)
            ZZAno = Right$(Fecha, 2)
            ZZNombre = WNombre
            ZZCuenta = ""
            
            ZZCuenta = ""
            ZDesCuenta = ""
                
            If Tipo1.Value = True Then
                
                ZProveedor = Trim(Proveedor.Text)
                ZTipo = WTipo
                ZLetra = WLetra
                ZPunto = WPunto
                ZNumero = WNumero
                ZCuenta = ""
                    
                Call Ceros(ZTipo, 2)
                Call Ceros(ZPunto, 4)
                Call Ceros(ZNumero, 8)
                    
                ZConcepto = ""
                WClave = ZProveedor + ZTipo + ZLetra + ZPunto + ZNumero
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM IvaComp"
                ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
                spIvaComp = ZSql
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    ZConcepto = Str$(rstIvaComp!Concepto)
                    rstIvaComp.Close
                End If
                    
                If Val(ZConcepto) <> 0 Then
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Conceptos"
                    ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + ZConcepto + "'"
                    spConceptos = ZSql
                    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstConceptos.RecordCount > 0 Then
                        ZCuenta = rstConceptos!Cuenta
                        rstConceptos.Close
                    End If
                    
                End If
                    
                ZZCuenta = ZCuenta
                    
            End If
                
            If Tipo3.Value = True Then
                ZZCuenta = WCuenta(IRow)
            End If
                
            If Trim(ZZCuenta) <> "" Then
                ZCuenta = ZZCuenta
                ZDesCuenta = ""
                ZZDescripcion = Left$(Trim(ZDesCuenta) + " - " + WDescripcion, 50)
            End If
                
            ZZRetib = Reteiva.Text
            ZZNroRet = Str$(WNroRet)
            ZZNroRet1 = "1"
                
            WProveedor = Trim(Proveedor.Text)
            NetoParcial = 0
            IvaParcial = 0
            
            If Val(WTipo) = 1 Or Val(WTipo) = 2 Or Val(WTipo) = 3 Then
            
                If WLetra <> "X" Then
                
                    WClave = WProveedor + WTipo + WLetra + WPunto + WNumero
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM IvaComp"
                    ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
                    spIvaComp = ZSql
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                        WTotalFactura = rstIvaComp!Neto + rstIvaComp!Exento + rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva27 + rstIvaComp!Iva105 + rstIvaComp!Ib + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible
                        If WTotalFactura <> 0 Then
                            WPorceFactura = WImporte / WTotalFactura
                                Else
                            WPorceFactura = 0
                        End If
                        NetoParcial = rstIvaComp!Neto * WPorceFactura
                        IvaParcial = (rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva105 + rstIvaComp!Iva27) * WPorceFactura
                        
                        rstIvaComp.Close
                    End If
                    
                    XBrutoIva = WImporte
                    XNetoIva = NetoParcial
                    XIvaIva = IvaParcial
                        
                    Call Redondeo(XBrutoIva)
                    Call Redondeo(XNetoIva)
                    Call Redondeo(XIvaIva)
                        
                    If WExepcion <> 0 Then
                        WSacaIva = XIvaIva * (WExepcion / 100)
                        XIvaIva = XIvaIva - WSacaIva
                        Call Redondeo(XIvaIva)
                    End If
                
                    XReteIva = XIvaIva * PorceRIva
                    Call Redondeo(XReteIva)
                            
                End If
            End If
                
            ZZTasa = Str$(PorceRIva * 100)
            ZZExepcion = Str$(WExepcion)
            ZZImpo1 = Str$(XBrutoIva)
            ZZImpo2 = Str$(XNetoIva)
            ZZImpo3 = Str$(XIvaIva)
            ZZImpo4 = Str$(XReteIva)
                
            If WLetra = "X" Or Val(WTipo) > 3 Then
                ZZImpo1 = "0"
                ZZImpo2 = "0"
                ZZImpo3 = "0"
                ZZImpo4 = "0"
            End If
            
            ZZCuenta = ZZNumero
                        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreOrd ("
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "TipoReg ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha1 ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Retencion ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Dia ,"
            ZSql = ZSql + "Mes ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Cuenta ,"
            ZSql = ZSql + "RetIb ,"
            ZSql = ZSql + "NroRet ,"
            ZSql = ZSql + "NroRet1 ,"
            ZSql = ZSql + "Tasa ,"
            ZSql = ZSql + "Impo1 ,"
            ZSql = ZSql + "Impo2 ,"
            ZSql = ZSql + "Exepcion ,"
            ZSql = ZSql + "Impo3 ,"
            ZSql = ZSql + "Impo4 ,"
            ZSql = ZSql + "Empresa ,"
            ZSql = ZSql + "NombreEmpresa ,"
            ZSql = ZSql + "DireccionEmpresa ,"
            ZSql = ZSql + "LocalidadEmpresa ,"
            ZSql = ZSql + "CuitEmpresa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZOrden + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZProveedor + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZTipoReg + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZFecha1 + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + ZZTotal + "',"
            ZSql = ZSql + "'" + ZZRetencion + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZDia + "',"
            ZSql = ZSql + "'" + ZZMes + "',"
            ZSql = ZSql + "'" + ZZAno + "',"
            ZSql = ZSql + "'" + ZZNombre + "',"
            ZSql = ZSql + "'" + ZZCuenta + "',"
            ZSql = ZSql + "'" + ZZRetib + "',"
            ZSql = ZSql + "'" + ZZNroRet + "',"
            ZSql = ZSql + "'" + ZZNroRet1 + "',"
            ZSql = ZSql + "'" + ZZTasa + "',"
            ZSql = ZSql + "'" + ZZImpo1 + "',"
            ZSql = ZSql + "'" + ZZImpo2 + "',"
            ZSql = ZSql + "'" + ZZExepcion + "',"
            ZSql = ZSql + "'" + ZZImpo3 + "',"
            ZSql = ZSql + "'" + ZZImpo4 + "',"
            ZSql = ZSql + "'" + WEmpresa + "',"
            ZSql = ZSql + "'" + WNombreEmpresa + "',"
            ZSql = ZSql + "'" + WDireccionEmpresa + "',"
            ZSql = ZSql + "'" + WLocalidadEmpresa + "',"
            ZSql = ZSql + "'" + WCuitEmpresa + "')"
                            
            spImpreOrd = ZSql
            Set rstImpreOrd = db.OpenRecordset(spImpreOrd, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
        
    For Ciclo = 1 To 50
        
        WVector1.Row = Ciclo
        
        WVector1.Col = 7
        WTipo = WVector1.Text
        
        WVector1.Col = 8
        WNumero = WVector1.Text
            
        WVector1.Col = 9
        WFecha = WVector1.Text
            
        WVector1.Col = 10
        WBanco = WVector1.Text
                    
        WVector1.Col = 11
        WDescripcion = WVector1.Text
            
        WVector1.Col = 12
        WImporte = WVector1.Text
            
        Call Ceros(WTipo, 2)
            
        If Val(WImporte) <> 0 Then
                    
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Orden.Text
            Call Ceros(Auxi1, 6)
                    
            ZZOrden = Orden.Text
            ZZRenglon = Str$(Renglon)
            ZZProveedor = Trim(Proveedor.Text)
            ZZfecha = Fecha.Text
            ZZTipoReg = "2"
            ZZTipo = WTipo
            ZZNumero = WNumero
            ZZFecha1 = WFecha
            ZZImporte = WImporte
            ZZDescripcion = WDescripcion
            ZZTotal = Creditos.Caption
            ZZRetencion = Retencion.Text
            ZZObservaciones = Observaciones.Text
            ZZDia = Left$(Fecha, 2)
            ZZMes = Mid$(Fecha, 4, 2)
            ZZAno = Right$(Fecha, 2)
            ZZNombre = WNombre
            Select Case Val(WTipo)
                Case 1
                    ZZCuenta = "04"
                Case 4
                    ZZCuenta = "03"
                Case Else
                    ZZCuenta = "01" + Right(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            End Select
                    
            ZZRetib = Reteiva.Text
            ZZNroRet = Str$(WNroRet)
            ZZNroRet1 = "1"
                        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpreOrd ("
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "TipoReg ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha1 ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Retencion ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Dia ,"
            ZSql = ZSql + "Mes ,"
            ZSql = ZSql + "Ano ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Cuenta ,"
            ZSql = ZSql + "RetIb ,"
            ZSql = ZSql + "NroRet ,"
            ZSql = ZSql + "NroRet1 ,"
            ZSql = ZSql + "Tasa ,"
            ZSql = ZSql + "Impo1 ,"
            ZSql = ZSql + "Impo2 ,"
            ZSql = ZSql + "Exepcion ,"
            ZSql = ZSql + "Impo3 ,"
            ZSql = ZSql + "Impo4 ,"
            ZSql = ZSql + "Empresa ,"
            ZSql = ZSql + "NombreEmpresa ,"
            ZSql = ZSql + "DireccionEmpresa ,"
            ZSql = ZSql + "LocalidadEmpresa ,"
            ZSql = ZSql + "CuitEmpresa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZOrden + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZProveedor + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZTipoReg + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZFecha1 + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + ZZTotal + "',"
            ZSql = ZSql + "'" + ZZRetencion + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZDia + "',"
            ZSql = ZSql + "'" + ZZMes + "',"
            ZSql = ZSql + "'" + ZZAno + "',"
            ZSql = ZSql + "'" + ZZNombre + "',"
            ZSql = ZSql + "'" + ZZCuenta + "',"
            ZSql = ZSql + "'" + ZZRetib + "',"
            ZSql = ZSql + "'" + ZZNroRet + "',"
            ZSql = ZSql + "'" + ZZNroRet1 + "',"
            ZSql = ZSql + "'" + ZZTasa + "',"
            ZSql = ZSql + "'" + ZZImpo1 + "',"
            ZSql = ZSql + "'" + ZZImpo2 + "',"
            ZSql = ZSql + "'" + ZZExepcion + "',"
            ZSql = ZSql + "'" + ZZImpo3 + "',"
            ZSql = ZSql + "'" + ZZImpo4 + "',"
            ZSql = ZSql + "'" + WEmpresa + "',"
            ZSql = ZSql + "'" + WNombreEmpresa + "',"
            ZSql = ZSql + "'" + WDireccionEmpresa + "',"
            ZSql = ZSql + "'" + WLocalidadEmpresa + "',"
            ZSql = ZSql + "'" + WCuitEmpresa + "')"
                            
            spImpreOrd = ZSql
            Set rstImpreOrd = db.OpenRecordset(spImpreOrd, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
            
    Next Ciclo
    
    
    
    
    
    
    
    Listado.WindowTitle = "Impresion de Ordenes de Pago"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT ImpreOrd.Orden, ImpreOrd.Proveedor, ImpreOrd.TipoReg, ImpreOrd.Tipo, ImpreOrd.Numero, ImpreOrd.Fecha1, ImpreOrd.Importe, ImpreOrd.Descripcion, ImpreOrd.Total, ImpreOrd.Retencion, ImpreOrd.Observaciones, ImpreOrd.Dia, ImpreOrd.Mes, ImpreOrd.Ano, ImpreOrd.Nombre, ImpreOrd.RetIb, ImpreOrd.NombreEmpresa, ImpreOrd.Empresa  " _
             + "From " _
             + DSQ + ".dbo.ImpreOrd ImpreOrd " _
             + "Where " _
             + "ImpreOrd.Orden >= 0 AND " _
             + "ImpreOrd.Orden <= 999999"
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Impreord.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.SelectionFormula = "{Impreord.Orden} in " + Orden.Text + " to " + Orden.Text
    
    Listado.ReportFileName = "Impreord.rpt"
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 1
    Listado.Action = 1
    Listado.CopiesToPrinter = 1

End Sub

Private Sub Impreret()

    Listado.WindowTitle = "Impresion de Retenciones de Ganancias"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT ImpreOrd.Orden, ImpreOrd.Fecha, ImpreOrd.TipoReg, ImpreOrd.Tipo, ImpreOrd.Fecha1, ImpreOrd.Total, ImpreOrd.Retencion, ImpreOrd.NroRet, ImpreOrd.NombreEmpresa, ImpreOrd.DireccionEmpresa, ImpreOrd.LocalidadEmpresa, ImpreOrd.CuitEmpresa, " _
            + "Proveedor.Nombre, Proveedor.Direccion, Proveedor.Cuit, Proveedor.Ganancia " _
            + "From " _
            + DSQ + ".dbo.ImpreOrd ImpreOrd, " _
            + DSQ + ".dbo.Proveedor Proveedor " _
            + "Where " _
            + "ImpreOrd.Proveedor = Proveedor.Proveedor AND " _
            + "ImpreOrd.Orden >= 0 AND " _
            + "ImpreOrd.Orden <= 999999"
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Impreord.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.SelectionFormula = "{Impreord.Orden} in " + Orden.Text + " to " + Orden.Text
    
    
    Listado.ReportFileName = "Impreretgan.rpt"

    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 2
    Listado.Action = 1
    Listado.CopiesToPrinter = 1

End Sub

Private Sub Impreretiva()

    Listado.WindowTitle = "Impresion de Retenciones de Iva"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Cuenta.Cuenta, Cuenta.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Cuenta Cuenta " _
                    + "Where " _
                    + "Cuenta.Cuenta >= " + Desde.Text + " AND " _
                    + "Cuenta.Cuenta <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Impreord.Orden} in " + Orden.Text + " to " + Orden.Text
    Listado.SelectionFormula = "{Impreord.Orden} in " + Orden.Text + " to " + Orden.Text
    
    Listado.ReportFileName = "Impreretiva.rpt"

    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.CopiesToPrinter = 2
    Listado.Action = 1
    Listado.CopiesToPrinter = 1

End Sub


Private Sub calcret_Click()

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
    
    
    PorceIva.Text = ConfigIva1
    PorceIva.Text = Pusing("###,###.##", PorceIva.Text)

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Parametro"
    ZSql = ZSql + " Where Parametro.Clave = 1"
    spParametro = ZSql
    Set rstParametro = db.OpenRecordset(spParametro, dbOpenSnapshot, dbSQLPassThrough)
    If rstParametro.RecordCount > 0 Then
        WMinimo1 = rstParametro!Minimo1
        WMinimo2 = rstParametro!Minimo2
        WMinimo3 = rstParametro!Minimo3
        WMinimo4 = rstParametro!Minimo4
        WEscala1 = rstParametro!Escala1
        WEscala2 = rstParametro!Escala2
        WEscala3 = rstParametro!Escala3
        WEscala4 = rstParametro!Escala4
        WEscala5 = rstParametro!Escala5
        XTasa1 = rstParametro!Tasa1
        XTasa2 = rstParametro!Tasa2
        XTasa3 = rstParametro!Tasa3
        XTasa4 = rstParametro!Tasa4
        XTasa5 = rstParametro!Tasa5
        WRetMinima = rstParametro!RetMinima
        WPorceBienes = rstParametro!PorceBienes / 100
        WPorceServicios = rstParametro!PorceServicios / 100
        WPorceTranspo = rstParametro!PorceTranspo / 100
        WMinimoIva = rstParametro!MinimoIva
        WIvaInscripto = rstParametro!IvaInscripto
        WIvaNoInscripto = rstParametro!IvaNoInscripto
        WTasaGen = rstParametro!TasaGen / 100
        WTasaBienes = rstParametro!TasaBienes / 100
        WTasaNoInscripto = rstParametro!TasaNoInscripto / 100
        rstParametro.Close
    End If
    
    XPara(0) = 0
    XPara(1) = WEscala1
    XPara(2) = WEscala2
    XPara(3) = WEscala3
    XPara(4) = WEscala4
    XPara(5) = WEscala5
    
    WTasa1(1) = XTasa1 / 100
    WTasa1(2) = XTasa2 / 100
    WTasa1(3) = XTasa3 / 100
    WTasa1(4) = XTasa4 / 100
    WTasa1(5) = XTasa5 / 100

    WRetencion = 0
    WReteIva = 0
    
    Call Calcula_Base_Retenciones
    
    Rem calculo de retencion de ganancias
    
    If Tipo1.Value = True Or Tipo2.Value = True Then
    
        If WTipoprv = 1 Or WTipoprv = 2 Or WTipoprv = 3 Or WTipoprv = 5 Then
        
            XBruto = Val(Debitos.Caption)
            Rem If WTipoiva = 3 And WTipoprv <> 3 Then
            Rem     XNeto = (XBruto / (1 + (Val(PorceIva.Text) / 100)))
            Rem         Else
            Rem      XNeto = XBruto
            Rem End If
            Rem XIva = XBruto - XNeto
            Rem XTBase = XNeto
            
            XNeto = NetoTotal
            XIva = IvaTotal
            XTBase = XNeto
            
            WFecha = Right$(Fecha.Text, 2) + Mid$(Fecha.Text, 4, 2)
            Auxi = Trim(Proveedor.Text)
            
            WClave = WFecha + Auxi
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Retencion"
            ZSql = ZSql + " Where Retencion.Clave = " + "'" + WClave + "'"
            spRetencion = ZSql
            Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
            If rstRetencion.RecordCount > 0 Then
            
                WNeto = rstRetencion!Neto
                WAnticipo = rstRetencion!Anticipo
                WBruto = rstRetencion!Bruto
                WIva = rstRetencion!Iva
                WRetenido = rstRetencion!Retenido
                rstRetencion.Close
                
                    Else
                    
                ZZfecha = WFecha
                ZZProveedor = Trim(Proveedor.Text)
                ZZNeto = "0"
                ZZAnticipo = "0"
                ZZBruto = "0"
                ZZIva = "0"
                ZZRetenido = "0"
                ZZClave = WFecha + Auxi
                    
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Retencion ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Proveedor ,"
                ZSql = ZSql + "Neto ,"
                ZSql = ZSql + "Retenido ,"
                ZSql = ZSql + "Anticipo ,"
                ZSql = ZSql + "Bruto ,"
                ZSql = ZSql + "Iva )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZZClave + "',"
                ZSql = ZSql + "'" + ZZfecha + "',"
                ZSql = ZSql + "'" + ZZProveedor + "',"
                ZSql = ZSql + "'" + ZZNeto + "',"
                ZSql = ZSql + "'" + ZZRetenido + "',"
                ZSql = ZSql + "'" + ZZAnticipo + "',"
                ZSql = ZSql + "'" + ZZBruto + "',"
                ZSql = ZSql + "'" + ZZIva + "')"
                            
                spRetencion = ZSql
                Set rstRetencion = db.OpenRecordset(spRetencion, dbOpenSnapshot, dbSQLPassThrough)
                    
                WNeto = 0
                WAnticipo = 0
                WBruto = 0
                WIva = 0
                WRetenido = 0
                
            End If
            
            Select Case WTipoprv
                Case 1
                    WMinimo = WMinimo1
                Case 2
                    WMinimo = WMinimo2
                Case 5
                    WMinimo = WMinimo4
                Case Else
                    WMinimo = WMinimo3
            End Select

            WAcupag = WNeto + XTBase
            WAuxi = WAcupag - WMinimo

            If WAuxi <= 0 Then
                WAuxi = 0
                WRetencion = 0
            End If
            
            WTasa = 0
            
            Select Case WTipoiva
                Case 1
                    Select Case WTipoprv
                        Case 2, 3
                            WTasa = WTasaNoInscripto
                        Case Else
                            WTasa = 0.1
                    End Select
                    WRetencion = WAuxi * WTasa
                Case Else
                    Select Case WTipoprv
                        Case 1, 5
                            WTasa = WTasaBienes
                            WRetencion = WAuxi * WTasa
                        Case 3
                            WTasa = WTasaGen
                            WRetencion = WAuxi * WTasa
                        Case 2
                            WRetencion = 0
                            WTope = 0
                            WTope1 = 0
                            
                            For da = 0 To 3
                                If WAuxi >= XPara(da) And WAuxi < XPara(da + 1) Then
                                    WTope1 = WAuxi
                                    WTope = XPara(da)
                                    WSum = WTope1 - WTope
                                    WSum = WSum * WTasa1(da + 1)
                                    WRetencion = WRetencion + WSum
                                End If
                                If WAuxi >= XPara(da + 1) Then
                                    WTope1 = XPara(da + 1)
                                    WTope = XPara(da)
                                    WSum = WTope1 - WTope
                                    WSum = WSum * WTasa1(da + 1)
                                    WRetencion = WRetencion + WSum
                                End If
                            Next da
                    End Select
            End Select

            WRetencion = WRetencion - WRetenido

            If WRetencion < WRetMinima Then
                WRetencion = 0
                        Else
                If WRetencion > XNeto Then
                        WRetencion = 0
                End If
            End If
                    
            Call Redondeo(WRetencion)
            Retencion.Text = WRetencion
            Retencion.Text = Pusing("###,###.##", Retencion.Text)
            
        End If
        
    End If
    
    Rem calculo de retencion de iva
    
    WReteIva = 0
    XReteIva = 0
    
    If ZPasa = 99 Then
    
    If Tipo1.Value = True Or Tipo2.Value = True Then
    
        If WTipoReteiva = 0 Or WTipoReteiva = 1 Or WTipoReteiva = 2 Or WTipoReteiva = 3 Then
        
            If WTipoiva = 3 Then
            
                Rem XBruto = 0
                Rem For IRow = 1 To 50
                Rem     If WVector1.TextMatrix(IRow, 2) <> "X" And Val(WVector1.TextMatrix(IRow, 1)) <= 3 Then
                Rem         XBruto = XBruto + Val(WVector1.TextMatrix(IRow, 5))
                Rem     End If
                Rem Next IRow
                Rem
                Rem XBruto = Val(Debitos.Caption)
                Rem XNeto = (XBruto / (1 + (Val(PorceIva.Text) / 100)))
                Rem XIva = XBruto - XNeto
                
                XBrutoIva = 0
                XNetoIva = 0
                XIvaIva = 0
                Erase WVectorIva
                
                For Ciclo = 1 To 50
                
                    If Tipo2.Value = True Then
                        WTipo = "05"
                            Else
                        WTipo = WVector1.TextMatrix(Ciclo, 1)
                        Call Ceros(WTipo, 2)
                    End If
                    
                    WLetra = WVector1.TextMatrix(Ciclo, 2)
                    WPunto = WVector1.TextMatrix(Ciclo, 3)
                    WNumero = WVector1.TextMatrix(Ciclo, 4)
                    WImporte = Val(WVector1.TextMatrix(Ciclo, 5))
            
                    If WImporte <> 0 Then
            
                        WProveedor = Trim(Proveedor.Text)
                        NetoParcial = 0
                        IvaParcial = 0
            
                        If Val(WTipo) = 1 Or Val(WTipo) = 2 Or Val(WTipo) = 3 Then
                        
                            If WLetra <> "X" Then
                            
                                WClave = WProveedor + WTipo + WLetra + WPunto + WNumero
                                ZSql = ""
                                ZSql = ZSql + "Select *"
                                ZSql = ZSql + " FROM IvaComp"
                                ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
                                spIvaComp = ZSql
                                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                                If rstIvaComp.RecordCount > 0 Then
                                    WTotalFactura = rstIvaComp!Neto + rstIvaComp!Exento + rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva27 + rstIvaComp!Iva105 + rstIvaComp!Ib + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible
                                    If WTotalFactura <> 0 Then
                                        WPorceFactura = WImporte / WTotalFactura
                                            Else
                                        WPorceFactura = 0
                                    End If
                                    NetoParcial = rstIvaComp!Neto * WPorceFactura
                                    IvaParcial = (rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva105 + rstIvaComp!Iva27) * WPorceFactura
                                    rstIvaComp.Close
                                End If
                    
                                XBrutoIva = XBrutoIva + WImporte
                                XNetoIva = XNetoIva + NetoParcial
                                XIvaIva = XIvaIva + IvaParcial
                                
                                WVectorIva(Ciclo, 1) = Str$(WImporte)
                                WVectorIva(Ciclo, 2) = Str$(NetoParcial)
                                WVectorIva(Ciclo, 3) = Str$(IvaParcial)
                                
                            End If
                        End If
                        
                    End If
                    
                Next Ciclo
                
                Call Redondeo(XBrutoIva)
                Call Redondeo(XNetoIva)
                Call Redondeo(XIvaIva)
                        
                If WExepcion <> 0 Then
                    WSacaIva = XIvaIva * (WExepcion / 100)
                    XIvaIva = XIvaIva - WSacaIva
                    Call Redondeo(XIvaIva)
                End If
                
                If XIvaIva > WMinimoIva Or WTipoReteiva = 3 Then
                
                    PorceRIva = 0
                
                    If XNetoIva < 10000 Then
                        PorceRIva = 1
                            Else
                        Select Case WTipoReteiva
                            Case 0
                                PorceRIva = WPorceBienes
                            Case 1
                                PorceRIva = WPorceServicios
                            Case 2
                                PorceRIva = WPorceTranspo
                            Case Else
                                PorceRIva = 1
                        End Select
                    End If
                        
                    XReteIva = XIvaIva * PorceRIva
                        
                End If
                    
                Call Redondeo(XReteIva)
                WReteIva = XReteIva
                
                Reteiva.Text = WReteIva
                Reteiva.Text = Pusing("###,###.##", Reteiva.Text)
                
            End If
            
        End If
        
    End If
    
    End If

End Sub


Private Sub Calcula_Base_Retenciones()

    NetoTotal = 0
    IvaTotal = 0
    FacturaTotal = 0

    For IRow = 1 To 50
        If Val(WVector1.TextMatrix(IRow, 5)) <> 0 Then
        
            WTipo = Left$(WVector1.TextMatrix(IRow, 1), 2)
            WLetra = Left$(WVector1.TextMatrix(IRow, 2), 1)
            WPunto = Left$(WVector1.TextMatrix(IRow, 3), 4)
            WNumero = Left$(WVector1.TextMatrix(IRow, 4), 8)
            WImporte = Val(WVector1.TextMatrix(IRow, 5))
            WProveedor = Trim(Proveedor.Text)
            
            Call Ceros(WTipo, 2)
            Call Ceros(WPunto, 4)
            Call Ceros(WNumero, 8)
            
            If Val(WTipo) = 1 Or Val(WTipo) = 2 Or Val(WTipo) = 3 Then
                If WLetra <> "X" Then
                
                    WClave = WProveedor + WTipo + WLetra + WPunto + WNumero
                        
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM IvaComp"
                    ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
                    spIvaComp = ZSql
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                        WTotalFactura = rstIvaComp!Neto + rstIvaComp!Exento + rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva27 + rstIvaComp!Iva105 + rstIvaComp!Ib + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible
                        If WTotalFactura <> 0 Then
                            WPorceFactura = WImporte / WTotalFactura
                                Else
                            WPorceFactura = 0
                        End If
                        NetoParcial = rstIvaComp!Neto * WPorceFactura
                        IvaParcial = (rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva105 + rstIvaComp!Iva27) * WPorceFactura
                        NetoTotal = NetoTotal + NetoParcial
                        IvaTotal = IvaTotal + IvaParcial
                        FacturaTotal = FacturaTotal + WImporte
                        rstIvaComp.Close
                    End If
                    
                End If
            End If
        End If
    Next IRow
    
    Call Redondeo(NetoTotal)
    Call Redondeo(IvaTotal)
    
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
                                IngresaItem = !Proveedor + " " + !Nombre
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Proveedor
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
                                IngresaItem = !Cuenta + " " + !Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Cuenta
                                WIndice.AddItem IngresaItem
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCuenta.Close
                End If
                
            Case 4
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
                                ZClaseCheque = IIf(IsNull(rstRecibos!ClaseCheque), "0", rstRecibos!ClaseCheque)
                                If Val(ZClaseCheque) <> 2 Then
                                    Auxi$ = Str$(rstRecibos!Importe2)
                                    Auxi$ = Mascara("###,###.##", Auxi$)
                                    Numero = Str$(Val(rstRecibos!Numero2))
                                    If Val(Numero) = Val(Ayuda.Text) Then
                                        Call Ceros(Numero, 8)
                                        IngresaItem = Numero + "    " + rstRecibos!Fecha2 + "      " + Auxi$ + "      " + rstRecibos!Banco2
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstRecibos!Clave
                                        WIndice.AddItem IngresaItem
                                    End If
                                End If
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
            WTexto3.MaxLength = WParametros(1, XColumna)
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

Private Sub Tipo1_Click()
    Banco.Text = ""
    DesBanco.Caption = ""
    LabelBanco.Visible = False
    Banco.Visible = False
    DesBanco.Visible = False
End Sub

Private Sub Tipo2_Click()
    Banco.Text = ""
    DesBanco.Caption = ""
    LabelBanco.Visible = False
    Banco.Visible = False
    DesBanco.Visible = False
End Sub

Private Sub Tipo3_Click()
    Banco.Text = ""
    DesBanco.Caption = ""
    LabelBanco.Visible = False
    Banco.Visible = False
    DesBanco.Visible = False
End Sub

Private Sub Tipo4_Click()
    LabelBanco.Visible = True
    Banco.Visible = True
    DesBanco.Visible = True
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,f5,f6,f9,F10
        Case 112, 113, 114, 115, 116, 117, 120, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl <> "X" Then
                If WControl = "S" Then
                    Call Control_Grilla
                End If
                Call StartEdit
            End If

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
            
        Rem F1,F2,F3,F4,f5,f6,f9,F10
        Case 112, 113, 114, 115, 116, 117, 120, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If Ingrecuenta.Visible = False And IngreCuenta1.Visible = False Then
                If WControl = "S" Then
                    Call Control_Grilla
                End If
                Call StartEdit
            End If
    
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
            
        Rem F1,F2,F3,F4,f5,f6,f9,F10
        Case 112, 113, 114, 115, 116, 117, 120, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
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
        Case 1, 2, 3, 4, 5, 7, 8, 9, 10, 11
            WVector1.Col = WVector1.Col + 1
        Case 6
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case 12
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 7
            
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()

    Call Suma_Datos
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    
    Select Case XColumna
        Case 1
            If WVector1.Text <> "" Then
                If Tipo1.Value = True Then
                    If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Then
                        Auxi$ = Str$(Val(WVector1.Text))
                        Call Ceros(Auxi$, 2)
                        WVector1.Text = Auxi$
                            Else
                        WControl = "N"
                    End If
                        Else
                    If Val(WVector1.Text) = 0 Then
                        Auxi$ = Str$(Val(WVector1.Text))
                        Call Ceros(Auxi$, 2)
                        WVector1.Text = Auxi$
                        WVector1.Col = WVector1.Col + 3
                            Else
                        WControl = "N"
                    End If
                End If
            End If
            
        Case 2
            If Tipo1.Value = True Then
                If WVector1.Text = "A" Or WVector1.Text = "C" Or WVector1.Text = "X" Or WVector1.Text = "Z" Then
                    Rem no hago anda
                        Else
                    WControl = "N"
                End If
            End If
            
        Case 3
            If Tipo1.Value = True Then
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 4)
                WVector1.Text = Auxi$
            End If
                
        Case 4
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
            If Tipo1.Value = True Then
            
                WProveedor = Trim(Proveedor.Text)
                WClave = WProveedor
                WVector1.Col = 2
                WClave = WClave + WVector1.Text
                WVector1.Col = 1
                WClave = WClave + WVector1.Text
                WVector1.Col = 3
                WClave = WClave + WVector1.Text
                WVector1.Col = 4
                WClave = WClave + WVector1.Text
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCtePrv"
                ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + WClave + "'"
                spCtaCtePrv = ZSql
                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCtePrv.RecordCount > 0 Then
                    WVector1.Col = 5
                    If Val(WVector1.Text) = 0 Then
                        WVector1.Text = rstCtaCtePrv!Saldo
                        WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        Call Suma_Datos
                    End If
                    WVector1.Col = 4
                    rstCtaCtePrv.Close
                        Else
                    WControl = "N"
                End If
                
            End If
                
        Case 5
            If Tipo1.Value = True Then
            
                WProveedor = Trim(Proveedor.Text)
                WClave = WProveedor
                WVector1.Col = 2
                WClave = WClave + WVector1.Text
                WVector1.Col = 1
                WClave = WClave + WVector1.Text
                WVector1.Col = 3
                WClave = WClave + WVector1.Text
                WVector1.Col = 4
                WClave = WClave + WVector1.Text
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCtePrv"
                ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + WClave + "'"
                spCtaCtePrv = ZSql
                Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCtePrv.RecordCount > 0 Then
                    Saldo = rstCtaCtePrv!Saldo
                    rstCtaCtePrv.Close
                        Else
                    Saldo = 0
                End If
                
                WVector1.Col = 5
                If Abs(Val(WVector1.Text)) > Abs(Saldo) Then
                    WVector1.Text = ""
                    WControl = "N"
                        Else
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                End If
                    Else
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
            End If
            
            If Tipo3.Value = True Then
                Cuenta.Text = WCuenta(WVector1.Row)
                Ingrecuenta.Visible = True
                Cuenta.SetFocus
            End If
            
            If Tipo4.Value = True Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Banco"
                ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
                spBanco = ZSql
                Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                If rstBanco.RecordCount > 0 Then
                    WCuenta(WVector1.Row) = rstBanco!Cuenta
                    rstBanco.Close
                        Else
                    WCuenta(WVector1.Row) = "999999"
                End If
            End If
     
        Case 7
            If WVector1.Text <> "" Then
                dada = Len(Trim(WVector1.Text))
                If Len(Trim(WVector1.Text)) = 29 Then
                
                    For IRow = 1 To 50
                        If WVector1.TextMatrix(IRow, 7) = "" Then
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
                                
                        WVector1.Col = 7
                        WVector1.Text = "03"
                            
                        WVector1.Col = 8
                        WVector1.Text = rstRecibos!Numero2
                        
                        WVector1.Col = 9
                        WVector1.Text = rstRecibos!Fecha2
                        
                        WVector1.Col = 10
                        WVector1.Text = ""
                            
                        WVector1.Col = 11
                        WVector1.Text = rstRecibos!Banco2
                        
                        WVector1.Col = 12
                        WVector1.Text = rstRecibos!Importe2
                        WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        
                        WVector1.Col = 13
                        WVector1.Text = rstRecibos!ClaveLectora
                                    
                        WVector1.Col = 7
                        WVector1.Text = "03"
                            
                        BajaCheque(WVector1.Row) = rstRecibos!Clave
                        
                        rstRecibos.Close
                                    
                        Call Suma_Datos
                                
                        WVector1.Col = 7
                        WVector1.Row = WVector1.Row + 1
                        WControl = "N"
                        Call StartEdit
                            
                    End If
                    
                        Else
                
                    If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Or Val(WVector1.Text) = 5 Then
                        Auxi$ = Str$(Val(WVector1.Text))
                        Call Ceros(Auxi$, 2)
                        WVector1.Text = Auxi$
                                
                        Select Case Val(WVector1.Text)
                            Case 1
                                WVector1.Col = 8
                                WVector1.Text = ""
                                WVector1.Col = 9
                                WVector1.Text = ""
                                WVector1.Col = 10
                                WVector1.Text = ""
                                WVector1.Col = 11
                                WVector1.Text = "Efectivo"
                            Case 4
                                WVector1.Col = 8
                                WVector1.Text = ""
                                WVector1.Col = 9
                                WVector1.Text = ""
                                WVector1.Col = 10
                                WVector1.Text = ""
                                WVector1.Col = 11
                                WVector1.Text = "Comp"
                            Case 5
                                WVector1.Col = 8
                                WVector1.Text = ""
                                WVector1.Col = 9
                                WVector1.Text = ""
                                WVector1.Col = 10
                                WVector1.Text = ""
                                WVector1.Col = 11
                                WVector1.Text = "Caja"
                                
                            Case Else
                        End Select
                                
                            Else
                                    
                        WControl = "N"
                                
                    End If
                    
                End If
            End If
                
        Case 8
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
                
        Case 9
        If Len(Trim(WTexto3.Text)) = 8 Then
            WTexto3.Text = Left$(WTexto3.Text, 6) + "20" + Right$(Trim(WTexto3.Text), 2)
        End If
            Call Valida_fecha1(WTexto3.Text, Auxi)
            WControl = "S"
            If Auxi <> "S" Then
                 WControl = "N"
            End If
                
        Case 10
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Banco = " + "'" + WVector1.Text + "'"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                WVector1.Col = 11
                WVector1.Text = rstBanco!Nombre
                rstBanco.Close
                    Else
                WControl = "N"
            End If

        Case 12
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            Call Suma_Datos
            Rem If Val(WVector1.TextMatrix(WVector1.Row, 7)) = 4 Then
            Rem     Cuenta1.Text = WCuenta1(WVector1.Row)
            Rem     IngreCuenta1.Visible = True
            Rem     Cuenta1.SetFocus
            Rem End If
            
        Case Else
    End Select
End Sub

Private Sub WVector1_DblClick()

    Exit Sub

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
    WVector1.Cols = 14
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
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Letra"
                WVector1.ColWidth(Ciclo) = 550
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Punto"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 6
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 1150
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 12
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 13
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case Else
                
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

Private Sub Proveedor_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corrientes"
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Banco_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corrientes"
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Cuenta_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corrientes"
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub CtaCte_Click()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corrientes"
    Opcion.ListIndex = 3
    
    Call Opcion_Click

End Sub

Private Sub Cheque_Click()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corrientes"
    Opcion.AddItem "Cheques de Terceros"
    Opcion.ListIndex = 4
    
    Call Opcion_Click

End Sub

Private Sub Cuenta1_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corrientes"
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Orden_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Banco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub



Private Sub Retencion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ReteIva_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Cuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cuenta1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Solicitud_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call CmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call CtaCte_Click
        Case 117
            Call Cheque_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub

Private Sub WTexto1_DblClick()
    If WVector1.Col = 1 Then
        WTexto1.Text = ""
        WVector1.TextMatrix(WVector1.Row, 1) = ""
        WVector1.TextMatrix(WVector1.Row, 2) = ""
        WVector1.TextMatrix(WVector1.Row, 3) = ""
        WVector1.TextMatrix(WVector1.Row, 4) = ""
        WVector1.TextMatrix(WVector1.Row, 5) = ""
        WVector1.TextMatrix(WVector1.Row, 6) = ""
    End If
    If WVector1.Col = 7 Then
        WTexto1.Text = ""
        WVector1.TextMatrix(WVector1.Row, 7) = ""
        WVector1.TextMatrix(WVector1.Row, 8) = ""
        WVector1.TextMatrix(WVector1.Row, 9) = ""
        WVector1.TextMatrix(WVector1.Row, 10) = ""
        WVector1.TextMatrix(WVector1.Row, 11) = ""
        WVector1.TextMatrix(WVector1.Row, 12) = ""
    End If
End Sub

Private Sub WTexto2_DblClick()
    If WVector1.Col = 1 Then
        WTexto1.Text = ""
        WVector1.TextMatrix(WVector1.Row, 1) = ""
        WVector1.TextMatrix(WVector1.Row, 2) = ""
        WVector1.TextMatrix(WVector1.Row, 3) = ""
        WVector1.TextMatrix(WVector1.Row, 4) = ""
        WVector1.TextMatrix(WVector1.Row, 5) = ""
        WVector1.TextMatrix(WVector1.Row, 6) = ""
    End If
    If WVector1.Col = 7 Then
        WTexto1.Text = ""
        WVector1.TextMatrix(WVector1.Row, 7) = ""
        WVector1.TextMatrix(WVector1.Row, 8) = ""
        WVector1.TextMatrix(WVector1.Row, 9) = ""
        WVector1.TextMatrix(WVector1.Row, 10) = ""
        WVector1.TextMatrix(WVector1.Row, 11) = ""
        WVector1.TextMatrix(WVector1.Row, 12) = ""
    End If
End Sub







