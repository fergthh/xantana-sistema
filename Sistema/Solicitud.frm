VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSolicitud 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de  Solicitud  de O.P"
   ClientHeight    =   7845
   ClientLeft      =   60
   ClientTop       =   585
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7845
   ScaleWidth      =   11880
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
      Left            =   8040
      MouseIcon       =   "Solicitud.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Cuenta Corriente de Proveedores"
      Top             =   2160
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
      Left            =   9000
      MouseIcon       =   "Solicitud.frx":0BD4
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":0EDE
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Cartera de Cheques"
      Top             =   2160
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
      Left            =   9960
      MouseIcon       =   "Solicitud.frx":136D
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":1677
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Impresion de Orden de Pago"
      Top             =   2160
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
      Left            =   4200
      MouseIcon       =   "Solicitud.frx":1EB9
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":21C3
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2160
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
      Left            =   5160
      MouseIcon       =   "Solicitud.frx":2A05
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":2D0F
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Elimina el Registro"
      Top             =   2160
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
      Left            =   6120
      MouseIcon       =   "Solicitud.frx":3551
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":385B
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2160
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
      Left            =   7080
      MouseIcon       =   "Solicitud.frx":409D
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":43A7
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Consulta de Datos"
      Top             =   2160
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
      Left            =   10920
      MouseIcon       =   "Solicitud.frx":4BE9
      MousePointer    =   99  'Custom
      Picture         =   "Solicitud.frx":4EF3
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Menu Principal"
      Top             =   2160
      Width           =   855
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
      Left            =   6120
      TabIndex        =   54
      Text            =   " "
      Top             =   1800
      Width           =   975
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
      Left            =   8640
      TabIndex        =   53
      Text            =   " "
      Top             =   1800
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
      Left            =   4800
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
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
         TabIndex        =   49
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label7 
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
         TabIndex        =   50
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
      Left            =   3240
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
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
         TabIndex        =   22
         Text            =   " "
         Top             =   360
         Width           =   1695
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
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2640
      TabIndex        =   47
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   46
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
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   45
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
      Index           =   10
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   44
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
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   43
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
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   42
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   41
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
      Index           =   6
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   40
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
      Index           =   5
      Left            =   3240
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
      Index           =   4
      Left            =   2640
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
      Left            =   2040
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
      Index           =   2
      Left            =   1440
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
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   4920
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
      TabIndex        =   33
      Top             =   4320
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1440
      TabIndex        =   32
      Top             =   4320
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
      Left            =   840
      TabIndex        =   31
      Top             =   4320
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
      Left            =   7200
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
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
      Left            =   6120
      TabIndex        =   28
      Text            =   " "
      Top             =   1440
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
      TabIndex        =   26
      Text            =   " "
      Top             =   1080
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
      TabIndex        =   18
      Text            =   " "
      Top             =   720
      Width           =   5415
   End
   Begin Crystal.CrystalReport LISTADO 
      Left            =   5400
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ordpago.rpt"
      WindowTitle     =   "Orden de Pago"
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
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   1560
      Width           =   4095
      Begin VB.OptionButton Tipo5 
         Caption         =   "Cheques Rechazados"
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
         TabIndex        =   24
         Top             =   960
         Width           =   2295
      End
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
         Left            =   2280
         TabIndex        =   23
         Top             =   600
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
         TabIndex        =   19
         Top             =   600
         Width           =   2055
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
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.TextBox Proveedor 
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
      MaxLength       =   11
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   735
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
      Left            =   7200
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
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
      Top             =   0
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5640
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
      ItemData        =   "Solicitud.frx":5735
      Left            =   7200
      List            =   "Solicitud.frx":573C
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   3615
      Left            =   0
      TabIndex        =   37
      Top             =   3600
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   6376
      _Version        =   327680
      BackColor       =   16777152
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
      Left            =   4680
      TabIndex        =   56
      Top             =   1800
      Width           =   1575
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
      Left            =   7200
      TabIndex        =   55
      Top             =   1800
      Width           =   1575
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
      TabIndex        =   52
      Top             =   3240
      Width           =   5535
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
      Top             =   3240
      Width           =   5535
   End
   Begin VB.Label Label6 
      Caption         =   "Retencion"
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
      Left            =   4680
      TabIndex        =   30
      Top             =   1440
      Width           =   1335
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
      TabIndex        =   27
      Top             =   1080
      Width           =   4575
   End
   Begin VB.Label bjm 
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
      TabIndex        =   25
      Top             =   1080
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
      TabIndex        =   17
      Top             =   720
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
      TabIndex        =   16
      Top             =   360
      Width           =   1455
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
      TabIndex        =   15
      Top             =   7440
      Width           =   1095
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
      TabIndex        =   14
      Top             =   7440
      Width           =   1095
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Forma de Pago : 1) Efectivo   2) Banco Propio   4) Varios"
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
      Left            =   5040
      TabIndex        =   13
      Top             =   7440
      Width           =   5175
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
      Left            =   2520
      TabIndex        =   9
      Top             =   360
      Width           =   4575
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
      Left            =   2520
      TabIndex        =   7
      Top             =   0
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
      Caption         =   "Nro Solicitud"
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
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "PrgSolicitud"
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
Private Numero As String
Private WNumero As String
Private WSaldo As Double
Private WRetencion As Double
Private WReteIva  As Double
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
Private AuxiFecha As String
Private WProveedor As String
Private WTipocta As Integer
Dim BajaCheque(100) As String
Private WCtaChequeRecha As String
Dim WMinimo1 As Double
Dim WMinimo2 As Double
Dim WMinimo3 As Double
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
    
    Call calcret_Click
    
    Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Retencion.Text) + Val(Reteiva.Text))
    
    Debitos.Caption = Pusing("###,###.##", Debitos.Caption)
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
    
End Sub

Private Sub Lee_Datos()

    Renglon = 0
    Debito = 0
    Credito = 0
    
    Erase BajaCheque
    Erase WCuenta
    Erase WCuenta1
    
    Do
        With rstSolicitud
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", Orden.Text + Auxi1
            If .NoMatch = False Then
                Select Case Val(!Tiporeg)
                    Case 1
                        Debito = Debito + 1
                        WVector1.Row = Debito
                        WVector1.Col = 1
                        WVector1.Text = !Tipo1
                        WVector1.Col = 2
                        WVector1.Text = !Letra1
                        WVector1.Col = 3
                        WVector1.Text = !Punto1
                        WVector1.Col = 4
                        WVector1.Text = !Numero1
                        WVector1.Col = 5
                        WVector1.Text = !Importe1
                        WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        WVector1.Col = 6
                        WVector1.Text = !Observaciones2
                        WCuenta(WVector1.Row) = !Cuenta
                    Case 2
                        Credito = Credito + 1
                        WVector1.Row = Credito
                        WVector1.Col = 7
                        WVector1.Text = !Tipo2
                        WVector1.Col = 8
                        WVector1.Text = !Numero2
                        WVector1.Col = 9
                        WVector1.Text = !Fecha2
                        WVector1.Col = 10
                        WVector1.Text = !Banco2
                        WVector1.Col = 11
                        If !Observaciones2 <> "" Then
                            WVector1.Text = !Observaciones2
                        End If
                        WVector1.Col = 12
                        WVector1.Text = !Importe2
                        WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        BajaCheque(WVector1.Row) = !ClaveCheque
                        WCuenta1(WVector1.Row) = !Cuenta
                    Case Else
                End Select
                    Else
                Exit Do
            End If
        End With
    Loop
End Sub

Sub Imprime_Datos()

    With rstProveedor
        .Index = "Proveedor"
        .Seek "=", Proveedor.Text
        If .NoMatch = False Then
            DesProveedor.Caption = !Nombre
            WPrvDireccion = !Direccion
            WPrvCuit = !Cuit
            WTipoprv = !Ganancia
            WTipoiva = !Iva
            WTipoReteiva = !Reteiva
            WExepcion = !PorceReteIva
                Else
            DesProveedor.Caption = ""
            WPrvDireccion = ""
            WPrvCuit = ""
            WTipoprv = 0
            WTipoiva = 0
            WTipoReteiva = 0
            WExepcion = 0
        End If
    End With
    With rstBanco
        .Index = "Banco"
        .Seek "=", Banco.Text
        If .NoMatch = False Then
            DesBanco.Caption = !Nombre
                Else
            DesBanco.Caption = ""
        End If
    End With
        
End Sub

Private Sub cmdAdd_Click()

    If Orden.Text <> "" Then
    
        If Proveedor.Text <> "" Or Tipo3.Value = True Or Tipo4.Value = True Or Tipo5.Value = True Then
    
            Auxi1 = Orden.Text
            Call Ceros(Auxi1, 6)
            Orden.Text = Auxi1
        
            With rstSolicitud
                Existe = "N"
                .Index = "Clave"
                Claveven$ = Orden.Text + "01"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Existe = "S"
                End If
            End With
    
            If Existe <> "S" Then
    
                Call Suma_Datos
        
                Debito = 0
                Credito = 0
        
                If Debito = Credito Then
        
                    If Val(Proveedor.Text) = 0 Then
                        Proveedor.Text = "0"
                    End If
                    If Val(Banco.Text) = 0 Then
                        Banco.Text = "0"
                    End If
    
                    With rstSolicitud
                
                        Renglon = 0
                        .Index = "Clave"
                        For IRow = 1 To 50
                            WRow = IRow
                            WVector1.Col = 5
                            WVector1.Row = IRow
                            If Val(WVector1.Text) <> 0 Then
                                .AddNew
                                Renglon = Renglon + 1
                                Auxi1 = Str$(Renglon)
                                Call Ceros(Auxi1, 2)
                                !Orden = Orden.Text
                                !Renglon = Auxi1
                                !Proveedor = Val(Proveedor.Text)
                                !Fecha = Fecha.Text
                                !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                                !Importe = Val(Debito)
                                !Retencion = Val(Retencion.Text)
                                !RetIva = Val(Reteiva.Text)
                                !Observaciones = Observaciones.Text
                                !Cuenta = ""
                                If Tipo1.Value = True Then
                                    !TipoOrd = "1"
                                End If
                                If Tipo2.Value = True Then
                                    !TipoOrd = "2"
                                End If
                                If Tipo3.Value = True Then
                                    !TipoOrd = "3"
                                    !Cuenta = WCuenta(IRow)
                                End If
                                If Tipo4.Value = True Then
                                    !TipoOrd = "4"
                                    !Cuenta = WCuenta(IRow)
                                End If
                                If Tipo5.Value = True Then
                                    !TipoOrd = "5"
                                    !Cuenta = WCuenta(IRow)
                                End If
                    
                                !Tiporeg = "1"
                                WVector1.Col = 1
                                !Tipo1 = Left$(WVector1.Text, 2)
                                WVector1.Col = 2
                                !Letra1 = Left$(WVector1.Text, 1)
                                WVector1.Col = 3
                                !Punto1 = Left$(WVector1.Text, 4)
                                WVector1.Col = 4
                                !Numero1 = Left$(WVector1.Text, 8)
                                WVector1.Col = 5
                                !Importe1 = Val(WVector1.Text)
                                WVector1.Col = 6
                                !Observaciones2 = Left$(WVector1.Text, 30)
                                !Tipo2 = ""
                                !Numero2 = ""
                                !Fecha2 = ""
                                !FechaOrd2 = ""
                                If Tipo4.Value = True Then
                                    !Banco2 = Val(Banco.Text)
                                        Else
                                    !Banco2 = 0
                                End If
                                !Importe2 = 0
                                !Clave = !Orden + !Renglon
                                !ClaveCheque = ""
                                !NroOrden = 0
                                
                                .Update
                                .Bookmark = .LastModified
                    
                                WLetra = !Letra1
                                WTipo = !Tipo1
                                WPunto = !Punto1
                                WNumero = !Numero1
                                WImporte = !Importe1
                    
                            End If
                
                            WVector1.Col = 12
                            If Val(WVector1.Text) <> 0 Then
                                .AddNew
                                Renglon = Renglon + 1
                                Auxi1 = Str$(Renglon)
                                Call Ceros(Auxi1, 2)
                                !Orden = Orden.Text
                                !Renglon = Auxi1
                                !Proveedor = Val(Proveedor.Text)
                                !Fecha = Fecha.Text
                                !fechaord = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                                !Importe = Val(Debito)
                                !Retencion = Val(Retencion.Text)
                                !RetIva = Val(Reteiva.Text)
                                !Observaciones = Observaciones.Text
                                If Tipo1.Value = True Then
                                    !TipoOrd = "1"
                                End If
                                If Tipo2.Value = True Then
                                    !TipoOrd = "2"
                                End If
                                If Tipo3.Value = True Then
                                    !TipoOrd = "3"
                                End If
                                If Tipo4.Value = True Then
                                    !TipoOrd = "4"
                                End If
                                If Tipo5.Value = True Then
                                    !TipoOrd = "5"
                                End If
                                !Tiporeg = "2"
                                !Tipo1 = ""
                                !Letra1 = ""
                                !Punto1 = ""
                                !Numero1 = ""
                                !Importe1 = 0
                                WVector1.Col = 7
                                !Tipo2 = Left$(WVector1.Text, 2)
                                WVector1.Col = 8
                                !Numero2 = Left$(WVector1.Text, 8)
                                WVector1.Col = 9
                                !Fecha2 = Left$(WVector1.Text, 10)
                                !FechaOrd2 = Right$(!Fecha2, 4) + Mid$(!Fecha2, 4, 2) + Left$(!Fecha2, 2)
                                WVector1.Col = 10
                                !Banco2 = Val(WVector1.Text)
                                WVector1.Col = 11
                                !Observaciones2 = Left$(WVector1.Text, 20)
                                WVector1.Col = 12
                                !Importe2 = Val(WVector1.Text)
                                If !Tipo2 = 3 Then
                                    !ClaveCheque = BajaCheque(WVector1.Row)
                                        Else
                                    !ClaveCheque = ""
                                End If
                                If Val(!Tipo2) = 4 Then
                                    !Cuenta = WCuenta1(IRow)
                                        Else
                                    !Cuenta = ""
                                End If
                                !Clave = !Orden + !Renglon
                                !NroOrden = 0
                                
                                .Update
                                .Bookmark = .LastModified
                    
                            End If
                
                        Next IRow
                        
                    End With
        
                    With rstEmpresa
                        .Index = "Empresa"
                        Claveven$ = Val(WEmpresa)
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            WCtaProveedor = !CtaProveedores
                            WCtaEfectivo = !CtaEfectivo
                            WCtaCheques = !CtaCheque
                            WCtaChequeRecha = !CtaChequeRecha
                            WAuxiliar = !Nombre
                        End If
                    End With
        
                    With rstAuxiliar
                        .Index = "Clave"
                        .Seek "=", 1
                        If .NoMatch = False Then
                            .Edit
                            !Nombre = WAuxiliar
                            .Update
                        End If
                    End With
        
                    Rem LISTADO.GroupSelectionFormula = "{Pagos.Orden} in " + Chr$(34) + Orden.Text + Chr$(34) + " to " + Chr$(34) + Orden.Text + Chr$(34)
                    Rem LISTADO.Destination = 1
                    Rem LISTADO.Action = 1
        
                    Call IMPREORDEN
                    Rem If Val(Retencion.Text) <> 0 Then
                    Rem     Call Impreret
                    Rem End If

                    Orden.SetFocus
                    Call CmdLimpiar_Click
        
                End If
        
            End If
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Orden.Text <> "" Then
    
        T$ = "Solicitud de Orden de Pagos"
        m$ = "Desea Borrar la Solicitud de Orden de pago "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
    
            With rstSolicitud
                For da = 1 To 99
                    Auxi1 = Str$(da)
                    Call Ceros(Auxi1, 2)
                    .Index = "Clave"
                    .Seek "=", Orden.Text + Auxi1
                    If .NoMatch = False Then
                        .Delete
                    End If
                Next da
            End With
            
            Call CmdLimpiar_Click
        
        End If
        
    End If
    
    
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector
    Erase BajaCheque

    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    Reteiva.Text = ""
    PorceIva.Text = ""
    
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    
    Debitos.Caption = ""
    Creditos.Caption = ""
    
    With rstConfiguracion
       .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            ConfigIva1 = Str$(!Iva1)
            ConfigIva2 = Str$(!Iva2)
            ConfigPercepcion = Str$(!Percepcion)
            ConfigPunto = !Punto
        End If
    End With
    
    PorceIva.Text = ConfigIva1
    PorceIva.Text = Pusing("###,###.##", PorceIva.Text)
    
    Orden.SetFocus
    
    With rstSolicitud
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Orden.Text = Mid$(Str$(!Orden + 1), 2, 6)
                Else
            Orden.Text = "1"
        End If
    End With
    
    Pantalla.Visible = False
    Opcion.Visible = False
    
    Ingrecuenta.Visible = False
    IngreCuenta1.Visible = False
    Erase WCuenta
    Erase WCuenta1

End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    
    With rstEmpresa
        .Close
    End With
    With rstProveedor
        .Close
    End With
    With rstSolicitud
        .Close
    End With
    With rstCtaCtePrv
        .Close
    End With
    With rstBanco
        .Close
    End With
    With rstCuenta
        .Close
    End With
    With rstRecibos
        .Close
    End With
    With rstImpreOrd
        .Close
    End With
    With rstIvaComp
        .Close
    End With
    
    DbsAdminis.Close
    Orden.SetFocus
    PrgSolicitud.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Cuenta
    OPEN_FILE_Proveedor
    OPEN_FILE_CtaCtePrv
    OPEN_FILE_Banco
    OPEN_FILE_Retencion
    OPEN_FILE_Solicitud
    OPEN_FILE_Recibos
    OPEN_FILE_ImpreOrd
    OPEN_FILE_Configuracion
    OPEN_FILE_Parametro
    OPEN_FILE_Ivacomp
End Sub

Private Sub Impresion_Click()
    Call IMPREORDEN
    Rem If Val(Retencion.Text) <> 0 Then
    Rem     Call Impreret
    Rem End If
    Rem Listado.GroupSelectionFormula = "{Pagos.Orden} in " + Chr$(34) + Orden.Text + Chr$(34) + " to " + Chr$(34) + Orden.Text + Chr$(34)
    Rem Listado.Destination = 1
    Rem Listado.Action = 1
End Sub

Private Sub Orden_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Auxi1 = Orden.Text
        Call Ceros(Auxi1, 6)
        Orden.Text = Auxi1
        
        With rstSolicitud
            Existe = "N"
            .Index = "Clave"
            Claveven$ = Orden.Text + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Existe = "S"
                Proveedor.Text = !Proveedor
                Fecha.Text = !Fecha
                Retencion.Text = !Retencion
                Retencion.Text = Pusing("###,###.##", Retencion.Text)
                Reteiva.Text = Str$(!RetIva)
                Reteiva.Text = Pusing("###,###.##", Reteiva.Text)
                Tipo1.Value = False
                Tipo2.Value = False
                Tipo3.Value = False
                Tipo4.Value = False
                Tipo5.Value = False
                Select Case Val(!TipoOrd)
                    Case 1
                        Tipo1.Value = True
                    Case 2
                        Tipo2.Value = True
                    Case 3
                        Tipo3.Value = True
                    Case 4
                        Tipo4.Value = True
                    Case 5
                        Tipo5.Value = True
                    Case Else
                End Select
                Observaciones.Text = !Observaciones
            End If
        End With
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
        Orden.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            If Tipo3.Value = True Or Tipo4.Value = True Or Tipo5.Value = True Then
                Observaciones.SetFocus
                    Else
                Proveedor.SetFocus
            End If
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            With rstProveedor
                .Index = "Proveedor"
                Claveven$ = Proveedor.Text
                .Seek "=", Proveedor.Text
                If .NoMatch Then
                    Proveedor.Text = Claveven$
                    Proveedor.SetFocus
                        Else
                    Proveedor.Text = !Proveedor
                    DesProveedor.Caption = !Nombre
                    WPrvDireccion = !Direccion
                    WPrvCuit = !Cuit
                    WTipoprv = !Ganancia
                    WTipoiva = !Iva
                    WTipoReteiva = !Reteiva
                    WExepcion = !PorceReteIva
                    Observaciones.SetFocus
                End If
            End With
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
        DesProveedor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            With rstBanco
                .Index = "Banco"
                Claveven$ = Val(Banco.Text)
                .Seek "=", Val(Banco.Text)
                If .NoMatch Then
                    Banco.Text = Claveven$
                    Banco.SetFocus
                        Else
                    Banco.Text = !Banco
                    DesBanco.Caption = !Nombre
                    WCtabanco = !Cuenta
                    WVector1.Col = 1
                    WVector1.Row = 1
                    Call StartEdit
                End If
            End With
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
     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Bancos"
     Opcion.AddItem "Cuentas Contables"
     Rem Opcion.AddItem "Cuenta Corrientes"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
    Ayuda.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            With rstProveedor
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 6)
                        IngresaItem = Auxi + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Ayuda.SetFocus
            
        Case 1
            Ayuda.Visible = True
            Ayuda.Text = ""
            With rstBanco
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = Str$(!Banco)
                        Call Ceros(Auxi, 4)
                        IngresaItem = Auxi + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Banco
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            Ayuda.SetFocus
            
        Case 2
            Ayuda.Visible = True
            Ayuda.Text = ""
            If Cuenta.Visible = True Or Cuenta1.Visible = True Then
                Ayuda.Visible = True
                Ayuda.Text = ""
                With rstCuenta
                    .Index = "Descripcion"
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
            End If
            Ayuda.SetFocus
            
        Case 3
            If Val(Proveedor.Text) <> 0 Then
            With rstCtaCtePrv
                .Index = "ClaveImpre"
                Auxi = Proveedor.Text
                .Seek ">", Auxi + Space$(100)
                If .NoMatch = False Then
                Do
                    If .EOF = False Then
                        If Proveedor.Text = !Proveedor Then
                            ZSaldo = !Saldo
                            Call Redondeo(ZSaldo)
                            If ZSaldo <> 0 Then
                                Auxi$ = Str$(ZSaldo)
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                IngresaItem = !Impre + " " + !Letra + " " + !Punto + " " + !Numero + " " + !Fecha + " " + Auxi$
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                                Else
                        Exit Do
                    End If
                Loop
                End If
            End With
            End If
     
        Case 4
            With rstRecibos
                .Index = "Fecha2"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(!Tiporeg) = 2 Then
                            If Val(!Tipo2) = 2 And !Estado2 <> "X" Then
                                Auxi$ = Str$(!Importe2)
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                Numero = Str$(Val(!Numero2))
                                Call Ceros(Numero, 6)
                                IngresaItem = Numero + "    " + !Fecha2 + "      " + Auxi$ + "      " + !Banco2
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
     
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
            With rstProveedor
                Indice = Pantalla.ListIndex
                .Index = "Proveedor"
                .Seek "=", WIndice.List(Indice)
                If .NoMatch = False Then
                    Proveedor.Text = !Proveedor
                    DesProveedor.Caption = !Nombre
                    WPrvDireccion = !Direccion
                    WPrvCuit = !Cuit
                    WTipoprv = !Ganancia
                    WTipoiva = !Iva
                    WTipoReteiva = !Reteiva
                    WExepcion = !PorceReteIva
                            Else
                    Proveedor.Text = ""
                End If
            End With
            Ayuda.Visible = False
            Pantalla.Visible = False
            Proveedor.SetFocus
            
        Case 1
            With rstBanco
                Indice = Pantalla.ListIndex
                .Index = "Banco"
                .Seek "=", WIndice.List(Indice)
                If .NoMatch = False Then
                    Banco.Text = !Banco
                    DesBanco.Caption = !Nombre
                            Else
                    Banco.Text = ""
                End If
            End With
            Ayuda.Visible = False
            Pantalla.Visible = False
            Banco.SetFocus
            
        Case 2
            If Cuenta.Visible = True Then
                With rstCuenta
                    Indice = Pantalla.ListIndex
                    .Index = "Cuenta"
                    .Seek "=", WIndice.List(Indice)
                    If .NoMatch = False Then
                        Cuenta.Text = !Cuenta
                        Rem   DesProveedor.Caption = !Nombre
                                Else
                        Cuenta.Text = ""
                    End If
                End With
                Ayuda.Visible = False
                Pantalla.Visible = False
                Cuenta.SetFocus
            End If
            If Cuenta1.Visible = True Then
                With rstCuenta
                    Indice = Pantalla.ListIndex
                    .Index = "Cuenta"
                    .Seek "=", WIndice.List(Indice)
                    If .NoMatch = False Then
                        Cuenta1.Text = !Cuenta
                        Rem   DesProveedor.Caption = !Nombre
                                Else
                        Cuenta1.Text = ""
                    End If
                End With
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
                    Compara2 = Proveedor.Text + WVector1.TextMatrix(IRow, 2)
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
        
                    With rstCtaCtePrv

                        Indice = Pantalla.ListIndex
                        .Index = "CtaCte"
                        .Seek "=", WIndice.List(Indice)
                        If .NoMatch = False Then
                
                            WVector1.Row = XRow
                            
                            WVector1.Col = 1
                            WVector1.Text = !Tipo
                    
                            WVector1.Col = 2
                            WVector1.Text = !Letra
                
                            WVector1.Col = 3
                            WVector1.Text = !Punto
                
                            WVector1.Col = 4
                            WVector1.Text = !Numero
                
                            WVector1.Col = 5
                            WVector1.Text = !Saldo
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    
                            WVector1.Col = 6
                            WVector1.Text = ""
                            
                            WVector1.Col = 1
                            WVector1.Text = !Tipo
                            
                            Call Suma_Datos
                            
                            WVector1.Col = 1
                            WVector1.Row = WVector1.Row + 1
                            Call StartEdit
                    
                        End If
                    End With
            
                End If
            
            End If
                
        Case 4
            For IRow = 1 To 50
                If WVector1.TextMatrix(IRow, 7) = "" Then
                    XRow = WVector1.Row
                    Exit For
                End If
            Next IRow
        
            With rstRecibos

                Indice = Pantalla.ListIndex
                .Index = "Clave"
                .Seek "=", WIndice.List(Indice)
                If .NoMatch = False Then
                
                    WVector1.Row = XRow
                        
                    WVector1.Col = 7
                    WVector1.Text = "03"
                    
                    WVector1.Col = 8
                    WVector1.Text = !Numero2
                
                    WVector1.Col = 9
                    WVector1.Text = !Fecha2
                
                    WVector1.Col = 10
                    WVector1.Text = ""
                    
                    WVector1.Col = 11
                    WVector1.Text = !Banco2
                
                    WVector1.Col = 12
                    WVector1.Text = !Importe2
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            
                    WVector1.Col = 7
                    WVector1.Text = "03"
                    
                    BajaCheque(WVector1.Row) = !Clave
                            
                    Call Suma_Datos
                        
                    WVector1.Col = 7
                    WVector1.Row = WVector1.Row + 1
                    Call StartEdit
                    
                End If
            End With
                
        Case Else
    End Select
    
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta.Text <> "" Then
            With rstCuenta
                .Index = "Cuenta"
                .Seek "=", Cuenta.Text
                If .NoMatch Then
                    Cuenta.SetFocus
                        Else
                    WCuenta(WVector1.Row) = Cuenta.Text
                    Ingrecuenta.Visible = False
                    WVector1.Col = WVector1.Col + 1
                    Call StartEdit
                End If
            End With
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
            With rstCuenta
                .Index = "Cuenta"
                .Seek "=", Cuenta1.Text
                If .NoMatch Then
                    Cuenta1.SetFocus
                        Else
                    WCuenta1(WVector1.Row) = Cuenta1.Text
                    IngreCuenta1.Visible = False
                    If WVector1.Row < WVector1.Rows - 1 Then
                        WVector1.Row = WVector1.Row + 1
                    End If
                    WVector1.Col = 7
                    Call StartEdit
                End If
            End With
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
    
    Orden.Text = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Tipo1.Value = True
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    
    Debitos.Caption = ""
    Creditos.Caption = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Retencion.Text = ""
    Reteiva.Text = ""
    PorceIva.Text = ""
    
    WLeyenda(1) = "Compra de Bienes"
    WLeyenda(2) = "Ejericio Prof. Lib. c/Aj.Inf."
    WLeyenda(3) = "Alquileres y Arrendamientos"
    
    With rstConfiguracion
       .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            ConfigIva1 = Str$(!Iva1)
            ConfigIva2 = Str$(!Iva2)
            ConfigPercepcion = Str$(!Percepcion)
            ConfigPunto = !Punto
        End If
    End With
    
    PorceIva.Text = ConfigIva1
    PorceIva.Text = Pusing("###,###.##", PorceIva.Text)
    
    With rstSolicitud
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Orden.Text = Mid$(Str$(!Orden + 1), 2, 6)
                Else
            Orden.Text = "1"
        End If
    End With
    
End Sub

Private Sub IMPREORDEN()

    T$ = "Impresion de la Solicitud de la Orden de Pago"
    m$ = "Desea realizar la impresion del comprobante"
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        With rstImpreOrd
            .Index = "Clave"
            .Seek ">=", 0
            Do
                If .EOF = False Then
                    .Delete
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With

        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With

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
                
                With rstCtaCtePrv
                
                    WProveedor = Proveedor.Text
                    Call Ceros(WProveedor, 6)
                    
                    WVector1.Col = 1
                    XTipo = WTipo
                        
                    WVector1.Col = 2
                    XLetra = WLetra
                    
                    WVector1.Col = 3
                    XPunto = WPunto
                    
                    WVector1.Col = 4
                    XNumero = WNumero
    
                    Indice = Pantalla.ListIndex
                    .Index = "CtaCte"
                    .Seek "=", WProveedor + XLetra + XTipo + XPunto + XNumero
                    If .NoMatch = False Then
                        WImpresion(Cantidad, 1) = !Fecha
                    End If
                    
                End With
                    
            End If
        Next IRow
        
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                WCtaProveedor = !CtaProveedores
                WCtaEfectivo = !CtaEfectivo
                WCtaCheques = !CtaCheque
            End If
        End With
        
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
                        With rstBanco
                            WVector1.Col = 10
                            .Index = "Banco"
                            .Seek "=", Val(WVector1.Text)
                            If .NoMatch = False Then
                                WCredito(Lugar, 1) = !Cuenta
                            End If
                        End With
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
            With rstProveedor
                .Index = "Proveedor"
                .Seek "=", Proveedor.Text
                If .NoMatch = False Then
                    WNombre = !Nombre
                End If
            End With
        End If
    
        Renglon = 0
        With rstImpreOrd
        
            .Index = "Clave"
    
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
                WImporte = Val(WVector1.Text)
            
                WVector1.Col = 6
                WDescripcion = WVector1.Text
            
                WFecha = "  /  /    "
            
                Call Ceros(WTipo, 2)
            
                If WImporte <> 0 Then
            
                    With rstCtaCtePrv
                        WProveedor = Proveedor.Text
                        Call Ceros(WProveedor, 6)
                        XTipo = WTipo
                        XLetra = WLetra
                        XPunto = WPunto
                        XNumero = WNumero
                        .Index = "CtaCte"
                        .Seek "=", WProveedor + XLetra + XTipo + XPunto + XNumero
                        If .NoMatch = False Then
                            WFecha = !Fecha
                        End If
                    End With
                    
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(Orden.Text)
                    Call Ceros(Auxi1, 6)
                    
                    .AddNew
                    !Orden = Val(Orden.Text)
                    !Renglon = Renglon
                    !Proveedor = Val(Proveedor.Text)
                    !Fecha = Fecha.Text
                    !Tiporeg = 1
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Fecha1 = WFecha
                    !Importe = WImporte
                    !Descripcion = WDescripcion
                    !Total = Val(Debitos.Caption)
                    !Retencion = Val(Retencion.Text)
                    !Retib = Val(Reteiva.Text)
                    !Observaciones = Observaciones.Text
                    !Dia = Val(Left$(Fecha, 2))
                    !Mes = Val(Mid$(Fecha, 4, 2))
                    !Ano = Val(Right$(Fecha, 2))
                    !Nombre = WNombre
                    !Cuenta = ""
                    !Nroret = WNroRet
                    !Nroret1 = WNroRet1
                    .Update
            
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
                WImporte = Val(WVector1.Text)
            
                Call Ceros(WTipo, 2)
                
                If WImporte <> 0 Then
                    
                    Renglon = Renglon + 1
                    Auxi = Str$(Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(Orden.Text)
                    Call Ceros(Auxi1, 6)
                    
                    .AddNew
                    !Orden = Val(Orden.Text)
                    !Renglon = Renglon
                    !Proveedor = Val(Proveedor.Text)
                    !Fecha = Fecha.Text
                    !Tiporeg = 2
                    !Tipo = WTipo
                    !Numero = WNumero
                    !Fecha1 = WFecha
                    !Importe = WImporte
                    !Descripcion = WDescripcion
                    !Total = Val(Debitos.Caption)
                    !Retencion = Val(Retencion.Text)
                    !Retib = Val(Reteiva.Text)
                    !Observaciones = Observaciones.Text
                    !Dia = Val(Left$(Fecha, 2))
                    !Mes = Val(Mid$(Fecha, 4, 2))
                    !Ano = Val(Right$(Fecha, 2))
                    !Nombre = WNombre
                    !Cuenta = ""
                    !Nroret = WNroRet
                    !Nroret1 = WNroRet1
                    .Update
            
                End If
                        
            Next Ciclo
        
        End With
    
        listado.ReportFileName = "Impresol.rpt"
        listado.DataFiles(0) = WEmpresa + "admi.mdb"
    
        listado.Destination = 1
        Rem LISTADO.Destination = 0
        listado.Action = 1
        
    End If

End Sub

Private Sub calcret_Click()

    With rstConfiguracion
       .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            ConfigIva1 = Str$(!Iva1)
            ConfigIva2 = Str$(!Iva2)
            ConfigPercepcion = Str$(!Percepcion)
            ConfigPunto = !Punto
        End If
    End With
    
    PorceIva.Text = ConfigIva1
    PorceIva.Text = Pusing("###,###.##", PorceIva.Text)

    With rstParametro
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            WMinimo1 = !Minimo1
            WMinimo2 = !Minimo2
            WMinimo3 = !Minimo3
            WEscala1 = !Escala1
            WEscala2 = !Escala2
            WEscala3 = !Escala3
            WEscala4 = !Escala4
            WEscala5 = !Escala5
            XTasa1 = !Tasa1
            XTasa2 = !Tasa2
            XTasa3 = !Tasa3
            XTasa4 = !Tasa4
            XTasa5 = !Tasa5
            WRetMinima = !RetMinima
            WPorceBienes = !PorceBienes / 100
            WPorceServicios = !PorceServicios / 100
            WPorceTranspo = !PorceTranspo / 100
            WMinimoIva = !MinimoIva
            WIvaInscripto = !IvaInscripto
            WIvaNoInscripto = !IvaNoInscripto
            WTasaGen = !TasaGen / 100
            WTasaBienes = !TasaBienes / 100
            WTasaNoInscripto = !TasaNoInscripto / 100
        End If
    End With
    
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
    
        If WTipoprv = 1 Or WTipoprv = 2 Or WTipoprv = 3 Then
        
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
            Auxi = Proveedor.Text
            Call Ceros(Auxi, 6)

            With rstRetencion
                .Index = "clave"
                Claveven$ = WFecha + Auxi
                .Seek "=", Claveven$
                If .NoMatch = True Then
                    .AddNew
                    !Fecha = WFecha
                    !Proveedor = Proveedor.Text
                    !Neto = 0
                    !Anticipo = 0
                    !Bruto = 0
                    !Iva = 0
                    !Retenido = 0
                    !Clave = !Fecha + Auxi
                    .Update
                    WNeto = 0
                    WAnticipo = 0
                    WBruto = 0
                    WIva = 0
                    WRetenido = 0
                        Else
                    WNeto = !Neto
                    WAnticipo = !Anticipo
                    WBruto = !Bruto
                    WIva = !Iva
                    WRetenido = !Retenido
                End If
            End With
            
            If WTipoprv = 1 Then
                WMinimo = WMinimo1
                    Else
                If WTipoprv = 2 Then
                    WMinimo = WMinimo2
                        Else
                    WMinimo = WMinimo3
                End If
            End If

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
                        Case 1
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
            
                        WProveedor = Proveedor.Text
                        Call Ceros(WProveedor, 6)
                        NetoParcial = 0
                        IvaParcial = 0
            
                        If Val(WTipo) = 1 Or Val(WTipo) = 2 Or Val(WTipo) = 3 Then
                        
                            If WLetra <> "X" Then
                            
                                With rstIvaComp
                                    .Index = "Clave"
                                    .Seek "=", WProveedor + WTipo + WLetra + WPunto + WNumero
                                    If .NoMatch = False Then
                                        WTotalFactura = !Neto + !Exento + !Iva21 + !Iva5 + !Iva27 + !Iva105 + !Ib + !ImpInterno + !ImpCombustible
                                        If WTotalFactura <> 0 Then
                                            WPorceFactura = WImporte / WTotalFactura
                                                Else
                                            WPorceFactura = 0
                                        End If
                                        NetoParcial = !Neto * WPorceFactura
                                        IvaParcial = (!Iva21 + !Iva5 + !Iva105 + !Iva27) * WPorceFactura
                                    End If
                                End With
                    
                                XBrutoIva = XBrutoIva + WImporte
                                XNetoIva = XNetoIva + NetoParcial
                                XIvaIva = XIvaIva + IvaParcial
                                
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
            WProveedor = Proveedor.Text
            
            Call Ceros(WTipo, 2)
            Call Ceros(WPunto, 4)
            Call Ceros(WNumero, 8)
            Call Ceros(WProveedor, 6)
            
            If Val(WTipo) = 1 Or Val(WTipo) = 2 Or Val(WTipo) = 3 Then
                If WLetra <> "X" Then
                    With rstIvaComp
                        .Index = "Clave"
                        .Seek "=", WProveedor + WTipo + WLetra + WPunto + WNumero
                        If .NoMatch = False Then
                            WTotalFactura = !Neto + !Exento + !Iva21 + !Iva5 + !Iva27 + !Iva105 + !Ib + !ImpInterno + !ImpCombustible
                            If WTotalFactura <> 0 Then
                                WPorceFactura = WImporte / WTotalFactura
                                    Else
                                WPorceFactura = 0
                            End If
                            NetoParcial = !Neto * WPorceFactura
                            IvaParcial = (!Iva21 + !Iva5 + !Iva105 + !Iva27) * WPorceFactura
                            NetoTotal = NetoTotal + NetoParcial
                            IvaTotal = IvaTotal + IvaParcial
                            FacturaTotal = FacturaTotal + WImporte
                        End If
                    End With
                End If
            End If
        End If
    Next IRow
    
    Call Redondeo(NetoTotal)
    Call Redondeo(IvaTotal)
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then
    
        WEspacios = Len(Ayuda.Text)
    
        Opcion.Visible = False

        Dim IngresaItem As String

        Pantalla.Clear
        WIndice.Clear

        XIndice = Opcion.ListIndex
    
        Select Case XIndice
            Case 0
                With rstProveedor
                    .Index = "Nombre"
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            da = Len(!Nombre) - WEspacios
                
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                    Auxi = Str$(!Proveedor)
                                    Call Ceros(Auxi, 6)
                                    IngresaItem = Auxi + " " + !Nombre
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = !Proveedor
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                    
                                    Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                
            Case 1
                With rstBanco
                    .Index = "Nombre"
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            da = Len(!Nombre) - WEspacios
                
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                    Auxi = Str$(!Banco)
                                    Call Ceros(Auxi, 4)
                                    IngresaItem = Auxi + "    " + !Nombre
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = !Banco
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                    
                                    Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                
            Case 2
                With rstCuenta
                    .Index = "Descripcion"
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            da = Len(!Descripcion) - WEspacios
                
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                    IngresaItem = !Cuenta + "    " + !Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = !Cuenta
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                    
                                    Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                
            Case Else
        
        End Select
    
    End If
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
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
                With rstCtaCtePrv
                    .Index = "CtaCte"
                    WProveedor = Proveedor.Text
                    Call Ceros(WProveedor, 6)
                    Claveven$ = WProveedor
                    WVector1.Col = 2
                    Claveven$ = Claveven$ + WVector1.Text
                    WVector1.Col = 1
                    Claveven$ = Claveven$ + WVector1.Text
                    WVector1.Col = 3
                    Claveven$ = Claveven$ + WVector1.Text
                    WVector1.Col = 4
                    Claveven$ = Claveven$ + WVector1.Text
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        WVector1.Col = 5
                        If Val(WVector1.Text) = 0 Then
                            WVector1.Text = !Saldo
                            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            Call Suma_Datos
                        End If
                        WVector1.Col = 4
                            Else
                        WControl = "N"
                    End If
                End With
            End If
                
        Case 5
            If Tipo1.Value = True Then
                With rstCtaCtePrv
                    .Index = "CtaCte"
                    WProveedor = Proveedor.Text
                    Call Ceros(WProveedor, 6)
                    Claveven$ = WProveedor
                    WVector1.Col = 2
                    Claveven$ = Claveven$ + WVector1.Text
                    WVector1.Col = 1
                    Claveven$ = Claveven$ + WVector1.Text
                    WVector1.Col = 3
                    Claveven$ = Claveven$ + WVector1.Text
                    WVector1.Col = 4
                    Claveven$ = Claveven$ + WVector1.Text
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                        Saldo = !Saldo
                            Else
                        Saldo = 0
                    End If
                End With
                
                WVector1.Col = 5
                If Val(WVector1.Text) > Saldo Then
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
                With rstBanco
                    .Index = "Banco"
                    Claveven$ = Val(Banco.Text)
                    .Seek "=", Val(Banco.Text)
                    If .NoMatch = False Then
                        WCuenta(WVector1.Row) = !Cuenta
                            Else
                        WCuenta(WVector1.Row) = "999999"
                    End If
                End With
            End If
            If Tipo5.Value = True Then
                With rstEmpresa
                    .Index = "Empresa"
                    Claveven$ = Val(WEmpresa)
                    .Seek "=", Claveven$
                    If .NoMatch = False Then
                            WCtaChequeRecha = !CtaChequeRecha
                    End If
                End With
                WCuenta(WVector1.Row) = WCtaChequeRecha
            End If
     
        Case 7
            If WVector1.Text <> "" Then
            If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Then
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
                        WVector1.Text = "Varios"
                        
                    Case Else
                End Select
                        
                    Else
                            
                WControl = "N"
                        
            End If
            End If
                
        Case 8
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
                
        Case 9
            Call Valida_fecha1(WTexto3.Text, Auxi)
            WControl = "S"
            If Auxi <> "S" Then
                 WControl = "N"
            End If
                
        Case 10
            With rstBanco
                .Index = "Banco"
                Claveven$ = WVector1.Text
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WVector1.Col = 11
                    WVector1.Text = !Nombre
                        Else
                    WControl = "N"
                End If
            End With

        Case 12
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            Call Suma_Datos
            If Val(WVector1.TextMatrix(WVector1.Row, 7)) = 4 Then
                Cuenta1.Text = WCuenta1(WVector1.Row)
                IngreCuenta1.Visible = True
                Cuenta1.SetFocus
            End If
            
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
    WVector1.Cols = 13
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
                WParametros(1, Ciclo) = 2
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

Private Sub Tipo5_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call cmdDelete_Click
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
            Call cmdClose_Click
        Case Else
    End Select
End Sub











