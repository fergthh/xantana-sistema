VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaCheque 
   Caption         =   "Consulta de Cheques Terceros"
   ClientHeight    =   7635
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11880
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
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   5040
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
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   6360
      Width           =   375
   End
   Begin VB.Frame DatosCheque 
      Height          =   4335
      Left            =   1320
      TabIndex        =   14
      Top             =   1560
      Visible         =   0   'False
      Width           =   8055
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   42
         Text            =   " "
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Clientes 
         BeginProperty Font 
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
         MaxLength       =   6
         TabIndex        =   38
         Text            =   " "
         Top             =   360
         Width           =   975
      End
      Begin VB.ComboBox EstadoCheque 
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
         TabIndex        =   34
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox NumeroCheque 
         BeginProperty Font 
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
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox CodigoBanco 
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
         MaxLength       =   6
         TabIndex        =   19
         Text            =   " "
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox SucursalCheque 
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
         MaxLength       =   6
         TabIndex        =   18
         Text            =   " "
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TipoCheque 
         BeginProperty Font 
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
         MaxLength       =   6
         TabIndex        =   17
         Text            =   " "
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox ClaseCheque 
         BeginProperty Font 
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
         MaxLength       =   6
         TabIndex        =   16
         Text            =   " "
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox ImporteCheque 
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
         TabIndex        =   15
         Text            =   " "
         Top             =   3240
         Width           =   1815
      End
      Begin MSMask.MaskEdBox FechaCheque 
         Height          =   285
         Left            =   2520
         TabIndex        =   21
         Top             =   1440
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
         Caption         =   "Cuit Firmante"
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
         Left            =   720
         TabIndex        =   43
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label DesClientes 
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
         Left            =   3720
         TabIndex        =   40
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cod. Cilente"
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
         Left            =   720
         TabIndex        =   39
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Estado"
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
         Left            =   720
         TabIndex        =   33
         Top             =   3600
         Width           =   1455
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
         Index           =   2
         Left            =   720
         TabIndex        =   31
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Fecha Cheque"
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
         Left            =   720
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codigo Banco"
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
         Left            =   720
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sucursal"
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
         Left            =   720
         TabIndex        =   28
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo de Cheque"
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
         Left            =   720
         TabIndex        =   27
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Clase Cheque"
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
         Left            =   720
         TabIndex        =   26
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label DesCodigoBanco 
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
         Left            =   3480
         TabIndex        =   25
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label lblLabels 
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
         Height          =   255
         Index           =   7
         Left            =   720
         TabIndex        =   24
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "0 - Terceros - 1 - Propio    "
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
         Left            =   3360
         TabIndex        =   23
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "0 - Portador   1 - A la Orden  2 - No a la Orden   "
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
         Left            =   3360
         TabIndex        =   22
         Top             =   2880
         Width           =   4335
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
      Index           =   11
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   6720
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
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3120
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
      TabIndex        =   35
      Top             =   2880
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
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   32
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
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3840
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
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5040
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4560
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4560
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4560
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4560
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4560
      Width           =   375
   End
   Begin VB.Frame BusquedaCheque 
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8655
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
         Left            =   6840
         MouseIcon       =   "consultacheque.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "consultacheque.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Menu Principal"
         Top             =   360
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
         Left            =   2280
         TabIndex        =   0
         Text            =   " "
         Top             =   720
         Width           =   1335
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
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Limpia "
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
         Left            =   4560
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5415
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9551
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgConsultaCheque"
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
Dim ZCheque(1000, 14) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Dim ZZNumeroCheque As String

Private Sub Acepta_Click()

    Call Limpia_Vector

    ZLugar = 0
    
    ZZNumero = Str(Val(Numero.Text))
    ZZNumero = Trim(ZZNumero)
    Call Ceros(ZZNumero, 8)

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + ZZNumero + "'"
    Rem ZSql = ZSql + " Where Recibos.Numero2 LIKE " + "'" + "%" + ZZNumero + "%" + "'"
    ZSql = ZSql + " Order by Clave"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZSuma = ZPrueba
                    ZLugar = ZLugar + 1
                    ZCheque(ZLugar, 1) = rstRecibos!Numero2
                    ZCheque(ZLugar, 2) = rstRecibos!Fecha2
                    ZCheque(ZLugar, 3) = rstRecibos!Banco2
                    ZCheque(ZLugar, 4) = Str$(rstRecibos!Importe2)
                    ZCheque(ZLugar, 5) = rstRecibos!Clave
                    ZCheque(ZLugar, 6) = IIf(IsNull(rstRecibos!CodigoBanco), "", rstRecibos!CodigoBanco)
                    ZCheque(ZLugar, 7) = rstRecibos!Estado2
                    ZCheque(ZLugar, 8) = rstRecibos!Cliente
                    ZCheque(ZLugar, 9) = rstRecibos!SucursalCheque
                    ZCheque(ZLugar, 10) = rstRecibos!TipoCheque
                    ZCheque(ZLugar, 11) = rstRecibos!ClaseCheque
                    ZCheque(ZLugar, 12) = rstRecibos!Destino
                    ZCheque(ZLugar, 13) = rstRecibos!Fecha
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    For Ciclo = 1 To ZLugar
        
        WVector1.Row = Ciclo
        ZZLugar = Ciclo
            
        WVector1.Col = 1
        WVector1.Text = ZCheque(Ciclo, 1)
        
        WVector1.Col = 2
        WVector1.Text = ZCheque(Ciclo, 13)
        
        WVector1.Col = 3
        WVector1.Text = ZCheque(Ciclo, 2)
        
        WVector1.Col = 4
        WVector1.Text = ZCheque(Ciclo, 6)
            
        WVector1.Col = 5
        WVector1.Text = ZCheque(Ciclo, 3)
        
        WVector1.Col = 6
        WVector1.Text = ZCheque(Ciclo, 4)
        WVector1.Text = Pusing("###,###.##", WVector1.Text)
        
        If ZCheque(Ciclo, 7) <> "X" Then
            WVector1.Col = 7
            WVector1.Text = "En Cartera"
                Else
            WVector1.Col = 7
            WVector1.Text = "Entregado"
        End If
        
        WVector1.Col = 8
        WVector1.Text = ZCheque(Ciclo, 9)
        
        Select Case ZCheque(Ciclo, 10)
            Case "0"
                WVector1.Col = 9
                WVector1.Text = "Propio"
            Case Else
                WVector1.Col = 9
                WVector1.Text = "Tercero"
        End Select
        
        Select Case ZCheque(Ciclo, 11)
            Case "2"
                WVector1.Col = 9
                WVector1.Text = WVector1.Text + "-" + "No Orden"
            Case "0"
                WVector1.Col = 9
                WVector1.Text = WVector1.Text + "-" + "Portador"
            Case Else
                WVector1.Col = 9
                WVector1.Text = WVector1.Text + "-" + "Orden"
        End Select
        
        WVector1.Col = 10
        WVector1.Text = ZCheque(Ciclo, 8)
        Rem ZSql = ""
        Rem ZSql = ZSql + "Select *"
        Rem ZSql = ZSql + " FROM Cliente"
        Rem ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZCheque(Ciclo, 8) + "'"
        Rem spCliente = ZSql
        Rem Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstCliente.RecordCount > 0 Then
        Rem     WVector1.Col = 10
        Rem     WVector1.Text = rstCliente!Razon
        Rem     rstCliente.Close
        Rem End If
            
        WVector1.Col = 11
        WVector1.Text = ZCheque(Ciclo, 12)
        
        WVector1.Col = 12
        WVector1.Text = ZCheque(Ciclo, 5)
        
        WVector1.Col = 13
        WVector1.Text = ZCheque(Ciclo, 14)
    
    Next Ciclo
        

End Sub

Private Sub CmdClose_Click()
    PrgConsultaCheque.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Form_Load()

    EstadoCheque.Clear
    
    EstadoCheque.AddItem "En cartera"
    EstadoCheque.AddItem "Entregado"
    
    EstadoCheque.ListIndex = 0

    Call Limpia_Vector
    Numero.Text = ""
    
End Sub

Private Sub Cancela_Click()

    Call Limpia_Vector
    Numero.Text = ""
    
End Sub


Private Sub numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Acepta_Click
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
End Sub

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 14
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
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
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
           Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 7
                WVector1.Text = "Estado"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 8
                WVector1.Text = "Suc."
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 9
                WVector1.Text = "Atributos"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 10
                WVector1.Text = "Cliente"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 11
                WVector1.Text = "Destino"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 12
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 13
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
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


Private Sub WVector1_DblClick()

    ZZClave = Trim(WVector1.TextMatrix(WVector1.Row, 12))
    
    If ZZClave <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Clave = " + "'" + ZZClave + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
        
            NumeroCheque.Text = rstRecibos!Numero2
            FechaCheque.Text = rstRecibos!Fecha2
            CodigoBanco.Text = IIf(IsNull(rstRecibos!CodigoBanco), "", rstRecibos!CodigoBanco)
            SucursalCheque.Text = IIf(IsNull(rstRecibos!SucursalCheque), "", rstRecibos!SucursalCheque)
            TipoCheque.Text = IIf(IsNull(rstRecibos!TipoCheque), "", rstRecibos!TipoCheque)
            ClaseCheque.Text = IIf(IsNull(rstRecibos!ClaseCheque), "", rstRecibos!ClaseCheque)
            ImporteCheque.Text = Str$(rstRecibos!Importe2)
            Clientes.Text = Trim(rstRecibos!Cliente)
            Cuit.Text = IIf(IsNull(rstRecibos!Cuit), "", rstRecibos!Cuit)
            
            If rstRecibos!Estado2 = "X" Then
                EstadoCheque.ListIndex = 1
                    Else
                EstadoCheque.ListIndex = 0
            End If
            
            DatosCheque.Visible = True
            NumeroCheque.SetFocus
        
            rstRecibos.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Bcra"
            ZSql = ZSql + " Where Bcra.Codigo = " + "'" + CodigoBanco.Text + "'"
            spBcra = ZSql
            Set rstBcra = db.OpenRecordset(spBcra, dbOpenSnapshot, dbSQLPassThrough)
            If rstBcra.RecordCount > 0 Then
                DesCodigoBanco.Caption = rstBcra!Descripcion
                rstBcra.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesClientes.Caption = rstCliente!Razon
                rstCliente.Close
                    Else
                DesClientes.Caption = ""
            End If
            
        End If
        
        DatosCheque.Visible = True
        Clientes.SetFocus
        
    End If
        
End Sub

Private Sub Clientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Clientes.Text) <> "" Then
            Auxi = UCase(Left$(Clientes.Text, 1))
            Auxi1 = Mid$(Clientes.Text, 2, 5)
            Call Ceros(Auxi1, 3)
            Clientes.Text = Auxi + "-" + Auxi1
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesClientes.Caption = rstCliente!Razon
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WProv = rstCliente!Provincia
            rstCliente.Close
            NumeroCheque.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Clientes.Text = ""
        DesClientes.Caption = ""
    End If
End Sub

Private Sub NumeroCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZZNumeroCheque = NumeroCheque.Text
        Call Ceros(ZZNumeroCheque, 8)
        NumeroCheque.Text = ZZNumeroCheque
        FechaCheque.SetFocus
    End If
    If KeyAscii = 27 Then
        NumeroCheque.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(FechaCheque.Text)) = 8 Then
            FechaCheque.Text = Left$(FechaCheque.Text, 6) + "20" + Right$(Trim(FechaCheque.Text), 2)
        End If
        Call Valida_fecha1(FechaCheque.Text, Auxi)
        If Auxi = "S" Then
            CodigoBanco.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaCheque.Text = "  /  /    "
    End If
End Sub

Private Sub CodigoBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Bcra"
        ZSql = ZSql + " Where Bcra.Codigo = " + "'" + CodigoBanco.Text + "'"
        spBcra = ZSql
        Set rstBcra = db.OpenRecordset(spBcra, dbOpenSnapshot, dbSQLPassThrough)
        If rstBcra.RecordCount > 0 Then
            DesCodigoBanco.Caption = rstBcra!Descripcion
            rstBcra.Close
            SucursalCheque.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        CodigoBanco.Text = ""
        DesCodigoBanco.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub SucursalCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoCheque.SetFocus
    End If
    If KeyAscii = 27 Then
        SucursalCheque.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub TipoCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoCheque.Text = UCase(TipoCheque.Text)
        If TipoCheque.Text = "0" Or TipoCheque.Text = "1" Then
            ClaseCheque.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        TipoCheque.Text = ""
    End If
End Sub

Private Sub ClaseCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ClaseCheque.Text = UCase(ClaseCheque.Text)
        If ClaseCheque.Text = "0" Or ClaseCheque.Text = "1" Or ClaseCheque.Text = "2" Then
            ImporteCheque.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        ClaseCheque.Text = ""
    End If
End Sub

Private Sub ImporteCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If EstadoCheque.ListIndex = 0 Then
            ZZEstado = ""
                Else
            ZZEstado = "X"
        End If
        
        ZZOrdFecha = Right$(FechaCheque.Text, 4) + Mid$(FechaCheque.Text, 4, 2) + Left$(FechaCheque.Text, 2)
            
        ZSql = ""
        ZSql = ZSql + "UPDATE Recibos SET "
        ZSql = ZSql + " Cliente = " + "'" + Clientes.Text + "',"
        ZSql = ZSql + " Numero2 = " + "'" + NumeroCheque.Text + "',"
        ZSql = ZSql + " Fecha2 = " + "'" + FechaCheque.Text + "',"
        ZSql = ZSql + " FechaOrd2 = " + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + " CodigoBanco = " + "'" + CodigoBanco + "',"
        ZSql = ZSql + " SucursalCheque = " + "'" + SucursalCheque.Text + "',"
        ZSql = ZSql + " TipoCheque = " + "'" + TipoCheque.Text + "',"
        ZSql = ZSql + " ClaseCheque = " + "'" + ClaseCheque.Text + "',"
        ZSql = ZSql + " Estado2 = " + "'" + ZZEstado + "',"
        ZSql = ZSql + " Importe2 = " + "'" + ImporteCheque.Text + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        
        If EstadoCheque.ListIndex = 0 Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Recibos SET "
            ZSql = ZSql + " Destino = " + "'" + "" + "',"
            ZSql = ZSql + " Orden = " + "'" + "0" + "',"
            ZSql = ZSql + " Deposito = " + "'" + "0" + "'"
            ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        DatosCheque.Visible = False
        Call numero_Keypress(13)
        
    End If
    If KeyAscii = 27 Then
        ImporteCheque.Text = ""
    End If
End Sub





