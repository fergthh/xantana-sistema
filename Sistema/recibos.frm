VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRecibos 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Recibos"
   ClientHeight    =   8250
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11880
   Begin VB.Frame DatosCheque 
      Height          =   3735
      Left            =   2040
      TabIndex        =   66
      Top             =   2040
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox NroCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5400
         MaxLength       =   20
         TabIndex        =   101
         Text            =   " "
         Top             =   1440
         Width           =   1815
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   86
         Text            =   " "
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox ClaveLectora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   85
         Text            =   " "
         Top             =   3360
         Visible         =   0   'False
         Width           =   975
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
         TabIndex        =   80
         Text            =   " "
         Top             =   2880
         Width           =   1815
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
         TabIndex        =   77
         Text            =   " "
         Top             =   2160
         Width           =   615
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
         TabIndex        =   75
         Text            =   " "
         Top             =   1800
         Width           =   615
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
         TabIndex        =   73
         Text            =   " "
         Top             =   1440
         Width           =   855
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
         TabIndex        =   71
         Text            =   " "
         Top             =   1080
         Width           =   855
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
         TabIndex        =   67
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin MSMask.MaskEdBox FechaCheque 
         Height          =   285
         Left            =   6120
         TabIndex        =   69
         Top             =   720
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
      Begin MSMask.MaskEdBox FechaEmision 
         Height          =   285
         Left            =   2520
         TabIndex        =   99
         Top             =   720
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
         Caption         =   "Nro Cuenta"
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
         Left            =   3600
         TabIndex        =   102
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label19 
         Caption         =   "Fecha Emision"
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
         TabIndex        =   100
         Top             =   720
         Width           =   1575
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
         TabIndex        =   87
         Top             =   2520
         Width           =   1455
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
         TabIndex        =   84
         Top             =   2160
         Width           =   4335
      End
      Begin VB.Label lblLabels 
         Caption         =   "0 - Terceros  1 - Propio   "
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
         TabIndex        =   83
         Top             =   1800
         Width           =   2175
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
         TabIndex        =   81
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
         TabIndex        =   79
         Top             =   1080
         Width           =   3495
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
         TabIndex        =   78
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
         TabIndex        =   76
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
         TabIndex        =   74
         Top             =   1440
         Width           =   1455
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
         TabIndex        =   72
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
         Left            =   4320
         TabIndex        =   70
         Top             =   720
         Width           =   1575
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
         TabIndex        =   68
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame PantaRete 
      Height          =   4455
      Left            =   1920
      TabIndex        =   51
      Top             =   1680
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox NroRetOtraIV 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   97
         Text            =   " "
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox RetOtraIV 
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   96
         Text            =   " "
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox NroRetOtraIII 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   94
         Text            =   " "
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox RetOtraIII 
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   93
         Text            =   " "
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox NroRetOtraII 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   91
         Text            =   " "
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox RetOtraII 
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   90
         Text            =   " "
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox RetSuss 
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   64
         Text            =   " "
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox NroRetSuss 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   63
         Text            =   " "
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Retganancias 
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   57
         Text            =   " "
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox RetOtra 
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   56
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox RetIva 
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
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   55
         Text            =   " "
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox NroRetganancias 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   54
         Text            =   " "
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox NroRetOtra 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   53
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox NroRetIva 
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
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   52
         Text            =   " "
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Ret. I.Brutos IV"
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
         TabIndex        =   98
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Ret. I.Brutos III"
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
         TabIndex        =   95
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label16 
         Caption         =   "Ret. I.Brutos II"
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
         TabIndex        =   92
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Ret. Suss"
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
         TabIndex        =   65
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Rte.Ganancias"
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
         TabIndex        =   62
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Ret. I.Brutos"
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
         TabIndex        =   61
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Ret.Iva"
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
         TabIndex        =   60
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
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
         Left            =   2160
         TabIndex        =   59
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Nro.Comprobante"
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
         Left            =   3840
         TabIndex        =   58
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame SeleccionaCheque 
      Height          =   3015
      Left            =   240
      TabIndex        =   88
      Top             =   1560
      Visible         =   0   'False
      Width           =   4095
      Begin VB.ListBox CargaCheque 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   240
         TabIndex        =   89
         Top             =   360
         Width           =   3735
      End
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
      Left            =   10920
      MouseIcon       =   "recibos.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "recibos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Impresion"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox TotalRete 
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
      Left            =   5160
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   48
      Text            =   " "
      Top             =   1200
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
      Left            =   10920
      MouseIcon       =   "recibos.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "recibos.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Menu Principal"
      Top             =   7200
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
      Left            =   10920
      MouseIcon       =   "recibos.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "recibos.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Consulta de Datos"
      Top             =   5040
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
      Left            =   10920
      MouseIcon       =   "recibos.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "recibos.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3960
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
      Left            =   10920
      MouseIcon       =   "recibos.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "recibos.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Elimina el Registro"
      Top             =   2880
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
      Left            =   10920
      MouseIcon       =   "recibos.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "recibos.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Ctacte 
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
      Left            =   10920
      MouseIcon       =   "recibos.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "recibos.frx":46D2
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Cuenta Corriente de Proveedores"
      Top             =   6120
      Width           =   855
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
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   5760
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
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   5760
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
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   5760
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
      TabIndex        =   34
      Top             =   5760
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   5760
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
      Left            =   6960
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
      Index           =   4
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   5280
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
      Left            =   6720
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
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
      TabIndex        =   28
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
      Index           =   2
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   27
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
      Index           =   1
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   5280
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
      Left            =   4080
      TabIndex        =   25
      Top             =   4800
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   24
      Top             =   5280
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
      Left            =   3360
      TabIndex        =   23
      Top             =   4800
      Width           =   375
   End
   Begin VB.Frame Ingrecuenta 
      Caption         =   "Ingreso de Cuenta Contable"
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
      Height          =   1095
      Left            =   3120
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
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
         Left            =   600
         MaxLength       =   10
         TabIndex        =   22
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
   End
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   20
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport listado 
      Left            =   6240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "recibo.rpt"
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
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   840
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Recibos"
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
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   5175
      Begin VB.OptionButton Tipo3 
         Caption         =   "Varios"
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
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton Tipo1 
         Caption         =   "Cobro "
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
         Left            =   480
         TabIndex        =   13
         Top             =   240
         Width           =   855
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
         Left            =   1920
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   1455
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
      Height          =   1425
      Left            =   6720
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4440
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
   Begin VB.TextBox Recibo 
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
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      ItemData        =   "recibos.frx":4F9C
      Left            =   6720
      List            =   "recibos.frx":4FA3
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   29
      Top             =   4800
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
      TabIndex        =   39
      Top             =   3120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7646
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
      Left            =   9240
      TabIndex        =   50
      Top             =   7920
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Total Retenciones"
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
      Left            =   3360
      TabIndex        =   49
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "VALORES RECIBIDOS"
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
      Left            =   4560
      TabIndex        =   41
      Top             =   2760
      Width           =   5895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "COMPROBANTES CANCELADOS"
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
      Left            =   120
      TabIndex        =   40
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  3) Factura    4) N/D   5 N/C"
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
      Left            =   360
      TabIndex        =   38
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Top             =   840
      Visible         =   0   'False
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
      Left            =   9240
      TabIndex        =   16
      Top             =   7560
      Width           =   1215
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
      Left            =   3360
      TabIndex        =   15
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  1) Ef.   2) Ch.   3) Dep.  4) Comp."
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
      Left            =   5880
      TabIndex        =   14
      Top             =   7560
      Width           =   3255
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
      Left            =   3480
      TabIndex        =   10
      Top             =   480
      Width           =   3495
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
      Left            =   3480
      TabIndex        =   8
      Top             =   120
      Width           =   975
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
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero de Recibo"
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
      TabIndex        =   4
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "PrgRecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Private Auxi1 As String
Private WSaldo As Double
Private Vector(20, 10) As String
Private Provincia(100) As String
Private m(20) As String
Private Impre1(100) As Single
Private Impre2(100) As Single
Private ImpreTipo(100) As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WPostal As String
Private WProvincia As String
Private WProv As String
Private WCuenta(20) As String
Private Debito As Double
Private Credito As Double
Private ZSaldo As Double
Dim ZMes As String
Dim ZAno As String

Dim ZZZRetOtra As Double
Dim ZZZRetOtraII As Double
Dim ZZZRetOtraIII As Double
Dim ZZZRetOtraIV As Double
Dim ZZAnticipo As Double



Dim ZZRecibo As String
Dim ZZRenglon As String
Dim ZZCliente As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZTipoRec As String
Dim ZZRetGanancias As String
Dim ZZRetIva As String
Dim ZZRetOtra As String
Dim ZZRetOtraII As String
Dim ZZRetOtraIII As String
Dim ZZRetOtraIV As String
Dim ZZRetSuss As String
Dim ZZNroRetganancias As String
Dim ZZNroRetIva As String
Dim ZZNroRetOtra As String
Dim ZZNroRetOtraII As String
Dim ZZNroRetOtraIII As String
Dim ZZNroRetOtraIV As String
Dim ZZNroRetSuss As String
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
Dim ZZFechaOrd2 As String
Dim ZZBanco2 As String
Dim ZZImporte2 As String
Dim ZZEstado2 As String
Dim ZZObservaciones As String
Dim ZZEmpresa As String
Dim ZZClave As String
Dim ZZImporte As String
Dim ZZCuenta As String
Dim ZZDestino As String
Dim ZZOrden As String
Dim ZZDeposito As String


Dim ZZLetra As String
Dim ZZTipo As String
Dim ZZPunto As String
Dim ZZNumero As String
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
Dim ZZProvincia As String
Dim ZZVendedor As String
Dim ZZCosto As String
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

Dim ZZNumeroCheque As String
Dim ZZVectorI(100, 6) As String
Dim ZZVectorII(100, 6) As String
Dim ZZVectorIII(100, 6) As String

Dim ZZRazon As String
Dim ZZPesosI As String
Dim ZZPesosII As String
Dim ZZFechaI As String
Dim zZNumeroI As String
Dim ZZImporteI As String
Dim ZZBanco As String
Dim ZZSucursal As String
Dim ZZNumeroII As String
Dim ZZFechaII As String
Dim ZZImporteII As String
Dim ZZEstructura As String
Dim ZZImporteIII As String

Dim XTexto2 As String
Dim XTexto1 As String

Dim ZCargaCheque(100) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Dim ZZAlta As Integer

Private Sub Suma_Datos()

    Debitos.Caption = ""
    Creditos.Caption = ""
    Diferencia.Caption = ""
    ZZAnticipo = 0
    
    Creditos.Caption = Str$(Val(Retganancias.Text) + Val(RetIva.Text) + Val(RetOtra.Text) + Val(RetOtraII.Text) + Val(RetOtraIII.Text) + Val(RetOtraIV.Text) + Val(RetSuss.Text))
    
    For IRow = 1 To 100
        If Val(WVector1.TextMatrix(IRow, 1)) = 99 Then
            ZZAnticipo = ZZAnticipo + Val(WVector1.TextMatrix(IRow, 5))
        End If
        Debitos.Caption = Str$(Val(Debitos.Caption) + Val(WVector1.TextMatrix(IRow, 5)))
        Creditos.Caption = Str$(Val(Creditos.Caption) + Val(WVector1.TextMatrix(IRow, 10)))
    Next IRow
    
    Debitos.Caption = Alinea("###,###.##", Debitos.Caption)
    Creditos.Caption = Alinea("###,###.##", Creditos.Caption)
    
    Diferencia.Caption = Str$(Val(Debitos.Caption) - Val(Creditos.Caption))
    Diferencia.Caption = Alinea("###,###.##", Diferencia.Caption)
    
    Rem If Val(Diferencia.Caption) = 0 Then
    Rem     If ZZAlta = 0 Then
    Rem         ZZAlta = 1
    Rem         T$ = "Ingresos de Recibos"
    Rem         M1$ = "Desea Grabar el Recibo"
    Rem         Respuestaaaaaa% = MsgBox(M1$, 32 + 4, T$)
    Rem         If Respuestaaaaaa% = 6 Then
    Rem             Call cmdAdd_Click
    Rem             WControl = "N"
    Rem             Recibo.SetFocus
    Rem         End If
    Rem     End If
    Rem End If
    
End Sub

Private Sub Lee_Datos()

    Call Limpia_Vector

    Renglon = 0
    Debito = 0
    Credito = 0
    CmdDelete.Enabled = True
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Recibo = " + "'" + recibo.Text + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Select Case Val(rstRecibos!Tiporeg)
                        Case 1
                            Debito = Debito + 1
                            WVector1.Row = Debito
                            WVector1.Col = 1
                            WVector1.Text = rstRecibos!Tipo1
                            WVector1.Col = 2
                            WVector1.Text = rstRecibos!Letra1
                            WVector1.Col = 3
                            WVector1.Text = rstRecibos!Punto1
                            WVector1.Col = 4
                            WVector1.Text = rstRecibos!Numero1
                            WVector1.Col = 5
                            WVector1.Text = rstRecibos!Importe1
                            WVector1.Text = Alinea("###,###.##", WVector1.Text)
                        Case 2
                            Credito = Credito + 1
                            WVector1.Row = Credito
                            WVector1.Col = 6
                            WVector1.Text = rstRecibos!Tipo2
                            WVector1.Col = 7
                            WVector1.Text = rstRecibos!Numero2
                            WVector1.Col = 8
                            WVector1.Text = rstRecibos!Fecha2
                            WVector1.Col = 9
                            WVector1.Text = rstRecibos!Banco2
                            WVector1.Col = 10
                            WVector1.Text = rstRecibos!Importe2
                            WVector1.Text = Alinea("###,###.##", WVector1.Text)
                            WVector1.Col = 11
                            WVector1.Text = rstRecibos!CodigoBanco
                            WVector1.Col = 12
                            WVector1.Text = rstRecibos!SucursalCheque
                            WVector1.Col = 13
                            WVector1.Text = rstRecibos!TipoCheque
                            WVector1.Col = 14
                            WVector1.Text = rstRecibos!ClaseCheque
                            WVector1.Col = 16
                            WVector1.Text = rstRecibos!ClaveLectora
                            WVector1.Col = 17
                            WVector1.Text = IIf(IsNull(rstRecibos!Cuit), "", rstRecibos!Cuit)
                            WVector1.Col = 18
                            WVector1.Text = IIf(IsNull(rstRecibos!NroCuenta), "", rstRecibos!NroCuenta)
                            WVector1.Col = 19
                            WVector1.Text = IIf(IsNull(rstRecibos!FechaEmision), "  /  /    ", rstRecibos!FechaEmision)
                        
                            If rstRecibos!Tipo2 = "02" And rstRecibos!Estado2 = "X" Then
                                CmdDelete.Enabled = False
                            End If
                        
                        
                        
                        Case Else
                    End Select
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstRecibos.Close
    End If
    
     
End Sub

Sub Verifica_datos()
    If Val(Retganancias.Text) = 0 Then
        Retganancias.Text = "0"
    End If
    If Val(RetIva.Text) = 0 Then
        RetIva.Text = "0"
    End If
    If Val(RetOtra.Text) = 0 Then
        RetOtra.Text = "0"
    End If
    If Val(RetOtraII.Text) = 0 Then
        RetOtraII.Text = "0"
    End If
    If Val(RetOtraIII.Text) = 0 Then
        RetOtraIII.Text = "0"
    End If
    If Val(RetOtraIV.Text) = 0 Then
        RetOtraIV.Text = "0"
    End If
    If Val(RetSuss.Text) = 0 Then
        RetSuss.Text = "0"
    End If
End Sub

Sub Format_datos()
    Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
    RetIva.Text = Alinea("###,###.##", RetIva.Text)
    RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
    RetOtraII.Text = Alinea("###,###.##", RetOtraII.Text)
    RetOtraIII.Text = Alinea("###,###.##", RetOtraIII.Text)
    RetOtraIV.Text = Alinea("###,###.##", RetOtraIV.Text)
    RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Clientes.Text = rstCliente!Cliente
        DesClientes.Caption = rstCliente!Fantasia
        WRazon = rstCliente!Fantasia
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WProv = rstCliente!Provincia
        rstCliente.Close
        Call Format_datos
    End If
    
End Sub

Private Sub CargaCheque_Click()
    

    ZZLugar = CargaCheque.ListIndex + 1
    ZZClave = ZCargaCheque(ZZLugar)
    SeleccionaCheque.Visible = False
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Clave = " + "'" + ZZClave + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        FechaCheque.Text = rstRecibos!Fecha2
        CodigoBanco.Text = rstRecibos!CodigoBanco
        SucursalCheque.Text = rstRecibos!SucursalCheque
        TipoCheque.Text = Trim(rstRecibos!TipoCheque)
        ClaseCheque.Text = Trim(rstRecibos!ClaseCheque)
        ImporteCheque.Text = Str$(rstRecibos!Importe2)
        Cuit.Text = IIf(IsNull(rstRecibos!Cuit), "", rstRecibos!Cuit)
        Cuit.Text = Trim(Cuit.Text)
        ImporteCheque.Text = Pusing("###,###.##", ImporteCheque.Text)
        rstRecibos.Close
    End If
    
    
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
    
    FechaCheque.SetFocus
    
End Sub

Private Sub cmdAdd_Click()


        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + recibo.Text + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            rstRecibos.Close
            M1$ = "Recibo ya existente"
            aaaaaa% = MsgBox(M1$, 0, "Ingreso de Recibos")
            Exit Sub
        End If


    Rem If WLicencia <> "1234-5678-ABCD-EFGH" And Val(Recibo.Text) > 10 Then
    Rem     WMsg$ = "La version del sistema es para un uso limitado de movimientos." + Chr$(13) + _
    REM          "El objetivo es el de verificar las opciones y el funcionamiento del mismo." + Chr$(13) + _
    REM          "Para poder utilizar el sistema sin limite de movimientos se debe adquirir la version definitiva."
    Rem     aaaaaa% = MsgBox(WMsg$, 0, "Sistema de Control de Gestion")
    Rem     Exit Sub
    Rem End If
    
    
    
    For IRow = 1 To 100
    
        WTipo = WVector1.TextMatrix(IRow, 1)
        WLetra = WVector1.TextMatrix(IRow, 2)
        WPunto = WVector1.TextMatrix(IRow, 3)
        WNumero = WVector1.TextMatrix(IRow, 4)
        WDebitos = Val(WVector1.TextMatrix(IRow, 5))
        
        If WDebitos <> 0 Then
        
            WClave = WLetra + WTipo + WPunto + WNumero + "01"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                If Trim(UCase(rstCtaCte!Cliente)) <> Trim(UCase(Clientes.Text)) Then
                    M1$ = "La factura " + WNumero + " no pertenece a este cliente"
                    aaaaaa% = MsgBox(M1$, 0, "Ingreso de Recibos")
                    Exit Sub
                End If
                rstCtaCte.Close
            End If
        
        End If
        
    Next IRow

    If recibo.Text <> "" And Fecha.Text <> "" Then
    
        Auxi1 = recibo.Text
        Call Ceros(Auxi1, 6)
        recibo.Text = Auxi1
        
        Existe = "N"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + recibo.Text + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            rstRecibos.Close
            M1$ = "Recibo ya existente"
            aaaaaa% = MsgBox(M1$, 0, "Ingreso de Recibos")
            Existe = "S"
        End If
    
        If Existe <> "S" Then
    
            Call Suma_Datos
        
            ZSumaUs = 0
            Debito = 0
            Credito = 0
            If Val(Debitos.Caption) <> 0 Then
                Debito = Val(Debitos.Caption)
            End If
        
            If Val(Creditos.Caption) <> 0 Then
                Credito = Val(Creditos.Caption)
            End If
        
            Call Redondeo(Debito)
            Call Redondeo(Credito)
            ZZDife = Abs(Debito - Credito)
        
            If ZZDife < 1 Or Tipo2.Value = True Or Tipo3.Value = True Then
    
                Renglon = 0
                For IRow = 1 To 100
        
                    If Tipo1.Value = True Then
                        
                        WRow = IRow
                        WVector1.Col = 5
                        WVector1.Row = IRow
                            
                        If Val(WVector1.Text) <> 0 Then
                        
                            Renglon = Renglon + 1
                            Auxi1 = Str$(Renglon)
                            Call Ceros(Auxi1, 2)
                            
                            ZZRecibo = recibo.Text
                            ZZRenglon = Auxi1
                            ZZCliente = Clientes.Text
                            ZZfecha = Fecha.Text
                            ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            If Tipo1.Value = True Then
                                ZZTipoRec = "1"
                            End If
                            If Tipo2.Value = True Then
                                ZZTipoRec = "2"
                            End If
                            If Tipo3.Value = True Then
                                ZZTipoRec = "3"
                            End If
                            ZZRetGanancias = Retganancias.Text
                            ZZRetIva = RetIva.Text
                            ZZRetOtra = RetOtra.Text
                            ZZRetOtraII = RetOtraII.Text
                            ZZRetOtraIII = RetOtraIII.Text
                            ZZRetOtraIV = RetOtraIV.Text
                            ZZRetSuss = RetSuss.Text
                            ZZNroRetganancias = NroRetganancias.Text
                            ZZNroRetIva = NroRetIva.Text
                            ZZNroRetOtra = NroRetOtra.Text
                            ZZNroRetOtraII = NroRetOtraII.Text
                            ZZNroRetOtraIII = NroRetOtraIII.Text
                            ZZNroRetOtraIV = NroRetOtraIV.Text
                            ZZNroRetSuss = NroRetSuss.Text
                            ZZRetencion = "0"
                            ZZTipoReg = "1"
                            
                            WVector1.Col = 1
                            ZZTipo1 = WVector1.Text
                            WVector1.Col = 2
                            ZZLetra1 = WVector1.Text
                            WVector1.Col = 3
                            ZZPunto1 = WVector1.Text
                            WVector1.Col = 4
                            ZZNumero1 = WVector1.Text
                            Call Ceros(ZZNumero1, 8)
                            WVector1.Col = 5
                            ZZImporte1 = WVector1.Text
                            ZZTipo2 = ""
                            ZZNumero2 = ""
                            ZZFecha2 = ""
                            ZZFechaOrd2 = ""
                            ZZBanco2 = ""
                            ZZImporte2 = 0
                            ZZEstado2 = ""
                            ZZObservaciones = Observaciones.Text
                            ZZEmpresa = WEmpresa
                            ZZClave = ZZRecibo + ZZRenglon
                            ZZImporte = Str$(Credito)
                            ZZCuenta = "1"
                            ZZDestino = ""
                            ZZOrden = "0"
                            ZZDeposito = "0"
                            
                    
                            WLetra = ZZLetra1
                            WTipo = ZZTipo1
                            WPunto = ZZPunto1
                            WNumero = ZZNumero1
                            WImporte = ZZImporte1
                            
                            Auxi$ = Clientes.Text
                            Call Ceros(Auxi$, 6)
                            Claveven$ = Auxi$
                            WClave = WLetra + WTipo + WPunto + WNumero + "01"
                            
                            ZZCodigoBanco = ""
                            ZZSucursalCheque = ""
                            ZZTipoCheque = ""
                            ZZClaseCheque = ""
                            ZZCuit = ""
                            
                            ZSql = ""
                            ZSql = ZSql + "INSERT INTO Recibos ("
                            ZSql = ZSql + "Clave ,"
                            ZSql = ZSql + "Recibo ,"
                            ZSql = ZSql + "Renglon ,"
                            ZSql = ZSql + "Cliente ,"
                            ZSql = ZSql + "Fecha ,"
                            ZSql = ZSql + "FechaOrd ,"
                            ZSql = ZSql + "TipoRec ,"
                            ZSql = ZSql + "RetGanancias ,"
                            ZSql = ZSql + "RetIva ,"
                            ZSql = ZSql + "RetOtra ,"
                            ZSql = ZSql + "RetOtraII ,"
                            ZSql = ZSql + "RetOtraIII ,"
                            ZSql = ZSql + "RetOtraIV ,"
                            ZSql = ZSql + "Retencion ,"
                            ZSql = ZSql + "TipoReg ,"
                            ZSql = ZSql + "Tipo1  ,"
                            ZSql = ZSql + "Letra1 ,"
                            ZSql = ZSql + "Punto1 ,"
                            ZSql = ZSql + "Numero1 ,"
                            ZSql = ZSql + "Importe1 ,"
                            ZSql = ZSql + "Tipo2 ,"
                            ZSql = ZSql + "Numero2 ,"
                            ZSql = ZSql + "Fecha2 ,"
                            ZSql = ZSql + "banco2 ,"
                            ZSql = ZSql + "Importe2 ,"
                            ZSql = ZSql + "Estado2 ,"
                            ZSql = ZSql + "Empresa ,"
                            ZSql = ZSql + "FechaOrd2 ,"
                            ZSql = ZSql + "Importe ,"
                            ZSql = ZSql + "Observaciones ,"
                            ZSql = ZSql + "Impolist ,"
                            ZSql = ZSql + "Impo1list ,"
                            ZSql = ZSql + "Destino ,"
                            ZSql = ZSql + "Cuenta ,"
                            ZSql = ZSql + "Orden ,"
                            ZSql = ZSql + "Deposito ,"
                            ZSql = ZSql + "CodigoBanco ,"
                            ZSql = ZSql + "SucursalCheque ,"
                            ZSql = ZSql + "TipoCheque ,"
                            ZSql = ZSql + "ClaseCheque ,"
                            ZSql = ZSql + "Cuit ,"
                            ZSql = ZSql + "NroRetGanancias ,"
                            ZSql = ZSql + "NroRetIva ,"
                            ZSql = ZSql + "NroRetOtra ,"
                            ZSql = ZSql + "NroRetOtraII ,"
                            ZSql = ZSql + "NroRetOtraIII ,"
                            ZSql = ZSql + "NroRetOtraIV ,"
                            ZSql = ZSql + "RetSuss ,"
                            ZSql = ZSql + "NroRetSuss )"
                            ZSql = ZSql + "Values ("
                            ZSql = ZSql + "'" + ZZClave + "',"
                            ZSql = ZSql + "'" + ZZRecibo + "',"
                            ZSql = ZSql + "'" + ZZRenglon + "',"
                            ZSql = ZSql + "'" + ZZCliente + "',"
                            ZSql = ZSql + "'" + ZZfecha + "',"
                            ZSql = ZSql + "'" + ZZFechaOrd + "',"
                            ZSql = ZSql + "'" + ZZTipoRec + "',"
                            ZSql = ZSql + "'" + ZZRetGanancias + "',"
                            ZSql = ZSql + "'" + ZZRetIva + "',"
                            ZSql = ZSql + "'" + ZZRetOtra + "',"
                            ZSql = ZSql + "'" + ZZRetOtraII + "',"
                            ZSql = ZSql + "'" + ZZRetOtraIII + "',"
                            ZSql = ZSql + "'" + ZZRetOtraIV + "',"
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
                            ZSql = ZSql + "'" + ZZEstado2 + "',"
                            ZSql = ZSql + "'" + ZZEmpresa + "',"
                            ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                            ZSql = ZSql + "'" + ZZImporte + "',"
                            ZSql = ZSql + "'" + ZZObservaciones + "',"
                            ZSql = ZSql + "'" + ZZImpoList + "',"
                            ZSql = ZSql + "'" + ZZImpo1list + "',"
                            ZSql = ZSql + "'" + ZZDestino + "',"
                            ZSql = ZSql + "'" + ZZCuenta + "',"
                            ZSql = ZSql + "'" + ZZOrden + "',"
                            ZSql = ZSql + "'" + ZZDeposito + "',"
                            ZSql = ZSql + "'" + ZZCodigoBanco + "',"
                            ZSql = ZSql + "'" + ZZSucursalCheque + "',"
                            ZSql = ZSql + "'" + ZZTipoCheque + "',"
                            ZSql = ZSql + "'" + ZZClaseCheque + "',"
                            ZSql = ZSql + "'" + ZZCuit + "',"
                            ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                            ZSql = ZSql + "'" + ZZNroRetIva + "',"
                            ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                            ZSql = ZSql + "'" + ZZNroRetOtraII + "',"
                            ZSql = ZSql + "'" + ZZNroRetOtraIII + "',"
                            ZSql = ZSql + "'" + ZZNroRetOtraIV + "',"
                            ZSql = ZSql + "'" + ZZRetSuss + "',"
                            ZSql = ZSql + "'" + ZZNroRetSuss + "')"
                                
                            spRecibos = ZSql
                            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)

                    
                            WLetra = ZZLetra1
                            WTipo = ZZTipo1
                            WPunto = ZZPunto1
                            WNumero = ZZNumero1
                            WImporte = ZZImporte1
                            
                            Auxi$ = Clientes.Text
                            Call Ceros(Auxi$, 6)
                            Claveven$ = Auxi$
                            WClave = WLetra + WTipo + WPunto + WNumero + "01"
                            
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM CtaCte"
                            ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                            spCtaCte = ZSql
                            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                            If rstCtaCte.RecordCount > 0 Then
                                ZSumaUs = ZSumaUs + rstCtaCte!Totalus
                                rstCtaCte.Close
                            End If
                            
                            ZSql = ""
                            ZSql = ZSql + "UPDATE CtaCte SET "
                            ZSql = ZSql + " Saldo = Saldo - " + "'" + ZZImporte1 + "',"
                            ZSql = ZSql + " SaldouS = 0"
                            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                            spCtaCte = ZSql
                            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                    End If
                    
                
                    WVector1.Col = 10
                    WVector1.Row = IRow
                    If Val(WVector1.Text) <> 0 Then
                    
                        Renglon = Renglon + 1
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        
                        ZZRecibo = recibo.Text
                        ZZRenglon = Auxi1
                        ZZCliente = Clientes.Text
                        ZZfecha = Fecha.Text
                        ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        
                        If Tipo1.Value = True Then
                            ZZTipoRec = "1"
                        End If
                        If Tipo2.Value = True Then
                            ZZTipoRec = "2"
                        End If
                        If Tipo3.Value = True Then
                            ZZTipoRec = "3"
                        End If
                        
                        ZZRetGanancias = Retganancias.Text
                        ZZRetIva = RetIva.Text
                        ZZRetOtra = RetOtra.Text
                        ZZRetOtraII = RetOtraII.Text
                        ZZRetOtraIII = RetOtraIII.Text
                        ZZRetOtraIV = RetOtraIV.Text
                        ZZRetSuss = RetSuss.Text
                        ZZNroRetganancias = NroRetganancias.Text
                        ZZNroRetIva = NroRetIva.Text
                        ZZNroRetOtra = NroRetOtra.Text
                        ZZNroRetOtraII = NroRetOtraII.Text
                        ZZNroRetOtraIII = NroRetOtraIII.Text
                        ZZNroRetOtraIV = NroRetOtraIV.Text
                        ZZNroRetSuss = NroRetSuss.Text
                        ZZRetencion = "0"
                        ZZTipoReg = "2"
                        ZZTipo1 = ""
                        ZZLetra1 = ""
                        ZZPunto1 = ""
                        ZZNumero1 = ""
                        ZZImporte1 = "0"
                        
                        WVector1.Col = 6
                        ZZTipo2 = WVector1.Text
                        WVector1.Col = 7
                        ZZNumero2 = WVector1.Text
                        WVector1.Col = 8
                        ZZFecha2 = WVector1.Text
                        ZZFechaOrd2 = Right$(ZZFecha2, 4) + Mid$(ZZFecha2, 4, 2) + Left$(ZZFecha2, 2)
                        ZZPeriodo = Right$(ZZFecha2, 4) + Mid$(ZZFecha2, 4, 2)
                        WVector1.Col = 9
                        ZZBanco2 = Left$(Trim(WVector1.Text), 20)
                        WVector1.Col = 10
                        ZZImporte2 = WVector1.Text
                        ZZEstado2 = "P"
                        ZZObservaciones = Observaciones.Text
                        ZZEmpresa = WEmpresa
                        ZZClave = ZZRecibo + ZZRenglon
                        ZZImporte = Str$(Credito)
                        If ZZTipo2 = 4 Then
                            ZZCuenta = WCuenta(IRow)
                                Else
                            ZCuenta = "1"
                        End If
                        ZZDestino = ""
                        ZZOrden = "0"
                        ZZDeposito = "0"
                        
                        ZZCodigoBanco = WVector1.TextMatrix(WVector1.Row, 11)
                        ZZSucursalCheque = WVector1.TextMatrix(WVector1.Row, 12)
                        ZZTipoCheque = WVector1.TextMatrix(WVector1.Row, 13)
                        ZZClaseCheque = WVector1.TextMatrix(WVector1.Row, 14)
                        ZZClaveLectora = WVector1.TextMatrix(WVector1.Row, 16)
                        ZZCuit = WVector1.TextMatrix(WVector1.Row, 17)
                        ZZNroCuenta = WVector1.TextMatrix(WVector1.Row, 18)
                        ZZFechaEmision = WVector1.TextMatrix(WVector1.Row, 19)

                        Rem ZSql = ""
                        Rem ZSql = ZSql + "Select *"
                        Rem ZSql = ZSql + " FROM Recibos"
                        Rem ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + ZZNumero2 + "'"
                        Rem ZSql = ZSql + " and Recibos.CodigoBanco = " + "'" + ZZCodigoBanco + "'"
                        Rem spRecibos = ZSql
                        Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                        Rem If rstRecibos.RecordCount > 0 Then
                        Rem     rstRecibos.Close
                        Rem     ZZDestino = "Cancelacion por Reingreso de Cheque"
                        Rem     ZZOrden = "999999"
                        Rem     ZZEstado2 = "X"
                        Rem End If
                            
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Recibos ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "Recibo ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Cliente ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "TipoRec ,"
                        ZSql = ZSql + "RetGanancias ,"
                        ZSql = ZSql + "RetIva ,"
                        ZSql = ZSql + "RetOtra ,"
                        ZSql = ZSql + "RetOtraII ,"
                        ZSql = ZSql + "RetOtraIII ,"
                        ZSql = ZSql + "RetOtraIV ,"
                        ZSql = ZSql + "Retencion ,"
                        ZSql = ZSql + "TipoReg ,"
                        ZSql = ZSql + "Tipo1  ,"
                        ZSql = ZSql + "Letra1 ,"
                        ZSql = ZSql + "Punto1 ,"
                        ZSql = ZSql + "Numero1 ,"
                        ZSql = ZSql + "Importe1 ,"
                        ZSql = ZSql + "Tipo2 ,"
                        ZSql = ZSql + "Numero2 ,"
                        ZSql = ZSql + "Fecha2 ,"
                        ZSql = ZSql + "banco2 ,"
                        ZSql = ZSql + "Importe2 ,"
                        ZSql = ZSql + "Estado2 ,"
                        ZSql = ZSql + "Empresa ,"
                        ZSql = ZSql + "FechaOrd2 ,"
                        ZSql = ZSql + "Periodo ,"
                        ZSql = ZSql + "Importe ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Impolist ,"
                        ZSql = ZSql + "Impo1list ,"
                        ZSql = ZSql + "Destino ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Orden ,"
                        ZSql = ZSql + "Deposito ,"
                        ZSql = ZSql + "CodigoBanco ,"
                        ZSql = ZSql + "SucursalCheque ,"
                        ZSql = ZSql + "TipoCheque ,"
                        ZSql = ZSql + "ClaseCheque ,"
                        ZSql = ZSql + "Cuit ,"
                        ZSql = ZSql + "NroCuenta ,"
                        ZSql = ZSql + "FechaEmision ,"
                        ZSql = ZSql + "ClaveLectora ,"
                        ZSql = ZSql + "NroRetGanancias ,"
                        ZSql = ZSql + "NroRetIva ,"
                        ZSql = ZSql + "NroRetOtra ,"
                        ZSql = ZSql + "NroRetOtraII ,"
                        ZSql = ZSql + "NroRetOtraIII ,"
                        ZSql = ZSql + "NroRetOtraIV ,"
                        ZSql = ZSql + "RetSuss ,"
                        ZSql = ZSql + "NroRetSuss )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + ZZClave + "',"
                        ZSql = ZSql + "'" + ZZRecibo + "',"
                        ZSql = ZSql + "'" + ZZRenglon + "',"
                        ZSql = ZSql + "'" + ZZCliente + "',"
                        ZSql = ZSql + "'" + ZZfecha + "',"
                        ZSql = ZSql + "'" + ZZFechaOrd + "',"
                        ZSql = ZSql + "'" + ZZTipoRec + "',"
                        ZSql = ZSql + "'" + ZZRetGanancias + "',"
                        ZSql = ZSql + "'" + ZZRetIva + "',"
                        ZSql = ZSql + "'" + ZZRetOtra + "',"
                        ZSql = ZSql + "'" + ZZRetOtraII + "',"
                        ZSql = ZSql + "'" + ZZRetOtraIII + "',"
                        ZSql = ZSql + "'" + ZZRetOtraIV + "',"
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
                        ZSql = ZSql + "'" + ZZEstado2 + "',"
                        ZSql = ZSql + "'" + ZZEmpresa + "',"
                        ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                        ZSql = ZSql + "'" + ZZPeriodo + "',"
                        ZSql = ZSql + "'" + ZZImporte + "',"
                        ZSql = ZSql + "'" + ZZObservaciones + "',"
                        ZSql = ZSql + "'" + ZZImpoList + "',"
                        ZSql = ZSql + "'" + ZZImpo1list + "',"
                        ZSql = ZSql + "'" + ZZDestino + "',"
                        ZSql = ZSql + "'" + ZZCuenta + "',"
                        ZSql = ZSql + "'" + ZZOrden + "',"
                        ZSql = ZSql + "'" + ZZDeposito + "',"
                        ZSql = ZSql + "'" + ZZCodigoBanco + "',"
                        ZSql = ZSql + "'" + ZZSucursalCheque + "',"
                        ZSql = ZSql + "'" + ZZTipoCheque + "',"
                        ZSql = ZSql + "'" + ZZClaseCheque + "',"
                        ZSql = ZSql + "'" + ZZCuit + "',"
                        ZSql = ZSql + "'" + ZZNroCuenta + "',"
                        ZSql = ZSql + "'" + ZZFechaEmision + "',"
                        ZSql = ZSql + "'" + ZZClaveLectora + "',"
                        ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                        ZSql = ZSql + "'" + ZZNroRetIva + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtraII + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtraIII + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtraIV + "',"
                        ZSql = ZSql + "'" + ZZRetSuss + "',"
                        ZSql = ZSql + "'" + ZZNroRetSuss + "')"
                            
                        spRecibos = ZSql
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)

                    End If
                
                Next IRow
        
                If Tipo1.Value = True Then
                
                    WLetra = "A"
                    WTipo = "06"
                    WPunto = "0000"
                    WNumero = "00" + recibo.Text
                
                    ZZLetra = WLetra
                    ZZTipo = WTipo
                    ZZPunto = WPunto
                    ZZNumero = WNumero
                    ZZRenglon = "01"
                    ZZCliente = Clientes.Text
                    ZZfecha = Fecha.Text
                    ZZEstado = "1"
                    ZZVencimiento = Fecha.Text
                    ZZTotal = Str$(Credito * -1)
                    ZZSaldo = Str$(ZZAnticipo * -1)
                    ZZTotalUs = Str$(ZSumaUs * -1)
                    ZZSaldoUs = Str$(ZZAnticipo * -1)
                    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    ZZOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    ZZImpre = "RC"
                    ZZNeto = "0"
                    ZZIva1 = "0"
                    ZZIva2 = "0"
                    ZZPedido = ""
                    ZZRemito = ""
                    ZZOrden = ""
                    ZZProvincia = WProv
                    ZZVendedor = WVendedor
                    ZZCosto = "0"
                    ZZImporte1 = "0"
                    ZZImporte2 = "0"
                    ZZImporte3 = "0"
                    ZZImporte4 = "0"
                    ZZImporte5 = "0"
                    ZZImporte6 = "0"
                    ZZImporte7 = "0"
                    Auxi = recibo.Text
                    Call Ceros(Auxi, 8)
                    ZZClave = WLetra + WTipo + WPunto + Auxi + "01"
                    ZZBusqueda = ""
                
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
                    ZSql = ZSql + "'" + ZZBusqueda + "')"
                            
                    spCtaCte = ZSql
                    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
        
                If Tipo2.Value = True Then
                
                    WLetra = "A"
                    WTipo = "07"
                    WPunto = "0000"
                    WNumero = "00" + recibo.Text
                    
                    ZZLetra = WLetra
                    ZZTipo = WTipo
                    ZZPunto = WPunto
                    ZZNumero = WNumero
                    ZZRenglon = "01"
                    ZZCliente = Clientes.Text
                    ZZfecha = Fecha.Text
                    ZZEstado = "1"
                    ZZVencimiento = Fecha.Text
                    ZZTotal = Str$(Credito * -1)
                    ZZSaldo = Str$(Credito * -1)
                    ZZTotalUs = Str$(Credito * -1)
                    ZZSaldoUs = Str$(Credito * -1)
                    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    ZZOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    ZZImpre = "AN"
                    ZZNeto = "0"
                    ZZIva1 = "0"
                    ZZIva2 = "0"
                    ZZPedido = ""
                    ZZRemito = ""
                    ZZOrden = ""
                    ZZProvincia = WProv
                    ZZVendedor = WVendedor
                    ZZCosto = "0"
                    ZZImporte1 = "0"
                    ZZImporte2 = "0"
                    ZZImporte3 = "0"
                    ZZImporte4 = "0"
                    ZZImporte5 = "0"
                    ZZImporte6 = "0"
                    ZZImporte7 = "0"
                    Auxi = recibo.Text
                    Call Ceros(Auxi, 8)
                    ZZClave = WLetra + WTipo + WPunto + Auxi + "01"
                    ZZBusqueda = ""
                
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
                    ZSql = ZSql + "'" + ZZBusqueda + "')"
                            
                    spCtaCte = ZSql
                    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                    
                    
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    
                    ZZRecibo = recibo.Text
                    ZZRenglon = Auxi1
                    ZZCliente = Clientes.Text
                    ZZfecha = Fecha.Text
                    ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    If Tipo1.Value = True Then
                        ZZTipoRec = "1"
                    End If
                    If Tipo2.Value = True Then
                        ZZTipoRec = "2"
                    End If
                    ZZRetGanancias = Retganancias.Text
                    ZZRetIva = RetIva.Text
                    ZZRetOtra = RetOtra.Text
                    ZZRetOtraII = RetOtraII.Text
                    ZZRetOtraIII = RetOtraIII.Text
                    ZZRetOtraIV = RetOtraIV.Text
                    ZZRetSuss = RetSuss.Text
                    ZZNroRetganancias = NroRetganancias.Text
                    ZZNroRetIva = NroRetIva.Text
                    ZZNroRetOtra = NroRetOtra.Text
                    ZZNroRetOtraII = NroRetOtraII.Text
                    ZZNroRetOtraIII = NroRetOtraIII.Text
                    ZZNroRetOtraIV = NroRetOtraIV.Text
                    ZZNroRetSuss = NroRetSuss.Text
                    ZZRetencion = "0"
                    ZZTipoReg = "1"
                    ZZTipo1 = "07"
                    ZZLetra1 = ""
                    ZZPunto1 = ""
                    ZZNumero1 = recibo.Text
                    ZZImporte1 = Str$(Credito)
                    ZZTipo2 = ""
                    ZZNumero2 = ""
                    ZZFecha2 = ""
                    ZZFechaOrd2 = ""
                    ZZBanco2 = ""
                    ZZImporte2 = "0"
                    ZZEstado2 = ""
                    ZZObservaciones = Observaciones.Text
                    ZZEmpresa = WEmpresa
                    ZZClave = ZZRecibo + ZZRenglon
                    ZZImporte = Str$(Credito)
                    ZZCuenta = "0"
                    ZZDestino = ""
                    ZZOrden = "0"
                    ZZDeposito = "0"
                    
                    ZZCodigoBanco = ""
                    ZZSucursalCheque = ""
                    ZZTipoCheque = ""
                    ZZClaseCheque = ""
                    ZZCuit = ""
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Recibos ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "Recibo ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Cliente ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "TipoRec ,"
                    ZSql = ZSql + "RetGanancias ,"
                    ZSql = ZSql + "RetIva ,"
                    ZSql = ZSql + "RetOtra ,"
                    ZSql = ZSql + "RetOtraII ,"
                    ZSql = ZSql + "RetOtraIII ,"
                    ZSql = ZSql + "RetOtraIV ,"
                    ZSql = ZSql + "Retencion ,"
                    ZSql = ZSql + "TipoReg ,"
                    ZSql = ZSql + "Tipo1  ,"
                    ZSql = ZSql + "Letra1 ,"
                    ZSql = ZSql + "Punto1 ,"
                    ZSql = ZSql + "Numero1 ,"
                    ZSql = ZSql + "Importe1 ,"
                    ZSql = ZSql + "Tipo2 ,"
                    ZSql = ZSql + "Numero2 ,"
                    ZSql = ZSql + "Fecha2 ,"
                    ZSql = ZSql + "banco2 ,"
                    ZSql = ZSql + "Importe2 ,"
                    ZSql = ZSql + "Estado2 ,"
                    ZSql = ZSql + "Empresa ,"
                    ZSql = ZSql + "FechaOrd2 ,"
                    ZSql = ZSql + "Importe ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Impolist ,"
                    ZSql = ZSql + "Impo1list ,"
                    ZSql = ZSql + "Destino ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Orden ,"
                    ZSql = ZSql + "Deposito ,"
                    ZSql = ZSql + "CodigoBanco ,"
                    ZSql = ZSql + "SucursalCheque ,"
                    ZSql = ZSql + "TipoCheque ,"
                    ZSql = ZSql + "ClaseCheque ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "NroRetGanancias ,"
                    ZSql = ZSql + "NroRetIva ,"
                    ZSql = ZSql + "NroRetOtra ,"
                    ZSql = ZSql + "NroRetOtraII ,"
                    ZSql = ZSql + "NroRetOtraIII ,"
                    ZSql = ZSql + "NroRetOtraIV ,"
                    ZSql = ZSql + "RetSuss ,"
                    ZSql = ZSql + "NroRetSuss )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + ZZClave + "',"
                    ZSql = ZSql + "'" + ZZRecibo + "',"
                    ZSql = ZSql + "'" + ZZRenglon + "',"
                    ZSql = ZSql + "'" + ZZCliente + "',"
                    ZSql = ZSql + "'" + ZZfecha + "',"
                    ZSql = ZSql + "'" + ZZFechaOrd + "',"
                    ZSql = ZSql + "'" + ZZTipoRec + "',"
                    ZSql = ZSql + "'" + ZZRetGanancias + "',"
                    ZSql = ZSql + "'" + ZZRetIva + "',"
                    ZSql = ZSql + "'" + ZZRetOtra + "',"
                    ZSql = ZSql + "'" + ZZRetOtraII + "',"
                    ZSql = ZSql + "'" + ZZRetOtraIII + "',"
                    ZSql = ZSql + "'" + ZZRetOtraIV + "',"
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
                    ZSql = ZSql + "'" + ZZEstado2 + "',"
                    ZSql = ZSql + "'" + ZZEmpresa + "',"
                    ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                    ZSql = ZSql + "'" + ZZImporte + "',"
                    ZSql = ZSql + "'" + ZZObservaciones + "',"
                    ZSql = ZSql + "'" + ZZImpoList + "',"
                    ZSql = ZSql + "'" + ZZImpo1list + "',"
                    ZSql = ZSql + "'" + ZZDestino + "',"
                    ZSql = ZSql + "'" + ZZCuenta + "',"
                    ZSql = ZSql + "'" + ZZOrden + "',"
                    ZSql = ZSql + "'" + ZZDeposito + "',"
                    ZSql = ZSql + "'" + ZZCodigoBanco + "',"
                    ZSql = ZSql + "'" + ZZSucursalCheque + "',"
                    ZSql = ZSql + "'" + ZZTipoCheque + "',"
                    ZSql = ZSql + "'" + ZZClaseCheque + "',"
                    ZSql = ZSql + "'" + ZZCuit + "',"
                    ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                    ZSql = ZSql + "'" + ZZNroRetIva + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtraII + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtraIII + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtraIV + "',"
                    ZSql = ZSql + "'" + ZZRetSuss + "',"
                    ZSql = ZSql + "'" + ZZNroRetSuss + "')"
                            
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
                End If
        
                If Tipo3.Value = True Then
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    
                    ZZRecibo = recibo.Text
                    ZZRenglon = Auxi1
                    ZZCliente = Clientes.Text
                       ZZfecha = Fecha.Text
                    ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    If Tipo1.Value = True Then
                        ZZTipoRec = "1"
                    End If
                    If Tipo2.Value = True Then
                        ZZTipoRec = "2"
                    End If
                    If Tipo3.Value = True Then
                        ZZTipoRec = "3"
                    End If
                    ZZRetGanancias = Retganancias.Text
                    ZZRetIva = RetIva.Text
                    ZZRetOtra = RetOtra.Text
                    ZZRetOtraII = RetOtraII.Text
                    ZZRetOtraIII = RetOtraIII.Text
                    ZZRetOtraIV = RetOtraIV.Text
                    ZZRetSuss = RetSuss.Text
                    ZZNroRetganancias = NroRetganancias.Text
                    ZZNroRetIva = NroRetIva.Text
                    ZZNroRetOtra = NroRetOtra.Text
                    ZZNroRetOtraII = NroRetOtraII.Text
                    ZZNroRetOtraIII = NroRetOtraIII.Text
                    ZZNroRetOtraIV = NroRetOtraIV.Text
                    ZZNroRetSuss = NroRetSuss.Text
                    ZZRetencion = "0"
                    ZZTipoReg = "1"
                    ZZTipo1 = "99"
                    ZZLetra1 = ""
                    ZZPunto1 = ""
                    ZZNumero1 = recibo.Text
                    ZZImporte1 = Str$(Credito)
                    ZZTipo2 = ""
                    ZZNumero2 = ""
                    ZZFecha2 = ""
                    ZZFechaOrd2 = ""
                    ZZBanco2 = ""
                    ZZImporte2 = "0"
                    ZZEstado2 = ""
                    ZZObservaciones = Observaciones.Text
                    ZZEmpresa = WEmpresa
                    ZZClave = ZZRecibo + ZZRenglon
                    ZZImporte = Str$(Credito)
                    ZZCuenta = Cuenta.Text
                    ZZDestino = ""
                    ZZOrden = "0"
                    ZZDeposito = "0"
                    
                    ZZCodigoBanco = ""
                    ZZSucursalCheque = ""
                    ZZTipoCheque = ""
                    ZZClaseCheque = ""
                    ZZCuit = ""
                            
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Recibos ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "Recibo ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Cliente ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "TipoRec ,"
                    ZSql = ZSql + "RetGanancias ,"
                    ZSql = ZSql + "RetIva ,"
                    ZSql = ZSql + "RetOtra ,"
                    ZSql = ZSql + "RetOtraII ,"
                    ZSql = ZSql + "RetOtraIII ,"
                    ZSql = ZSql + "RetOtraIV ,"
                    ZSql = ZSql + "Retencion ,"
                    ZSql = ZSql + "TipoReg ,"
                    ZSql = ZSql + "Tipo1  ,"
                    ZSql = ZSql + "Letra1 ,"
                    ZSql = ZSql + "Punto1 ,"
                    ZSql = ZSql + "Numero1 ,"
                    ZSql = ZSql + "Importe1 ,"
                    ZSql = ZSql + "Tipo2 ,"
                    ZSql = ZSql + "Numero2 ,"
                    ZSql = ZSql + "Fecha2 ,"
                    ZSql = ZSql + "banco2 ,"
                    ZSql = ZSql + "Importe2 ,"
                    ZSql = ZSql + "Estado2 ,"
                    ZSql = ZSql + "Empresa ,"
                    ZSql = ZSql + "FechaOrd2 ,"
                    ZSql = ZSql + "Importe ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Impolist ,"
                    ZSql = ZSql + "Impo1list ,"
                    ZSql = ZSql + "Destino ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Orden ,"
                    ZSql = ZSql + "Deposito ,"
                    ZSql = ZSql + "CodigoBanco ,"
                    ZSql = ZSql + "SucursalCheque ,"
                    ZSql = ZSql + "TipoCheque ,"
                    ZSql = ZSql + "ClaseCheque ,"
                    ZSql = ZSql + "Cuit ,"
                    ZSql = ZSql + "NroRetGanancias ,"
                    ZSql = ZSql + "NroRetIva ,"
                    ZSql = ZSql + "NroRetOtra ,"
                    ZSql = ZSql + "NroRetOtraII ,"
                    ZSql = ZSql + "NroRetOtraIII ,"
                    ZSql = ZSql + "NroRetOtraIV ,"
                    ZSql = ZSql + "RetSuss ,"
                    ZSql = ZSql + "NroRetSuss )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + ZZClave + "',"
                    ZSql = ZSql + "'" + ZZRecibo + "',"
                    ZSql = ZSql + "'" + ZZRenglon + "',"
                    ZSql = ZSql + "'" + ZZCliente + "',"
                    ZSql = ZSql + "'" + ZZfecha + "',"
                    ZSql = ZSql + "'" + ZZFechaOrd + "',"
                    ZSql = ZSql + "'" + ZZTipoRec + "',"
                    ZSql = ZSql + "'" + ZZRetGanancias + "',"
                    ZSql = ZSql + "'" + ZZRetIva + "',"
                    ZSql = ZSql + "'" + ZZRetOtra + "',"
                    ZSql = ZSql + "'" + ZZRetOtraII + "',"
                    ZSql = ZSql + "'" + ZZRetOtraIII + "',"
                    ZSql = ZSql + "'" + ZZRetOtraIV + "',"
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
                    ZSql = ZSql + "'" + ZZEstado2 + "',"
                    ZSql = ZSql + "'" + ZZEmpresa + "',"
                    ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                    ZSql = ZSql + "'" + ZZImporte + "',"
                    ZSql = ZSql + "'" + ZZObservaciones + "',"
                    ZSql = ZSql + "'" + ZZImpoList + "',"
                    ZSql = ZSql + "'" + ZZImpo1list + "',"
                    ZSql = ZSql + "'" + ZZDestino + "',"
                    ZSql = ZSql + "'" + ZZCuenta + "',"
                    ZSql = ZSql + "'" + ZZOrden + "',"
                    ZSql = ZSql + "'" + ZZDeposito + "',"
                    ZSql = ZSql + "'" + ZZCodigoBanco + "',"
                    ZSql = ZSql + "'" + ZZSucursalCheque + "',"
                    ZSql = ZSql + "'" + ZZTipoCheque + "',"
                    ZSql = ZSql + "'" + ZZClaseCheque + "',"
                    ZSql = ZSql + "'" + ZZCuit + "',"
                    ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                    ZSql = ZSql + "'" + ZZNroRetIva + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtraII + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtraIII + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtraIV + "',"
                    ZSql = ZSql + "'" + ZZRetSuss + "',"
                    ZSql = ZSql + "'" + ZZNroRetSuss + "')"
                            
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
        
                T$ = "Ingresos de Recibos"
                M1$ = "Desea imprimir el comprobante"
                Respuestaaaaaa% = MsgBox(M1$, 32 + 4, T$)
                If Respuestaaaaaa% = 6 Then
                    Call Impresion_Recibo
                End If
    
                mm$ = "Grabacion realizada"
                aaa% = MsgBox(mm$, 0, "Archivo de Recibos")

                Call CmdLimpiar_Click
    
                
                recibo.SetFocus
                
                    Else
                    
                M1$ = "Los Valores del Recibo no Balancean"
                aaaaaa% = MsgBox(M1$, 0, "Ingreso de Recibos")
            
            End If
        
        End If
        
    End If
                
    
End Sub

Private Sub CmdDelete_Click()
    If recibo.Text <> "" Then
    
        T$ = "Ingresos de Recibos"
        M1$ = "Desea Anular el Recibo"
        Respuestaaaaaa% = MsgBox(M1$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
                
            For da = 1 To 99
            
                Auxi1 = Str$(da)
                Call Ceros(Auxi1, 2)
                WClave = recibo.Text + Auxi1
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Recibos"
                ZSql = ZSql + " Where Recibos.Clave = " + "'" + WClave + "'"
                spRecibos = ZSql
                Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibos.RecordCount > 0 Then
                
                    WTipoRec = rstRecibos!TipoRec
                    WLetra = rstRecibos!Letra1
                    WTipo = rstRecibos!Tipo1
                    WPunto = rstRecibos!Punto1
                    WNumero = rstRecibos!Numero1
                    WImporte = rstRecibos!Importe1
                    WTipoReg = rstRecibos!Tiporeg
                    
                    rstRecibos.Close
                    
                    If Val(WTipoReg) = 1 Then
                    
                        WClave = WLetra + WTipo + WPunto + WNumero + "01"
                    
                        ZSql = ""
                        ZSql = ZSql + "UPDATE CtaCte SET "
                        ZSql = ZSql + " Saldo = Saldo + " + "'" + Str$(WImporte) + "'"
                        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                        spCtaCte = ZSql
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                End If
                    
            Next da
            
            ZSql = ""
            ZSql = ZSql + "DELETE Recibos"
            ZSql = ZSql + " Where Recibos.Recibo = " + "'" + recibo.Text + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
            If Val(WTipoRec) = 1 Then
            
                WLetra = "A"
                WTipo = "06"
                WPunto = "0000"
                WNumero = "00" + recibo.Text
                WClave = WLetra + WTipo + WPunto + WNumero + "01"
                
                ZSql = ""
                ZSql = ZSql + "DELETE CtaCte"
                ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
        
            If Val(WTipoRec) = 2 Then
            
                WLetra = "A"
                WTipo = "07"
                WPunto = "0000"
                WNumero = "00" + recibo.Text
                WClave = WLetra + WTipo + WPunto + WNumero + "01"
                
                ZSql = ""
                ZSql = ZSql + "DELETE CtaCte"
                ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
            Call CmdLimpiar_Click
                        
        End If
        
    End If
    
    recibo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector

    CmdDelete.Enabled = True
    recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = ""
    RetIva.Text = ""
    RetOtra.Text = ""
    RetOtraII.Text = ""
    RetOtraIII.Text = ""
    RetOtraIV.Text = ""
    RetSuss.Text = ""
    recibo.SetFocus
    Debitos.Caption = ""
    Creditos.Caption = ""
    Diferencia.Caption = ""
    Cuenta.Text = ""
    NroRetganancias.Text = ""
    NroRetOtra.Text = ""
    NroRetOtraII.Text = ""
    NroRetOtraIII.Text = ""
    NroRetOtraIV.Text = ""
    NroRetSuss.Text = ""
    NroRetIva.Text = ""
    ZZAlta = 0
    
    TotalRete.Text = ""
    PantaRete.Visible = False
    
    Ingrecuenta.Visible = False
    Erase WCuenta
    Pantalla.Visible = False
    Opcion.Visible = False
                
    recibo.Text = "1"
    ZSql = ""
    ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where TipoRec = 1 or TipoRec = 2"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        rstRecibos.MoveLast
        ZUltimo = IIf(IsNull(rstRecibos!ReciboMayor), "0", rstRecibos!ReciboMayor)
        recibo.Text = ZUltimo + 1
        rstRecibos.Close
    End If
                
    
End Sub

Private Sub CmdClose_Click()
    PrgRecibos.Hide
    Unload Me
    If ZZSistema = 1 Then
        MenuAdminis.Show
            Else
        MenuVen.Show
    End If
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi1 = recibo.Text
        Call Ceros(Auxi1, 6)
        recibo.Text = Auxi1
        Existe = "N"
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + recibo.Text + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
        
            Existe = "S"
            
            Clientes.Text = rstRecibos!Cliente
            Observaciones.Text = rstRecibos!Observaciones
            Fecha.Text = rstRecibos!Fecha
            Retganancias.Text = Str$(rstRecibos!Retganancias)
            RetIva.Text = Str$(rstRecibos!RetIva)
            RetOtra.Text = Str$(rstRecibos!RetOtra)
            
            ZZZRetOtraII = IIf(IsNull(rstRecibos!RetOtraII), "0", rstRecibos!RetOtraII)
            RetOtraII.Text = Str$(ZZZRetOtraII)
            
            ZZZRetOtraIII = IIf(IsNull(rstRecibos!RetOtraIII), "0", rstRecibos!RetOtraIII)
            RetOtraIII.Text = Str$(ZZZRetOtraIII)
            
            ZZZRetOtraIV = IIf(IsNull(rstRecibos!RetOtraIV), "0", rstRecibos!RetOtraIV)
            RetOtraIV.Text = Str$(ZZZRetOtraIV)
            
            RetSuss.Text = Str$(rstRecibos!RetSuss)
            Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
            RetIva.Text = Alinea("###,###.##", RetIva.Text)
            RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
            RetOtraII.Text = Alinea("###,###.##", RetOtraII.Text)
            RetOtraIII.Text = Alinea("###,###.##", RetOtraIII.Text)
            RetOtraIV.Text = Alinea("###,###.##", RetOtraIV.Text)
            RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
            NroRetganancias.Text = IIf(IsNull(rstRecibos!NroRetganancias), "", rstRecibos!NroRetganancias)
            NroRetIva.Text = IIf(IsNull(rstRecibos!NroRetIva), "", rstRecibos!NroRetIva)
            NroRetOtra.Text = IIf(IsNull(rstRecibos!NroRetOtra), "", rstRecibos!NroRetOtra)
            NroRetOtraII.Text = IIf(IsNull(rstRecibos!NroRetOtraII), "", rstRecibos!NroRetOtraII)
            NroRetOtraIII.Text = IIf(IsNull(rstRecibos!NroRetOtraIII), "", rstRecibos!NroRetOtraIII)
            NroRetOtraIV.Text = IIf(IsNull(rstRecibos!NroRetOtraIV), "", rstRecibos!NroRetOtraIV)
            NroRetSuss.Text = IIf(IsNull(rstRecibos!NroRetSuss), "", rstRecibos!NroRetSuss)
            TotalRete.Text = Str$(Val(Retganancias.Text) + Val(RetOtra.Text) + Val(RetOtraII.Text) + Val(RetOtraIII.Text) + Val(RetOtraIV.Text) + Val(RetIva.Text) + Val(RetSuss.Text))
            TotalRete.Text = Alinea("###,###.##", TotalRete.Text)
            Tipo1.Value = True
            Tipo2.Value = False
            Select Case Val(rstRecibos!TipoRec)
                Case 1
                    Tipo1.Value = True
                Case 2
                    Tipo2.Value = True
                Case 3
                    Tipo3.Value = True
                Case Else
            End Select
            
            rstRecibos.Close
            
        End If
                
        If Existe = "S" Then
            ZZAlta = 1
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
        recibo.Text = ""
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
            Clientes.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Clientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Clientes.Text <> "" Then
        
            Clientes.Text = UCase(Trim(Clientes.Text))
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesClientes.Caption = rstCliente!Fantasia
                WRazon = rstCliente!Fantasia
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WProv = rstCliente!Provincia
                rstCliente.Close
                
                Opcion.Clear
                Opcion.AddItem "Clientes"
                Opcion.AddItem "Cuentas Contables"
                Opcion.AddItem "Cuentas Corrientes"
                Opcion.ListIndex = 2
                Call Opcion_Click
    
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
                
                    Else
                    
                Clientes.SetFocus
                
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        Clientes.Text = ""
        DesClientes.Caption = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
       Observaciones.Text = ""
    End If
End Sub

Private Sub Retganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
        NroRetganancias.SetFocus
    End If
    If KeyAscii = 27 Then
        Retganancias.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetIva.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetganancias.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetIva.Text = Alinea("###,###.##", RetIva.Text)
        NroRetIva.SetFocus
    End If
    If KeyAscii = 27 Then
        RetIva.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtra.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetIva.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetOtra_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
        NroRetOtra.SetFocus
    End If
    If KeyAscii = 27 Then
        RetOtra.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetSuss.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetOtra.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
        NroRetSuss.SetFocus
    End If
    If KeyAscii = 27 Then
        RetSuss.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetsuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtraII.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetSuss.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetOtraII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtraII.Text = Alinea("###,###.##", RetOtraII.Text)
        NroRetOtraII.SetFocus
    End If
    If KeyAscii = 27 Then
        RetOtraII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetOtraII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtraIII.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetOtraII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetOtraIII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtraIII.Text = Alinea("###,###.##", RetOtraIII.Text)
        NroRetOtraIII.SetFocus
    End If
    If KeyAscii = 27 Then
        RetOtraIII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetOtraIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtraIV.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetOtraIII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetOtraIV_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtraIV.Text = Alinea("###,###.##", RetOtraIV.Text)
        NroRetOtraIV.SetFocus
    End If
    If KeyAscii = 27 Then
        RetOtraIV.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetOtraIV_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TotalRete.Text = Str$(Val(Retganancias.Text) + Val(RetOtra.Text) + Val(RetOtraII.Text) + Val(RetOtraIII.Text) + Val(RetOtraIV.Text) + Val(RetIva.Text) + Val(RetSuss.Text))
        TotalRete.Text = Alinea("###,###.##", TotalRete.Text)
        PantaRete.Visible = False
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        NroRetOtraIV.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
                rstCuenta.Close
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
            WCuenta(WVector1.Row) = Cuenta1.Text
            Ingrecuenta.Visible = False
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 6
            Call StartEdit
        End If
    End If
    If KeyAscii = 27 Then
        Cuenta1.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corriestes"
    Opcion.AddItem "Banco"

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
            
            Ayuda.SetFocus
            
        Case 1, 4
            Ayuda.Visible = True
            Ayuda.Text = ""
            
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
            
            Ayuda.SetFocus
           
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Cliente = " + "'" + Clientes.Text + "'"
            ZSql = ZSql + " Order by CtaCte.Numero"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                With rstCtaCte
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            ZSaldo = rstCtaCte!Saldo
                            Call Redondeo(ZSaldo)
                        
                            If ZSaldo <> 0 Then
                                ZSaldo = rstCtaCte!Saldous
                                Call Redondeo(ZSaldo)
                                Auxi2 = Str$(ZSaldo)
                                Auxi2 = Mascara("###,###.##", Auxi2)
                                ZSaldo = rstCtaCte!Saldo
                                Call Redondeo(ZSaldo)
                                Auxi = Str$(ZSaldo)
                                Auxi = Mascara("###,###.##", Auxi)
                                Auxi1 = Str$(rstCtaCte!Numero)
                                Call Ceros(Auxi1, 6)
                                IngresaItem = rstCtaCte!Impre + " " + Auxi1 + " " + rstCtaCte!Fecha + " " + Auxi
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstCtaCte!Clave
                                WIndice.AddItem IngresaItem
                            End If
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCtaCte.Close
            End If
            
        Case 3
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
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Select Case XIndice
        Case 0
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Clientes.Text = WIndice.List(Indice)
            Clientes.Text = UCase(Trim(Clientes.Text))
            Call Clientes_KeyPress(13)
            
        Case 1
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Cuenta.Text = WIndice.List(Indice)
            Cuenta.Text = UCase(Trim(Cuenta.Text))
            Cuenta.SetFocus
            
        Case 4
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Cuenta1.Text = WIndice.List(Indice)
            Cuenta1.Text = UCase(Trim(Cuenta1.Text))
            Cuenta1.SetFocus
            
        Case 3
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            WVector1.Col = 9
            WVector1.Text = WIndice.List(Indice)
            Call StartEdit
            
        Case 2
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For IRow = 1 To 100
                WVector1.Row = IRow
                WVector1.Col = 2
                Compara2 = WVector1.Text
                WVector1.Col = 1
                Compara2 = Compara2 + WVector1.Text
                WVector1.Col = 3
                Compara2 = Compara2 + WVector1.Text
                WVector1.Col = 4
                Compara2 = Compara2 + WVector1.Text + "01"
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next IRow
            
            If Entra = "S" Then
            
                For IRow = 1 To 100
                    WVector1.Row = IRow
                    WVector1.Col = 1
                    If WVector1.Text = "" Then
                        XRow = WVector1.Row
                        Exit For
                    End If
                Next IRow
                
                Indice = Pantalla.ListIndex
                WClave = WIndice.List(Indice)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCte"
                ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
            
                    WVector1.Row = XRow
                    WVector1.Col = 1
                    Auxi = rstCtaCte!Tipo
                    Call Ceros(Auxi, 2)
                    WVector1.Text = Auxi
                    
                    WVector1.Row = XRow
                    WVector1.Col = 2
                    WVector1.Text = rstCtaCte!Letra
                
                    WVector1.Row = XRow
                    WVector1.Col = 3
                    Auxi = rstCtaCte!Punto
                    Call Ceros(Auxi, 4)
                    WVector1.Text = Auxi
                
                    WVector1.Row = XRow
                    WVector1.Col = 4
                    Auxi = rstCtaCte!Numero
                    Call Ceros(Auxi, 8)
                    WVector1.Text = Auxi
                                            
                    WParidad = rstCtaCte!Paridad
                    WSaldo = rstCtaCte!Saldo
                    Rem If WParidad <> 0 Then
                    Rem     WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    Rem End If
                    
                    WVector1.Row = XRow
                    WVector1.Col = 5
                    WVector1.Text = Str$(WSaldo)
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    
                    rstCtaCte.Close
                    
                    WVector1.Row = XRow
                    WVector1.Col = 5
                    
                End If
                    
                Call Suma_Datos
            
            End If
                
            WVector1.Row = XRow
            WVector1.Col = 1
            Call StartEdit
                
        Case Else
    End Select
    
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
            ZSql = ZSql + " Where Cliente.Fantasia LIKE " + "'" + "%" + ZAyuda + "%" + "'"
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
            
        Case 1, 4
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
            
        Case 3
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

Private Sub Form_Load()

    Call Limpia_Vector
 
    Provincia$(0) = "Capital Federal"
    Provincia$(1) = "Buenos Aires"
    Provincia$(2) = "Catamarca"
    Provincia$(3) = "Cordoba"
    Provincia$(4) = "Corrientes"
    Provincia$(5) = "Chaco"
    Provincia$(6) = "Chubut"
    Provincia$(7) = "Entre Rios"
    Provincia$(8) = "Formosa"
    Provincia$(9) = "Jujuy"
    Provincia$(10) = "La Pampa"
    Provincia$(11) = "La Rioja"
    Provincia$(12) = "Mendoza"
    Provincia$(13) = "Misiones"
    Provincia$(14) = "Neuquen"
    Provincia$(15) = "Rio Negro"
    Provincia$(16) = "Salta"
    Provincia$(17) = "San Juan"
    Provincia$(18) = "San Luis"
    Provincia$(19) = "Santa Cruz"
    Provincia$(20) = "Santa Fe"
    Provincia$(21) = "Santiago del Estero"
    Provincia$(22) = "Tucuman"
    Provincia$(23) = "Tierra del Fuego"
    Provincia$(24) = "Exterior"
    Provincia$(25) = ""
     
    ImpreTipo$(1) = "FC"
     
    Tipo1.Value = True
    Tipo2.Value = False
    
    Retganancias.Text = ""
    RetIva.Text = ""
    RetOtra.Text = ""
    RetOtraII.Text = ""
    RetOtraIII.Text = ""
    RetOtraIV.Text = ""
    RetSuss.Text = ""

    recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = ""
    RetIva.Text = ""
    RetOtra.Text = ""
    RetOtraII.Text = ""
    RetOtraIII.Text = ""
    RetOtraIV.Text = ""
    RetSuss.Text = ""
    Debitos.Caption = ""
    Creditos.Caption = ""
    Diferencia.Caption = ""
    Observaciones.Text = ""
    Cuenta.Text = ""
    NroRetganancias.Text = ""
    NroRetOtra.Text = ""
    NroRetOtraII.Text = ""
    NroRetOtraIII.Text = ""
    NroRetOtraIV.Text = ""
    NroRetSuss.Text = ""
    NroRetIva.Text = ""
    ZZAlta = 0
    
    TotalRete.Text = ""
    PantaRete.Visible = False
    
    recibo.Text = "1"
    ZSql = ""
    ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where TipoRec = 1 or TipoRec = 2"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        rstRecibos.MoveLast
        ZUltimo = IIf(IsNull(rstRecibos!ReciboMayor), "0", rstRecibos!ReciboMayor)
        recibo.Text = ZUltimo + 1
        rstRecibos.Close
    End If
            
End Sub

Private Sub Impresion_Click()

    T$ = "Ingresos de Recibos"
    M1$ = "Desea imprimir el comprobante"
    Respuestaaaaaa% = MsgBox(M1$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
        Call Impresion_Recibo
    End If

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

Private Sub TotalRete_DblClick()
    PantaRete.Visible = True
    Retganancias.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If Ingrecuenta.Visible = False Then
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
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If Ingrecuenta.Visible = False Then
                If WControl = "S" Then
                    Call Control_Grilla
                    Call StartEdit
                End If
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
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
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
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case 10
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 6
            
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    
    WVector1.SetFocus
    GridEditText KeyAscii
    
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
    WVector1.Cols = 20
    WVector1.FixedRows = 1
    WVector1.Rows = 201
    
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
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Punto"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "#,###,###.##"
            Case 6
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Numero Ch."
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 20
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 12
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 13
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 14
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 15
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 60
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 16
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 17
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 18
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 19
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
                
            Case Else
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To 10
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
    Rem modificar el tama�o de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub
    
Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Val(WVector1.Text) <> 0 Then
                If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                        Else
                    If Val(WVector1.Text) = 99 Then
                        WVector1.Col = 4
                            Else
                        WControl = "N"
                    End If
                End If
            End If
        Case 2, 3
            WVector1.Col = XColumna
        Case 4
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
            
            WVector1.Col = 2
            Claveven$ = WVector1.Text
            WVector1.Col = 1
            Claveven$ = Claveven$ + WVector1.Text
            WVector1.Col = 3
            Claveven$ = Claveven$ + WVector1.Text
            WVector1.Col = 4
            WClave = Claveven$ + WVector1.Text + "01"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                WVector1.Col = 5
                XRow = WVector1.Row
                If Val(WVector1.Text) = 0 Then
                    WParidad = rstCtaCte!Paridad
                    WSaldo = rstCtaCte!Saldo
                    If WParidad <> 0 Then
                        WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    End If
                    WVector1.Text = WSaldo
                    Call Suma_Datos
                End If
                WVector1.Col = 4
                rstCtaCte.Close
                    Else
                WControl = "N"
            End If
            
            
        Case 5
            WVector1.Col = 2
            Claveven$ = WVector1.Text
            WVector1.Col = 1
            Claveven$ = Claveven$ + WVector1.Text
            ZZTipo = WVector1.Text
            WVector1.Col = 3
            Claveven$ = Claveven$ + WVector1.Text
            WVector1.Col = 4
            WClave = Claveven$ + WVector1.Text + "01"
                    
            If Val(ZZTipo) <> 99 Then
                        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCte"
                ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                    WParidad = rstCtaCte!Paridad
                    WSaldo = rstCtaCte!Saldo
                    If WParidad <> 0 Then
                        WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    End If
                    Saldo = Alinea("###,###.##", Str$(WSaldo))
                    rstCtaCte.Close
                        Else
                    Saldo = 0
                End If
                    
                WVector1.Col = 5
                If Val(WVector1.Text) > Val(Saldo) Then
                    WVector1.Text = ""
                    WControl = "N"
                        Else
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    Call Suma_Datos
                End If
                    
                    Else
            
                WVector1.Text = Alinea("###,###.##", WVector1.Text)
                Call Suma_Datos
            End If
            
        Case 6
            WVector1.Text = Trim(WVector1.Text)
            If Len(WVector1.Text) = 29 Then
                
                WVector1.TextMatrix(WVector1.Row, 7) = Mid$(WVector1.Text, 11, 8)
                WVector1.TextMatrix(WVector1.Row, 8) = ""
                WVector1.TextMatrix(WVector1.Row, 9) = ""
                WVector1.TextMatrix(WVector1.Row, 10) = ""
                WVector1.TextMatrix(WVector1.Row, 11) = Mid$(WVector1.Text, 1, 3)
                WVector1.TextMatrix(WVector1.Row, 12) = Mid$(WVector1.Text, 4, 3)
                WVector1.TextMatrix(WVector1.Row, 13) = ""
                WVector1.TextMatrix(WVector1.Row, 14) = ""
                WVector1.TextMatrix(WVector1.Row, 15) = ""
                WVector1.TextMatrix(WVector1.Row, 16) = WVector1.Text
                WVector1.Text = "2"
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Bcra"
                ZSql = ZSql + " Where Bcra.Codigo = " + "'" + WVector1.TextMatrix(WVector1.Row, 11) + "'"
                spBcra = ZSql
                Set rstBcra = db.OpenRecordset(spBcra, dbOpenSnapshot, dbSQLPassThrough)
                If rstBcra.RecordCount > 0 Then
                    WVector1.TextMatrix(WVector1.Row, 9) = rstBcra!Descripcion
                    rstBcra.Close
                End If
                
                    Else
                    
                If Val(WVector1.Text) = 0 Then
                    WVector1.Text = "2"
                End If
                
            End If
        
            If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Or Val(WVector1.Text) = 99 Then
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 2)
                WVector1.Text = Auxi$
                Select Case Val(WVector1.Text)
                    Case 1, 4, 99
                        WVector1.Col = 7
                        WVector1.Text = ""
                        WVector1.Col = 8
                        WVector1.Text = ""
                        WVector1.Col = 9
                        WVector1.Text = ""
                    Case 2
                        NumeroCheque.Text = WVector1.TextMatrix(WVector1.Row, 7)
                        If Trim(WVector1.TextMatrix(WVector1.Row, 8)) = "" Then
                            FechaCheque.Text = "  /  /    "
                                Else
                            FechaCheque.Text = WVector1.TextMatrix(WVector1.Row, 8)
                        End If
                        DesCodigoBanco.Caption = WVector1.TextMatrix(WVector1.Row, 9)
                        ImporteCheque.Text = WVector1.TextMatrix(WVector1.Row, 10)
                        CodigoBanco.Text = WVector1.TextMatrix(WVector1.Row, 11)
                        SucursalCheque.Text = WVector1.TextMatrix(WVector1.Row, 12)
                        TipoCheque.Text = WVector1.TextMatrix(WVector1.Row, 13)
                        ClaseCheque.Text = WVector1.TextMatrix(WVector1.Row, 14)
                        ClaveLectora.Text = WVector1.TextMatrix(WVector1.Row, 16)
                        Cuit.Text = WVector1.TextMatrix(WVector1.Row, 17)
                        NroCuenta.Text = WVector1.TextMatrix(WVector1.Row, 18)
                        If Len(WVector1.TextMatrix(WVector1.Row, 19)) = 10 Then
                            FechaEmision.Text = WVector1.TextMatrix(WVector1.Row, 19)
                                Else
                            FechaEmision.Text = "  /  /    "
                        End If
                        
                        DatosCheque.Visible = True
                        NumeroCheque.SetFocus
                        WControl = "N"
                    Case Else
                End Select
                    Else
                WControl = "N"
            End If
            
        Case 7
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
            
        Case 8
            Call Valida_fecha1(WVector1.Text, Auxi)
            If Auxi <> "S" Then
                WControl = "N"
            End If
            
        Case 10
            IRow = WVector1.Row
            WVector1.Col = 6
            XTipo = WVector1.Text
            WVector1.Col = 10
            WVector1.Text = Alinea("###,###.##", WVector1.Text)
            Call Suma_Datos
            WVector1.Row = IRow
            Rem If Val(WVector1.TextMatrix(WVector1.Row, 6)) = 4 Then
            Rem     Ingrecuenta.Visible = True
            Rem     Cuenta1.Text = WCuenta(WVector1.Row)
            Rem     Cuenta1.SetFocus
            Rem End If
        
        Case Else
    End Select
End Sub

Private Sub Clientes_DblClick()

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Corrientes"
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Cuenta_DblClick()

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Corrientes"
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub CtaCte_Click()

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Corrientes"
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub Cuenta1_DblClick()

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Corrientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.ListIndex = 4
    
    Call Opcion_Click

End Sub



Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Recibo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Clientes_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetGanancias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetOtra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetOtraII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetOtraIII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetOtraIV_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetSuss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub RetIva_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub

Private Sub NumeroCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZZNumeroCheque = NumeroCheque.Text
        Call Ceros(ZZNumeroCheque, 8)
        NumeroCheque.Text = ZZNumeroCheque
        
        Rem CargaCheque.Clear
        Rem Erase ZCargaCheque
        Rem ZLugarCheque = 0
        
        Rem ZSql = ""
        Rem ZSql = ZSql + "Select *"
        Rem ZSql = ZSql + " FROM Recibos"
        Rem ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + NumeroCheque.Text + "'"
        Rem Rem ZSql = ZSql + " and Recibos.Estado2 = " + "'" + "X" + "'"
        Rem spRecibos = ZSql
        Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstRecibos.RecordCount > 0 Then
        Rem     With rstRecibos
        Rem         .MoveFirst
        Rem         Do
        Rem             If .EOF = False Then
        Rem                 ZLugarCheque = ZLugarCheque + 1
        Rem                 ZCargaCheque(ZLugarCheque) = rstRecibos!Clave
        Rem                 CargaCheque.AddItem rstRecibos!Fecha2 + " " + Str$(rstRecibos!Importe2)
        Rem                 .MoveNext
        Rem                     Else
        Rem                 Exit Do
        Rem             End If
        Rem         Loop
        Rem     End With
        Rem     rstRecibos.Close
        Rem End If
    
        Rem If ZLugarCheque > 1 Then
        Rem     SeleccionaCheque.Visible = True
        Rem
        Rem     Exit Sub
        Rem             Else
        Rem     ZSql = ""
        Rem     ZSql = ZSql + "Select *"
        Rem     ZSql = ZSql + " FROM Recibos"
        Rem     ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + NumeroCheque.Text + "'"
        Rem     Rem ZSql = ZSql + " and Recibos.Estado2 = " + "'" + "X" + "'"
        Rem     spRecibos = ZSql
        Rem     Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        Rem     If rstRecibos.RecordCount > 0 Then
        Rem         FechaCheque.Text = rstRecibos!Fecha2
        Rem         CodigoBanco.Text = rstRecibos!CodigoBanco
        Rem         SucursalCheque.Text = rstRecibos!SucursalCheque
        Rem         TipoCheque.Text = Trim(rstRecibos!TipoCheque)
        Rem         ClaseCheque.Text = Trim(rstRecibos!ClaseCheque)
        Rem         ImporteCheque.Text = Str$(rstRecibos!Importe2)
        Rem         Cuit.Text = IIf(IsNull(rstRecibos!Cuit), "", rstRecibos!Cuit)
        Rem         Cuit.Text = Trim(Cuit.Text)
        Rem         ImporteCheque.Text = Pusing("###,###.##", ImporteCheque.Text)
        Rem         rstRecibos.Close
        Rem         ZSql = ""
        Rem         ZSql = ZSql + "Select *"
        Rem         ZSql = ZSql + " FROM Bcra"
        Rem         ZSql = ZSql + " Where Bcra.Codigo = " + "'" + CodigoBanco.Text + "'"
        Rem         spBcra = ZSql
        Rem         Set rstBcra = db.OpenRecordset(spBcra, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstBcra.RecordCount > 0 Then
        Rem             DesCodigoBanco.Caption = rstBcra!Descripcion
        Rem             rstBcra.Close
        Rem             SucursalCheque.SetFocus
        Rem         End If
        Rem     End If
        Rem End If
                    
        
        FechaEmision.SetFocus
    End If
    If KeyAscii = 27 Then
        NumeroCheque.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaEmision_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(FechaEmision.Text)) = 8 Then
            FechaEmision.Text = Left$(FechaEmision.Text, 6) + "20" + Right$(Trim(FechaEmision.Text), 2)
        End If
        Call Valida_fecha1(FechaEmision.Text, Auxi)
        If Auxi = "S" Then
            FechaCheque.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaEmision.Text = "  /  /    "
    End If
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
        NroCuenta.SetFocus
    End If
    If KeyAscii = 27 Then
        SucursalCheque.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoCheque.SetFocus
    End If
    If KeyAscii = 27 Then
        NroCuenta.Text = ""
    End If
End Sub

Private Sub TipoCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoCheque.Text = UCase(TipoCheque.Text)
        
        If TipoCheque.Text = "1" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cuit.Text = Trim(rstCliente!Cuit)
                rstCliente.Close
            End If
        End If
        
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
        ClaseCheque.Text = UCase(Trim(ClaseCheque.Text))
        If ClaseCheque.Text = "0" Or ClaseCheque.Text = "1" Or ClaseCheque.Text = "2" Then
            Cuit.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        ClaseCheque.Text = ""
    End If
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuit.Text = UCase(Trim(Cuit.Text))
        ImporteCheque.SetFocus
    End If
    If KeyAscii = 27 Then
        Cuit.Text = ""
    End If
End Sub

Private Sub ImporteCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZEntra = "S"
    
        Rem If ZZNivel = 1 Then
        Rem     txtUserName = "SA"
        Rem     txtPassword = "Sw58125812"
        Rem     txtOdbc = "FraganciasII"
        Rem     strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem     Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem End If
                    
        Rem ZSql = ""
        Rem ZSql = ZSql + "Select *"
        Rem ZSql = ZSql + " FROM Recibos"
        Rem ZSql = ZSql + " Where Recibos.Numero2 = " + "'" + NumeroCheque.Text + "'"
        Rem ZSql = ZSql + " and Recibos.CodigoBanco = " + "'" + CodigoBanco.Text + "'"
        Rem spRecibos = ZSql
        Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstRecibos.RecordCount > 0 Then
        Rem     rstRecibos.Close
        Rem     T$ = "Ingresos de Recibos"
        Rem     M1$ = "Cheque ya registrado, confirma su ingreso"
        Rem     Respuestaaaaaa% = MsgBox(M1$, 32 + 4, T$)
        Rem     If Respuestaaaaaa% = 6 Then
        Rem         ZEntra = "S"
        Rem             Else
        Rem         ZEntra = "N"
        Rem     End If
        Rem End If
    
        Rem If ZZNivel = 1 Then
        Rem     txtUserName = "SA"
        Rem     txtPassword = "Sw58125812"
        Rem     txtOdbc = "Fragancias"
        Rem     strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem     Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem End If
        
        If ZEntra = "S" Then
            ImporteCheque.Text = Pusing("###,###.##", ImporteCheque.Text)
            WVector1.TextMatrix(WVector1.Row, 7) = NumeroCheque.Text
            WVector1.TextMatrix(WVector1.Row, 8) = FechaCheque.Text
            WVector1.TextMatrix(WVector1.Row, 9) = DesCodigoBanco.Caption
            WVector1.TextMatrix(WVector1.Row, 10) = ImporteCheque.Text
            WVector1.TextMatrix(WVector1.Row, 11) = CodigoBanco.Text
            WVector1.TextMatrix(WVector1.Row, 12) = SucursalCheque.Text
            WVector1.TextMatrix(WVector1.Row, 13) = TipoCheque.Text
            WVector1.TextMatrix(WVector1.Row, 14) = ClaseCheque.Text
            WVector1.TextMatrix(WVector1.Row, 16) = ClaveLectora.Text
            WVector1.TextMatrix(WVector1.Row, 17) = Cuit.Text
            WVector1.TextMatrix(WVector1.Row, 18) = NroCuenta.Text
            WVector1.TextMatrix(WVector1.Row, 19) = FechaEmision.Text
            DatosCheque.Visible = False
            Call Suma_Datos
            WVector1.Row = WVector1.Row + 1
            WVector1.Col = 6
            Call StartEdit
        End If
        
    End If
    If KeyAscii = 27 Then
        ImporteCheque.Text = ""
    End If
End Sub











Private Sub Impresion_Recibo()

    ZSql = ""
    ZSql = ZSql + "DELETE ImpreRecibo"
    spImpreRecibo = ZSql
    Set rstImpreRecibo = db.OpenRecordset(spImpreRecibo, dbOpenSnapshot, dbSQLPassThrough)

    Erase ZZVectorI
    Erase ZZVectorII
    Erase ZZVectorIII
    
    ZLugarI = 0
    ZLugarII = 0
    ZLugarIII = 0
    
    Call Numtolet
    ZSumaCheque = 0
    
    For IRow = 1 To 100

        WRow = IRow
        
        WVector1.Col = 5
        WVector1.Row = IRow
            
        If Val(WVector1.TextMatrix(IRow, 5)) <> 0 Then
        
            ZLugarI = ZLugarI + 1
            
            ZZVectorI(ZLugarI, 1) = ""
            ZZVectorI(ZLugarI, 2) = WVector1.TextMatrix(IRow, 4)
            ZZVectorI(ZLugarI, 3) = WVector1.TextMatrix(IRow, 5)
            
            ZZTipo1 = WVector1.TextMatrix(IRow, 1)
            ZZLetra1 = WVector1.TextMatrix(IRow, 2)
            ZZPunto1 = WVector1.TextMatrix(IRow, 3)
            ZZNumero1 = WVector1.TextMatrix(IRow, 4)
            
            ZZClave = ZZLetra1 + ZZTipo1 + ZZPunto1 + ZZNumero1 + "01"
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Clave = " + "'" + ZZClave + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                ZZVectorI(ZLugarI, 1) = rstCtaCte!Fecha
                rstCtaCte.Close
            End If
        
        End If
        
        WVector1.Col = 10
        WVector1.Row = IRow
        If Val(WVector1.Text) <> 0 Then
        
            ZZTipo2 = WVector1.TextMatrix(IRow, 6)
            ZZNumero2 = WVector1.TextMatrix(IRow, 7)
            ZZFecha2 = WVector1.TextMatrix(IRow, 8)
            ZZBanco2 = WVector1.TextMatrix(IRow, 9)
            ZZImporte2 = WVector1.TextMatrix(IRow, 10)
            ZZSucursal = WVector1.TextMatrix(IRow, 12)
            
            If Val(ZZTipo2) = 2 Then
            
                ZLugarII = ZLugarII + 1
                
                ZZVectorII(ZLugarII, 1) = ZZBanco2
                ZZVectorII(ZLugarII, 2) = ZZSucursal
                ZZVectorII(ZLugarII, 3) = ZZNumero2
                ZZVectorII(ZLugarII, 4) = ZZFecha2
                ZZVectorII(ZLugarII, 5) = ZZImporte2
                
                ZSumaCheque = ZSumaCheque + Val(ZZImporte2)
                
                    Else
            
                ZLugarIII = ZLugarIII + 1
                
                If Val(ZZTipo2) = 1 Then
                    ZZVectorIII(ZLugarIII, 1) = "Efectivo"
                    ZZVectorIII(ZLugarIII, 2) = ZZImporte2
                        Else
                    If Val(ZZTipo2) = 99 Then
                        ZZVectorIII(ZLugarIII, 1) = "Ajuste"
                        ZZVectorIII(ZLugarIII, 2) = ZZImporte2
                            Else
                        ZZVectorIII(ZLugarIII, 1) = "Compensacion"
                        ZZVectorIII(ZLugarIII, 2) = ZZImporte2
                    End If
                End If
            End If

        End If
    
    Next IRow
    
    If Val(Retganancias.Text) <> 0 Then
        ZLugarIII = ZLugarIII + 1
        ZZVectorIII(ZLugarIII, 1) = "Ret.Ganancias"
        ZZVectorIII(ZLugarIII, 2) = Retganancias.Text
    End If
    
    If Val(RetIva.Text) <> 0 Then
        ZLugarIII = ZLugarIII + 1
        ZZVectorIII(ZLugarIII, 1) = "Ret.Iva"
        ZZVectorIII(ZLugarIII, 2) = RetIva.Text
    End If
    
    ZZZRetOtra = Val(RetOtra.Text) + Val(RetOtraII.Text) + Val(RetOtraIII.Text) + Val(RetOtraIV.Text)
    If ZZZRetOtra <> 0 Then
        ZLugarIII = ZLugarIII + 1
        ZZVectorIII(ZLugarIII, 1) = "Ret.Ing.Brutos"
        ZZVectorIII(ZLugarIII, 2) = Str$(ZZZRetOtra)
    End If
    
    If Val(RetSuss.Text) <> 0 Then
        ZLugarIII = ZLugarIII + 1
        ZZVectorIII(ZLugarIII, 1) = "Ret.SUSS"
        ZZVectorIII(ZLugarIII, 2) = RetSuss.Text
    End If
    
    If Val(ZSumaCheque) <> 0 Then
        ZLugarIII = ZLugarIII + 1
        ZZVectorIII(ZLugarIII, 1) = "Cheques"
        ZZVectorIII(ZLugarIII, 2) = Str$(ZSumaCheque)
    End If

    ZZCanti = ZLugarI
    If ZLugarII > ZZCanti Then
        ZZCanti = ZLugarII
    End If
    If ZLugarIII > ZZCanti Then
        ZZCanti = ZLugarIII
    End If
    
    If ZZCanti < 20 Then
        ZZCanti = 20
            Else
        If ZZCanti < 61 Then
            ZZCanti = 60
                Else
            ZZCanti = 99
        End If
    End If
    
    For ZZCiclo = 1 To ZZCanti
    
        ZZRecibo = recibo.Text
        ZZRenglon = Str$(ZZCiclo)
        ZZfecha = Fecha.Text
        ZZRazon = DesClientes.Caption
        ZZPesosI = XTexto1
        ZZPesosII = XTexto2
        ZZTotal = Debitos.Caption
        If Val(ZZVectorI(ZZCiclo, 3)) <> 0 Then
            ZZFechaI = ZZVectorI(ZZCiclo, 1)
            zZNumeroI = ZZVectorI(ZZCiclo, 2)
            ZZImporteI = ZZVectorI(ZZCiclo, 3)
                Else
            ZZFechaI = ""
            zZNumeroI = ""
            ZZImporteI = ""
        End If
        If Val(ZZVectorII(ZZCiclo, 5)) <> 0 Then
            ZZBanco = ZZVectorII(ZZCiclo, 1)
            ZZSucursal = ZZVectorII(ZZCiclo, 2)
            ZZNumeroII = ZZVectorII(ZZCiclo, 3)
            ZZFechaII = ZZVectorII(ZZCiclo, 4)
            ZZImporteII = ZZVectorII(ZZCiclo, 5)
                Else
            ZZBanco = ""
            ZZSucursal = ""
            ZZNumeroII = ""
            ZZFechaII = ""
            ZZImporteII = ""
        End If
        If Val(ZZVectorIII(ZZCiclo, 2)) <> 0 Then
            ZZEstructura = ZZVectorIII(ZZCiclo, 1)
            ZZImporteIII = ZZVectorIII(ZZCiclo, 2)
                Else
            ZZEstructura = ""
            ZZImporteIII = ""
        End If
    
        ZZCopia = "1"
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreRecibo ("
        ZSql = ZSql + "Copia ,"
        ZSql = ZSql + "Recibo ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "PesosI ,"
        ZSql = ZSql + "PesosII ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "FechaI ,"
        ZSql = ZSql + "NumeroI ,"
        ZSql = ZSql + "ImporteI ,"
        ZSql = ZSql + "Banco ,"
        ZSql = ZSql + "Sucursal ,"
        ZSql = ZSql + "NumeroII  ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "ImporteII ,"
        ZSql = ZSql + "Estructura ,"
        ZSql = ZSql + "ImporteIII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZCopia + "',"
        ZSql = ZSql + "'" + ZZRecibo + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZRazon + "',"
        ZSql = ZSql + "'" + ZZPesosI + "',"
        ZSql = ZSql + "'" + ZZPesosII + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZFechaI + "',"
        ZSql = ZSql + "'" + zZNumeroI + "',"
        ZSql = ZSql + "'" + ZZImporteI + "',"
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + ZZSucursal + "',"
        ZSql = ZSql + "'" + ZZNumeroII + "',"
        ZSql = ZSql + "'" + ZZFechaII + "',"
        ZSql = ZSql + "'" + ZZImporteII + "',"
        ZSql = ZSql + "'" + ZZEstructura + "',"
        ZSql = ZSql + "'" + ZZImporteIII + "')"
            
        spImpreRecibo = ZSql
        Set rstImpreRecibo = db.OpenRecordset(spImpreRecibo, dbOpenSnapshot, dbSQLPassThrough)
        
        
    
        ZZCopia = "2"
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreRecibo ("
        ZSql = ZSql + "Copia ,"
        ZSql = ZSql + "Recibo ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "PesosI ,"
        ZSql = ZSql + "PesosII ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "FechaI ,"
        ZSql = ZSql + "NumeroI ,"
        ZSql = ZSql + "ImporteI ,"
        ZSql = ZSql + "Banco ,"
        ZSql = ZSql + "Sucursal ,"
        ZSql = ZSql + "NumeroII  ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "ImporteII ,"
        ZSql = ZSql + "Estructura ,"
        ZSql = ZSql + "ImporteIII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZCopia + "',"
        ZSql = ZSql + "'" + ZZRecibo + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZRazon + "',"
        ZSql = ZSql + "'" + ZZPesosI + "',"
        ZSql = ZSql + "'" + ZZPesosII + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZFechaI + "',"
        ZSql = ZSql + "'" + zZNumeroI + "',"
        ZSql = ZSql + "'" + ZZImporteI + "',"
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + ZZSucursal + "',"
        ZSql = ZSql + "'" + ZZNumeroII + "',"
        ZSql = ZSql + "'" + ZZFechaII + "',"
        ZSql = ZSql + "'" + ZZImporteII + "',"
        ZSql = ZSql + "'" + ZZEstructura + "',"
        ZSql = ZSql + "'" + ZZImporteIII + "')"
            
        spImpreRecibo = ZSql
        Set rstImpreRecibo = db.OpenRecordset(spImpreRecibo, dbOpenSnapshot, dbSQLPassThrough)
        
        
    Next ZZCiclo
    

    Listado.WindowTitle = "Impresion de Recibo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT ImpreRecibo.Copia, ImpreRecibo.Recibo, ImpreRecibo.Renglon, ImpreRecibo.Fecha, ImpreRecibo.Razon, ImpreRecibo.PesosI, ImpreRecibo.PesosII, ImpreRecibo.Total, ImpreRecibo.FechaI, ImpreRecibo.NumeroI, ImpreRecibo.ImporteI, ImpreRecibo.Banco, ImpreRecibo.Sucursal, ImpreRecibo.NumeroII, ImpreRecibo.FechaII, ImpreRecibo.ImporteII, ImpreRecibo.Estructura, ImpreRecibo.ImporteIII " _
            + "From " _
            + DSQ + ".dbo.ImpreRecibo ImpreRecibo " _
            + "Where  " _
            + "ImpreRecibo.Recibo >= 0 AND " _
            + "ImpreRecibo.Recibo <= 999999"
    
    Listado.Connect = Connect()
    
    Uno = "{ImpreRecibo.Recibo} in " + "0" + " to " + "999999"
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    Listado.Destination = 0
    If ZZCanti = 99 Then
        Listado.ReportFileName = "ImpreReciboII.rpt"
        Listado.Action = 1
            Else
        Listado.ReportFileName = "ImpreRecibo.rpt"
        Listado.Action = 1
    End If

End Sub



Private Sub Numtolet()

    'Convertir en letras el n�mero en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = ""
    sCentimos = "centavos"
    
    Numero = CStr(Val(Debitos.Caption))
    
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

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then
        WTexto1.Text = ""
        WVector1.TextMatrix(WVector1.Row, 1) = ""
        WVector1.TextMatrix(WVector1.Row, 2) = ""
        WVector1.TextMatrix(WVector1.Row, 3) = ""
        WVector1.TextMatrix(WVector1.Row, 4) = ""
        WVector1.TextMatrix(WVector1.Row, 5) = ""
    End If
    If WVector1.Col = 6 Then
        WTexto1.Text = ""
        WVector1.TextMatrix(WVector1.Row, 6) = ""
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
        WTexto2.Text = ""
        WVector1.TextMatrix(WVector1.Row, 1) = ""
        WVector1.TextMatrix(WVector1.Row, 2) = ""
        WVector1.TextMatrix(WVector1.Row, 3) = ""
        WVector1.TextMatrix(WVector1.Row, 4) = ""
        WVector1.TextMatrix(WVector1.Row, 5) = ""
    End If
    If WVector1.Col = 6 Then
        WTexto2.Text = ""
        WVector1.TextMatrix(WVector1.Row, 6) = ""
        WVector1.TextMatrix(WVector1.Row, 7) = ""
        WVector1.TextMatrix(WVector1.Row, 8) = ""
        WVector1.TextMatrix(WVector1.Row, 9) = ""
        WVector1.TextMatrix(WVector1.Row, 10) = ""
        WVector1.TextMatrix(WVector1.Row, 11) = ""
        WVector1.TextMatrix(WVector1.Row, 12) = ""
    End If
    
End Sub
