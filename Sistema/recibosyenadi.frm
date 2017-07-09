VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgRecibosYenadi 
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
      TabIndex        =   72
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox Descuento 
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
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   70
      Text            =   " "
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Frame PantaRete 
      Height          =   2895
      Left            =   1680
      TabIndex        =   44
      Top             =   2520
      Visible         =   0   'False
      Width           =   9015
      Begin VB.TextBox Juridiccion 
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
         MaxLength       =   15
         TabIndex        =   63
         Text            =   " "
         Top             =   1800
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
         Text            =   " "
         Top             =   1320
         Width           =   1335
      End
      Begin MSMask.MaskEdBox FechaRetGanancias 
         Height          =   285
         Left            =   5520
         TabIndex        =   59
         Top             =   840
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
      Begin MSMask.MaskEdBox FechaRetIva 
         Height          =   285
         Left            =   5520
         TabIndex        =   60
         Top             =   1320
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
      Begin MSMask.MaskEdBox FechaRetOtra 
         Height          =   285
         Left            =   5520
         TabIndex        =   61
         Top             =   1800
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
      Begin MSMask.MaskEdBox FechaRetSuss 
         Height          =   285
         Left            =   5520
         TabIndex        =   62
         Top             =   2280
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
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Juridiccion"
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
         Left            =   7320
         TabIndex        =   65
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
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
         Left            =   5520
         TabIndex        =   64
         Top             =   360
         Width           =   1215
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
         TabIndex        =   58
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
         Top             =   360
         Width           =   1815
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
      Index           =   8
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   5760
      Width           =   375
   End
   Begin VB.TextBox Letra 
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
      TabIndex        =   66
      Text            =   " "
      Top             =   480
      Width           =   375
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
      Left            =   5280
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   42
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
      MouseIcon       =   "recibosyenadi.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "recibosyenadi.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   41
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
      MouseIcon       =   "recibosyenadi.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "recibosyenadi.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   40
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
      MouseIcon       =   "recibosyenadi.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "recibosyenadi.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   39
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
      MouseIcon       =   "recibosyenadi.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "recibosyenadi.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   38
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
      MouseIcon       =   "recibosyenadi.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "recibosyenadi.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   37
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
      MouseIcon       =   "recibosyenadi.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "recibosyenadi.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   36
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
      Index           =   7
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   32
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
      TabIndex        =   31
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
      TabIndex        =   30
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
      TabIndex        =   29
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
      TabIndex        =   28
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
      TabIndex        =   26
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
      TabIndex        =   25
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
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   4800
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   19
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
         TabIndex        =   20
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
      TabIndex        =   18
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Crystal.CrystalReport listado 
      Left            =   5880
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
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
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
         Visible         =   0   'False
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
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
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
      TabIndex        =   2
      Top             =   480
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
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   1
      Text            =   " "
      Top             =   480
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
      ItemData        =   "recibosyenadi.frx":4450
      Left            =   6720
      List            =   "recibosyenadi.frx":4457
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   27
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
      TabIndex        =   34
      Top             =   3120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7646
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label5 
      Caption         =   "Total Bonificacion"
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
      TabIndex        =   71
      Top             =   1560
      Width           =   1695
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
      Left            =   4080
      TabIndex        =   68
      Top             =   7560
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   8280
      TabIndex        =   67
      Top             =   7560
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
      TabIndex        =   43
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "COMPROBANTES CANCELADOS"
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
      TabIndex        =   35
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
      TabIndex        =   33
      Top             =   7560
      Visible         =   0   'False
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
      TabIndex        =   17
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
      TabIndex        =   15
      Top             =   840
      Width           =   1455
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
      Left            =   6840
      TabIndex        =   14
      Top             =   7560
      Width           =   1215
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
      Left            =   3000
      TabIndex        =   10
      Top             =   120
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
      Left            =   3480
      TabIndex        =   8
      Top             =   480
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod. Cliente"
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
      Top             =   120
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
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "PrgRecibosYenadi"
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
Dim ZZConversion As Double
Dim WImpo As Double
Dim WIva As Integer

Dim ZZRecibo As String
Dim ZZRenglon As String
Dim ZZCliente As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZFechaRetIva As String
Dim ZZFechaOrdRetIva As String
Dim ZZFechaRetSuss As String
Dim ZZFechaOrdRetSuss As String
Dim ZZFechaRetOtra As String
Dim ZZFechaOrdRetOtra As String
Dim ZZFechaRetGanancias As String
Dim ZZFechaOrdRetGanancias As String
Dim ZZJuridiccion As String
Dim ZZTipoRec As String
Dim ZZRetGanancias As String
Dim ZZRetIva As String
Dim ZZRetOtra As String
Dim ZZRetSuss As String
Dim ZZNroRetganancias As String
Dim ZZNroRetIva As String
Dim ZZNroRetOtra As String
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
Dim WWNeto As Double
Dim WWIva1 As Double
Dim WWIva2 As Double
Dim WWImporte As Double

Dim XTexto1 As String
Dim XTexto2 As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub Suma_Datos()

    Debitos.Caption = ""
    Creditos.Caption = ""
    Diferencia.Caption = ""
    
    ZRetencion = Str$(Val(Retganancias.Text) + Val(RetIva.Text) + Val(RetOtra.Text) + Val(RetSuss.Text))
    ZDescuento = Val(Descuento.Text)
    ZDebito = 0
    ZImpoRecibo = 0
    ZImpoDtoRecibo = 0
    
    If Val(Descuento.Text) <> 0 Then
        WImpo = Val(Descuento.Text)
        ZZPartida = WVector1.TextMatrix(1, 9)
        If Val(WEmpresa) = 1 Then
            If ZZPartida = "V" Then
                WImpo = WImpo * 0.547511
            End If
            If ZZPartida = "W" Then
                WImpo = WImpo * 0.376947
            End If
            If ZZPartida = "M" Then
                WImpo = WImpo * 0.547511
            End If
            If ZZPartida = "Z" Then
                WImpo = WImpo * 0.376947
            End If
        End If
        Call Redondeo(WImpo)
        ZImpoDtoRecibo = WImpo
    End If
    
    For IRow = 1 To 100
        WDebitos = Val(WVector1.TextMatrix(IRow, 7))
        WImpoRecibo = Val(WVector1.TextMatrix(IRow, 8))
        ZDebito = ZDebito + WDebitos
        ZImpoRecibo = ZImpoRecibo + WImpoRecibo
    Next IRow
    
    Debitos.Caption = Str$(ZDebito - Val(ZRetencion) - ZDescuento)
    Debitos.Caption = Alinea("###,###.##", Debitos.Caption)
    
    Rem Diferencia.Caption = Str$(ZImpoRecibo - ZRetencion - ZImpoDtoRecibo)
    Diferencia.Caption = Str$(ZImpoRecibo - ZImpoDtoRecibo)
    Diferencia.Caption = Alinea("###,###.##", Diferencia.Caption)
    
    Creditos.Caption = Alinea("###,###.##", Creditos.Caption)
    
    
End Sub

Private Sub Lee_Datos()

    Call Limpia_Vector
    
    CmdDelete.Enabled = True

    Renglon = 0
    Debito = 0
    Credito = 0
    
    Do
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        WClave = Recibo.Text + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Clave = " + "'" + WClave + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
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
                    WVector1.Text = rstRecibos!ImporteI
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    WVector1.Col = 6
                    WVector1.Text = rstRecibos!ImporteII
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    WVector1.Col = 7
                    WVector1.Text = rstRecibos!Importe1
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    WVector1.Col = 8
                    WVector1.Text = rstRecibos!ImporteIII
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    WVector1.Col = 9
                    WVector1.Text = rstRecibos!Partida
                Case 2
                    Rem Credito = Credito + 1
                    Rem WVector1.Row = Credito
                    Rem WVector1.Col = 6
                    Rem WVector1.Text = rstRecibos!Tipo2
                    Rem WVector1.Col = 7
                    Rem WVector1.Text = rstRecibos!Numero2
                    Rem WVector1.Col = 8
                    Rem WVector1.Text = rstRecibos!Fecha2
                    Rem WVector1.Col = 9
                    Rem WVector1.Text = rstRecibos!Banco2
                    Rem WVector1.Col = 10
                    Rem WVector1.Text = rstRecibos!Importe2
                    Rem WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    Rem If rstRecibos!Estado2 = "X" Then
                    Rem     CmdDelete.Enabled = False
                    Rem End If
                    
                Case Else
            End Select
            rstRecibos.Close
                Else
            Exit Do
        End If
    Loop
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
    If Val(RetSuss.Text) = 0 Then
        RetSuss.Text = "0"
    End If
End Sub

Sub Format_datos()
    Retganancias.Text = Alinea("###,###.##", Retganancias.Text)
    RetIva.Text = Alinea("###,###.##", RetIva.Text)
    RetOtra.Text = Alinea("###,###.##", RetOtra.Text)
    RetSuss.Text = Alinea("###,###.##", RetSuss.Text)
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.razon = " + "'" + Clientes.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Clientes.Text = rstCliente!Cliente
        DesClientes.Caption = rstCliente!Razon
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WProv = rstCliente!Provincia
        WIva = Val(rstCliente!Iva)
        rstCliente.Close
        Call Format_datos
    End If
    
End Sub

Private Sub cmdAdd_Click()

    Rem If WLicencia <> "1234-5678-ABCD-EFGH" And Val(Recibo.Text) > 10 Then
    Rem     WMsg$ = "La version del sistema es para un uso limitado de movimientos." + Chr$(13) + _
    REM          "El objetivo es el de verificar las opciones y el funcionamiento del mismo." + Chr$(13) + _
    REM          "Para poder utilizar el sistema sin limite de movimientos se debe adquirir la version definitiva."
    Rem     A% = MsgBox(WMsg$, 0, "Sistema de Control de Gestion")
    Rem     Exit Sub
    Rem End If
    

    If Recibo.Text <> "" And Fecha.Text <> "" Then
    
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        Existe = "N"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Recibo.Text + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            rstRecibos.Close
            m1$ = "Recibo ya existente"
            A% = MsgBox(m1$, 0, "Ingreso de Recibos")
            Existe = "S"
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
        
            Call Redondeo(Debito)
            Call Redondeo(Credito)
            ZDiferencia = Abs(Debito - Credito)
        
            Rem If ZDiferencia < 0.02 Or Tipo2.Value = True Or Tipo3.Value = True Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cliente"
                ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
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
                    rstCliente.Close
                End If
    
                ZZTotalRecibo = 0
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
                            
                            ZZRecibo = Recibo.Text
                            ZZRenglon = Auxi1
                            ZZCliente = Clientes.Text
                            ZZfecha = Fecha.Text
                            ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            
                            ZZFechaRetGanancias = FechaRetGanancias.Text
                            ZZFechaOrdRetGanancias = Right$(FechaRetGanancias.Text, 4) + Mid$(FechaRetGanancias.Text, 4, 2) + Left$(FechaRetGanancias.Text, 2)
                            ZZFechaRetIva = FechaRetIva.Text
                            ZZFechaOrdRetIva = Right$(FechaRetIva.Text, 4) + Mid$(FechaRetIva.Text, 4, 2) + Left$(FechaRetIva.Text, 2)
                            ZZFechaRetOtra = FechaRetOtra.Text
                            ZZFechaOrdRetOtra = Right$(FechaRetOtra.Text, 4) + Mid$(FechaRetOtra.Text, 4, 2) + Left$(FechaRetOtra.Text, 2)
                            ZZFechaRetSuss = FechaRetSuss.Text
                            ZZFechaOrdRetSuss = Right$(FechaRetSuss.Text, 4) + Mid$(FechaRetSuss.Text, 4, 2) + Left$(FechaRetSuss.Text, 2)
                            ZZJuridiccion = Juridiccion.Text
                            
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
                            ZZRetSuss = RetSuss.Text
                            ZZNroRetganancias = NroRetganancias.Text
                            ZZNroRetIva = NroRetIva.Text
                            ZZNroRetOtra = NroRetOtra.Text
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
                            WVector1.Col = 5
                            ZZImporte1 = WVector1.Text
                            WVector1.Col = 6
                            ZZZImporte2 = WVector1.Text
                            WVector1.Col = 7
                            ZZImporte3 = WVector1.Text
                            WVector1.Col = 8
                            ZZImporte4 = WVector1.Text
                            WVector1.Col = 9
                            ZZPartida = WVector1.Text
                            
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
                            ZZImporte = Str$(Debito)
                            ZZCuenta = "1"
                            ZZDestino = ""
                            ZZOrden = "0"
                            ZZDeposito = "0"
                            ZZDescuento = Descuento.Text
                            
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
                            ZSql = ZSql + "Retencion ,"
                            ZSql = ZSql + "Descuento ,"
                            ZSql = ZSql + "TipoReg ,"
                            ZSql = ZSql + "Tipo1  ,"
                            ZSql = ZSql + "Letra1 ,"
                            ZSql = ZSql + "Punto1 ,"
                            ZSql = ZSql + "Numero1 ,"
                            ZSql = ZSql + "ImporteI ,"
                            ZSql = ZSql + "ImporteII ,"
                            ZSql = ZSql + "Importe1 ,"
                            ZSql = ZSql + "ImporteIII ,"
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
                            ZSql = ZSql + "Partida ,"
                            ZSql = ZSql + "Cuenta ,"
                            ZSql = ZSql + "Orden ,"
                            ZSql = ZSql + "Deposito ,"
                            ZSql = ZSql + "NroRetGanancias ,"
                            ZSql = ZSql + "NroRetIva ,"
                            ZSql = ZSql + "NroRetOtra ,"
                            ZSql = ZSql + "RetSuss ,"
                            ZSql = ZSql + "NroRetSuss ,"
                            ZSql = ZSql + "FechaRetIva ,"
                            ZSql = ZSql + "OrdFechaRetIva ,"
                            ZSql = ZSql + "FechaRetSuss ,"
                            ZSql = ZSql + "OrdFechaRetSuss ,"
                            ZSql = ZSql + "FechaRetOtra ,"
                            ZSql = ZSql + "OrdFechaRetOtra ,"
                            ZSql = ZSql + "FechaRetGanancias ,"
                            ZSql = ZSql + "OrdFechaRetGanancias ,"
                            ZSql = ZSql + "Juridiccion )"
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
                            ZSql = ZSql + "'" + ZZRetencion + "',"
                            ZSql = ZSql + "'" + ZZDescuento + "',"
                            ZSql = ZSql + "'" + ZZTipoReg + "',"
                            ZSql = ZSql + "'" + ZZTipo1 + "',"
                            ZSql = ZSql + "'" + ZZLetra1 + "',"
                            ZSql = ZSql + "'" + ZZPunto1 + "',"
                            ZSql = ZSql + "'" + ZZNumero1 + "',"
                            ZSql = ZSql + "'" + ZZImporte1 + "',"
                            ZSql = ZSql + "'" + ZZZImporte2 + "',"
                            ZSql = ZSql + "'" + ZZImporte3 + "',"
                            ZSql = ZSql + "'" + ZZImporte4 + "',"
                            ZSql = ZSql + "'" + ZZTipo2 + "',"
                            ZSql = ZSql + "'" + ZZNumero2 + "',"
                            ZSql = ZSql + "'" + ZZFecha2 + "',"
                            ZSql = ZSql + "'" + ZZBanco2 + "',"
                            ZSql = ZSql + "'" + ZZImporte2 + "',"
                            ZSql = ZSql + "'" + ZZEstado + "',"
                            ZSql = ZSql + "'" + ZZEmpresa + "',"
                            ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                            ZSql = ZSql + "'" + ZZImporte + "',"
                            ZSql = ZSql + "'" + ZZObservaciones + "',"
                            ZSql = ZSql + "'" + ZZImpoList + "',"
                            ZSql = ZSql + "'" + ZZImpo1list + "',"
                            ZSql = ZSql + "'" + ZZDestino + "',"
                            ZSql = ZSql + "'" + ZZPartida + "',"
                            ZSql = ZSql + "'" + ZZCuenta + "',"
                            ZSql = ZSql + "'" + ZZOrden + "',"
                            ZSql = ZSql + "'" + ZZDeposito + "',"
                            ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                            ZSql = ZSql + "'" + ZZNroRetIva + "',"
                            ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                            ZSql = ZSql + "'" + ZZRetSuss + "',"
                            ZSql = ZSql + "'" + ZZNroRetSuss + "',"
                            ZSql = ZSql + "'" + ZZFechaRetIva + "',"
                            ZSql = ZSql + "'" + ZZFechaOrdRetIva + "',"
                            ZSql = ZSql + "'" + ZZFechaRetSuss + "',"
                            ZSql = ZSql + "'" + ZZFechaOrdRetSuss + "',"
                            ZSql = ZSql + "'" + ZZFechaRetOtra + "',"
                            ZSql = ZSql + "'" + ZZFechaOrdRetOtra + "',"
                            ZSql = ZSql + "'" + ZZFechaRetGanancias + "',"
                            ZSql = ZSql + "'" + ZZFechaOrdRetGanancias + "',"
                            ZSql = ZSql + "'" + ZZJuridiccion + "')"
                                
                            spRecibos = ZSql
                            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)

                    
                            WLetra = ZZLetra1
                            WTipo = ZZTipo1
                            WPunto = ZZPunto1
                            WNumero = ZZNumero1
                            WImporte = ZZImporte3
                            
                            Auxi$ = Clientes.Text
                            Call Ceros(Auxi$, 6)
                            Claveven$ = Auxi$
                            WClave = WLetra + WTipo + WPunto + WNumero + "01"
                                
                            WSaldo = Val(WImporte)
                            If Val(WEmpresa) = 1 Then
                                If ZZPartida = "V" Then
                                    WSaldo = WSaldo * 0.547511
                                End If
                                If ZZPartida = "W" Then
                                    WSaldo = WSaldo * 0.376947
                                End If
                                If ZZPartida = "M" Then
                                    WSaldo = WSaldo * 0.547511
                                End If
                                If ZZPartida = "Z" Then
                                    WSaldo = WSaldo * 0.376947
                                End If
                            End If
                            Call Redondeo(WSaldo)
                            WResta = WSaldo
                            ZZTotalRecibo = ZZTotalRecibo + WResta

                            ZSql = ""
                            ZSql = ZSql + "UPDATE CtaCte SET "
                            ZSql = ZSql + " Saldo = Saldo - " + "'" + Str$(WResta) + "'"
                            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                            spCtaCte = ZSql
                            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        
                        End If
                    End If
                    
                
                    Rem WVector1.Col = 10
                    Rem WVector1.Row = iRow
                    Rem If Val(WVector1.Text) <> 0 Then
                    Rem
                    Rem     Renglon = Renglon + 1
                    Rem     Auxi1 = Str$(Renglon)
                    Rem     Call Ceros(Auxi1, 2)
                    Rem
                    Rem     ZZRecibo = Recibo.Text
                    Rem     ZZRenglon = Auxi1
                    Rem     ZZCliente = Clientes.Text
                    Rem     ZZfecha = Fecha.Text
                    Rem     ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    Rem
                    Rem     ZZFechaRetGanancias = FechaRetGanancias.Text
                    Rem     ZZFechaOrdRetGanancias = Right$(FechaRetGanancias.Text, 4) + Mid$(FechaRetGanancias.Text, 4, 2) + Left$(FechaRetGanancias.Text, 2)
                    Rem     ZZFechaRetIva = FechaRetIva.Text
                    Rem     ZZFechaOrdRetIva = Right$(FechaRetIva.Text, 4) + Mid$(FechaRetIva.Text, 4, 2) + Left$(FechaRetIva.Text, 2)
                    Rem     ZZFechaRetOtra = FechaRetOtra.Text
                    Rem     ZZFechaOrdRetOtra = Right$(FechaRetOtra.Text, 4) + Mid$(FechaRetOtra.Text, 4, 2) + Left$(FechaRetOtra.Text, 2)
                    Rem     ZZFechaRetSuss = FechaRetSuss.Text
                    Rem     ZZFechaOrdRetSuss = Right$(FechaRetSuss.Text, 4) + Mid$(FechaRetSuss.Text, 4, 2) + Left$(FechaRetSuss.Text, 2)
                    Rem     ZZJuridiccion = Juridiccion.Text
                    Rem
                    Rem     If Tipo1.Value = True Then
                    Rem         ZZTipoRec = "1"
                    Rem     End If
                    Rem     If Tipo2.Value = True Then
                    Rem         ZZTipoRec = "2"
                    Rem     End If
                    Rem     If Tipo3.Value = True Then
                    Rem         ZZTipoRec = "3"
                    Rem     End If
                    Rem
                    Rem     ZZRetGanancias = Retganancias.Text
                    Rem     ZZRetIva = RetIva.Text
                    Rem     ZZRetOtra = RetOtra.Text
                    Rem     ZZRetSuss = RetSuss.Text
                    Rem     ZZNroRetganancias = NroRetganancias.Text
                    Rem     ZZNroRetIva = NroRetIva.Text
                    Rem     ZZNroRetOtra = NroRetOtra.Text
                    Rem     ZZNroRetSuss = NroRetSuss.Text
                    Rem     ZZRetencion = "0"
                    Rem     ZZTipoReg = "2"
                    Rem     ZZTipo1 = ""
                    Rem     ZZLetra1 = ""
                    Rem     ZZPunto1 = ""
                    Rem     ZZNumero1 = ""
                    Rem     ZZImporte1 = "0"
                    Rem     ZZPartida = ""
                    Rem
                    Rem     WVector1.Col = 6
                    Rem     ZZTipo2 = WVector1.Text
                    Rem     WVector1.Col = 7
                    Rem     ZZNumero2 = WVector1.Text
                    Rem     WVector1.Col = 8
                    Rem     ZZFecha2 = WVector1.Text
                    Rem     ZZFechaOrd2 = Right$(ZZFecha2, 4) + Mid$(ZZFecha2, 4, 2) + Left$(ZZFecha2, 2)
                    Rem     WVector1.Col = 9
                    Rem     ZZBanco2 = WVector1.Text
                    Rem     WVector1.Col = 10
                    Rem     ZZImporte2 = WVector1.Text
                    Rem     ZZEstado2 = "P"
                    Rem     ZZObservaciones = Observaciones.Text
                    Rem     ZZEmpresa = WEmpresa
                    Rem     ZZClave = ZZRecibo + ZZRenglon
                    Rem     ZZImporte = Str$(Debito)
                    Rem     If ZZTipo2 = 4 Then
                    Rem         ZZCuenta = WCuenta(iRow)
                    Rem             Else
                    Rem         ZCuenta = "1"
                    Rem     End If
                    Rem     ZZDestino = ""
                    Rem     ZZOrden = "0"
                    Rem     ZZDeposito = "0"
                    Rem
                    Rem     ZSql = ""
                    Rem     ZSql = ZSql + "INSERT INTO Recibos ("
                    Rem     ZSql = ZSql + "Clave ,"
                    Rem     ZSql = ZSql + "Recibo ,"
                    Rem     ZSql = ZSql + "Renglon ,"
                    Rem     ZSql = ZSql + "Cliente ,"
                    Rem     ZSql = ZSql + "Fecha ,"
                    Rem     ZSql = ZSql + "FechaOrd ,"
                    Rem     ZSql = ZSql + "TipoRec ,"
                    Rem     ZSql = ZSql + "RetGanancias ,"
                    Rem     ZSql = ZSql + "RetIva ,"
                    Rem     ZSql = ZSql + "RetOtra ,"
                    Rem     ZSql = ZSql + "Retencion ,"
                    Rem     ZSql = ZSql + "TipoReg ,"
                    Rem     ZSql = ZSql + "Tipo1  ,"
                    Rem     ZSql = ZSql + "Letra1 ,"
                    Rem     ZSql = ZSql + "Punto1 ,"
                    Rem     ZSql = ZSql + "Numero1 ,"
                    Rem     ZSql = ZSql + "Importe1 ,"
                    Rem     ZSql = ZSql + "Tipo2 ,"
                    Rem     ZSql = ZSql + "Numero2 ,"
                    Rem     ZSql = ZSql + "Fecha2 ,"
                    Rem     ZSql = ZSql + "banco2 ,"
                    Rem     ZSql = ZSql + "Importe2 ,"
                    Rem     ZSql = ZSql + "Estado2 ,"
                    Rem     ZSql = ZSql + "Empresa ,"
                    Rem     ZSql = ZSql + "FechaOrd2 ,"
                    Rem     ZSql = ZSql + "Importe ,"
                    Rem     ZSql = ZSql + "Observaciones ,"
                    Rem     ZSql = ZSql + "Impolist ,"
                    Rem     ZSql = ZSql + "Impo1list ,"
                    Rem     ZSql = ZSql + "Destino ,"
                    Rem     ZSql = ZSql + "Partida ,"
                    Rem     ZSql = ZSql + "Cuenta ,"
                    Rem     ZSql = ZSql + "Orden ,"
                    Rem     ZSql = ZSql + "Deposito ,"
                    Rem     ZSql = ZSql + "NroRetGanancias ,"
                    Rem     ZSql = ZSql + "NroRetIva ,"
                    Rem     ZSql = ZSql + "NroRetOtra ,"
                    Rem     ZSql = ZSql + "RetSuss ,"
                    Rem     ZSql = ZSql + "NroRetSuss ,"
                    Rem     ZSql = ZSql + "FechaRetIva ,"
                    Rem     ZSql = ZSql + "OrdFechaRetIva ,"
                    Rem     ZSql = ZSql + "FechaRetSuss ,"
                    Rem     ZSql = ZSql + "OrdFechaRetSuss ,"
                    Rem     ZSql = ZSql + "FechaRetOtra ,"
                    Rem     ZSql = ZSql + "OrdFechaRetOtra ,"
                    Rem     ZSql = ZSql + "FechaRetGanancias ,"
                    Rem     ZSql = ZSql + "OrdFechaRetGanancias ,"
                    Rem     ZSql = ZSql + "Juridiccion )"
                    Rem     ZSql = ZSql + "Values ("
                    Rem     ZSql = ZSql + "'" + ZZClave + "',"
                    Rem     ZSql = ZSql + "'" + ZZRecibo + "',"
                    Rem     ZSql = ZSql + "'" + ZZRenglon + "',"
                    Rem     ZSql = ZSql + "'" + ZZCliente + "',"
                    Rem     ZSql = ZSql + "'" + ZZfecha + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaOrd + "',"
                    Rem     ZSql = ZSql + "'" + ZZTipoRec + "',"
                    Rem     ZSql = ZSql + "'" + ZZRetGanancias + "',"
                    Rem     ZSql = ZSql + "'" + ZZRetIva + "',"
                    Rem     ZSql = ZSql + "'" + ZZRetOtra + "',"
                    Rem     ZSql = ZSql + "'" + ZZRetencion + "',"
                    Rem     ZSql = ZSql + "'" + ZZTipoReg + "',"
                    Rem     ZSql = ZSql + "'" + ZZTipo1 + "',"
                    Rem     ZSql = ZSql + "'" + ZZLetra1 + "',"
                    Rem     ZSql = ZSql + "'" + ZZPunto1 + "',"
                    Rem     ZSql = ZSql + "'" + ZZNumero1 + "',"
                    Rem     ZSql = ZSql + "'" + ZZImporte1 + "',"
                    Rem     ZSql = ZSql + "'" + ZZTipo2 + "',"
                    Rem     ZSql = ZSql + "'" + ZZNumero2 + "',"
                    Rem     ZSql = ZSql + "'" + ZZFecha2 + "',"
                    Rem     ZSql = ZSql + "'" + ZZBanco2 + "',"
                    Rem     ZSql = ZSql + "'" + ZZImporte2 + "',"
                    Rem     ZSql = ZSql + "'" + ZZEstado + "',"
                    Rem     ZSql = ZSql + "'" + ZZEmpresa + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
                    Rem     ZSql = ZSql + "'" + ZZImporte + "',"
                    Rem     ZSql = ZSql + "'" + ZZObservaciones + "',"
                    Rem     ZSql = ZSql + "'" + ZZImpoList + "',"
                    Rem     ZSql = ZSql + "'" + ZZImpo1list + "',"
                    Rem     ZSql = ZSql + "'" + ZZDestino + "',"
                    Rem     ZSql = ZSql + "'" + ZZPartida + "',"
                    Rem     ZSql = ZSql + "'" + ZZCuenta + "',"
                    Rem     ZSql = ZSql + "'" + ZZOrden + "',"
                    Rem     ZSql = ZSql + "'" + ZZDeposito + "',"
                    Rem     ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                    Rem     ZSql = ZSql + "'" + ZZNroRetIva + "',"
                    Rem     ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                    Rem     ZSql = ZSql + "'" + ZZRetSuss + "',"
                    Rem     ZSql = ZSql + "'" + ZZNroRetSuss + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaRetIva + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaOrdRetIva + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaRetSuss + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaOrdRetSuss + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaRetOtra + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaOrdRetOtra + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaRetGanancias + "',"
                    Rem     ZSql = ZSql + "'" + ZZFechaOrdRetGanancias + "',"
                    Rem     ZSql = ZSql + "'" + ZZJuridiccion + "')"
                    Rem
                    Rem     spRecibos = ZSql
                    Rem     Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    Rem
                    Rem End If
                
                Next IRow
        
                If Tipo1.Value = True Then
                
                    WWNeto = 0
                    WWIva1 = 0
                    WWIva2 = 0
                
                    If Val(Descuento.Text) <> 0 Then
                        ZImpoDtoRecibo = Val(Descuento.Text)
                        If Val(Descuento.Text) <> 0 Then
                            WImpo = Val(Descuento.Text)
                            ZZPartida = WVector1.TextMatrix(1, 9)
                            If Val(WEmpresa) = 1 Then
                                If ZZPartida = "V" Then
                                    WImpo = WImpo * 0.547511
                                End If
                                If ZZPartida = "W" Then
                                    WImpo = WImpo * 0.376947
                                End If
                                If ZZPartida = "M" Then
                                    WImpo = WImpo * 0.547511
                                End If
                                If ZZPartida = "Z" Then
                                    WImpo = WImpo * 0.376947
                                End If
                            End If
                            Call Redondeo(WImpo)
                            ZImpoDtoRecibo = WImpo
                        End If
            
                        WWImporte = ZImpoDtoRecibo
                        Select Case WIva
                            Case 6
                                WWNeto = WWImporte
                                Call Redondeo(WWNeto)
                                WWIva1 = 0
                                WWIva2 = 0
                            Case Else
                                WWNeto = WWImporte / 1.21
                                Call Redondeo(WWNeto)
                                WWIva1 = WWImporte - WWNeto
                                Call Redondeo(WWIva1)
                                WWIva2 = 0
                            End Select
                        End If
                    End If
                
                
                
                
                
                
                
                
                    ZZLetra = "A"
                    ZZTipo = "06"
                    ZZPunto = "0000"
                    ZZNumero = "00" + Recibo.Text
                    ZZRenglon = "01"
                    ZZCliente = Clientes.Text
                    ZZfecha = Fecha.Text
                    ZZEstado = "1"
                    ZZVencimiento = Fecha.Text
                    ZZTotal = Str$(ZZTotalRecibo * -1)
                    ZZSaldo = "0"
                    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    ZZOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    ZZImpre = "RC"
                    ZZNeto = Str$(WWNeto * -1)
                    ZZNetoTotal = "0"
                    ZZIva1 = Str$(WWIva1 * -1)
                    ZZIva2 = Str$(WWIva2 * -1)
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
                    Auxi = Recibo.Text
                    Call Ceros(Auxi, 8)
                    ZZClave = ZZLetra + ZZTipo + "0000" + Auxi + "01"
                    
                    ZZBusqueda = ""
                    ZZDescuento1 = "0"
                    ZZDescuento2 = "0"
                    ZZDescuento3 = "0"
                    ZZPartida = ""
                    ZZPago = ""
                    ZZLista = ""
                    ZZLinea = ""
                    ZZOCompra = ""
                    ZZCampana = ""
                    ZZDespacho1 = ""
                    ZZDespacho2 = ""
                    ZZBase1 = ""
                    ZZBase2 = ""
                    
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
                    ZSql = ZSql + "'" + ZZBusqueda + "')"
                            
                    spCtaCte = ZSql
                    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
        
                If Tipo2.Value = True Then
                
                    WLetra = "A"
                    WTipo = "07"
                    WPunto = "0000"
                    WNumero = "00" + Recibo.Text
                    
                    ZZLetra = WLetra
                    ZZTipo = WTipo
                    ZZPunto = WPunto
                    ZZNumero = WNumero
                    ZZRenglon = "01"
                    ZZCliente = Clientes.Text
                    ZZfecha = Fecha.Text
                    ZZEstado = "1"
                    ZZVencimiento = Fecha.Text
                    ZZTotal = Str$(Debito * -1)
                    ZZSaldo = Str$(Debito * -1)
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
                    Auxi = Recibo.Text
                    Call Ceros(Auxi, 8)
                    ZZClave = WLetra + WTipo + WPunto + Auxi + "01"
                    
                    ZZBusqueda = ""
                    ZZDescuento1 = "0"
                    ZZDescuento2 = "0"
                    ZZDescuento3 = "0"
                    ZZPartida = ""
                    ZZPago = ""
                    ZZLista = ""
                    ZZLinea = ""
                    ZZOCompra = ""
                    ZZCampana = ""
                    ZZDespacho1 = ""
                    ZZDespacho2 = ""
                    ZZBase1 = ""
                    ZZBase2 = ""
                    
                
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
                    
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    
                    ZZRecibo = Recibo.Text
                    ZZRenglon = Auxi1
                    ZZCliente = Clientes.Text
                    ZZfecha = Fecha.Text
                    ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    
                    ZZFechaRetGanancias = FechaRetGanancias.Text
                    ZZFechaOrdRetGanancias = Right$(FechaRetGanancias.Text, 4) + Mid$(FechaRetGanancias.Text, 4, 2) + Left$(FechaRetGanancias.Text, 2)
                    ZZFechaRetIva = FechaRetIva.Text
                    ZZFechaOrdRetIva = Right$(FechaRetIva.Text, 4) + Mid$(FechaRetIva.Text, 4, 2) + Left$(FechaRetIva.Text, 2)
                    ZZFechaRetOtra = FechaRetOtra.Text
                    ZZFechaOrdRetOtra = Right$(FechaRetOtra.Text, 4) + Mid$(FechaRetOtra.Text, 4, 2) + Left$(FechaRetOtra.Text, 2)
                    ZZFechaRetSuss = FechaRetSuss.Text
                    ZZFechaOrdRetSuss = Right$(FechaRetSuss.Text, 4) + Mid$(FechaRetSuss.Text, 4, 2) + Left$(FechaRetSuss.Text, 2)
                    ZZJuridiccion = Juridiccion.Text
                    
                    If Tipo1.Value = True Then
                        ZZTipoRec = "1"
                    End If
                    If Tipo2.Value = True Then
                        ZZTipoRec = "2"
                    End If
                    ZZRetGanancias = Retganancias.Text
                    ZZRetIva = RetIva.Text
                    ZZRetOtra = RetOtra.Text
                    ZZRetSuss = RetSuss.Text
                    ZZNroRetganancias = NroRetganancias.Text
                    ZZNroRetIva = NroRetIva.Text
                    ZZNroRetOtra = NroRetOtra.Text
                    ZZNroRetSuss = NroRetSuss.Text
                    ZZRetencion = "0"
                    ZZTipoReg = "1"
                    ZZTipo1 = "07"
                    ZZLetra1 = ""
                    ZZPunto1 = ""
                    ZZNumero1 = Recibo.Text
                    ZZImporte1 = Str$(Debito)
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
                    ZZImporte = Str$(Debito)
                    ZZCuenta = "0"
                    ZZDestino = ""
                    ZZOrden = "0"
                    ZZDeposito = "0"
                    
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
                    ZSql = ZSql + "NroRetGanancias ,"
                    ZSql = ZSql + "NroRetIva ,"
                    ZSql = ZSql + "NroRetOtra ,"
                    ZSql = ZSql + "RetSuss ,"
                    ZSql = ZSql + "NroRetSuss ,"
                    ZSql = ZSql + "FechaRetIva ,"
                    ZSql = ZSql + "OrdFechaRetIva ,"
                    ZSql = ZSql + "FechaRetSuss ,"
                    ZSql = ZSql + "OrdFechaRetSuss ,"
                    ZSql = ZSql + "FechaRetOtra ,"
                    ZSql = ZSql + "OrdFechaRetOtra ,"
                    ZSql = ZSql + "FechaRetGanancias ,"
                    ZSql = ZSql + "OrdFechaRetGanancias ,"
                    ZSql = ZSql + "Juridiccion )"
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
                    ZSql = ZSql + "'" + ZZEstado + "',"
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
                    ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                    ZSql = ZSql + "'" + ZZNroRetIva + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                    ZSql = ZSql + "'" + ZZRetSuss + "',"
                    ZSql = ZSql + "'" + ZZNroRetSuss + "',"
                    ZSql = ZSql + "'" + ZZFechaRetIva + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetIva + "',"
                    ZSql = ZSql + "'" + ZZFechaRetSuss + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetSuss + "',"
                    ZSql = ZSql + "'" + ZZFechaRetOtra + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetOtra + "',"
                    ZSql = ZSql + "'" + ZZFechaRetGanancias + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetGanancias + "',"
                    ZSql = ZSql + "'" + ZZJuridiccion + "')"
                            
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
                End If
        
                If Tipo3.Value = True Then
                
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    
                    ZZRecibo = Recibo.Text
                    ZZRenglon = Auxi1
                    ZZCliente = Clientes.Text
                    ZZfecha = Fecha.Text
                    ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            
                    ZZFechaRetGanancias = FechaRetGanancias.Text
                    ZZFechaOrdRetGanancias = Right$(FechaRetGanancias.Text, 4) + Mid$(FechaRetGanancias.Text, 4, 2) + Left$(FechaRetGanancias.Text, 2)
                    ZZFechaRetIva = FechaRetIva.Text
                    ZZFechaOrdRetIva = Right$(FechaRetIva.Text, 4) + Mid$(FechaRetIva.Text, 4, 2) + Left$(FechaRetIva.Text, 2)
                    ZZFechaRetOtra = FechaRetOtra.Text
                    ZZFechaOrdRetOtra = Right$(FechaRetOtra.Text, 4) + Mid$(FechaRetOtra.Text, 4, 2) + Left$(FechaRetOtra.Text, 2)
                    ZZFechaRetSuss = FechaRetSuss.Text
                    ZZFechaOrdRetSuss = Right$(FechaRetSuss.Text, 4) + Mid$(FechaRetSuss.Text, 4, 2) + Left$(FechaRetSuss.Text, 2)
                    ZZJuridiccion = Juridiccion.Text
                    
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
                    ZZRetSuss = RetSuss.Text
                    ZZNroRetganancias = NroRetganancias.Text
                    ZZNroRetIva = NroRetIva.Text
                    ZZNroRetOtra = NroRetOtra.Text
                    ZZNroRetSuss = NroRetSuss.Text
                    ZZRetencion = "0"
                    ZZTipoReg = "1"
                    ZZTipo1 = "99"
                    ZZLetra1 = ""
                    ZZPunto1 = ""
                    ZZNumero1 = Recibo.Text
                    ZZImporte1 = Str$(Debito)
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
                    ZZImporte = Str$(Debito)
                    ZZCuenta = Cuenta.Text
                    ZZDestino = ""
                    ZZOrden = "0"
                    ZZDeposito = "0"
                            
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
                    ZSql = ZSql + "NroRetGanancias ,"
                    ZSql = ZSql + "NroRetIva ,"
                    ZSql = ZSql + "NroRetOtra ,"
                    ZSql = ZSql + "RetSuss ,"
                    ZSql = ZSql + "NroRetSuss ,"
                    ZSql = ZSql + "FechaRetIva ,"
                    ZSql = ZSql + "OrdFechaRetIva ,"
                    ZSql = ZSql + "FechaRetSuss ,"
                    ZSql = ZSql + "OrdFechaRetSuss ,"
                    ZSql = ZSql + "FechaRetOtra ,"
                    ZSql = ZSql + "OrdFechaRetOtra ,"
                    ZSql = ZSql + "FechaRetGanancias ,"
                    ZSql = ZSql + "OrdFechaRetGanancias ,"
                    ZSql = ZSql + "Juridiccion )"
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
                    ZSql = ZSql + "'" + ZZEstado + "',"
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
                    ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                    ZSql = ZSql + "'" + ZZNroRetIva + "',"
                    ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                    ZSql = ZSql + "'" + ZZRetSuss + "',"
                    ZSql = ZSql + "'" + ZZNroRetSuss + "',"
                    ZSql = ZSql + "'" + ZZFechaRetIva + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetIva + "',"
                    ZSql = ZSql + "'" + ZZFechaRetSuss + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetSuss + "',"
                    ZSql = ZSql + "'" + ZZFechaRetOtra + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetOtra + "',"
                    ZSql = ZSql + "'" + ZZFechaRetGanancias + "',"
                    ZSql = ZSql + "'" + ZZFechaOrdRetGanancias + "',"
                    ZSql = ZSql + "'" + ZZJuridiccion + "')"
                            
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
        
                Rem If Val(WEmpresa) = 1 Then
                    Call Impresion_Click
                Rem End If

                Call CmdLimpiar_Click
                Clientes.SetFocus
                
            Rem         Else
            Rem
            Rem    M1$ = "Los Valores del Recibo no Balancean"
            Rem    A% = MsgBox(M1$, 0, "Ingreso de Recibos")
            Rem
            Rem End If
        
        Rem End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Recibo.Text <> "" Then
    
        T$ = "Ingresos de Recibos"
        m1$ = "Desea Anular el Recibo"
        Respuesta% = MsgBox(m1$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            For DA = 1 To 99
            
                Auxi1 = Str$(DA)
                Call Ceros(Auxi1, 2)
                WClave = Recibo.Text + Auxi1
            
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
                    WImporteI = rstRecibos!ImporteI
                    WTipoReg = rstRecibos!Tiporeg
                    
                    rstRecibos.Close
                    
                    If Val(WTipoReg) = 1 Then
                    
                        WClave = WLetra + WTipo + WPunto + WNumero + "01"
                    
                        ZSql = ""
                        ZSql = ZSql + "UPDATE CtaCte SET "
                        ZSql = ZSql + " Saldo = Saldo + " + "'" + Str$(WImporteI) + "'"
                        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                        spCtaCte = ZSql
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                End If
                    
            Next DA
            
            ZSql = ""
            ZSql = ZSql + "DELETE Recibos"
            ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Recibo.Text + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
            If Val(WTipoRec) = 1 Then
            
                WLetra = "A"
                WTipo = "06"
                WPunto = "0000"
                WNumero = "00" + Recibo.Text
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
                WNumero = "00" + Recibo.Text
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
    
    Clientes.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector

    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Observaciones.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = ""
    RetIva.Text = ""
    RetOtra.Text = ""
    RetSuss.Text = ""
    Debitos.Caption = ""
    Creditos.Caption = ""
    Diferencia.Caption = ""
    Cuenta.Text = ""
    NroRetganancias.Text = ""
    NroRetOtra.Text = ""
    NroRetSuss.Text = ""
    NroRetIva.Text = ""
    FechaRetGanancias.Text = "  /  /    "
    FechaRetOtra.Text = "  /  /    "
    FechaRetSuss.Text = "  /  /    "
    FechaRetIva.Text = "  /  /    "
    Juridiccion.Text = ""
    Letra.Text = ""
    
    TotalRete.Text = ""
    Descuento.Text = ""
    PantaRete.Visible = False
    
    Ingrecuenta.Visible = False
    Erase WCuenta
    Pantalla.Visible = False
    Opcion.Visible = False
    
    Recibo.Text = ""
    
    Rem Recibo.Text = "1"
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    Rem ZSql = ZSql + " FROM Recibos"
    Rem spRecibos = ZSql
    Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstRecibos.RecordCount > 0 Then
    Rem     rstRecibos.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstRecibos!ReciboMayor), "0", rstRecibos!ReciboMayor)
    Rem     Recibo.Text = ZUltimo + 1
     Rem    rstRecibos.Close
    Rem End If
    
    Clientes.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgRecibos.Hide
    Unload Me
    Menu.Show
End Sub



Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        Existe = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Recibo.Text + "'"
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
            RetSuss.Text = Str$(rstRecibos!RetSuss)
            NroRetganancias.Text = IIf(IsNull(rstRecibos!NroRetganancias), "", rstRecibos!NroRetganancias)
            NroRetIva.Text = IIf(IsNull(rstRecibos!NroRetIva), "", rstRecibos!NroRetIva)
            NroRetOtra.Text = IIf(IsNull(rstRecibos!NroRetOtra), "", rstRecibos!NroRetOtra)
            NroRetSuss.Text = IIf(IsNull(rstRecibos!NroRetSuss), "", rstRecibos!NroRetSuss)
            TotalRete.Text = Str$(Val(Retganancias.Text) + Val(RetOtra.Text) + Val(RetIva.Text) + Val(RetSuss.Text))
            TotalRete.Text = Alinea("###,###.##", TotalRete.Text)
            Descuento.Text = Str$(rstRecibos!Descuento)
            Descuento.Text = Alinea("###,###.##", Descuento.Text)
            
            FechaRetIva.Text = IIf(IsNull(rstRecibos!FechaRetIva), "  /  /    ", rstRecibos!FechaRetIva)
            FechaRetSuss.Text = IIf(IsNull(rstRecibos!FechaRetSuss), "  /  /    ", rstRecibos!FechaRetSuss)
            FechaRetOtra.Text = IIf(IsNull(rstRecibos!FechaRetOtra), "  /  /    ", rstRecibos!FechaRetOtra)
            FechaRetGanancias.Text = IIf(IsNull(rstRecibos!FechaRetGanancias), "  /  /    ", rstRecibos!FechaRetGanancias)
            Juridiccion.Text = IIf(IsNull(rstRecibos!Juridiccion), "0", rstRecibos!Juridiccion)
            
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
        Recibo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
        
            ZMes = Mid$(Fecha.Text, 4, 2)
            ZAno = Right(Fecha.Text, 4)
    
            Call Ceros(ZMes, 2)
            Call Ceros(ZAno, 4)
    
            ZClave = ZAno + ZMes
            ZEstado = 0
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cierre"
            ZSql = ZSql + " Where Cierre.Clave = " + "'" + ZClave + "'"
            spCierre = ZSql
            Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
            If rstCierre.RecordCount > 0 Then
                ZEstado = rstCierre!Estado
                rstCierre.Close
            End If
    
            If ZEstado = 1 Then
                mm$ = "El mes ya se encuentra cerrado"
                A% = MsgBox(mm$, 0, "Archivo de Ingresos de Cobranzas")
                    Else
                Clientes.SetFocus
            End If
        
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
                WCodIva = rstCliente!Iva
                WIva = Val(rstCliente!Iva)
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
                rstCliente.Close
                
                If Letra.Text = "A" Then
                
                    Recibo.Text = "100001"
                    ZBusqueda = "199999"
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Recibos"
                    ZSql = ZSql + " Where Recibos.Recibo <= " + "'" + "199999" + "'"
                    ZSql = ZSql + " Order by Recibo"
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibos.RecordCount > 0 Then
                        rstRecibos.MoveLast
                        Recibo.Text = rstRecibos!Recibo + 1
                        rstRecibos.Close
                    End If
                    
                        Else
                
                    Recibo.Text = "200001"
                    ZBusqueda = "299999"
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Recibos"
                    ZSql = ZSql + " Where Recibos.Recibo <= " + "'" + "299999" + "'"
                    ZSql = ZSql + " and Recibos.Recibo >= " + "'" + "200001" + "'"
                    ZSql = ZSql + " Order by Recibo"
                    spRecibos = ZSql
                    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstRecibos.RecordCount > 0 Then
                        rstRecibos.MoveLast
                        Recibo.Text = rstRecibos!Recibo + 1
                        rstRecibos.Close
                    End If
                        
                End If
                
                Observaciones.SetFocus
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
        WVector1.Col = 7
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
       Observaciones.Text = ""
    End If
End Sub

Private Sub Retganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Retganancias.Text = Pusing("###,###.##", Retganancias.Text)
        If Val(Retganancias.Text) = 0 Then
            RetIva.SetFocus
                Else
            RetIva.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Retganancias.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetganancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaRetGanancias.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetganancias.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaRetGanancias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(FechaRetGanancias.Text, Auxi)
        If Auxi = "S" Then
            RetIva.SetFocus
                Else
            FechaRetGanancias.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaRetGanancias.Text = "  /  /    "
    End If
End Sub

Private Sub RetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetIva.Text = Pusing("###,###.##", RetIva.Text)
        If Val(RetIva.Text) = 0 Then
            RetOtra.SetFocus
                Else
            RetOtra.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        RetIva.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaRetIva.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetIva.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaRetIva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(FechaRetIva.Text, Auxi)
        If Auxi = "S" Then
            RetOtra.SetFocus
                Else
            FechaRetIva.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaRetIva.Text = "  /  /    "
    End If
End Sub

Private Sub RetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetOtra.Text = Pusing("###,###.##", RetOtra.Text)
        If Val(RetOtra.Text) = 0 Then
            RetSuss.SetFocus
                Else
            RetSuss.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        RetOtra.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaRetOtra.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetOtra.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaRetOtra_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(FechaRetOtra.Text, Auxi)
        If Auxi = "S" Then
            Juridiccion.SetFocus
                Else
            FechaRetOtra.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaRetOtra.Text = "  /  /    "
    End If
End Sub

Private Sub Juridiccion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetSuss.SetFocus
    End If
    If KeyAscii = 27 Then
        Juridiccion.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RetSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        RetSuss.Text = Pusing("###,###.##", RetSuss.Text)
        If Val(RetSuss.Text) = 0 Then
            TotalRete.Text = Str$(Val(Retganancias.Text) + Val(RetOtra.Text) + Val(RetIva.Text) + Val(RetSuss.Text))
            TotalRete.Text = Pusing("###,###.##", TotalRete.Text)
            PantaRete.Visible = False
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            TotalRete.Text = Str$(Val(Retganancias.Text) + Val(RetOtra.Text) + Val(RetIva.Text) + Val(RetSuss.Text))
            TotalRete.Text = Pusing("###,###.##", TotalRete.Text)
            PantaRete.Visible = False
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
        End If
    End If
    If KeyAscii = 27 Then
        RetSuss.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento.Text = Pusing("###,###.##", Descuento.Text)
        Call Suma_Datos
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Descuento.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroRetsuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FechaRetSuss.SetFocus
    End If
    If KeyAscii = 27 Then
        NroRetSuss.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaRetSuss_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(FechaRetSuss.Text, Auxi)
        If Auxi = "S" Then
            TotalRete.Text = Str$(Val(Retganancias.Text) + Val(RetOtra.Text) + Val(RetIva.Text) + Val(RetSuss.Text))
            TotalRete.Text = Alinea("###,###.##", TotalRete.Text)
            PantaRete.Visible = False
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            FechaRetSuss.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaRetSuss.Text = "  /  /    "
    End If
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
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta1.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                rstCuenta.Close
                WCuenta(WVector1.Row) = Cuenta1.Text
                Ingrecuenta.Visible = False
                If WVector1.Row < WVector1.Rows - 1 Then
                    WVector1.Row = WVector1.Row + 1
                End If
                WVector1.Col = 6
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

Private Sub Consulta_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    Opcion.Clear

    Opcion.AddItem "Clientes"
    Rem Opcion.AddItem "Cuentas Contables"
    Rem Opcion.AddItem "Cuenta Corriestes"

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
            
            Ayuda.SetFocus
            
        Case 1, 3
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
            ZSql = ZSql + " Order by CtaCte.Tipo, CtaCte.Numero"
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
                                Auxi = Str$(ZSaldo)
                                Auxi = Mascara("###,###.##", Auxi$)
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
            Call Clientes_KeyPress(13)
            
        Case 1
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Cuenta.Text = WIndice.List(Indice)
            Cuenta.SetFocus
            
        Case 3
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Cuenta1.Text = WIndice.List(Indice)
            Cuenta1.SetFocus
            
        Case 2
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For IRow = 1 To 100
                WVector1.Row = IRow
                WVector1.Col = 1
                Compara2 = WVector1.Text
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
                    
                    
                                            
                    WPartida = rstCtaCte!Partida
                    WParidad = rstCtaCte!Paridad
                    WSaldoOri = rstCtaCte!Saldo
                    WSaldo = rstCtaCte!Saldo
                    Rem If WParidad <> 0 Then
                    Rem     WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    Rem End If
                    
                    If Val(WEmpresa) = 1 Then
                        If WPartida = "V" Then
                            WSaldo = WSaldo * 1.82644629
                        End If
                        If WPartida = "W" Then
                            Rem WSaldo = WSaldo * 1.2655893
                            WSaldo = WSaldo * 2.6528926
                        End If
                        If WPartida = "M" Then
                            WSaldo = WSaldo * 1.82644629
                        End If
                        If WPartida = "Z" Then
                            Rem WSaldo = WSaldo * 1.2655893
                            WSaldo = WSaldo * 2.6528926
                        End If
                    End If
                    
                    Call Redondeo(WSaldo)
                    
                    WVector1.Row = XRow
                    
                    WVector1.Col = 5
                    WVector1.Text = Str$(WSaldoOri)
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    
                    WVector1.Col = 6
                    WVector1.Text = Str$(WSaldo)
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    
                    WVector1.Col = 7
                    WVector1.Text = Str$(WSaldo)
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    
                    WVector1.Col = 8
                    WVector1.Text = Str$(WSaldoOri)
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    
                    WVector1.Col = 9
                    WVector1.Text = rstCtaCte!Partida
                    
                    rstCtaCte.Close
                    
                    Call Suma_Datos
                    
                    WVector1.Row = XRow
                    WVector1.Col = 6
                    
                End If
            
            End If
                
            Call Suma_Datos
            
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
            
        Case 1, 3
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
            
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Form_Load()

    PrgRecibos.Caption = "Ingreso de Recibos: " + WNombreEmpresa
    
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
    RetSuss.Text = ""

    Recibo.Text = ""
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Tipo1.Value = True
    Tipo2.Value = False
    Retganancias.Text = ""
    RetIva.Text = ""
    RetOtra.Text = ""
    RetSuss.Text = ""
    Debitos.Caption = ""
    Creditos.Caption = ""
    Diferencia.Caption = ""
    Observaciones.Text = ""
    Cuenta.Text = ""
    NroRetganancias.Text = ""
    NroRetOtra.Text = ""
    NroRetSuss.Text = ""
    NroRetIva.Text = ""
    FechaRetGanancias.Text = "  /  /    "
    FechaRetOtra.Text = "  /  /    "
    FechaRetSuss.Text = "  /  /    "
    FechaRetIva.Text = "  /  /    "
    Juridiccion.Text = ""
    Letra.Text = ""
    
    TotalRete.Text = ""
    Descuento.Text = ""
    PantaRete.Visible = False
    
    Recibo.Text = ""
    
    Rem Recibo.Text = "1"
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    Rem ZSql = ZSql + " FROM Recibos"
    Rem spRecibos = ZSql
    Rem Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstRecibos.RecordCount > 0 Then
    Rem     rstRecibos.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstRecibos!ReciboMayor), "0", rstRecibos!ReciboMayor)
    Rem     Recibo.Text = ZUltimo + 1
    Rem     rstRecibos.Close
    Rem End If

End Sub

Private Sub Impresion_Click()

    T$ = "Ingresos de Recibos"
    m1$ = "Desea imprimir el comprobante"
    Respuesta% = MsgBox(m1$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        Open "Lpt1" For Output As #99 Len = 255
        Rem Open "Dada.txt" For Output As #99 Len = 255
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
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

        Print #99, Tab(1); Chr$(18);
        Print #99, Tab(34); Fecha.Text;
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, Tab(30); "Vendedor : "; Val(WVendedor)
        Print #99, ""
        Print #99, Left$(WRazon, 20);
        Print #99, Tab(25); WDireccion
        Print #99, Left$(WLocalidad, 20);
        Print #99, "      C.P.:"; WPostal
        Print #99, ""
        Print #99, ""
        Print #99, ""

        If WIva = 1 Or WIva = 2 Then
            Print #99, Tab(1); "Responsable Inscripto";
                Else
            Select Case WIva
                Case 1
                    Print #99, Tab(20); "X";
                Case 4
                    Print #99, Tab(13); "X";
                Case 5
                    Print #99, Tab(7); "X";
                Case Else
            End Select
        End If

        Print #99, Tab(30); Left$(WCuit, 15);
        Print #99, Tab(47); Clientes.Text
        Print #99, ""
        Print #99, ""
        Print #99, ""
        
        Rem m# = Impresion#
        Rem GoSub 4630
        ZZConversion = Val(Diferencia.Caption)
        Call Numtolet
        
        Print #99, Tab(1); "PESOS : "; XTexto1
        Print #99, Tab(1); XTexto2
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""

        WTotal = 0
        ZPasa = 0

        For Ciclo = 1 To 9
        
            If Val(WVector1.TextMatrix(Ciclo, 1)) <> 0 Then
                Select Case Val(WVector1.TextMatrix(Ciclo, 1))
                    Case 1, 3
                        Print #99, Tab(1); "FACTURA";
                    Case 2, 5
                        Print #99, Tab(1); "NOTA DE CREDITO";
                    Case Else
                        Print #99, Tab(1); "NOTA DE DEBITO";
                End Select
                Impre = Pusing("######", WVector1.TextMatrix(Ciclo, 4))
                Print #99, Tab(20); Impre;
                Impre = Pusing("###,###,###.##", WVector1.TextMatrix(Ciclo, 5))
                Print #99, Tab(30); Impre;
                
                    Else
            
                If ZPasa = 0 And Val(Descuento.Text) <> 0 Then
                    ZPasa = 1
                    
                    ZImpoDtoRecibo = Val(Descuento.Text)
                    If Val(Descuento.Text) <> 0 Then
                        WImpo = Val(Descuento.Text)
                        ZZPartida = WVector1.TextMatrix(1, 9)
                        If Val(WEmpresa) = 1 Then
                            If ZZPartida = "V" Then
                                WImpo = WImpo * 0.547511
                            End If
                            If ZZPartida = "W" Then
                                WImpo = WImpo * 0.376947
                            End If
                            If ZZPartida = "M" Then
                                WImpo = WImpo * 0.547511
                            End If
                            If ZZPartida = "Z" Then
                                WImpo = WImpo * 0.376947
                            End If
                        End If
                        Call Redondeo(WImpo)
                        ZImpoDtoRecibo = WImpo
                    End If
                    
                    Print #99, Tab(1); "BONIFICACION";
                    Impre = Pusing("######", Right$(Recibo.Text, 5))
                    Print #99, Tab(20); Impre;
                    Impre = Pusing("###,###,###.##", Str$(ZImpoDtoRecibo))
                    Print #99, Tab(30); Impre;
                        
                End If
            
            End If
            
            Print #99, ""
            
        Next Ciclo

        If Val(RetIva.Text) <> 0 Then
            Print #99, Using; "Ret. Iva : ";
            Impre = Pusing("###,###,###.##", RetIva.Text)
            Print #99, Impre;
        End If
        
        If Val(RetSuss.Text) <> 0 Then
            Print #99, Using; "  Ret.Suss : ";
            Impre = Pusing("###,###,###.##", RetSuss.Text)
            Print #99, Impre
                Else
            Print #99, ""
        End If

        If Val(Retganancias.Text) <> 0 Then
            Print #99, Using; "Ret.Ganancia : ";
            Impre = Pusing("###,###,###.##", Retganancias.Text)
            Print #99, Impre
                Else
            Print #99, ""
        End If

        If Val(RetOtra.Text) <> 0 Then
            Print #99, Using; "Ret.Ingresos Brutos : ";
            Impre = Pusing("###,###,###.##", RetOtra.Text)
            Print #99, Impre
                Else
            Print #99, ""
        End If


        WWNeto = 0
        WWIva1 = 0
        WWIva2 = 0
        WWImporte = 0
        
        If Val(Descuento.Text) <> 0 Then
            Signo$ = " "
            
            ZImpoDtoRecibo = Val(Descuento.Text)
            If Val(Descuento.Text) <> 0 Then
                WImpo = Val(Descuento.Text)
                ZZPartida = WVector1.TextMatrix(1, 9)
                If Val(WEmpresa) = 1 Then
                    If ZZPartida = "V" Then
                        WImpo = WImpo * 0.547511
                    End If
                    If ZZPartida = "W" Then
                        WImpo = WImpo * 0.376947
                    End If
                    If ZZPartida = "M" Then
                        WImpo = WImpo * 0.547511
                    End If
                    If ZZPartida = "Z" Then
                        WImpo = WImpo * 0.376947
                    End If
                End If
                Call Redondeo(WImpo)
                ZImpoDtoRecibo = WImpo
            End If
            
            WWImporte = ZImpoDtoRecibo
            Select Case WIva
                Case 1, 2
                    WWNeto = WWImporte / 1.21
                    Call Redondeo(WWNeto)
                    WWIva1 = WWImporte - WWNeto
                    Call Redondeo(WWIva1)
                    WWIva2 = 0
                Case Else
                    WWNeto = WWImporte
                    Call Redondeo(WWNeto)
                    WWIva1 = 0
                    WWIva2 = 0
            End Select
        End If

        If WIva = 1 Or WIva = 2 Then
        
            Auxi1 = Str$(WWImporte - WWIva1 - WWIva2)
            Impre = Pusing("###,###,###.##", Auxi1)
            Print #99, Tab(12); Impre;
            
            Auxi1 = Diferencia.Caption
            Impre = Pusing("###,###,###.##", Auxi1)
            Print #99, Tab(26); " $ "; Impre
            
            Print #99, ""
            Print #99, ""
            
            Auxi1 = Str$(WWIva1)
            Impre = Pusing("###,###,###.##", Auxi1)
            Print #99, Tab(12); Impre
            
            Print #99, ""
            Print #99, ""
            
            Auxi1 = Str$(WWIva2)
            Impre = Pusing("###,###,###.##", Auxi1)
            Print #99, Tab(12); Impre
            
            Print #99, ""
            
            Auxi1 = Str$(WWImporte)
            Impre = Pusing("###,###,###.##", Auxi1)
            Print #99, Tab(12); Impre
            
                Else
                
            Print #99, ""
            
            Auxi1 = Diferencia.Caption
            Impre = Pusing("###,###,###.##", Auxi1)
            Print #99, Tab(26); Impre
            
            Print #99, ""
            Print #99, ""
            
            Auxi1 = Str$(WWImporte)
            Impre = Pusing("###,###,###.##", Auxi1)
            Print #99, Tab(12); Impre
            
            Print #99, ""
            Print #99, ""
            Print #99, ""
            Print #99, ""
            Print #99, ""
            
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
        Case 7
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
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Letra"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Punto"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "#,###,###.##"
            Case 6
                WVector1.Text = "Total"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Cancela"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "$ Recibo"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Partida"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 1
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
        Case 7
            If Val(WVector1.Text) > Val(WVector1.TextMatrix(WVector1.Row, 6)) Then
                WVector1.Text = ""
                    Else
                WImpo = Val(WVector1.Text)
                ZZPartida = WVector1.TextMatrix(1, 9)
                If Val(WEmpresa) = 1 Then
                    If ZZPartida = "V" Then
                        WImpo = WImpo * 0.547511
                    End If
                    If ZZPartida = "W" Then
                        WImpo = WImpo * 0.376947
                    End If
                    If ZZPartida = "M" Then
                        WImpo = WImpo * 0.547511
                    End If
                    If ZZPartida = "Z" Then
                        WImpo = WImpo * 0.376947
                    End If
                End If
                Call Redondeo(WImpo)
                WVector1.TextMatrix(WVector1.Row, 8) = Str$(WImpo)
                WVector1.TextMatrix(WVector1.Row, 8) = Pusing("###,###.##", WVector1.TextMatrix(WVector1.Row, 8))
            End If
            Call Suma_Datos
    
        Case 101
            If Val(WVector1.Text) <> 0 Then
                If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                        Else
                    WControl = "N"
                End If
            End If
            
        Case 102, 103
            WVector1.Col = XColumna
            
        Case 104
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
                XRow = WVector1.Row
                If Val(WVector1.Text) = 0 Then
                    WParidad = rstCtaCte!Paridad
                    WSaldo = rstCtaCte!Saldo
                    If WParidad <> 0 Then
                        WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    End If
                    WVector1.Col = 5
                    WVector1.Text = WSaldo
                    WVector1.Col = 6
                    WVector1.Text = WSaldo
                    WVector1.Col = 7
                    WVector1.Text = WSaldo
                    Call Suma_Datos
                End If
                WVector1.Col = 9
                WVector1.Text = rstCtaCte!Partida
                WVector1.Col = 6
                rstCtaCte.Close
                    Else
                WControl = "N"
            End If
            
        Case 105
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
                WParidad = rstCtaCte!Paridad
                WSaldo = rstCtaCte!Saldo
                If WParidad <> 0 Then
                    WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                End If
                Saldo = Alinea("###,###.##", Str$(WSaldo))
                WVector1.Col = 9
                WVector1.Text = rstCtaCte!Partida
                WVector1.Col = 4
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
            
        Case 106
            If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Or Val(WVector1.Text) = 4 Then
                Auxi$ = Str$(Val(WVector1.Text))
                Call Ceros(Auxi$, 2)
                WVector1.Text = Auxi$
                Select Case Val(WVector1.Text)
                    Case 1
                        WVector1.Col = 7
                        WVector1.Text = ""
                        WVector1.Col = 8
                        WVector1.Text = ""
                        WVector1.Col = 9
                        WVector1.Text = ""
                    Case Else
                End Select
                    Else
                WControl = "N"
            End If
            
        Case 107
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
            
        Case 108
            Call Valida_fecha1(WVector1.Text, Auxi)
            If Auxi <> "S" Then
                WControl = "N"
            End If
            
        Case 110
            IRow = WVector1.Row
            WVector1.Col = 6
            XTipo = WVector1.Text
            WVector1.Col = 10
            WVector1.Text = Alinea("###,###.##", WVector1.Text)
            Call Suma_Datos
            WVector1.Row = IRow
            If Val(WVector1.TextMatrix(WVector1.Row, 6)) = 4 Then
                Ingrecuenta.Visible = True
                Cuenta1.Text = WCuenta(WVector1.Row)
                Cuenta1.SetFocus
            End If
        
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

Private Sub Ctacte_Click()

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
    Opcion.ListIndex = 3
    
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

Private Sub FechaRetIva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub FechaRetSuss_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub FechaRetOtra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Juridiccion_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call cmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call Ctacte_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = ""
    sCentimos = "centavos"
    
    Numero = Str$(ZZConversion)
    
    XTexto1 = Numero2Letra(Numero, , sMoneda & " ", sCentimos & " ")
    XTexto1 = XTexto1 + Space$(100)
    
    Pasa = 0
    
    For DA = 60 To 1 Step -1
        If Mid$(XTexto1, DA, 1) = Space$(1) Then
            Pasa = 1
        End If
        If Pasa = 1 Then
            If Mid$(XTexto1, DA, 1) <> Space$(1) Then
                Exit For
            End If
        End If
    Next DA
    
    XTexto2 = Mid$(XTexto1, DA + 2, 100)
    XTexto1 = Left$(XTexto1, DA)
    
End Sub



















