VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPto 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Cotizaciones de Clientes"
   ClientHeight    =   9630
   ClientLeft      =   0
   ClientTop       =   540
   ClientWidth     =   15270
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
   ScaleHeight     =   9630
   ScaleWidth      =   15270
   Visible         =   0   'False
   Begin VB.Frame DatosClientes 
      Height          =   4935
      Left            =   1800
      TabIndex        =   52
      Top             =   1080
      Visible         =   0   'False
      Width           =   7695
      Begin VB.CommandButton CierraDatos 
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
         MouseIcon       =   "Pto.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Pto.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Razon 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   60
         Top             =   360
         Width           =   3135
      End
      Begin VB.TextBox Direccion 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   59
         Text            =   " "
         Top             =   720
         Width           =   3135
      End
      Begin VB.TextBox Localidad 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   58
         Text            =   " "
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox Postal 
         BeginProperty Font 
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
         MaxLength       =   15
         TabIndex        =   57
         Text            =   " "
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Telefono 
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
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   56
         Text            =   " "
         Top             =   2160
         Width           =   3135
      End
      Begin VB.TextBox Contacto 
         BeginProperty Font 
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
         MaxLength       =   100
         TabIndex        =   55
         Text            =   " "
         Top             =   2520
         Width           =   4695
      End
      Begin VB.ComboBox ProvinciaCli 
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
         Left            =   2760
         TabIndex        =   54
         Text            =   " "
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Fantasia 
         BeginProperty Font 
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
         MaxLength       =   50
         TabIndex        =   53
         Top             =   2880
         Width           =   3135
      End
      Begin VB.Label lblLabels 
         Caption         =   "Razon Social"
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
         TabIndex        =   68
         Top             =   2880
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Direccion"
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
         TabIndex        =   67
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Poblaci 
         Caption         =   "Localidad"
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
         TabIndex        =   66
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label8 
         Caption         =   "Codigo Postal"
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
         TabIndex        =   65
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Telefono"
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
         TabIndex        =   64
         Top             =   2160
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Contacto"
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
         TabIndex        =   63
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Provincia"
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
         TabIndex        =   62
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         Caption         =   "Denominacion Comercial"
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
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame xdf 
      Height          =   1335
      Left            =   9000
      TabIndex        =   39
      Top             =   8160
      Width           =   5055
      Begin VB.Label Label17 
         Caption         =   "U$S"
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
         Index           =   3
         Left            =   120
         TabIndex        =   50
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label17 
         Caption         =   "$"
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
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   480
         Width           =   375
      End
      Begin VB.Label NetoII 
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
         Left            =   600
         TabIndex        =   48
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Iva1II 
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
         Left            =   2040
         TabIndex        =   47
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label TotalII 
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
         Left            =   3600
         TabIndex        =   46
         Top             =   840
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
         Left            =   3600
         TabIndex        =   45
         Top             =   480
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
         Left            =   2040
         TabIndex        =   44
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   3720
         TabIndex        =   43
         Top             =   240
         Width           =   1215
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
         Left            =   2040
         TabIndex        =   42
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label17 
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
         Index           =   0
         Left            =   600
         TabIndex        =   41
         Top             =   240
         Width           =   1335
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
         Left            =   600
         TabIndex        =   40
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CommandButton IngreNotas 
      Caption         =   "Datos Clientes"
      Height          =   615
      Left            =   11160
      TabIndex        =   38
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton Panta 
      Caption         =   "Pantalla F1"
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
      Left            =   12840
      MouseIcon       =   "Pto.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "Pto.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Impresion por Pantalla"
      Top             =   6960
      Width           =   855
   End
   Begin VB.CommandButton ImpresionII 
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
      Left            =   12840
      MouseIcon       =   "Pto.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "Pto.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Impresion"
      Top             =   5880
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
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   35
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
      Index           =   9
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   34
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
      Index           =   8
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   4800
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
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   32
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
      Index           =   6
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   4800
      Width           =   375
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
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   30
      Top             =   840
      Width           =   8055
   End
   Begin VB.Frame PantallaConfirma 
      Height          =   1335
      Left            =   2520
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
      Begin VB.TextBox Confirma 
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
         Left            =   3240
         MaxLength       =   6
         TabIndex        =   27
         Text            =   " "
         Top             =   600
         Width           =   735
      End
      Begin VB.Label fhfg 
         Caption         =   "Confirma los datos ingresados"
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
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   2895
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
      Index           =   5
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3000
      Width           =   375
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
      Left            =   13800
      MouseIcon       =   "Pto.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "Pto.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Menu Principal"
      Top             =   5880
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
      Left            =   11880
      MouseIcon       =   "Pto.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "Pto.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Consulta de Datos"
      Top             =   5880
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
      Left            =   10920
      MouseIcon       =   "Pto.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "Pto.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Limpia la pantalla"
      Top             =   5880
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
      Left            =   9960
      MouseIcon       =   "Pto.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "Pto.frx":46D2
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Elimina el Registro"
      Top             =   5880
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
      Left            =   9000
      MouseIcon       =   "Pto.frx":4F14
      MousePointer    =   99  'Custom
      Picture         =   "Pto.frx":521E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox Cliente 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1575
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
      Left            =   1800
      MaxLength       =   8
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Width           =   1095
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
      TabIndex        =   14
      Top             =   3240
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   360
      Left            =   3480
      TabIndex        =   13
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
      TabIndex        =   12
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
      Index           =   2
      Left            =   4800
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
      Index           =   3
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   9
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
      TabIndex        =   8
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
      Left            =   240
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   8655
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "FactuForPro.rpt"
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
      Left            =   600
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
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
      Left            =   9960
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1500
      ItemData        =   "Pto.frx":5A60
      Left            =   240
      List            =   "Pto.frx":5A67
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   8655
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   15
      Top             =   3240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   393216
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
      Top             =   1200
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777152
   End
   Begin VB.Label DesClienteII 
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
      Left            =   6360
      TabIndex        =   51
      Top             =   480
      Width           =   4695
   End
   Begin VB.Label Label7 
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
      Height          =   285
      Left            =   120
      TabIndex        =   29
      Top             =   840
      Width           =   1575
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
      Left            =   3600
      TabIndex        =   18
      Top             =   120
      Width           =   855
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
      Left            =   6360
      TabIndex        =   17
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   2415
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
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "PrgPto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WCodIva As String
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private WTipo As String
Private WTipoIva As String
Private WProvincia As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private Mes(0 To 30) As String
Private XIndice As Single
Private XArticulo As String
Private XTexto1 As String
Private XTexto2 As String

Dim ZZNumero As String
Dim ZZRenglon As String
Dim ZZArticulo As String
Dim ZZCantidad As String
Dim ZZPrecio As String
Dim ZZImporte As String
Dim ZZFacturado As String
Dim ZZCliente As String
Dim ZZfecha As String
Dim ZZImporte1 As String
Dim ZZImporte2 As String
Dim ZZImporte3 As String
Dim ZZImporte4 As String
Dim ZZOrdFecha As String
Dim ZZObservaciones As String
Dim ZZFecEntrega As String
Dim ZZOrdFecEntrega As String
Dim ZZCotiza As String
Dim ZZAjuste As String
Dim ZZClave As String
Dim ZZGrupo As String
Dim ZCantidad As String
Dim ZCodigo As String
Dim ZObservaciones As String
Dim ZZPedido As String
Dim ZZVERSION As String

Dim WWArti As String
Dim WWCanti As Double
Dim WWLinea As String
Dim WWTipo As String
Dim WWFragancia As String
Dim WWCalidad As String
Dim WWTamano As String
Dim WWDto As Double
Dim WWPrecio As Double
Dim WWImporte As Double
Dim WWImporteII As Double

Dim WWTope1 As Double
Dim WWValor1 As Double
Dim WWTope2 As Double
Dim WWValor2 As Double
Dim WWTope3 As Double
Dim WWValor3 As Double
Dim WWTope4 As Double
Dim WWValor4 As Double
Dim WWDesde As String
Dim WWHasta As String
Dim WWOrdDesde As String
Dim WWOrdHasta As String
Dim WWMoneda As Integer

Dim WWWWTope1 As Double
Dim WWWWValor1 As Double
Dim WWWWTope2 As Double
Dim WWWWValor2 As Double
Dim WWWWTope3 As Double
Dim WWWWValor3 As Double
Dim WWWWTope4 As Double
Dim WWWWValor4 As Double
Dim WWWWDesde As String
Dim WWWWHasta As String
Dim WWWWOrdDesde As String
Dim WWWWOrdHasta As String

Dim ControlPrecioI As Integer
Dim ControlPrecioII As Integer

Dim Impodto As Double
Dim WWParidad As Double
Dim WFecha As String
Dim WVencimiento As String
Dim WPlazo1 As Integer


Dim ZVector(1000, 10) As String
Dim ZGraba(100, 30) As String


Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Dim ZZDeriva As Integer



Private Sub Command1_Click()
    DatosClientes.Visible = True
End Sub

Private Sub CierraDatos_Click()
    DatosClientes.Visible = False

End Sub

Private Sub Consulta_Click()

    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulos"

    Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    If ZZPasaProcesoII <> 0 Then
        ZZCodigo = Trim(ZZPasaLinea) + "-" + Trim(ZZPasaTipo) + "-" + Trim(ZZPasaFragancia) + "-" + Trim(ZZPasaCalidad) + "-" + Trim(ZZPasaTama�o)
        WVector1.Col = 1
        WVector1.Text = ZZCodigo
        Call StartEdit
        ZZPasaProcesoII = 0
        Call WTexto1_KeyDown(13, 0)
    End If

End Sub

Private Sub ImpresionII_Click()
    ZZDeriva = 1
    Call Impresion
End Sub

Private Sub IngreNotas_Click()
    DatosClientes.Visible = True
    Razon.SetFocus
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
                            IngresaItem = !Codigo + " " + !DescripcionII
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

Private Sub CmdClose_Click()
    PrgPto.Hide
    Unload Me
    PrgControlPedido.Show
End Sub

Private Sub Graba_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        rstCliente.Close
            Else
        m$ = "Cliente Inexistente"
        aaaaaa% = MsgBox(m$, 0, "Carga de Cotizaciones")
        Exit Sub
    End If


    For A = 1 To 99
        Articulo = UCase(WVector1.TextMatrix(A, 1))
        WWCantidad = Val(WVector1.TextMatrix(A, 3))
        If WWCantidad <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                    Else
                m$ = "Producto Inexistente"
                aaaaaa% = MsgBox(m$, 0, "Carga de Cotizaciones")
                Exit Sub
            End If
        End If
    Next A
    

Rem dada
Rem dada
Rem dada

    
    If Val(Numero.Text) <> 0 Then
        ZSql = ""
        ZSql = ZSql + "DELETE Pto"
        ZSql = ZSql + " Where Pto.Numero = " + "'" + Numero.Text + "'"
        spPto = ZSql
        Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
            Else
        ZSql = ""
        ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
        ZSql = ZSql + " FROM Pto"
        spPto = ZSql
        Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
        If rstPto.RecordCount > 0 Then
            rstPto.MoveLast
            ZUltimo = IIf(IsNull(rstPto!NumeroMayor), "0", rstPto!NumeroMayor)
            Numero.Text = ZUltimo + 1
            rstPto.Close
        End If
    End If

    Renglon = 0
    WRenglon = 0
        
    For A = 1 To 99
        
        WRenglon = WRenglon + 1
            
        WVector1.Row = WRenglon
            
        WVector1.Col = 1
        Articulo = UCase(WVector1.Text)
                    
        WVector1.Col = 3
        WWCantidad = Val(WVector1.Text)
                    
        WVector1.Col = 4
        WWTipoMoneda = WVector1.Text
                    
        WVector1.Col = 5
        WWDto = Val(WVector1.Text)
                    
        WVector1.Col = 6
        WWPrecio = Val(WVector1.Text)
                    
        WVector1.Col = 7
        WWPrecioII = Val(WVector1.Text)
                    
        WVector1.Col = 8
        WWImporte = Val(WVector1.Text)
                    
        WVector1.Col = 9
        WWImporteII = Val(WVector1.Text)
                    
        WVector1.Col = 10
        WWObserva = WVector1.Text
        
        If WWCantidad <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZLinea = rstArticulo!Linea
                ZZTipo = rstArticulo!Tipo
                ZZFragancia = rstArticulo!Fragancia
                ZZCalidad = rstArticulo!Calidad
                ZZTamano = rstArticulo!Tamano
                rstArticulo.Close
            End If
        
        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Numero.Text)
            Call Ceros(Auxi1, 8)
                    
            ZZNumero = Numero.Text
            ZZRenglon = Str$(Renglon)
            ZZArticulo = Articulo
            ZZCantidad = Str$(WWCantidad)
            ZZCliente = UCase(Cliente.Text)
            ZZImporte = Str$(WWImporte)
            ZZImporteII = Str$(WWImporteII)
            ZZPrecio = Str$(WWPrecio)
            ZZPrecioII = Str$(WWPrecioII)
            ZZDto = Str$(WWDto)
            If WWTipoMoneda = "$" Then
                ZZMoneda = "0"
                    Else
                ZZMoneda = "1"
            End If
            ZZObserva = WWObserva
            ZZImporte1 = "0"
            ZZImporte2 = "0"
            ZZImporte3 = "0"
            ZZImporte4 = "0"
            ZZfecha = Fecha.Text
            ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZZObservaciones = Observaciones.Text
            ZZMarca = Marca
            
            ZZClave = Auxi1 + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Pto ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "LInea ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "FRagancia ,"
            ZSql = ZSql + "CAlidad ,"
            ZSql = ZSql + "Tamano ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "MOneda ,"
            ZSql = ZSql + "Dto ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "PrecioII ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "ImporteII ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Importe1 ,"
            ZSql = ZSql + "Importe2 ,"
            ZSql = ZSql + "Importe3 ,"
            ZSql = ZSql + "Importe4 ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Observa ,"
            ZSql = ZSql + "Marca ,"
            ZSql = ZSql + "Razon ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Provincia ,"
            ZSql = ZSql + "Postal ,"
            ZSql = ZSql + "Telefono ,"
            ZSql = ZSql + "Contacto ,"
            ZSql = ZSql + "Fantasia )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZLinea + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZFragancia + "',"
            ZSql = ZSql + "'" + ZZCalidad + "',"
            ZSql = ZSql + "'" + ZZTamano + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZMoneda + "',"
            ZSql = ZSql + "'" + ZZDto + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZPrecioII + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZImporteII + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZImporte1 + "',"
            ZSql = ZSql + "'" + ZZImporte2 + "',"
            ZSql = ZSql + "'" + ZZImporte3 + "',"
            ZSql = ZSql + "'" + ZZImporte4 + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZObserva + "',"
            ZSql = ZSql + "'" + ZZMarca + "',"
            ZSql = ZSql + "'" + Razon.Text + "',"
            ZSql = ZSql + "'" + Direccion.Text + "',"
            ZSql = ZSql + "'" + Localidad.Text + "',"
            ZSql = ZSql + "'" + ProvinciaCli.Text + "',"
            ZSql = ZSql + "'" + Postal.Text + "',"
            ZSql = ZSql + "'" + Telefono.Text + "',"
            ZSql = ZSql + "'" + Contacto.Text + "',"
            ZSql = ZSql + "'" + Fantasia.Text + "')"
            
            spPto = ZSql
            Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
                                        
    Next A


    T$ = "Impresion de Cotizaciones"
    m$ = "Desea Imprimir la Cotizacion"
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
        ZZDeriva = 1
        Call Impresion
    End If
    
    Rem Call Limpia_Click
    
    m$ = "Grabacion realizada"
    aaaaaaaaaa% = MsgBox(m$, 0, "Archivo de Cotizaciones")
    
    
    Numero.SetFocus
        
End Sub

Private Sub Impresion()

    Rem Exit Sub

    
    Listado.WindowTitle = "Impresion de Proforma"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Pto.Numero, Pto.Articulo, Pto.Cantidad, Pto.Precio, Pto.Fecha, Pto.Observaciones, Pto.Moneda, Pto.PrecioII, Pto.Razon, Pto.Direccion, " _
            + "Articulo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Pto Pto, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "Pto.Articulo = Articulo.Codigo AND " _
            + "Pto.Numero >= " + Numero.Text + " AND " _
            + "Pto.Numero <= " + Numero.Text
    
    Listado.Connect = Connect()
    
    Uno = "{Pto.Numero} in " + Numero.Text + " to " + Numero.Text
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    Listado.CopiesToPrinter = 2
    Listado.Destination = 0
    
    Listado.ReportFileName = "ImprePto.rpt"
    
    Listado.Action = 1

End Sub

Private Sub cmdDelete_Click()

    T$ = "Baja de Comprobantes"
    m$ = "Desea Borrar el Comprobante "
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
    
        ZSql = ""
        ZSql = ZSql + "DELETE Pto"
        ZSql = ZSql + " Where Pto.Numero = " + "'" + Numero.Text + "'"
        spPto = ZSql
        Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
        
        Call Limpia_Click
        Numero.SetFocus
        
    End If

End Sub

Private Sub Limpia_Click()

    Call Limpia_Vector

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    DesClienteII.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    ProvinciaCli.ListIndex = 0
    Postal.Text = ""
    Telefono.Text = ""
    Contacto.Text = ""
    Fantasia.Text = ""
    
    Renglon = 0
    
    Numero.Text = ""
    ZSql = ""
    ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    ZSql = ZSql + " FROM Pto"
    spPto = ZSql
    Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
    If rstPto.RecordCount > 0 Then
        rstPto.MoveLast
        ZUltimo = IIf(IsNull(rstPto!NumeroMayor), "0", rstPto!NumeroMayor)
        Numero.Text = ZUltimo + 1
        rstPto.Close
    End If
    
    Numero.SetFocus

End Sub

Private Sub Panta_Click()
    ZZDeriva = 0
    Call Impresion
End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            indice = Pantalla.ListIndex
            Cliente.Text = WIndice.List(indice)
            Call Cliente_KeyPress(13)
            
        Case 1
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(indice)
            
            WTexto1.Visible = False
            WTexto2.Visible = False
            WVector1.TextMatrix(WVector1.Row, 1) = WIndice.List(indice)
            WTexto1.Text = WIndice.List(indice)
            Call WTexto1_KeyDown(13, 0)
            
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
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
    
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
    

    ProvinciaCli.Clear
    
    ProvinciaCli.AddItem "0 - Capital Federal"
    ProvinciaCli.AddItem "1 - Buenos Aires"
    ProvinciaCli.AddItem "2 - Catamarca"
    ProvinciaCli.AddItem "3 - Cordoba"
    ProvinciaCli.AddItem "4 - Corrientes"
    ProvinciaCli.AddItem "5 - Chaco"
    ProvinciaCli.AddItem "6 - Chubut"
    ProvinciaCli.AddItem "7 - Entre Rios"
    ProvinciaCli.AddItem "8 - Formosa"
    ProvinciaCli.AddItem "9 - Jujuy"
    ProvinciaCli.AddItem "10 - La Pampa"
    ProvinciaCli.AddItem "11 - La Rioja"
    ProvinciaCli.AddItem "12 - Mendoza"
    ProvinciaCli.AddItem "13 - Misiones"
    ProvinciaCli.AddItem "14 - Neuquen"
    ProvinciaCli.AddItem "15 - Rio Negro"
    ProvinciaCli.AddItem "16 - Salta"
    ProvinciaCli.AddItem "17 - San Juan"
    ProvinciaCli.AddItem "18 - San Luis"
    ProvinciaCli.AddItem "19 - Santa Cruz"
    ProvinciaCli.AddItem "20 - Santa Fe"
    ProvinciaCli.AddItem "21 - Santiago del Estero"
    ProvinciaCli.AddItem "22 - Tucuman"
    ProvinciaCli.AddItem "23 - Tierra del Fuego"
    ProvinciaCli.AddItem "24 - Exterior"
    ProvinciaCli.AddItem ""
    
    ProvinciaCli.ListIndex = 0
        
    
    
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    DesClienteII.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    ProvinciaCli.ListIndex = 0
    Postal.Text = ""
    Telefono.Text = ""
    Contacto.Text = ""
    Fantasia.Text = ""
    WWParidad = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + "1" + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        WWParidad = rstDolar!Paridad
    End If
    
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
    
    
    Numero.Text = ""
    ZSql = ""
    ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    ZSql = ZSql + " FROM Pto"
    spPto = ZSql
    Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
    If rstPto.RecordCount > 0 Then
        rstPto.MoveLast
        ZUltimo = IIf(IsNull(rstPto!NumeroMayor), "0", rstPto!NumeroMayor)
        Numero.Text = ZUltimo + 1
        rstPto.Close
    End If
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    
    Erase ZVector

    Renglon = 0
    WNeto = 0
    ControlPrecioI = 0
    
    For WRenglon = 1 To 99
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
            
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pto"
        ZSql = ZSql + " Where Pto.Clave = " + "'" + WClave + "'"
        spPto = ZSql
        Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
        If rstPto.RecordCount > 0 Then
            
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                    
            ZZArticulo = Trim(rstPto!Articulo)
            ZZCantidad = rstPto!Cantidad
            ZZLinea = rstPto!Linea
            ZZTipo = rstPto!Tipo
            ZZFragancia = rstPto!Fragancia
            ZZCalidad = rstPto!Calidad
            ZZTamano = rstPto!Tamano
            ZZObserva = Trim(rstPto!observa)
            ZZMarca = rstPto!Marca
            
            rstPto.Close
            
            Canti = ZZCantidad
            
            WVector1.Col = 1
            WVector1.Text = Trim(ZZArticulo)
            Auxi1 = ZZArticulo
                
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", Str$(ZZCantidad))
            WWCantidad = ZZCantidad
            
            
            
            WWArti = Trim(ZZArticulo)
            WWCanti = ZZCantidad
            
            WWLinea = ZZLinea
            WWTipo = ZZTipo
            WWFragancia = ZZFragancia
            WWCalidad = ZZCalidad
            WWTamano = ZZTamano
            
            ControlPrecioII = 0
            Call Calcula_Costo
            
            If WWMoneda = 0 Then
                WVector1.Col = 4
                WVector1.Text = "$"
                WWPrecioII = WWPrecio / WWParidad
                    Else
                WVector1.Col = 4
                WVector1.Text = "U$S"
                WWPrecioII = WWPrecio
                WWPrecio = WWPrecio * WWParidad
            End If
                
            WWImporte = WWPrecio * WWCanti
            WWImporteII = WWImporte / WWParidad
            
            
            If ControlPrecioII = 0 Then
                WVector1.Col = 5
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWDto))
                WVector1.Col = 6
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecio))
                WVector1.Col = 7
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecioII))
                WVector1.Col = 8
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWImporte))
                WVector1.Col = 9
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWImporteII))
                    Else
                WVector1.Col = 5
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWDto))
                WVector1.Col = 6
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecio))
                WVector1.Col = 7
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecioII))
                WVector1.Col = 8
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWImporte))
                WVector1.Col = 8
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWImporteII))
            End If
            
                
            WVector1.Col = 10
            WVector1.Text = Trim(ZZObserva)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstArticulo!DescripcionII
                rstArticulo.Close
            End If
            
        End If
    
    Next WRenglon
    
    Call Calcula_Click
    
End Sub

Sub Numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pto"
        ZSql = ZSql + " Where Pto.Numero = " + "'" + Numero.Text + "'"
        spPto = ZSql
        Set rstPto = db.OpenRecordset(spPto, dbOpenSnapshot, dbSQLPassThrough)
        If rstPto.RecordCount > 0 Then
    
            Fecha.Text = rstPto!Fecha
            Cliente.Text = rstPto!Cliente
            Observaciones.Text = rstPto!Observaciones
            DesCliente.Caption = rstPto!Razon
            DesClienteII.Caption = rstPto!Fantasia
            Razon.Text = rstPto!Razon
            Direccion.Text = rstPto!Direccion
            Localidad.Text = rstPto!Localidad
            Postal.Text = rstPto!Postal
            Telefono.Text = rstPto!Telefono
            Contacto.Text = rstPto!Contacto
            Fantasia.Text = rstPto!Fantasia
            
            rstPto.Close
            
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                
                Else
                    
            Rem Cliente.SetFocus
               
        End If
            
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            Cliente.Text = rstCliente!Cliente
            DesCliente.Caption = rstCliente!Fantasia
            DesClienteII.Caption = rstCliente!Razon
            
            Razon.Text = Trim(rstCliente!Razon)
            Direccion.Text = Trim(rstCliente!Direccion)
            Localidad.Text = Trim(rstCliente!Localidad)
            Postal.Text = Trim(rstCliente!Postal)
            Telefono.Text = Trim(rstCliente!Telefono)
            Contacto.Text = Trim(rstCliente!Observaciones)
            Fantasia.Text = Trim(rstCliente!Fantasia)
            ProvinciaCli.ListIndex = rstCliente!Provincia
            
            rstCliente.Close
            Fecha.SetFocus
                Else
            Cliente.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
        DesClienteII.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            aaaaaa% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
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


Private Sub Razon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Direccion.SetFocus
    End If
    If KeyAscii = 27 Then
        Razon.Text = ""
    End If
End Sub

Private Sub Direccion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Localidad.SetFocus
    End If
    If KeyAscii = 27 Then
        Direccion.Text = ""
    End If
End Sub

Private Sub Localidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Postal.SetFocus
    End If
    If KeyAscii = 27 Then
        Localidad.Text = ""
    End If
End Sub

Private Sub Postal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Telefono.SetFocus
    End If
    If KeyAscii = 27 Then
        Postal.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Telefono_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Contacto.SetFocus
    End If
    If KeyAscii = 27 Then
        Telefono.Text = ""
    End If
End Sub

Private Sub Contacto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fantasia.SetFocus
    End If
    If KeyAscii = 27 Then
        Contacto.Text = ""
    End If
End Sub

Private Sub Fantasia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Razon.SetFocus
    End If
    If KeyAscii = 27 Then
        Fantasia.Text = ""
    End If
End Sub


Private Sub Confirma_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Confirma.Text = Trim(UCase(Confirma.Text))
        If Confirma.Text = "S" Or Confirma.Text = "N" Or Confirma.Text = "/" Or Confirma.Text = "?" Then
            PantallaConfirma.Visible = False
            If Confirma.Text <> "N" Then
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Confirma.Text = ""
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

    On Error GoTo WError
    
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

    Exit Sub
    
WError:
    Resume Next


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
                Call Control_wvector1
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
        Case 9
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
                If Trim(WVector1.TextMatrix(WVector1.Row, 1)) = "" Then
                    WVector1.TextMatrix(WVector1.Row, 1) = WVector1.TextMatrix(WVector1.Row - 1, 1)
                End If
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
        
            Articulo = UCase(WVector1.Text)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                If rstArticulo!Activo = 0 Then
                
                    WWArti = WVector1.Text
                    WWCanti = Val(WVector1.TextMatrix(WVector1.Row, 3))
                    
                    WWLinea = rstArticulo!Linea
                    WWTipo = rstArticulo!Tipo
                    WWFragancia = rstArticulo!Fragancia
                    WWCalidad = rstArticulo!Calidad
                    WWTamano = rstArticulo!Tamano
                    
                    WWDescripcion = rstArticulo!DescripcionII
                    
                    rstArticulo.Close
                    
                    Call Calcula_Costo
                    
                    If WWMoneda = 0 Then
                        WVector1.Col = 4
                        WVector1.Text = "$"
                        WWPrecioII = WWPrecio / WWParidad
                            Else
                        WVector1.Col = 4
                        WVector1.Text = "U$S"
                        WWPrecioII = WWPrecio
                        WWPrecio = WWPrecio * WWParidad
                    End If
                        
                    WWImporte = WWPrecio * WWCanti
                    WWImporteII = WWImporte / WWParidad
                    
                    WVector1.Col = 5
                    WVector1.Text = Pusing("###,###.##", Str$(WWDto))
                    
                    WVector1.Col = 6
                    WVector1.Text = Pusing("###,###.##", Str$(WWPrecio))
                    
                    Rem dada
                    WVector1.Col = 7
                    WVector1.Text = Pusing("###,###.##", Str$(WWPrecioII))
                    
                    WVector1.Col = 8
                    WVector1.Text = Pusing("###,###.##", Str$(WWImporte))
                    
                    WVector1.Col = 9
                    WVector1.Text = Pusing("###,###.##", Str$(WWImporteII))
                    
                    WFecha = Fecha.Text
                    WPlazo1 = 7
                    Call Calcula_vencimiento(WFecha, WPlazo1, WVencimiento)
                    
                    WVector1.Col = 10
                    WVector1.Text = WVencimiento
                    
                    WVector1.Col = 2
                    WVector1.Text = WWDescripcion
                    
                        Else
                
                
                    m$ = "Articulo Inactivo"
                    aaaaaa% = MsgBox(m$, 0, "Carga de Cotizaciones")
                    rstArticulo.Close
                    WControl = "N"
                
                End If
                
                    Else
                    
                WControl = "N"
                If WVector1.Text = "" Then
                    ZZPasaProcesoII = 3
                    ZZPasaCliente = Cliente.Text
                    prgBusquedaArtiCliente.Show
                End If
                
            End If
            
        Case 3
            WWArti = WVector1.TextMatrix(WVector1.Row, 1)
            WWCanti = Val(WVector1.Text)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WWArti + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WWLinea = rstArticulo!Linea
                WWTipo = rstArticulo!Tipo
                WWFragancia = rstArticulo!Fragancia
                WWCalidad = rstArticulo!Calidad
                WWTamano = rstArticulo!Tamano
                
                Call Calcula_Costo
                
                If WWMoneda = 0 Then
                    WVector1.Col = 4
                    WVector1.Text = "$"
                    WWPrecioII = WWPrecio / WWParidad
                        Else
                    WVector1.Col = 4
                    WVector1.Text = "U$S"
                    WWPrecioII = WWPrecio
                    WWPrecio = WWPrecio * WWParidad
                End If
                    
                WWImporte = WWPrecio * WWCanti
                WWImporteII = WWImporte / WWParidad
                        
                WVector1.Col = 5
                WVector1.Text = Pusing("###,###.##", Str$(WWDto))
                
                WVector1.Col = 6
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecio))
                
                Rem dada
                WVector1.Col = 7
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecioII))
                
                WVector1.Col = 8
                WVector1.Text = Pusing("###,###.##", Str$(WWImporte))
                
                WVector1.Col = 9
                WVector1.Text = Pusing("###,###.##", Str$(WWImporteII))
                
                If WWMoneda = 0 Then
                    WVector1.Col = 4
                    WVector1.Text = "$"
                        Else
                    WVector1.Col = 4
                    WVector1.Text = "U$S"
                End If
                
                
            End If
        
        
            WWCantidad = Val(WVector1.Text)
            WWTipoMoneda = WVector1.TextMatrix(WVector1.Row, 4)
            WWPrecio = Val(WVector1.TextMatrix(WVector1.Row, 6))
            
            WWImporte = WWPrecio * WWCanti
            WWImporteII = WWImporte / WWParidad
            
            WVector1.Col = 8
            WVector1.Text = Pusing("###,###.##", Str$(WWImporte))
            
            WVector1.Col = 9
            WVector1.Text = Pusing("###,###.##", Str$(WWImporteII))
            
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
    
    Call Calcula_Click
    
    End If
    
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then
        ZZPasaProcesoII = 3
        ZZPasaCliente = Cliente.Text
        prgBusquedaArtiCliente.Show
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
    WVector1.Cols = 11
    WVector1.FixedRows = 1
    WVector1.Rows = 100
    
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Codigo"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 25
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Mon."
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 5
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Dto"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Precio $"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Precio U$S"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Total $"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "Total U$S"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 100
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho
    
    For Ciclo = 1 To 99
        WVector1.TextMatrix(Ciclo, 0) = Trim(Str$(Ciclo))
    Next Ciclo

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

Private Sub Cliente_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulo"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call cmdDelete_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Calcula_Costo()


    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ControlPrecioII = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WWNroLista = Str$(rstCliente!NroLista)
        rstCliente.Close
    End If


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ClienteLista"
    ZSql = ZSql + " Where ClienteLista.Cliente = " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " and ClienteLista.LInea = " + "'" + WWLinea + "'"
    ZSql = ZSql + " and ClienteLista.Tipo = " + "'" + WWTipo + "'"
        
    spClienteLista = ZSql
    Set rstClientelista = db.OpenRecordset(spClienteLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientelista.RecordCount > 0 Then
        WWNroLista = Str$(rstClientelista!Lista)
        rstClientelista.Close
    End If

    

    WWTope1 = 0
    WWValor1 = 0
    WWTope2 = 0
    WWValor2 = 0
    WWTope3 = 0
    WWValor3 = 0
    WWTope4 = 0
    WWValor4 = 0
    WWDesde = "00/00/0000"
    WWHasta = "00/00/0000"
    WWOrdDesde = "00000000"
    WWOrdHasta = "00000000"
    WWMoneda = 0


    WWNroLista = Trim(WWNroLista)
    
    ZZLee = "S"

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Lista = " + "'" + WWNroLista + "'"
    ZSql = ZSql + " and Precios.LInea = " + "'" + WWLinea + "'"
    ZSql = ZSql + " and Precios.Tipo = " + "'" + WWTipo + "'"
    ZSql = ZSql + " and Precios.fragancia = " + "'" + WWFragancia + "'"
    ZSql = ZSql + " and Precios.Calidad = " + "'" + WWCalidad + "'"
    ZSql = ZSql + " and Precios.Tamano = " + "'" + WWTamano + "'"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        WWTope1 = rstPrecios!Tope1
        WWValor1 = rstPrecios!Valor1
        WWTope2 = rstPrecios!Tope2
        WWValor2 = rstPrecios!Valor2
        WWTope3 = rstPrecios!Tope3
        WWValor3 = rstPrecios!Valor3
        WWTope4 = rstPrecios!Tope4
        WWValor4 = rstPrecios!Valor4
        WWDesde = rstPrecios!Desde
        WWHasta = rstPrecios!Hasta
        WWOrdDesde = rstPrecios!OrdDesde
        WWOrdHasta = rstPrecios!OrdHasta
        WWMoneda = rstPrecios!Moneda
        rstPrecios.Close
        
        ZZLee = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.LInea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and Articulo.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and Articulo.fragancia = " + "'" + WWFragancia + "'"
        ZSql = ZSql + " and Articulo.Calidad = " + "'" + WWCalidad + "'"
        ZSql = ZSql + " and Articulo.Tamano = " + "'" + WWTamano + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZZActivo = rstArticulo!Activo
            rstArticulo.Close
        End If
        
        If (WWValor1 = 0 And WWValor2 = 0 And WWValor3 = 0 And WWValor4 = 0) Or ZZActivo = 1 Then
            ZZLee = "S"
        End If
        
        If ZZOrdFecha > WWOrdHasta Then
            ZZLee = "S"
        End If
                
    End If
    
    If ZZLee = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + WWNroLista + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + WWCalidad + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + WWTamano + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
        
            WWTope1 = rstPrecios!Tope1
            WWValor1 = rstPrecios!Valor1
            WWTope2 = rstPrecios!Tope2
            WWValor2 = rstPrecios!Valor2
            WWTope3 = rstPrecios!Tope3
            WWValor3 = rstPrecios!Valor3
            WWTope4 = rstPrecios!Tope4
            WWValor4 = rstPrecios!Valor4
            WWDesde = rstPrecios!Desde
            WWHasta = rstPrecios!Hasta
            WWOrdDesde = rstPrecios!OrdDesde
            WWOrdHasta = rstPrecios!OrdHasta
            WWMoneda = rstPrecios!Moneda
            rstPrecios.Close
            
            ZZLee = "N"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.LInea = " + "'" + WWLinea + "'"
            ZSql = ZSql + " and Articulo.Tipo = " + "'" + WWTipo + "'"
            ZSql = ZSql + " and Articulo.fragancia = " + "'" + WWFragancia + "'"
            ZSql = ZSql + " and Articulo.Calidad = " + "'" + WWCalidad + "'"
            ZSql = ZSql + " and Articulo.Tamano = " + "'" + WWTamano + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZActivo = rstArticulo!Activo
                rstArticulo.Close
            End If
            
            If (WWValor1 = 0 And WWValor2 = 0 And WWValor3 = 0 And WWValor4 = 0) Or ZZActivo = 1 Then
                ZZLee = "S"
            End If
            
            If ZZOrdFecha > WWOrdHasta Then
                ZZLee = "S"
            End If
            
        End If
    End If
                
    If ZZLee = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + WWNroLista + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + "" + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + WWTamano + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
        
            WWTope1 = rstPrecios!Tope1
            WWValor1 = rstPrecios!Valor1
            WWTope2 = rstPrecios!Tope2
            WWValor2 = rstPrecios!Valor2
            WWTope3 = rstPrecios!Tope3
            WWValor3 = rstPrecios!Valor3
            WWTope4 = rstPrecios!Tope4
            WWValor4 = rstPrecios!Valor4
            WWDesde = rstPrecios!Desde
            WWHasta = rstPrecios!Hasta
            WWOrdDesde = rstPrecios!OrdDesde
            WWOrdHasta = rstPrecios!OrdHasta
            WWMoneda = rstPrecios!Moneda
            rstPrecios.Close
            
            ZZLee = "N"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.LInea = " + "'" + WWLinea + "'"
            ZSql = ZSql + " and Articulo.Tipo = " + "'" + WWTipo + "'"
            ZSql = ZSql + " and Articulo.fragancia = " + "'" + WWFragancia + "'"
            ZSql = ZSql + " and Articulo.Calidad = " + "'" + WWCalidad + "'"
            ZSql = ZSql + " and Articulo.Tamano = " + "'" + WWTamano + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZActivo = rstArticulo!Activo
                rstArticulo.Close
            End If
            
            If (WWValor1 = 0 And WWValor2 = 0 And WWValor3 = 0 And WWValor4 = 0) Or ZZActivo = 1 Then
                ZZLee = "S"
            End If
            
            If ZZOrdFecha > WWOrdHasta Then
                ZZLee = "S"
            End If
            
        End If
    End If
                

    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)

    If ZZOrdFecha > WWOrdHasta Then
        Rem m$ = "Precio No Activo"
        Rem aaaaaa% = MsgBox(m$, 0, "Precios")
        ControlPrecioII = 1
        Rem WWPrecio = 0
        Rem WWDto = 0
        Rem Exit Sub
    End If
    
    If WWCanti < WWTope1 Then
        WWPrecio = WWValor1
            Else
        If WWCanti < WWTope2 Then
            WWPrecio = WWValor2
                Else
            If WWCanti < WWTope3 Then
                WWPrecio = WWValor3
                    Else
                WWPrecio = WWValor4
            End If
        End If
    End If


    ZZEntra = "N"

    WWWWTope1 = 0
    WWWWValor1 = 0
    WWWWTope2 = 0
    WWWWValor2 = 0
    WWWWTope3 = 0
    WWWWValor3 = 0
    WWWWTope4 = 0
    WWWWValor4 = 0
    WWWWDesde = 0
    WWWWHasta = 0
    WWWWOrdDesde = "000000000"
    WWWWOrdHasta = "000000000"
    
    WWDto = 0



    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ClienteBonifica"
    ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
    ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
    ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + WWFragancia + "'"
    ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + WWCalidad + "'"
    ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + WWTamano + "'"
    ZSql = ZSql + " Order by ClienteBonifica.orddesde"
    spClienteBonifica = ZSql
    Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstClienteBonifica.RecordCount > 0 Then
        With rstClienteBonifica
            .MoveLast
            ZZComparaI = rstClienteBonifica!OrdHasta
            ZZComparaII = rstClienteBonifica!OrdHasta
            Rem If ZZComparaI <> "" And ZZComparaII = "" Then
            Rem     ZZComparaII = "20991231"
            Rem End If
            If ZZComparaII > ZZOrdFecha Then
                ZZEntra = "S"
                WWWWTope1 = rstClienteBonifica!Tope1
                WWWWValor1 = rstClienteBonifica!Valor1
                WWWWTope2 = rstClienteBonifica!Tope2
                WWWWValor2 = rstClienteBonifica!Valor2
                WWWWTope3 = rstClienteBonifica!Tope3
                WWWWValor3 = rstClienteBonifica!Valor3
                WWWWTope4 = rstClienteBonifica!Tope4
                WWWWValor4 = rstClienteBonifica!Valor4
                WWWWDesde = rstClienteBonifica!Desde
                WWWWHasta = rstClienteBonifica!Hasta
                WWWWOrdDesde = rstClienteBonifica!OrdDesde
                WWWWOrdHasta = rstClienteBonifica!OrdHasta
            End If
            rstClienteBonifica.Close
        End With
        
    End If

    
    If ZZEntra = "N" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + WWCalidad + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + WWTamano + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
        
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + WWTamano + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
        
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
        
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + "" + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
            
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + "" + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
            
        End If
    End If


    If ZZOrdFecha > WWOrdDesde Or ZZOrdFecha < WWOrdHasta Then
        
        If WWCanti < WWWWTope1 Then
            WWDto = WWWWValor1
                Else
            If WWCanti < WWWWTope2 Then
                WWDto = WWWWValor2
                    Else
                If WWDto < WWWWTope3 Then
                    WWDto = WWWWValor3
                        Else
                    WWDto = WWWWValor4
                End If
            End If
        End If
        
    End If

    If WWDto <> 0 Then
        WWPrecio = WWValor1
        If WWValor2 > WWPrecio Then
            WWPrecio = WWValor2
        End If
        If WWValor3 > WWPrecio Then
            WWPrecio = WWValor3
        End If
        If WWValor4 > WWPrecio Then
            WWPrecio = WWValor4
        End If
    End If

    WImpoDto = 0
    WDescuento = WWDto
    If WDescuento <> 0 Then
        WImpoDto = WWPrecio * WDescuento / 100
        Call Redondeo(WImpoDto)
        WWPrecio = WWPrecio - WImpoDto
    End If

    If ControlPrecioII = 1 Then
        WWPrecio = 0
        WWDto = 0
    End If


End Sub

Private Sub Calcula_Click()

    WNeto = 0
    WNetoII = 0
    Rem dada
    
    For A = 1 To 50
    
        WImporte = Val(WVector1.TextMatrix(A, 8))
        WImporteII = Val(WVector1.TextMatrix(A, 9))
        
        WNeto = WNeto + WImporte
        WNetoII = WNetoII + WImporteII
        
    Next A
    

    WIva1 = 0
    WIva2 = 0
    
    WIva1 = WNeto * 0.21
    WTotal = WNeto + WIva1
    
    Neto.Caption = Str$(WNeto)
    Iva1.Caption = Str$(WIva1)
    Total.Caption = Str$(WTotal)
    
    Neto.Caption = Pusing("###,###.##", Neto.Caption)
    Iva1.Caption = Pusing("###,###.##", Iva1.Caption)
    Total.Caption = Pusing("###,###.##", Total.Caption)


    WIva1 = 0
    WIva2 = 0
    
    WIva1 = WNetoII * 0.21
    WTotal = WNetoII + WIva1
    
    NetoII.Caption = Str$(WNetoII)
    Iva1II.Caption = Str$(WIva1)
    TotalII.Caption = Str$(WTotal)
    
    NetoII.Caption = Pusing("###,###.##", NetoII.Caption)
    Iva1II.Caption = Pusing("###,###.##", Iva1II.Caption)
    TotalII.Caption = Pusing("###,###.##", TotalII.Caption)


End Sub

