VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFacturaRemito 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturas"
   ClientHeight    =   8415
   ClientLeft      =   600
   ClientTop       =   750
   ClientWidth     =   13950
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
   ScaleHeight     =   8415
   ScaleWidth      =   13950
   Visible         =   0   'False
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   4680
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton BorraRenglon 
      Caption         =   "Borra Renglon"
      Height          =   615
      Left            =   11160
      TabIndex        =   65
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Frame PantaDolar 
      Height          =   2895
      Left            =   2280
      TabIndex        =   62
      Top             =   1800
      Visible         =   0   'False
      Width           =   7335
      Begin VB.ComboBox TipoParidadOtro 
         Height          =   360
         Left            =   1920
         TabIndex        =   64
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Selecciones el tipo de cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   63
         Top             =   360
         Width           =   6855
      End
   End
   Begin VB.TextBox ValorDolar 
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
      Left            =   12720
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   60
      Top             =   600
      Width           =   1095
   End
   Begin VB.ComboBox TipoParidad 
      Height          =   360
      Left            =   10680
      TabIndex        =   59
      Top             =   0
      Width           =   2055
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
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   3000
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
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   3480
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
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   3240
      Width           =   375
   End
   Begin VB.CheckBox Entregada 
      Caption         =   "Entregada"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10200
      TabIndex        =   55
      Top             =   840
      Width           =   1455
   End
   Begin VB.CheckBox Contado 
      Caption         =   "Contado"
      Height          =   255
      Left            =   10200
      TabIndex        =   54
      Top             =   480
      Width           =   1455
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
      Left            =   2520
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.ListBox Pantalla 
      Height          =   3420
      ItemData        =   "FacturaRemito.frx":0000
      Left            =   2520
      List            =   "FacturaRemito.frx":0007
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   7335
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
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   4320
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
      Height          =   285
      Left            =   8880
      MaxLength       =   8
      TabIndex        =   49
      Text            =   " "
      Top             =   480
      Width           =   1215
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
      TabIndex        =   48
      Top             =   4200
      Width           =   375
   End
   Begin VB.TextBox Expreso 
      BeginProperty Font 
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
      MaxLength       =   50
      TabIndex        =   46
      Text            =   " "
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Anula 
      Caption         =   "Anula"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12960
      MouseIcon       =   "FacturaRemito.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":031F
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Elimina el Registro"
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Pago 
      BeginProperty Font 
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
      MaxLength       =   6
      TabIndex        =   40
      Text            =   " "
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton PedidoAyuda 
      Caption         =   "Pedido F5"
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
      Left            =   12960
      MouseIcon       =   "FacturaRemito.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":0E6B
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Pedidos de Clientes"
      Top             =   5160
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
      Left            =   12960
      MouseIcon       =   "FacturaRemito.frx":1735
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":1A3F
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Impresion"
      Top             =   6240
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
      Left            =   12960
      MouseIcon       =   "FacturaRemito.frx":2281
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":258B
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   960
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
      Height          =   855
      Left            =   9120
      MouseIcon       =   "FacturaRemito.frx":2DCD
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":30D7
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Elimina el Registro"
      Top             =   4680
      Visible         =   0   'False
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
      Left            =   12960
      MouseIcon       =   "FacturaRemito.frx":3919
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":3C23
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Limpia la pantalla"
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
      Left            =   12960
      MouseIcon       =   "FacturaRemito.frx":4465
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":476F
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Consulta de Datos"
      Top             =   4080
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
      Left            =   12960
      MouseIcon       =   "FacturaRemito.frx":4FB1
      MousePointer    =   99  'Custom
      Picture         =   "FacturaRemito.frx":52BB
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Menu Principal"
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox Pedido 
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
      MaxLength       =   8
      TabIndex        =   30
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   21
      Top             =   7440
      Width           =   9255
      Begin VB.Label Label5 
         Caption         =   "Dto"
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
         Left            =   1680
         TabIndex        =   53
         Top             =   240
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
         Left            =   1680
         TabIndex        =   52
         Top             =   480
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
         Left            =   3240
         TabIndex        =   44
         Top             =   480
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
         Left            =   3240
         TabIndex        =   43
         Top             =   240
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
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Iva"
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
         Left            =   4680
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
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
         Index           =   0
         Left            =   6240
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
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
         Left            =   7920
         TabIndex        =   26
         Top             =   240
         Width           =   1215
      End
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
         Left            =   120
         TabIndex        =   25
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
         Left            =   4680
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
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
         Left            =   6240
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
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
         Left            =   7800
         TabIndex        =   22
         Top             =   480
         Width           =   1335
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
      Left            =   9240
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
      Left            =   8520
      MaxLength       =   4
      TabIndex        =   18
      Top             =   120
      Width           =   615
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
      Left            =   8040
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
   Begin Crystal.CrystalReport Listado 
      Left            =   11400
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Factura.rpt"
      CopiesToPrinter =   3
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
      Height          =   2160
      Left            =   4320
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   3615
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1455
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1455
      _ExtentX        =   2566
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
      Height          =   6135
      Left            =   120
      TabIndex        =   32
      Top             =   1320
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   10821
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
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
      Left            =   9480
      TabIndex        =   70
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Deuda Cta.Cte"
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
      Left            =   9480
      TabIndex        =   69
      Top             =   7560
      Width           =   1335
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
      Left            =   3000
      TabIndex        =   68
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Dolar"
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
      Left            =   11760
      TabIndex        =   61
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label14 
      Caption         =   "Remito "
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
      Left            =   8160
      TabIndex        =   50
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Expreso"
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
      Left            =   5040
      TabIndex        =   47
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Cond. Pago"
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
      TabIndex        =   42
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label DesPago 
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
      Left            =   2280
      TabIndex        =   41
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label13 
      Caption         =   "Pedido"
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
      TabIndex        =   31
      Top             =   480
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
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   3015
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
      Width           =   1335
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
Attribute VB_Name = "PrgFacturaRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WNeto As Double
Private WSubTotal As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoDto1 As Double
Private WImpoDto2 As Double
Private WImpoDto3 As Double
Private WDescuento As Double
Private WCodIva As String
Private WPorceIva As Double
Private WDias As Integer
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
Private WPlazo1 As Integer
Private WVencimiento As String
Dim WPedido(1000) As String
Dim WSaldo As Double
Dim CantiFac As Integer
Dim CantiRem As Integer
Dim CantiArti As Integer
Dim ZMes As String
Dim ZAno As String
Dim ZZCambia As String
Dim ZZPasaImpre As Integer
Dim ZZPasaDatos(100, 15) As String
Dim ZZZCantidad As Double


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
Dim ZZDescuento As String
Dim ZZPartida As String
Dim ZZPago As String
Dim ZBaja As String
Dim ZZDireccionExpreso As String

Dim ControlPrecioI As Integer
Dim ControlPrecioII As Integer

Dim ZZCantidad As String
Dim ZZCantidadII As String

Dim WVector(100, 10) As String
Dim ZZVector(100, 10) As String

Dim WWArticulo As String
Dim WWDescripcion As String
Dim WWCantidad As String
Dim WWImpre As Double
Dim WWIva(10) As Double

Dim ZFactuImpre(1000, 5) As String
Dim ZZImpreBarra As String
Dim ZZImpreBarraII As String

Dim XXPrecio As Double
Dim XXDto As Double
Dim XXComision As Double
Dim XXImpoComision As Double
Dim ZZZLetra As String
Dim ZZZPunto As String
Dim ZZZNumeroFac As String




Dim WWArti As String
Dim WWCanti As Double
Dim WWLinea As String
Dim WWTipo As String
Dim WWFragancia As String
Dim WWCalidad As String
Dim WWTamano As String
Dim WWDto As Double
Dim WWPrecio As Double
Dim WWPrecioII As Double
Dim WWPrecioSB As Double
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

Dim Impodto As Double
Dim WWParidad As Double
Dim WSParcial As Double
Dim WParcial As Double

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
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

Private Sub Anula_Click()

    T$ = "Anulacion de Comprobantes"
    m$ = "Desea Anular el Comprobante "
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then

        T$ = "Anulacion de Comprobantes"
        m$ = "Esta Seguro que Desea Anular el Comprobante "
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
        
            ZBaja = "N"
            If Val(Pedido.Text) <> 0 Then
                T$ = "Baja de Comprobantes"
                m$ = "Desea Restaurar el saldo del pedido"
                Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
                If Respuestaaaaaa% = 6 Then
                    ZBaja = "S"
                End If
            End If
        
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "01"
                
            Claveven$ = Letra.Text + WTipo + WPunto + Auxi + "01"
               
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "FraganciasII"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
               
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Clave = " + "'" + Claveven$ + "'"
            spCtaCte = ZSql
            Set �Cte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
            
                ZSaldo = rstCtaCte!Saldo
                ZTotal = rstCtaCte!Total
                rstCtaCte.Close
                
                If ZSaldo <> ZTotal Then
                
                    m$ = "El comprobante se encuentra total o parcialmente cancelado"
                    aaaaaa% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
                    
                    txtUserName = "SA"
                    txtPassword = "Sw58125812"
                    txtOdbc = "Fragancias"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                    Exit Sub
                    
                End If
                
            End If
            
            Erase WVector
            Erase ZZVector
        
            For WRenglon = 1 To 99
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "01" + Auxi + Auxi1
                
                ZZZLetra = Letra.Text
                
                ZZZPunto = Punto.Text
                Call Ceros(ZZZPunto, 1)
                
                ZZZNumeroFac = Numero.Text
                Call Ceros(ZZZNumeroFac, 6)
                
                ZZClave = "01" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi1
                
                
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + ZZClave + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                
                    ZZVector(WRenglon, 1) = rstEstadistica!Articulo
                    ZZVector(WRenglon, 2) = Str$(rstEstadistica!Cantidad)
                    ZZVector(WRenglon, 4) = IIf(IsNull(rstEstadistica!ClavePedido), "", rstEstadistica!ClavePedido)
                        
                    rstEstadistica.Close
                    
                End If
                
            Next WRenglon
            
            ZSql = ""
            ZSql = ZSql + "DELETE Estadistica"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CtaCte SET"
            ZSql = ZSql + " Total = 0 ,"
            ZSql = ZSql + " Saldo = 0 ,"
            ZSql = ZSql + " TotalUs = 0 ,"
            ZSql = ZSql + " SaldoUs = 0 ,"
            ZSql = ZSql + " Neto = 0 ,"
            ZSql = ZSql + " NetoTotal = 0 ,"
            ZSql = ZSql + " Iva1 = 0 ,"
            ZSql = ZSql + " Iva2 = 0"
            ZSql = ZSql + " Where Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            ZSql = ""
            ZSql = ZSql + "DELETE MovimientoInsumo"
            ZSql = ZSql + " Where MovimientoInsumo.Tipo = " + "'" + "6" + "'"
            ZSql = ZSql + " and MovimientoInsumo.Numero = " + "'" + Numero.Text + "'"
            spMovimientoInsumo = ZSql
            Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
            Erase WVector
        
            For WRenglon = 1 To 99
            
                Articulo = ZZVector(WRenglon, 1)
                Cantidad = Val(ZZVector(WRenglon, 2))
                CantidadII = Val(ZZVector(WRenglon, 3))
                ZZClavePedido = ZZVector(WRenglon, 4)
                
                WWInsumoII = ""
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WWInsumoII = IIf(IsNull(rstArticulo!InsumoII), "", rstArticulo!InsumoII)
                    rstArticulo.Close
                End If
                                                
                If Trim(WInsumoII) = "" Then
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " StockIV = StockIV + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Insumo SET "
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " StockII = StockII + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + WInsumoII + "'"
                    spInsumo = ZSql
                    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
                ZSql = ""
                ZSql = ZSql + "UPDATE Pedido SET "
                ZSql = ZSql + " Facturado = Facturado - " + "'" + Str$(Cantidad) + "'"
                ZSql = ZSql + " Where Clave = " + "'" + ZZClavePedido + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Next WRenglon
            
            Call Limpia_Click
            
            Cliente.SetFocus
            
        End If
        
    End If


End Sub



Private Sub Impresion_FacturaFe()


    ZSql = ""
    ZSql = ZSql + "DELETE Factura"
    spFactura = ZSql
    Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZZProvincia = rstCliente!Provincia
        ZZCodIva = rstCliente!Iva
        ZZRazon = rstCliente!Fantasia
        ZZDireccion = rstCliente!Direccion
        ZZLocalidad = rstCliente!Localidad
        ZZPostal = rstCliente!Postal
        ZZCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZZCuitII = ""
    
    
    
    
    ZZLetra = "X"
    ZZTipo = "01"
    ZZPunto = "0001"
    Auxi1 = Numero.Text
    Call Ceros(Auxi1, 8)
    ZZFactura = Auxi1
    ZZfecha = Fecha.Text
    ZZCliente = Cliente.Text
    ZZNombre = Trim(ZZRazon)
    ZZDireccion = Trim(ZZDireccion) + " - " + Trim(ZZLocalidad) + " - " + Trim(Provincia(ZZProvincia))
    ZZLocalidad = Trim(ZZLocalidad) + " - " + Trim(Provincia(ZZProvincia))
    ZZLocalidad = Left$(ZZDireccion, 50)
    ZZLocalidad = Left$(ZZLocalidad, 50)
    ZZPartida = ""
    ZZNeto = Neto.Caption
    Rem ZZDto = Dto.Caption
    ZZNeto1 = SubTotal.Caption
    ZZIva1 = Iva1.Caption
    ZZIva2 = Dto.Caption
    ZZTotal = Total.Caption
    ZZImprepago = Left$(DesPago.Caption, 35)
    ZZImpreIva = Iva(Val(ZZCodIva))
    ZZPorceDto = 0
    ZZPostal = WPostal
    
    ZZLugarFactura = 0
    
    Call Numtolet
    
    For A = 1 To 99
    
        If Trim(WVector1.TextMatrix(A, 1)) <> "" Then
            
            ZZLugarFactura = ZZLugarFactura + 1
    
            ZZRenglon = Str$(ZZLugarFactura)
            Auxi1 = ZZRenglon
            Call Ceros(Auxi1, 2)
            ZZRenglon = Auxi1
            
            ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
            
            ZZItem = Str$(A)
            
            ZZArticulo = WVector1.TextMatrix(A, 1)
            ZZDescripcion = WVector1.TextMatrix(A, 2)
            ZZZCantidad = Val(WVector1.TextMatrix(A, 3))
            ZZZDto = Val(WVector1.TextMatrix(A, 4))
            ZZZPrecio = Val(WVector1.TextMatrix(A, 5))
            ZZZImporte = ZZZPrecio * ZZZCantidad
            ZZZPrecioSB = Val(WVector1.TextMatrix(A, 10))
            ZZZMoneda = Val(WVector1.TextMatrix(A, 11))
            
            ZZCantidad = Str$(ZZZCantidad)
            ZZPrecio = Str$(ZZZPrecio)
            ZZPrecioII = Str$(ZZZPrecioSB)
            ZZImporte = Str$(ZZZImporte)
            ZZDto = Str$(ZZZDto)
            ZZPorceIva = Str$(WWParidad)
            ZZDias = ZZZMoneda
            
            If Trim(ZZArticulo) = "" Then
                ZZItem = ""
                ZZArticulo = ""
                ZZDescripcion = ""
                ZZCantidad = ""
                ZZPrecio = ""
                ZZPrecioII = ""
                ZZImporte = ""
                ZZDto = ""
                ZZPorceIva = ""
                ZZDias = ""
            End If
            
            ZZDescriII = ""
            ZZCantiII = ""
            
            Call Numtolet
            ZZImpre1 = XTexto1
            ZZImpre2 = XTexto2
            
            Auxi2 = Numero.Text
            Call Ceros(Auxi2, 8)
            ZZImpre3 = Auxi2
            ZZImpre4 = "FACTURA"
            
            ZZRemito = Remito.Text
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZZZLinea = rstArticulo!Linea
                ZZZZTipo = rstArticulo!Tipo
                ZZZZCalidad = rstArticulo!Calidad
                ZZZZTamano = rstArticulo!Tamano
                rstArticulo.Close
            End If
            ZZCorte = ZZZZLinea + ZZZZTipo + ZZZZCalidad + ZZZZTamano
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Factura ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Factura ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "fecha ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Postal ,"
            ZSql = ZSql + "Partida ,"
            ZSql = ZSql + "Cuit  ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Dto ,"
            ZSql = ZSql + "Neto ,"
            ZSql = ZSql + "Neto1 ,"
            ZSql = ZSql + "Iva1 ,"
            ZSql = ZSql + "Iva2 ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Item ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Corte ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "Imprepago ,"
            ZSql = ZSql + "CondIva ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "Impre1 ,"
            ZSql = ZSql + "Impre3 ,"
            ZSql = ZSql + "Impre4 ,"
            ZSql = ZSql + "Cae ,"
            ZSql = ZSql + "VtoCae ,"
            ZSql = ZSql + "ImpreBarra ,"
            ZSql = ZSql + "ImpreBarraII ,"
            ZSql = ZSql + "DescriII ,"
            ZSql = ZSql + "CantiII ,"
            ZSql = ZSql + "PrecioII ,"
            ZSql = ZSql + "Dias ,"
            ZSql = ZSql + "PorceIva ,"
            ZSql = ZSql + "PordeDto )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZPunto + "',"
            ZSql = ZSql + "'" + ZZFactura + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZNombre + "',"
            ZSql = ZSql + "'" + Left$(ZZDireccion, 50) + "',"
            ZSql = ZSql + "'" + Left$(ZZLocalidad, 50) + "',"
            ZSql = ZSql + "'" + ZZPostal + "',"
            ZSql = ZSql + "'" + ZZPartida + "',"
            ZSql = ZSql + "'" + ZZCuit + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZDto + "',"
            ZSql = ZSql + "'" + ZZNeto + "',"
            ZSql = ZSql + "'" + ZZNeto1 + "',"
            ZSql = ZSql + "'" + ZZIva1 + "',"
            ZSql = ZSql + "'" + ZZIva2 + "',"
            ZSql = ZSql + "'" + ZZTotal + "',"
            ZSql = ZSql + "'" + ZZItem + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZCorte + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZImprepago + "',"
            ZSql = ZSql + "'" + ZZImpreIva + "',"
            ZSql = ZSql + "'" + ZZRemito + "',"
            ZSql = ZSql + "'" + ZZImpre1 + "',"
            ZSql = ZSql + "'" + ZZImpre3 + "',"
            ZSql = ZSql + "'" + ZZImpre4 + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + ZZImpreBarra + "',"
            ZSql = ZSql + "'" + ZZImpreBarraII + "',"
            ZSql = ZSql + "'" + ZZDescriII + "',"
            ZSql = ZSql + "'" + ZZCantiII + "',"
            ZSql = ZSql + "'" + ZZPrecioII + "',"
            ZSql = ZSql + "'" + Str$(ZZDias) + "',"
            ZSql = ZSql + "'" + ZZPorceIva + "',"
            ZSql = ZSql + "'" + Str$(ZZPorceDto) + "')"
                                    
            spFactura = ZSql
            Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
    Next A
    
    
    For A = ZZLugarFactura + 1 To 30

        ZZLugarFactura = ZZLugarFactura + 1

        ZZRenglon = Str$(ZZLugarFactura)
        Auxi1 = ZZRenglon
        Call Ceros(Auxi1, 2)
        ZZRenglon = Auxi1
        
        ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
        
        ZZItem = ""
        ZZArticulo = ""
        ZZDescripcion = ""
        ZZCantidad = ""
        ZZPrecio = ""
        ZZImporte = ""
        
        Call Numtolet
        ZZImpre1 = XTexto1
        ZZImpre2 = XTexto2
        
        Auxi2 = Numero.Text
        Call Ceros(Auxi2, 8)
        ZZImpre3 = Auxi2
        ZZImpre4 = "FACTURA"
        
        ZZRemito = Remito.Text
    
        ZZDescriII = ""
        ZZCantiII = ""
        ZZPrecioII = ""
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Factura ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Factura ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Cuit  ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Dto ,"
        ZSql = ZSql + "Neto ,"
        ZSql = ZSql + "Neto1 ,"
        ZSql = ZSql + "Iva1 ,"
        ZSql = ZSql + "Iva2 ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Imprepago ,"
        ZSql = ZSql + "CondIva ,"
        ZSql = ZSql + "Remito ,"
        ZSql = ZSql + "Impre1 ,"
        ZSql = ZSql + "Impre3 ,"
        ZSql = ZSql + "Impre4 ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "VtoCae ,"
        ZSql = ZSql + "ImpreBarra ,"
        ZSql = ZSql + "ImpreBarraII ,"
        ZSql = ZSql + "DescriII ,"
        ZSql = ZSql + "CantiII ,"
        ZSql = ZSql + "PrecioII ,"
        ZSql = ZSql + "PorceIva ,"
        ZSql = ZSql + "PordeDto )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZFactura + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZNombre + "',"
        ZSql = ZSql + "'" + Left$(ZZDireccion, 50) + "',"
        ZSql = ZSql + "'" + Left$(ZZLocalidad, 50) + "',"
        ZSql = ZSql + "'" + ZZPostal + "',"
        ZSql = ZSql + "'" + ZZPartida + "',"
        ZSql = ZSql + "'" + ZZCuit + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZImporte + "',"
        ZSql = ZSql + "'" + ZZDto + "',"
        ZSql = ZSql + "'" + ZZNeto + "',"
        ZSql = ZSql + "'" + ZZNeto1 + "',"
        ZSql = ZSql + "'" + ZZIva1 + "',"
        ZSql = ZSql + "'" + ZZIva2 + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + ZZArticulo + "',"
        ZSql = ZSql + "'" + ZZCantidad + "',"
        ZSql = ZSql + "'" + ZZPrecio + "',"
        ZSql = ZSql + "'" + ZZImprepago + "',"
        ZSql = ZSql + "'" + ZZImpreIva + "',"
        ZSql = ZSql + "'" + ZZRemito + "',"
        ZSql = ZSql + "'" + ZZImpre1 + "',"
        ZSql = ZSql + "'" + ZZImpre3 + "',"
        ZSql = ZSql + "'" + ZZImpre4 + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + ZZImpreBarra + "',"
        ZSql = ZSql + "'" + ZZImpreBarraII + "',"
        ZSql = ZSql + "'" + ZZDescriII + "',"
        ZSql = ZSql + "'" + ZZCantiII + "',"
        ZSql = ZSql + "'" + ZZPrecioII + "',"
        ZSql = ZSql + "'" + ZZPorceIva + "',"
        ZSql = ZSql + "'" + Str$(ZZPorceDto) + "')"
                                
        spFactura = ZSql
        Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)

    Next A
    
    
    Listado.WindowTitle = "Impresion de Proforma"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Factura.Factura, Factura.Fecha, Factura.Cliente, Factura.Importe, Factura.Dto, Factura.Articulo, Factura.Cantidad, Factura.Precio, Factura.Dias, Factura.PorceIva, Factura.PrecioII, " _
            + "Cliente.Cliente, Cliente.Razon, Cliente.Cuit, " _
            + "Articulo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Factura Factura, " _
            + DSQ + ".dbo.Cliente Cliente, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "Factura.Cliente = Cliente.Cliente AND " _
            + "Factura.Articulo = Articulo.Codigo AND " _
            + "Factura.Factura >= ' ' AND " _
            + "Factura.Factura <= '99999999'"
    
    Listado.Connect = Connect()
    
    Uno = "{Factura.Factura} in '' to '99999999'"
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    Listado.CopiesToPrinter = 2
    Rem Listado.Destination = 0
    
    Listado.ReportFileName = "ImpreRemito.rpt"
    
    Listado.Action = 1

End Sub

Private Sub BorraRenglon_Click()

    ZZDesderow = WVector1.Row
    ZZHastaRow = WVector1.RowSel
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
    
    For ZZCiclo = ZZDesderow To ZZHastaRow
        WVector1.Row = ZZCiclo
        For Ciclo = 1 To WVector1.Cols - 1
            Rem WVector1.Col = Ciclo
            Rem WVector1.Text = ""
            WVector1.TextMatrix(ZZCiclo, Ciclo) = ""
        Next Ciclo
    Next ZZCiclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WAuxi1 = WVector1.TextMatrix(Ciclo, 1)
        WAuxi3 = WVector1.TextMatrix(Ciclo, 3)
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WBorra(EntraVector, Ciclo1) = WVector1.TextMatrix(Ciclo, Ciclo1)
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            If da = 4 Or da = 5 Or da = 6 Then
                ZZControlPrecioII = WBorra(Ciclo, 9)
                If Val(ZZControlPrecioII) = 0 Then
                    WVector1.Col = da
                    WVector1.CellBackColor = &HFFFFC0
                    WVector1.Text = WBorra(Ciclo, da)
                        Else
                    WVector1.Col = da
                    WVector1.CellBackColor = &HFF00FF
                    WVector1.Text = WBorra(Ciclo, da)
                End If
                    Else
                WVector1.Col = da
                WVector1.Text = WBorra(Ciclo, da)
            End If
        Next da
    Next Ciclo
    
    Call Calcula_Click

End Sub

Private Sub Consulta_Click()

    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Pedidos"
    Opcion.AddItem "Condicion de Pago"
    Opcion.AddItem "Articulos"

    Opcion.Visible = True
     
 End Sub

Private Sub Entregada_Click()

    If Entregada.Value = 1 Then
    
        ZZLineas = 0
        Erase ZZPasaDatos
        
        For CicloII = 1 To 99
            
            ZZArticulo = WVector1.TextMatrix(CicloII, 1)
            ZZDesArticulo = WVector1.TextMatrix(CicloII, 2)
            ZZCantidad = WVector1.TextMatrix(CicloII, 3)
            ZZPrecio = WVector1.TextMatrix(CicloII, 4)
            ZZImporte = WVector1.TextMatrix(CicloII, 5)
            ZZStock = WVector1.TextMatrix(CicloII, 6)
            ZZClavePedido = WVector1.TextMatrix(CicloII, 8)
            ZZPrecioSB = WVector1.TextMatrix(CicloII, 10)
            
            If Val(ZZCantidad) <> 0 Then
                
                    
                ZZLineas = ZZLineas + 1
                ZZPasaDatos(ZZLineas, 1) = ZZArticulo
                ZZPasaDatos(ZZLineas, 2) = ZZDesArticulo
                ZZPasaDatos(ZZLineas, 3) = ZZCantidad
                ZZPasaDatos(ZZLineas, 4) = ZZPrecio
                ZZPasaDatos(ZZLineas, 5) = ZZImporte
                ZZPasaDatos(ZZLineas, 6) = ZZStock
                ZZPasaDatos(ZZLineas, 7) = "0"
                ZZPasaDatos(ZZLineas, 8) = ZZClavePedido
                ZZPasaDatos(ZZLineas, 10) = ZZPrecioSB
                
            End If
        
        Next CicloII

    
        ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
        
        For WRenglon = 1 To 99
                
            Articulo = ZZPasaDatos(WRenglon, 1)
            Cantidad = Val(ZZPasaDatos(WRenglon, 3))
            ClavePedido = ZZPasaDatos(WRenglon, 8)
                
            If Cantidad <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Pedido SET "
                ZSql = ZSql + " Entregado = Entregado + " + "'" + Str$(Cantidad) + "',"
                ZSql = ZSql + " Marca = " + "'" + "" + "'"
                ZSql = ZSql + " Where Clave = " + "'" + ClavePedido + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
                
        Next WRenglon
        
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
                
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
                
        WTipo = "01"
                
        Claveven$ = Letra.Text + WTipo + WPunto + Auxi + "01"
            
        If ZZNivelFactura = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
            
        ZSql = ""
        ZSql = ZSql + "UPDATE Ctacte SET "
        ZSql = ZSql + " Entregada = Entregada + " + "'" + "1" + "'"
        ZSql = ZSql + " Where Clave = " + "'" + Claveven$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
        If ZZNivelFactura = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        m$ = "Factura Actualizada"
        aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
        
    End If
End Sub

Private Sub Impresion_Click()

    Rem Call Impresion_Factura_Reimpre
    
    T$ = "Emision de Facturas"
    m$ = "Desea reimprimir la factura"
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
        ZZPasaImpre = 1
        Call Impresion_FacturaFe
    End If
    
    Rem T$ = "Emision de Facturas"
    Rem m$ = "Desea reimprimir el remito"
    Rem Rem Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    Rem If Respuestaaaaaa% = 6 Then
    Rem     ZZPasaImpre = 1
    Rem     Call Impresion_remito
    Rem End If

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
            Erase WPedido
            LugarPedido = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Cliente = " + "'" + Cliente + "'"
            ZSql = ZSql + " Order by Pedido.Numero"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Saldo = rstPedido!Cantidad - rstPedido!facturado
                            If Saldo > 0 Then
                                Entra = "S"
                                For Ciclo = 1 To LugarPedido
                                    If Val(WPedido(Ciclo)) = rstPedido!Numero Then
                                        Entra = "N"
                                        Exit For
                                    End If
                                Next Ciclo
                                If Entra = "S" Then
                                    LugarPedido = LugarPedido + 1
                                    WPedido(LugarPedido) = Str$(rstPedido!Numero)
                                    WNumero = Str$(rstPedido!Numero)
                                    Call Ceros(WNumero, 8)
                                    IngresaItem = WNumero + " " + rstPedido!Fecha + " " + rstPedido!Observaciones
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstPedido!Numero
                                    WIndice.AddItem IngresaItem
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPedido.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Order by CondPago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
            End If
            
        Case 3
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
    
    For A = 1 To 99
    
        WCantidad = Val(WVector1.TextMatrix(A, 3))
        WPrecio = Val(WVector1.TextMatrix(A, 5))
        
        If Letra.Text = "B" Then
            WWImpre = WPrecio * (1 + (ConfigIva1) / 100)
            Call Redondeo(WWImpre)
            WPrecio = WWImpre
        End If
        
        WParcial = (WPrecio * WCantidad)
        Call Redondeo(WParcial)
        WNeto = WNeto + WParcial
        
    Next A
    
    Call Calcula_Importe
    
End Sub

Private Sub Calcula_Importe()

    WIva1 = 0
    WIva2 = 0
    WSubTotal = WNeto
        
    Rem If WPorceIva > 0 Then
    Rem     WIva1 = WNeto * ((WPorceIva) / 100)
    Rem     Call Redondeo(WIva1)
    Rem End If
    
    
    
    WWIva(1) = WIva1
    WWIva(2) = WIva2
    
    WImpoDto = 0
    WDescuento = WPorceIva
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    
    WTotal = WNeto + WIva1 + WIva2
    
    SubTotal.Caption = Str$(WSubTotal)
    Dto.Caption = Str$(WImpoDto)
    Neto.Caption = Str$(WNeto)
    Iva1.Caption = Str$(WIva1)
    Iva2.Caption = Str$(WIva2)
    Total.Caption = Str$(WTotal)
    
    SubTotal.Caption = Pusing("###,###.##", SubTotal.Caption)
    Neto.Caption = Pusing("###,###.##", Neto.Caption)
    Iva1.Caption = Pusing("###,###.##", Iva1.Caption)
    Iva2.Caption = Pusing("###,###.##", Iva2.Caption)
    Total.Caption = Pusing("###,###.##", Total.Caption)

End Sub

Private Sub CmdClose_Click()
    PrgFacturaRemito.Hide
    Unload Me
    If ZZPasaProcesoFactura = 0 Then
        MenuVen.Show
            Else
        PrgControlPedido.Show
    End If
End Sub

Private Sub Graba_Click()

    If Val(ValorDolar.Text) = 0 Then
        m$ = "Se debe seleccionar el tipo de cambio correspondiente"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Fcaturas")
        Exit Sub
    End If

    T$ = "Grabacion de Presupuesto"
    m$ = "Desea emitir el presupuesto "
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
    
        ZZLineas = 0
        Erase ZZPasaDatos
        
        For CicloII = 1 To 99
            
            ZZArticulo = WVector1.TextMatrix(CicloII, 1)
            ZZDesArticulo = WVector1.TextMatrix(CicloII, 2)
            ZZCantidad = WVector1.TextMatrix(CicloII, 3)
            ZZDto = WVector1.TextMatrix(CicloII, 4)
            ZZPrecio = WVector1.TextMatrix(CicloII, 5)
            ZZImporte = WVector1.TextMatrix(CicloII, 6)
            ZZClavePedido = WVector1.TextMatrix(CicloII, 8)
            ZZPrecioSB = WVector1.TextMatrix(CicloII, 10)
            ZZMoneda = WVector1.TextMatrix(CicloII, 11)
            
            If Val(ZZCantidad) <> 0 Then
                
                    
                ZZLineas = ZZLineas + 1
                ZZPasaDatos(ZZLineas, 1) = ZZArticulo
                ZZPasaDatos(ZZLineas, 2) = ZZDesArticulo
                ZZPasaDatos(ZZLineas, 3) = ZZCantidad
                ZZPasaDatos(ZZLineas, 4) = ZZPrecio
                ZZPasaDatos(ZZLineas, 5) = ZZImporte
                ZZPasaDatos(ZZLineas, 8) = ZZClavePedido
                ZZPasaDatos(ZZLineas, 10) = ZZPrecioSB
                ZZPasaDatos(ZZLineas, 11) = ZZMoneda
                ZZPasaDatos(ZZLineas, 12) = ZZDto
                
            End If
        
        Next CicloII
    
        Rem If ZZLineas > 30 Then
        Rem     m$ = "La factura a emitor supera los 30 renglones"
        Rem     aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
        Rem     Exit Sub
        Rem End If
        
        Call Calcula_Click
        
        WNeto = Val(Neto.Caption)
        WIva1 = Val(Iva1.Caption)
        WIva2 = Val(Iva2.Caption)
        WTotal = Val(Total.Caption)
        
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
                
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
                
        WTipo = "01"
                
        Claveven$ = Letra.Text + WTipo + WPunto + Auxi + "01"
        
        If ZZNivelFactura = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Clave = " + "'" + Claveven$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            rstCtaCte.Close
            
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            m$ = "Factura ya emitida"
            aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
            
            Exit Sub
            
        End If
        
        Call Calcula_Click
        
        WNeto = Val(Neto.Caption)
        WIva1 = Val(Iva1.Caption)
        WIva2 = Val(Iva2.Caption)
        WTotal = Val(Total.Caption)
        
        Pasa = "S"
        
        If Letra.Text = "B" Then
            WNeto = Val(Total.Caption) / (1 + ((ConfigIva1) / 100))
            Call Redondeo(WNeto)
            WIva1 = WTotal - WNeto
            Neto.Caption = Str$(WNeto)
            Iva1.Caption = Str$(WIva1)
        End If
            
        Rem If Trim(Cae.Text) <> "" Then
        Rem     Exit Sub
        Rem End If
        
        Rem If Trim(Cae.Text) = "" Then
        Rem     Call Calcula_Cae
        Rem     If Trim(Cae.Text) = "" Then
        Rem         Exit Sub
        Rem     End If
        Rem End If
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CondPago"
        ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstCondPago.RecordCount > 0 Then
            WPlazo1 = Str$(rstCondPago!Dias)
            rstCondPago.Close
        End If
        
        WFecha = Fecha.Text
        Call Calcula_vencimiento(WFecha, WPlazo1, WVencimiento)
            
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
                
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
        
        ZZTipo = "01"
        ZZImpre = "FC"
                
        ZZPunto = WPunto
        ZZLetra = Letra.Text
        ZZNumero = Auxi
        ZZRenglon = "01"
        ZZCliente = UCase(Cliente.Text)
        ZZfecha = Fecha.Text
        ZZEstado = "0"
        ZZVencimiento = WVencimiento
        ZZTotal = Str$(WTotal)
        ZZSaldo = Str$(WTotal)
        If Letra.Text = "B" Then
            WNeto = WTotal / (1 + ((ConfigIva1) / 100))
            Call Redondeo(WNeto)
            WIva1 = WTotal - WNeto
            WIva2 = 0
            ZZNeto = Str$(WNeto)
            ZZIva1 = Str$(WIva1)
            ZZIva2 = Str$(WIva2)
                Else
            ZZNeto = Str$(WNeto)
            ZZIva1 = Str$(WIva1)
            ZZIva2 = Str$(WIva2)
        End If
        
        If Contado.Value = 1 Then
            ZZSaldo = 0
        End If
        
        ZZExento = "0"
        ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        ZZOrdVencimiento = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
        ZZPedido = Pedido.Text
        ZZRemito = ""
        ZZOrden = ""
        ZZProvincia = WProvincia
        ZZVendedor = ""
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
        ZZParidad = Str$(WWParidad)
        ZZRemito1 = ""
        ZZRemito2 = ""
        ZZBusqueda = ZZLetra + WPunto + Auxi
        
        ZZDescuento = ""
        ZZPago = Pago.Text
        ZZPartida = ""
        ZZExpreso = Expreso.Text
        ZZTipoIva = ""
        ZZComision = ""
        ZZRemito = Remito.Text
        
        ZZContado = Str$(Contado.Value)
        ZZEntregada = Str$(Entregada.Value)
        
        ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
        
        ZZLinea = ""
        
        ZZNetoTotal = ZZNeto
        
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
        ZSql = ZSql + "Exento ,"
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
        ZSql = ZSql + "Descuento ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Pago ,"
        ZSql = ZSql + "Linea ,"
        ZSql = ZSql + "Expreso ,"
        ZSql = ZSql + "TipoIva ,"
        ZSql = ZSql + "Comision ,"
        ZSql = ZSql + "NroRemito ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "VtoCae ,"
        ZSql = ZSql + "Contado ,"
        ZSql = ZSql + "Entregada ,"
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
        ZSql = ZSql + "'" + ZZExento + "',"
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
        ZSql = ZSql + "'" + ZZDescuento + "',"
        ZSql = ZSql + "'" + ZZPartida + "',"
        ZSql = ZSql + "'" + ZZPago + "',"
        ZSql = ZSql + "'" + ZZLinea + "',"
        ZSql = ZSql + "'" + ZZExpreso + "',"
        ZSql = ZSql + "'" + ZZTipoIva + "',"
        ZSql = ZSql + "'" + ZZComision + "',"
        ZSql = ZSql + "'" + ZZRemito + "',"
        ZSql = ZSql + "'" + ZZCae + "',"
        ZSql = ZSql + "'" + ZZVtoCae + "',"
        ZSql = ZSql + "'" + ZZContado + "',"
        ZSql = ZSql + "'" + ZZEntregada + "',"
        ZSql = ZSql + "'" + ZZBusqueda + "')"
                                
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        
        Renglon = 0
        WRenglon = 0
        ZZSumaComision = 0
        ZZRenglonMov = 0
            
        For A = 1 To 99
            
            WRenglon = WRenglon + 1
            
            Articulo = UCase(ZZPasaDatos(WRenglon, 1))
            DesArticulo = ZZPasaDatos(WRenglon, 2)
            Cantidad = Val(ZZPasaDatos(WRenglon, 3))
            Precio = Val(ZZPasaDatos(WRenglon, 4))
            Preciosalva = Val(ZZPasaDatos(WRenglon, 4))
            ZZZPedido = ZZPasaDatos(WRenglon, 7)
            ZZZClavePedido = ZZPasaDatos(WRenglon, 8)
            PrecioSb = Val(ZZPasaDatos(WRenglon, 10))
            Moneda = ZZPasaDatos(WRenglon, 11)
            Dto = ZZPasaDatos(WRenglon, 12)
                
            If Cantidad <> 0 Then
                        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                            
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
                
    
                ZZTipo = "01"
                ZZNumero = Numero.Text
                ZZRenglon = Renglon
                ZZArticulo = Articulo
                ZZDescripcion = DesArticulo
                ZZCantidad = Str$(Cantidad)
                ZZCantidadII = Str$(Cantidad)
                ZZPrecio = Str$(Precio)
                ZZPrecioSalva = Str$(Preciosalva)
                ZZPrecioUs = Str$(XXPrecio)
                ZZImporte = Str$(Precio * Cantidad)
                ZZImporteUs = Str$(XXPrecio * Cantidad)
                ZZCliente = Cliente.Text
                If Val(Moneda) = 0 Then
                    ZZParidad = "1"
                        Else
                    ZZParidad = ValorDolar.Text
                End If
                ZZVendedor = "0"
                ZZRubro = "0"
                ZZLinea = "0"
                ZZCosto2 = Str$(PrecioSb)
                ZZCoeficiente = "0"
                ZZPedido = ZZZPedido
                ZZClavePedido = ZZZClavePedido
                ZZfecha = Fecha.Text
                ZZImporte1 = "0"
                ZZImporte2 = "0"
                ZZImporte3 = "0"
                ZZImporte4 = "0"
                ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                ZZWArticulo = ""
                ZZRemito = ""
                
                ZZZLetra = Letra.Text
                
                ZZZPunto = Punto.Text
                Call Ceros(ZZZPunto, 1)
                
                ZZZNumeroFac = Numero.Text
                Call Ceros(ZZZNumeroFac, 6)
                
                ZZClave = "01" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi
                
                ZZWDate = Date$
                ZZClaveCtacte = "01" + Auxi1 + "01"
                
                ZZImprefactura = "FACTURA"
                ZZNroFactura = Auxi1
                ZZTalle = Talle
                ZZColor = XXColor
                ZZCuenta = WCuenta
                ZZDescuento = Dto
                ZZPartida = ""
                
                ZZCantidadII = ZZCantidad
                
                ZZPrecioII = Str$(XXPrecio)
                
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
                ZSql = ZSql + "PrecioSalva ,"
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
                ZSql = ZSql + "Comision ,"
                ZSql = ZSql + "TipoComision ,"
                ZSql = ZSql + "Coeficiente ,"
                ZSql = ZSql + "Pedido ,"
                ZSql = ZSql + "ClavePedido ,"
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
                ZSql = ZSql + "TipoII ,"
                ZSql = ZSql + "ClaveCtacte ,"
                ZSql = ZSql + "Imprefactura ,"
                ZSql = ZSql + "NroFactura ,"
                ZSql = ZSql + "Descuento ,"
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
                ZSql = ZSql + "'" + ZZPrecioSalva + "',"
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
                ZSql = ZSql + "'" + ZZComision + "',"
                ZSql = ZSql + "'" + ZZTipoComision + "',"
                ZSql = ZSql + "'" + ZZCoeficiente + "',"
                ZSql = ZSql + "'" + ZZPedido + "',"
                ZSql = ZSql + "'" + ZZClavePedido + "',"
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
                ZSql = ZSql + "'" + ZZTipoII + "',"
                ZSql = ZSql + "'" + ZZClaveCtacte + "',"
                ZSql = ZSql + "'" + ZZImprefactura + "',"
                ZSql = ZSql + "'" + ZZNroFactura + "',"
                ZSql = ZSql + "'" + ZZDescuento + "',"
                ZSql = ZSql + "'" + ZZPartida + "')"
                                
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                            
            End If
                                            
        Next A
        
        If ZZNivelFactura = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        
        ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
        
        For WRenglon = 1 To 99
                
            Articulo = ZZPasaDatos(WRenglon, 1)
            Cantidad = Val(ZZPasaDatos(WRenglon, 3))
            ClavePedido = ZZPasaDatos(WRenglon, 8)
                
            If Cantidad <> 0 Then
            
                If Cantidad <> 0 Then
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Pedido SET "
                    ZSql = ZSql + " Facturado = Facturado + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " Marca = " + "'" + "" + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + ClavePedido + "'"
                    spPedido = ZSql
                    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    
                
                    WWInsumoII = ""
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Articulo"
                    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WWInsumoII = IIf(IsNull(rstArticulo!InsumoII), "", rstArticulo!InsumoII)
                        rstArticulo.Close
                    End If
                                                    
                                                    
                    Rem dada
                    Rem dada
                    Rem dada
                    Rem dada
                    If Trim(WWInsumoII) = "" Then
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Stock = Stock - " + "'" + Str$(Cantidad) + "',"
                        ZSql = ZSql + " StockIV = StockIV - " + "'" + Str$(Cantidad) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            Else
                            
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Insumo SET "
                        ZSql = ZSql + " Stock = Stock - " + "'" + Str$(Cantidad) + "',"
                        ZSql = ZSql + " StockII = StockII - " + "'" + Str$(Cantidad) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + WInsumoII + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                        
                    
                        Rem
                        Rem doty de alta el movimiento en listado
                        Rem
                        
                        ZZTipoMov = "06"
                        ZZNumeroMov = Numero.Text
                        ZZRenglonMov = ZZRenglonMov + 1
                        
                        Auxi1 = Numero.Text
                        Call Ceros(Auxi1, 6)
                        Auxi2 = Str$(ZZRenglonMov)
                        Call Ceros(Auxi2, 2)
                        ZZClaveMov = ZZTipoMov + Auxi1 + Auxi2
                        
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO MovimientoInsumo ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "Tipo ,"
                        ZSql = ZSql + "Numero ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Insumo ,"
                        ZSql = ZSql + "Cantidad ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "OrdFecha ,"
                        ZSql = ZSql + "Deposito ,"
                        ZSql = ZSql + "DepositoII ,"
                        ZSql = ZSql + "Concepto )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + ZZClaveMov + "',"
                        ZSql = ZSql + "'" + ZZTipoMov + "',"
                        ZSql = ZSql + "'" + ZZNumeroMov + "',"
                        ZSql = ZSql + "'" + Str$(ZZRenglonMov) + "',"
                        ZSql = ZSql + "'" + WInsumoII + "',"
                        ZSql = ZSql + "'" + Str$(Cantidad * -1) + "',"
                        ZSql = ZSql + "'" + ZZfecha + "',"
                        ZSql = ZSql + "'" + ZZOrdFecha + "',"
                        ZSql = ZSql + "'" + "1" + "',"
                        ZSql = ZSql + "'" + "0" + "',"
                        ZSql = ZSql + "'" + "0" + "')"
                                        
                        spMovimientoInsumo = ZSql
                        Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    If Entregada.Value = 1 Then
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Pedido SET "
                        ZSql = ZSql + " Entregada = Entregada + " + "'" + Str$(Cantidad) + "',"
                        ZSql = ZSql + " Marca = " + "'" + "" + "'"
                        ZSql = ZSql + " Where Clave = " + "'" + ClavePedido + "'"
                        spPedido = ZSql
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                
                End If
                
            End If
                
        Next WRenglon
        
        WOrdUtimaCompra = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        ZSql = ""
        ZSql = ZSql + "UPDATE Cliente SET "
        ZSql = ZSql + " UltimaCompra = " + "'" + Fecha.Text + "',"
        ZSql = ZSql + " OrdUltimaCompra = " + "'" + WOrdUltimaCompra + "'"
        ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                
        ZZPasaImpre = 0
        
        T$ = "Emision de Facturas"
        m$ = "Desea imprimir el presupuesto"
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
            Call Impresion_FacturaFe
            Rem Call Impresion_remito
        End If
        
        
        
    
        Rem T$ = "Emision de Facturas"
        Rem m$ = "Desea Imprimir la Factura"
        Rem Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        Rem If Respuestaaaaaa% = 6 Then
        Rem     Call WImpresion
        Rem End If
            
        Rem Call Limpia_Click
        
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Fcaturas")
        
        Cliente.SetFocus
        
        Rem If ZZPasaProcesoFactura = 0 Then
        Rem     Cliente.SetFocus
        Rem         Else
        Rem     Call cmdClose_Click
        Rem End If
        
    End If
        
End Sub

Private Sub cmdDelete_Click()

    T$ = "Baja de Comprobantes"
    m$ = "Desea Borrar el Comprobante "
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then

        T$ = "Baja de Comprobantes"
        m$ = "Esta seguro que Desea Borrar el Comprobante "
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
        
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "01"
                
            Claveven$ = Letra.Text + WTipo + WPunto + Auxi + "01"

            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txt7Password = "Sw58125812"
                txtOdbc = "FraganciasII"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
               
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Clave = " + "'" + Claveven$ + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
            
                ZSaldo = rstCtaCte!Saldo
                ZTotal = rstCtaCte!Total
                rstCtaCte.Close
                
                If ZSaldo <> ZTotal Then
                
                    m$ = "El comprobante se encuentra total o parcialmente cancelado"
                    aaaaaa% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
                    
                    If ZZNivelFactura = 1 Then
                        txtUserName = "SA"
                        txtPassword = "Sw58125812"
                        txtOdbc = "Fragancias"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End If
                    
                    Exit Sub
                    
                End If
            End If
        
            ZBaja = "N"
            If Val(Pedido.Text) <> 0 Then
                T$ = "Baja de Comprobantes"
                m$ = "Desea Restaurar el saldo del pedido"
                Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
                If Respuestaaaaaa% = 6 Then
                    ZBaja = "S"
                End If
            End If
        
            Erase WVector
            Erase ZZVector
        
            For WRenglon = 1 To 99
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "01" + Auxi + Auxi1
                
                ZZZLetra = Letra.Text
                
                ZZZPunto = Punto.Text
                Call Ceros(ZZZPunto, 1)
                
                ZZZNumeroFac = Numero.Text
                Call Ceros(ZZZNumeroFac, 6)
                
                ZZClave = "01" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi1
                
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + ZZClave + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                
                    ZZVector(WRenglon, 1) = rstEstadistica!Articulo
                    ZZVector(WRenglon, 2) = Str$(rstEstadistica!Cantidad)
                    ZZVector(WRenglon, 4) = IIf(IsNull(rstEstadistica!ClavePedido), "", rstEstadistica!ClavePedido)
                        
                    rstEstadistica.Close
                    
                End If
                
            Next WRenglon
            
            ZSql = ""
            ZSql = ZSql + "DELETE Estadistica"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            
            ZSql = ""
            ZSql = ZSql + "DELETE CtaCte"
            ZSql = ZSql + " Where Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            
            ZSql = ""
            ZSql = ZSql + "DELETE MovimientoInsumo"
            ZSql = ZSql + " Where MovimientoInsumo.Tipo = " + "'" + "6" + "'"
            ZSql = ZSql + " and MovimientoInsumo.Numero = " + "'" + Numero.Text + "'"
            spMovimientoInsumo = ZSql
            Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
            For WRenglon = 1 To 99
            
                Articulo = ZZVector(WRenglon, 1)
                Cantidad = Val(ZZVector(WRenglon, 2))
                ZZClavePedido = ZZVector(WRenglon, 4)
                
                
                WWInsumoII = ""
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WWInsumoII = IIf(IsNull(rstArticulo!InsumoII), "", rstArticulo!InsumoII)
                    rstArticulo.Close
                End If
                                                
                If Trim(WInsumoII) = "" Then
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " StockIV = StockIV + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Insumo SET "
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " StockII = StockII + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + WInsumoII + "'"
                    spInsumo = ZSql
                    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
                ZSql = ""
                ZSql = ZSql + "UPDATE Pedido SET "
                ZSql = ZSql + " Facturado = Facturado - " + "'" + Str$(Cantidad) + "'"
                ZSql = ZSql + " Where Clave = " + "'" + ZZClavePedido + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                
            Next WRenglon
            
            Call Limpia_Click
            
            Cliente.SetFocus
            
        End If
    
    End If

End Sub

Private Sub Limpia_Click()

    Call Limpia_Vector

    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    DesClienteII.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Expreso.Text = ""
    Remito.Text = ""
    
    Renglon = 0
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
    Contado.Value = 0
    Entregada.Value = 0
    
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
            Indice = Pantalla.ListIndex
            Pedido.Text = WIndice.List(Indice)
            Call Pedido_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            Pago.Text = WIndice.List(Indice)
            Call Pago_Keypress(13)
            
        Case 3
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            
            ZPasa = "S"
            If Val(Pedido.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Pedido"
                ZSql = ZSql + " Where Numero = " + "'" + Pedido.Text + "'"
                ZSql = ZSql + " and Articulo = " + "'" + Claveven$ + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                If rstPedido.RecordCount > 0 Then
                    rstPedido.Close
                        Else
                    m$ = "El articulo no esta en el pedido"
                    aaaaaa% = MsgBox(m$, 0, "Carga de Articulos")
                    ZPasa = "N"
                End If
                
            End If
            
            If ZPasa = "S" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Claveven$ + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 1
                    WVector1.Text = rstArticulo!Codigo
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    WVector1.Col = 4
                    WVector1.Text = Str$(rstArticulo!Precio)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    WVector1.Col = 6
                    WVector1.Text = Str$(rstArticulo!Stock)
                    WVector1.Col = 3
                    rstArticulo.Close
                    Call StartEdit
                End If
                
                    Else
                    
                WVector1.Col = 1
                WVector1.Text = ""
                WVector1.Col = 2
                WVector1.Text = ""
                WVector1.Col = 4
                WVector1.Text = ""
                    
                WVector1.Col = 1
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
    Provincia(23) = "1Tierra del Fuego"
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
    
    TipoParidad.Clear
    
    TipoParidad.AddItem "Oficial"
    TipoParidad.AddItem "Paralelo"
    TipoParidad.AddItem ""
    
    TipoParidad.ListIndex = 2
    
    
    TipoParidadOtro.Clear
    
    TipoParidadOtro.AddItem "Oficial"
    TipoParidadOtro.AddItem "Paralelo"
    TipoParidadOtro.AddItem ""
    
    TipoParidadOtro.ListIndex = 2
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"
    
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    DesClienteII.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Expreso.Text = ""
    Remito.Text = ""
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
    Contado.Value = 0
    Entregada.Value = 0
    
    WWParidad = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + "1" + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        WWParidad = rstDolar!ParidadII
    End If
    ValorDolar.Text = Pusing("###,###.##", Str$(WWParidad))
    
    ValorDolar.Text = "1"
    
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
        Rem ConfigPunto = rstConfiguracion!Punto
        ConfigPunto = 2
        CantiFac = rstConfiguracion!CantiFac
        CantiRem = rstConfiguracion!CantiRem
        CantiArti = rstConfiguracion!CantiArti
        rstConfiguracion.Close
    End If
    
    If ZZPasaProcesoFactura = 1 Then
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Numero = " + "'" + ZZPasaPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Cliente.Text = rstPedido!Cliente
            rstPedido.Close
            Call Cliente_KeyPress(13)
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                DesPago.Caption = Trim(rstCondPago!Nombre)
                rstCondPago.Close
            End If
            
            Pedido.Text = ZZPasaPedido
            PantaDolar.Visible = True
            Rem TipoParidadOtro.SetFocus
            
            Rem Call Pedido_KeyPress(13)
            Rem Call Lee_Pedido
            Rem WVector1.Col = 1
            Rem WVector1.Row = 1
            Rem Call StartEdit
            
                Else
            m$ = "Pedido Inexistente"
            aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
            Exit Sub
        End If
    
    End If
    
End Sub

Private Sub Form_LoadOtro()

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
    Provincia(23) = "1Tierra del Fuego"
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
    
    TipoParidad.Clear
    
    TipoParidad.AddItem "Oficial"
    TipoParidad.AddItem "Paralelo"
    TipoParidad.AddItem ""
    
    TipoParidad.ListIndex = 2
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"
    
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    DesClienteII.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Expreso.Text = ""
    Remito.Text = ""
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
    Contado.Value = 0
    Entregada.Value = 0
    
    WWParidad = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + "1" + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        WWParidad = rstDolar!ParidadII
    End If
    ValorDolar.Text = Pusing("###,###.##", Str$(WWParidad))
    
    ValorDolar.Text = "1"
    
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
        Rem ConfigPunto = rstConfiguracion!Punto
        ConfigPunto = 2
        CantiFac = rstConfiguracion!CantiFac
        CantiRem = rstConfiguracion!CantiRem
        CantiArti = rstConfiguracion!CantiArti
        rstConfiguracion.Close
    End If
    
    If ZZPasaProcesoFactura = 1 Then
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Numero = " + "'" + ZZPasaPedido + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Cliente.Text = rstPedido!Cliente
            rstPedido.Close
            Call Cliente_KeyPress(13)
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                DesPago.Caption = Trim(rstCondPago!Nombre)
                rstCondPago.Close
            End If
            
            m$ = "Recuerde que debe seleccionar el tipo de cambio"
            aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
            
            
            
            Pedido.Text = ZZPasaPedido
            Call Pedido_KeyPress(13)
            Call Lee_Pedido
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
            
                Else
            m$ = "Pedido Inexistente"
            aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
            Exit Sub
        End If
    
    End If
    
End Sub


Private Sub Proceso_Click()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    If ZZNivelFactura = 1 Then
        txtUserName = "SA"
        txtPassword = "Sw58125812"
        txtOdbc = "FraganciasII"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If
    
    For WRenglon = 1 To 99
        
        If Val(Punto.Text) = 1 Then
    
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "01" + Auxi + Auxi1
            
                Else
            
            ZZZLetra = Letra.Text
            
            ZZZPunto = Punto.Text
            Call Ceros(ZZZPunto, 1)
            
            ZZZNumeroFac = Numero.Text
            Call Ceros(ZZZNumeroFac, 6)
                 
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "01" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi1
                
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
            WVector1.Text = Pusing("###,###", Str$(rstEstadistica!Cantidad))
                
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###", Str$(rstEstadistica!Descuento))
                
            WVector1.Col = 5
            WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Precio))
            
            WVector1.Col = 6
            WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Precio * rstEstadistica!CantidadII))
            
            ZZPedido = IIf(IsNull(rstEstadistica!Pedido), "", rstEstadistica!Pedido)
            ZZClavePedido = IIf(IsNull(rstEstadistica!ClavePedido), "", rstEstadistica!ClavePedido)
            
            WVector1.Col = 7
            WVector1.Text = ZZPedido
            
            WVector1.Col = 8
            WVector1.Text = ZZClavePedido
                
            WVector1.Col = 10
            WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Costo2))
            
            rstEstadistica.Close
                
        End If
    
    Next WRenglon
    
    If ZZNivelFactura = 1 Then
        txtUserName = "SA"
        txtPassword = "Sw58125812"
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If

    For WRenglon = 1 To Renglon
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WVector1.TextMatrix(WRenglon, 2) = rstArticulo!Descripcion
            rstArticulo.Close
        End If
            
    Next WRenglon

    Call Calcula_Click
    
    Graba.Enabled = True

End Sub

Private Sub ValorDolar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    
    On Error GoTo WError
    
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
        
            DesCliente.Caption = Trim(rstCliente!Fantasia)
            DesClienteII.Caption = Trim(rstCliente!Razon)
            Pago.Text = rstCliente!Condicion
            Expreso.Text = rstCliente!Expreso
            WProvincia = rstCliente!Provincia
            WCodIva = rstCliente!Iva
            WPorceIva = rstCliente!PorceIva
            WRazon = Trim(rstCliente!Fantasia)
            WDireccion = Trim(rstCliente!Direccion)
            WLocalidad = Trim(rstCliente!Localidad)
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            Select Case Val(WCodIva)
                Case 1, 2
                    Letra.Text = "A"
                Case Else
                    Letra.Text = "B"
            End Select
            If ZZNivelFactura = 1 Then
                Letra.Text = "X"
            End If
            ZMarca = IIf(IsNull(rstCliente!Marca), "0", rstCliente!Marca)
            
            Rem If Letra.Text = "B" Then
            Rem     m$ = "COLOQUE EL FORMULARIO B"
            Rem     aaaaaa% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Rem End If
            
            rstCliente.Close
            Call Lee_CtaCte
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                DesPago.Caption = rstCondPago!Nombre
                rstCondPago.Close
                    Else
                DesPago.Caption = ""
            End If
            
            WPunto = Str(ConfigPunto)
            Call Ceros(WPunto, 4)
            Punto.Text = WPunto
                
            Numero.Text = "1"
            WTipo = "01"
    
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "FraganciasII"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select CtaCte.Letra, CtaCte.Punto, CtaCte.Numero, CtaCte.Tipo, CtaCTe.Remito"
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
                                Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 3 Then
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
            
            ZRemito1 = "0"
            ZRemito2 = "0"
            
            If ZZNivelFactura = 0 Then
                
                ZSql = ""
                ZSql = ZSql + "Select CtaCte.Letra, CtaCte.Punto, CtaCte.Numero, CtaCte.Tipo, CtaCTe.NroRemito"
                ZSql = ZSql + " FROM Ctacte"
                ZSql = ZSql + " Where Ctacte.Punto = " + "'" + Punto.Text + "'"
                ZSql = ZSql + " and Ctacte.Tipo = " + "'" + "01" + "'"
                ZSql = ZSql + " and Ctacte.Letra = " + "'" + "A" + "'"
                ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
                ZSql = ZSql + " Order by Ctacte.Numero"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                    With rstCtaCte
                        .MoveLast
                        Do
                            If .BOF = False Then
                        
                                If Punto.Text = rstCtaCte!Punto Then
                                    Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                    If Val(rstCtaCte!Tipo) = 1 Then
                                        ZRemito1 = Str$(rstCtaCte!NroRemito + 1)
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
                
                ZSql = ""
                ZSql = ZSql + "Select CtaCte.Letra, CtaCte.Punto, CtaCte.Numero, CtaCte.Tipo, CtaCTe.NroRemito"
                ZSql = ZSql + " FROM Ctacte"
                ZSql = ZSql + " Where Ctacte.Punto = " + "'" + Punto.Text + "'"
                ZSql = ZSql + " and Ctacte.Tipo = " + "'" + "01" + "'"
                ZSql = ZSql + " and Ctacte.Letra = " + "'" + "B" + "'"
                ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
                ZSql = ZSql + " Order by Ctacte.Numero"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                    With rstCtaCte
                        .MoveLast
                        Do
                            If .BOF = False Then
                        
                                If Punto.Text = rstCtaCte!Punto Then
                                    Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                    If Val(rstCtaCte!Tipo) = 1 Then
                                        ZRemito2 = Str$(rstCtaCte!NroRemito + 1)
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
    
            End If
    
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            If Val(ZRemito1) > Val(ZRemito2) Then
                Remito.Text = ZRemito1
                    Else
                Remito.Text = ZRemito2
            End If
            
            Pedido.SetFocus
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM HistorialCliente"
            ZSql = ZSql + " Where HistorialCliente.Cliente = " + "'" + Cliente.Text + "'"
            spHistorialCliente = ZSql
            Set rstHistorialCliente = db.OpenRecordset(spHistorialCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstHistorialCliente.RecordCount > 0 Then
                rstHistorialCliente.Close
                ZZPasaCliente = Cliente.Text
                ZZPasaProceso = 0
                PrgHistorialClienteConsulta.Show
            End If
            
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
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Numero.Text
        Call Ceros(WPunto, 4)
        
        Numero.Text = "1"
        WTipo = "01"
    
        If ZZNivelFactura = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
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
                            Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                            If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 3 Then
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
    
        If ZZNivelFactura = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
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
            
        WTipo = "01"
            
        Claveven$ = Letra.Text + WTipo + WPunto + Auxi + "01"
    
        If ZZNivelFactura = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
           
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Clave = " + "'" + Claveven$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            
            Fecha.Text = rstCtaCte!Fecha
            Cliente.Text = rstCtaCte!Cliente
            Pedido.Text = Str$(Val(rstCtaCte!Pedido))
            Pago.Text = rstCtaCte!Pago
            Expreso.Text = rstCtaCte!Expreso
            Remito.Text = rstCtaCte!NroRemito
            Contado.Value = rstCtaCte!Contado
            Entregada.Value = rstCtaCte!Entregada
            ValorDolar.Text = Str$(rstCtaCte!Paridad)

            rstCtaCte.Close
    
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Fantasia
                DesClienteII.Caption = rstCliente!Razon
                Rem Descuento.Text = Str$(rstCliente!Descuento)
                Rem Descuento.Text = Pusing("###,###.##", Descuento.Text)
                WProvincia = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WPorceIva = rstCliente!PorceIva
                WRazon = rstCliente!Fantasia
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                rstCliente.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                DesPago.Caption = Trim(rstCondPago!Nombre)
                rstCondPago.Close
            End If
            
            Call Proceso_Click
                
                Else
    
            If ZZNivelFactura = 1 Then
                txtUserName = "SA"
                txtPassword = "Sw58125812"
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End If
                    
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
            Pedido.SetFocus
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

Private Sub Pedido_KeyPress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
        If Val(Pedido.Text) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Numero = " + "'" + Pedido.Text + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZZCliente = rstPedido!Cliente
                
                rstPedido.Close
                If Trim(UCase(ZZCliente)) <> Trim(UCase(Cliente.Text)) Then
                    m$ = "El cliente informado no concuerda con el del pedido"
                    aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
                    Exit Sub
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CondPago"
                ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
                spCondPago = ZSql
                Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
                If rstCondPago.RecordCount > 0 Then
                    DesPago.Caption = Trim(rstCondPago!Nombre)
                    rstCondPago.Close
                End If
                
                    Else
                m$ = "Pedido Inexistente"
                aaaaaa% = MsgBox(m$, 0, "Ingreso de Facturas")
                Exit Sub
            End If
        
            Expreso.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Pedido.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
    
    Exit Sub
    
WError:
    Resume Next
End Sub

Private Sub Pago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CondPago"
        ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstCondPago.RecordCount > 0 Then
            DesPago.Caption = rstCondPago!Nombre
            rstCondPago.Close
            Expreso.SetFocus
                Else
            Pago.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Pago.Text = ""
        DesPago.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub EXPRESO_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Lee_Pedido
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Expreso.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by CondPago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
            End If
    
        Case 3
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

Private Sub Reconstruccion_Click()
    FechaRecon.Text = "  /  /    "
    PantaRecon.Visible = True
    FechaRecon.SetFocus
End Sub

Private Sub TipoParidad_Change()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + "1" + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        Select Case TipoParidad.ListIndex
            Case 0
                WWParidad = rstDolar!Paridad
            Case Else
                WWParidad = rstDolar!ParidadII
        End Select
    End If
    ValorDolar.Text = Pusing("###,###.##", Str$(WWParidad))
    Call Calcula_Click
End Sub

Private Sub TipoParidad_Click()
    WWParidad = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + "1" + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        Select Case TipoParidad.ListIndex
            Case 0
                WWParidad = rstDolar!Paridad
            Case Else
                WWParidad = rstDolar!ParidadII
        End Select
    End If
    ValorDolar.Text = Pusing("###,###.##", Str$(WWParidad))
    Call Pedido_KeyPress(13)
    Call Lee_Pedido
    WVector1.Col = 1
    WVector1.Row = 1
    Call StartEdit
End Sub

Private Sub TipoParidadOtro_Click()
    If TipoParidadOtro.ListIndex = 0 Or TipoParidadOtro.ListIndex = 1 Then
        PantaDolar.Visible = False
        TipoParidad.ListIndex = TipoParidadOtro.ListIndex
        Rem Call TipoParidad_Click
    End If
End Sub

Private Sub TipoParidadOtro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PantaDolar.Visible = False
        TipoParidad.ListIndex = TipoParidadOtro.ListIndex
        Rem Call TipoParidad_Click
    End If
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,f5,f9,F10
        Case 112, 113, 114, 115, 116, 120, 121
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
            
        Rem F1,F2,F3,F4,f5,f9,F10
        Case 112, 113, 114, 115, 116, 120, 121
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
            
        Rem F1,F2,F3,F4,f5,f9,F10
        Case 112, 113, 114, 115, 116, 120, 121
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
        Case 3
            Rem If WVector1.Row < WVector1.Rows - 1 Then
            If WVector1.Row < 99 Then
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
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Numero = " + "'" + Pedido.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + WVector1.Text + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZPasa = "S"
                rstPedido.Close
                    Else
                Rem m$ = "El articulo no esta en el pedido"
                Rem aaaaaa% = MsgBox(m$, 0, "Carga de Articulos")
                Rem ZPasa = "N"
                Rem WControl = "N"
                ZPasa = "S"
            End If
            
            If ZPasa = "S" Then
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    Rem WVector1.Col = 4
                    Rem WVector1.Text = Str$(rstArticulo!Precio)
                    Rem WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    Rem WVector1.Col = 6
                    Rem WVector1.Text = Str$(rstArticulo!Stock)
                    WVector1.Col = 2
                    rstArticulo.Close
                        Else
                    WControl = "N"
                End If
        
            End If
            
        Case 3
            WCantidad = Val(WVector1.Text)
        
            WVector1.TextMatrix(WVector1.Row, 6) = Str$(WCantidad * Val(WVector1.TextMatrix(WVector1.Row, 5)))
            WVector1.TextMatrix(WVector1.Row, 6) = Pusing("###,###.##", WVector1.TextMatrix(WVector1.Row, 6))
            
        Case 5
            WCantidad = Val(WVector1.TextMatrix(WVector1.Row, 3))
        
            WVector1.TextMatrix(WVector1.Row, 6) = Str$(WCantidad * Val(WVector1.TextMatrix(WVector1.Row, 5)))
            WVector1.TextMatrix(WVector1.Row, 6) = Pusing("###,###.##", WVector1.TextMatrix(WVector1.Row, 6))
            
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

    Opcion.Clear
    
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 3
    
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
    WVector1.Cols = 12
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 25
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
                WVector1.Text = "Dto"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 5
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 6
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 7
                WVector1.Text = "Pedido"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 100
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 100
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 100
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 100
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
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
    Opcion.AddItem "Pedidos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Pedido_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Pedidos"
    Opcion.AddItem "Condicion"
    Opcion.AddItem "Articulo"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Pago_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Pedidos"
    Opcion.AddItem "Condicion"
    Opcion.AddItem "Articulo"

    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub PedidoAyuda_Click()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulo"
    Opcion.AddItem "Pedidos"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Numtolet()

    'Convertir en letras el n�mero en Text1
    
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

Private Sub Vencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Pedido_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Descuento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Remito_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Pago_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 116
            Call PedidoAyuda_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub


Private Sub Impresion_remito()


    ZSql = ""
    ZSql = ZSql + "DELETE Factura"
    spFactura = ZSql
    Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZZProvincia = rstCliente!Provincia
        ZZCodIva = rstCliente!Iva
        ZZRazon = rstCliente!Fantasia
        ZZDireccion = rstCliente!Direccion
        ZZLocalidad = rstCliente!Localidad
        ZZPostal = rstCliente!Postal
        ZZCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZZCuitII = ""
    
    
    
    
    ZZLetra = "X"
    ZZTipo = "01"
    ZZPunto = "0001"
    Auxi1 = Numero.Text
    Call Ceros(Auxi1, 8)
    ZZFactura = Auxi1
    ZZfecha = Fecha.Text
    ZZCliente = Cliente.Text
    ZZNombre = Trim(ZZRazon)
    ZZDireccion = Trim(ZZDireccion) + " - " + Trim(ZZLocalidad) + " - " + Trim(Provincia(ZZProvincia))
    ZZLocalidad = Trim(ZZLocalidad) + " - " + Trim(Provincia(ZZProvincia))
    ZZLocalidad = Left$(ZZDireccion, 50)
    ZZLocalidad = Left$(ZZLocalidad, 50)
    ZZPartida = ""
    ZZNeto = Neto.Caption
    ZZDto = Dto.Caption
    ZZNeto1 = SubTotal.Caption
    ZZIva1 = Iva1.Caption
    ZZIva2 = Iva2.Caption
    ZZTotal = Total.Caption
    ZZImprepago = Left$(DesPago.Caption, 35)
    ZZImpreIva = Iva(Val(ZZCodIva))
    ZZPorceIva = "21"
    ZZPorceDto = 0
    ZZPostal = WPostal
    
    ZZLugarFactura = 0
    
    Call Numtolet
    
    For A = 1 To 99
    
        If Trim(WVector1.TextMatrix(A, 1)) <> "" Then
            
            ZZLugarFactura = ZZLugarFactura + 1
    
            ZZRenglon = Str$(ZZLugarFactura)
            Auxi1 = ZZRenglon
            Call Ceros(Auxi1, 2)
            ZZRenglon = Auxi1
            
            ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
            
            ZZItem = Str$(A)
            
            ZZArticulo = WVector1.TextMatrix(A, 1)
            ZZDescripcion = WVector1.TextMatrix(A, 2)
            ZZZCantidad = Val(WVector1.TextMatrix(A, 3))
            ZZZPrecio = Val(WVector1.TextMatrix(A, 5))
            ZZZImporte = ZZZPrecio * ZZZCantidad
            
            ZZCantidad = Str$(ZZZCantidad)
            ZZPrecio = Str$(ZZZPrecio)
            ZZImporte = Str$(ZZZImporte)
            
            If Trim(ZZArticulo) = "" Then
                ZZItem = ""
                ZZArticulo = ""
                ZZDescripcion = ""
                ZZCantidad = ""
                ZZPrecio = ""
                ZZImporte = ""
            End If
            
            ZZDescriII = ""
            ZZCantiII = ""
            ZZPrecioII = ""
            
            Call Numtolet
            ZZImpre1 = XTexto1
            ZZImpre2 = XTexto2
            
            Auxi2 = Numero.Text
            Call Ceros(Auxi2, 8)
            ZZImpre3 = Auxi2
            ZZImpre4 = "FACTURA"
            
            ZZRemito = Remito.Text
            
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Factura ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Factura ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "fecha ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Postal ,"
            ZSql = ZSql + "Partida ,"
            ZSql = ZSql + "Cuit  ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Dto ,"
            ZSql = ZSql + "Neto ,"
            ZSql = ZSql + "Neto1 ,"
            ZSql = ZSql + "Iva1 ,"
            ZSql = ZSql + "Iva2 ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Item ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "Imprepago ,"
            ZSql = ZSql + "CondIva ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "Impre1 ,"
            ZSql = ZSql + "Impre3 ,"
            ZSql = ZSql + "Impre4 ,"
            ZSql = ZSql + "Cae ,"
            ZSql = ZSql + "VtoCae ,"
            ZSql = ZSql + "ImpreBarra ,"
            ZSql = ZSql + "ImpreBarraII ,"
            ZSql = ZSql + "DescriII ,"
            ZSql = ZSql + "CantiII ,"
            ZSql = ZSql + "PrecioII ,"
            ZSql = ZSql + "PorceIva ,"
            ZSql = ZSql + "PordeDto )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZPunto + "',"
            ZSql = ZSql + "'" + ZZFactura + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZNombre + "',"
            ZSql = ZSql + "'" + Left$(ZZDireccion, 50) + "',"
            ZSql = ZSql + "'" + Left$(ZZLocalidad, 50) + "',"
            ZSql = ZSql + "'" + ZZPostal + "',"
            ZSql = ZSql + "'" + ZZPartida + "',"
            ZSql = ZSql + "'" + ZZCuit + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZDto + "',"
            ZSql = ZSql + "'" + ZZNeto + "',"
            ZSql = ZSql + "'" + ZZNeto1 + "',"
            ZSql = ZSql + "'" + ZZIva1 + "',"
            ZSql = ZSql + "'" + ZZIva2 + "',"
            ZSql = ZSql + "'" + ZZTotal + "',"
            ZSql = ZSql + "'" + ZZItem + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZImprepago + "',"
            ZSql = ZSql + "'" + ZZImpreIva + "',"
            ZSql = ZSql + "'" + ZZRemito + "',"
            ZSql = ZSql + "'" + ZZImpre1 + "',"
            ZSql = ZSql + "'" + ZZImpre3 + "',"
            ZSql = ZSql + "'" + ZZImpre4 + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + ZZImpreBarra + "',"
            ZSql = ZSql + "'" + ZZImpreBarraII + "',"
            ZSql = ZSql + "'" + ZZDescriII + "',"
            ZSql = ZSql + "'" + ZZCantiII + "',"
            ZSql = ZSql + "'" + ZZPrecioII + "',"
            ZSql = ZSql + "'" + ZZPorceIva + "',"
            ZSql = ZSql + "'" + Str$(ZZPorceDto) + "')"
                                    
            spFactura = ZSql
            Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
    Next A
    
    
    For A = ZZLugarFactura + 1 To 30

        ZZLugarFactura = ZZLugarFactura + 1

        ZZRenglon = Str$(ZZLugarFactura)
        Auxi1 = ZZRenglon
        Call Ceros(Auxi1, 2)
        ZZRenglon = Auxi1
        
        ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
        
        ZZItem = ""
        ZZArticulo = ""
        ZZDescripcion = ""
        ZZCantidad = ""
        ZZPrecio = ""
        ZZImporte = ""
        
        Call Numtolet
        ZZImpre1 = XTexto1
        ZZImpre2 = XTexto2
        
        Auxi2 = Numero.Text
        Call Ceros(Auxi2, 8)
        ZZImpre3 = Auxi2
        ZZImpre4 = "FACTURA"
        
        ZZRemito = Remito.Text
    
        ZZDescriII = ""
        ZZCantiII = ""
        ZZPrecioII = ""
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Factura ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Factura ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Cuit  ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Dto ,"
        ZSql = ZSql + "Neto ,"
        ZSql = ZSql + "Neto1 ,"
        ZSql = ZSql + "Iva1 ,"
        ZSql = ZSql + "Iva2 ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Imprepago ,"
        ZSql = ZSql + "CondIva ,"
        ZSql = ZSql + "Remito ,"
        ZSql = ZSql + "Impre1 ,"
        ZSql = ZSql + "Impre3 ,"
        ZSql = ZSql + "Impre4 ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "VtoCae ,"
        ZSql = ZSql + "ImpreBarra ,"
        ZSql = ZSql + "ImpreBarraII ,"
        ZSql = ZSql + "DescriII ,"
        ZSql = ZSql + "CantiII ,"
        ZSql = ZSql + "PrecioII ,"
        ZSql = ZSql + "PorceIva ,"
        ZSql = ZSql + "PordeDto )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZFactura + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZNombre + "',"
        ZSql = ZSql + "'" + Left$(ZZDireccion, 50) + "',"
        ZSql = ZSql + "'" + Left$(ZZLocalidad, 50) + "',"
        ZSql = ZSql + "'" + ZZPostal + "',"
        ZSql = ZSql + "'" + ZZPartida + "',"
        ZSql = ZSql + "'" + ZZCuit + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZImporte + "',"
        ZSql = ZSql + "'" + ZZDto + "',"
        ZSql = ZSql + "'" + ZZNeto + "',"
        ZSql = ZSql + "'" + ZZNeto1 + "',"
        ZSql = ZSql + "'" + ZZIva1 + "',"
        ZSql = ZSql + "'" + ZZIva2 + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + ZZArticulo + "',"
        ZSql = ZSql + "'" + ZZCantidad + "',"
        ZSql = ZSql + "'" + ZZPrecio + "',"
        ZSql = ZSql + "'" + ZZImprepago + "',"
        ZSql = ZSql + "'" + ZZImpreIva + "',"
        ZSql = ZSql + "'" + ZZRemito + "',"
        ZSql = ZSql + "'" + ZZImpre1 + "',"
        ZSql = ZSql + "'" + ZZImpre3 + "',"
        ZSql = ZSql + "'" + ZZImpre4 + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + ZZImpreBarra + "',"
        ZSql = ZSql + "'" + ZZImpreBarraII + "',"
        ZSql = ZSql + "'" + ZZDescriII + "',"
        ZSql = ZSql + "'" + ZZCantiII + "',"
        ZSql = ZSql + "'" + ZZPrecioII + "',"
        ZSql = ZSql + "'" + ZZPorceIva + "',"
        ZSql = ZSql + "'" + Str$(ZZPorceDto) + "')"
                                
        spFactura = ZSql
        Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)

    Next A
    
    
    Listado.WindowTitle = "Impresion de Proforma"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Factura.Factura, Factura.Renglon, Factura.Fecha, Factura.Cliente, Factura.Nombre, Factura.Direccion, Factura.Localidad, Factura.Cuit, Factura.Descripcion, Factura.Neto, Factura.Dto, Factura.Neto1, Factura.Iva1, Factura.Iva2, Factura.Total, Factura.Imprepago, Factura.CondIva, Factura.Item, Factura.Articulo, Factura.Cantidad, Factura.Precio, Factura.PordeDto, Factura.Postal " _
            + "From " _
            + DSQ + ".dbo.Factura Factura " _
            + "Where " _
            + "Factura.Item >= 0 AND " _
            + "Factura.Item <= 99"
    
    Listado.Connect = Connect()
    
    Uno = "{Factura.Item} in 0 to 99"
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    Listado.CopiesToPrinter = 2
    Rem Listado.Destination = 0
    
        
    Listado.ReportFileName = "Impreremito.rpt"
    
    Listado.Action = 1

End Sub

Private Sub Lee_Pedido()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    ControlPrecioI = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Cliente = " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " and Pedido.Fabrica > Pedido.Facturado"
    ZSql = ZSql + " Order by Pedido.Clave"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZZCantidad = rstPedido!fabrica - rstPedido!facturado
                    Call Redondeo(ZZZCantidad)
                
                    If ZZZCantidad <> 0 Then
                
                        Canti = ZZZCantidad
                        
                        Renglon = Renglon + 1
                        
                        If Renglon <= 99 Then
                                
                            WVector1.Row = Renglon
                                    
                            WVector1.Col = 1
                            WVector1.Text = Trim(rstPedido!Articulo)
                            Auxi1 = rstPedido!Articulo
                                
                            WVector1.Col = 3
                            WVector1.Text = Pusing("###,###.##", Str$(ZZZCantidad))
                                
                            Rem WVector1.Col = 4
                            Rem WVector1.Text = Pusing("###,###.##", Str$(rstPedido!Dto))
                                
                            Rem WVector1.Col = 5
                            Rem WVector1.Text = Pusing("###,###.##", Str$(rstPedido!Precio))
                                
                            Rem WVector1.Col = 6
                            Rem WVector1.Text = Pusing("###,###.##", Str$(rstPedido!Importe))
                            
                            WVector1.Col = 7
                            WVector1.Text = rstPedido!Numero
                            
                            WVector1.Col = 8
                            WVector1.Text = rstPedido!Clave
                            
                                Else
                                
                            m$ = "Hay mas item en condiciones de facturar"
                            aaaaaa% = MsgBox(m$, 0, "Emision de facturas")
                            Exit Do
                        
                        End If
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstPedido.Close
        
    End If
    
    For Ciclo = 1 To 99
    
        If Trim(WVector1.TextMatrix(Ciclo, 1)) <> "" Then
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.TextMatrix(Ciclo, 1) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.TextMatrix(Ciclo, 2) = rstArticulo!Descripcion
                ZZLinea = rstArticulo!Linea
                ZZTipo = rstArticulo!Tipo
                ZZFragancia = rstArticulo!Fragancia
                ZZCalidad = rstArticulo!Calidad
                ZZTamano = rstArticulo!Tamano
                rstArticulo.Close
            End If
            
            
            WWArti = Trim(WVector1.TextMatrix(Ciclo, 1))
            WWCanti = Val(WVector1.TextMatrix(Ciclo, 3))
            
            WWLinea = ZZLinea
            WWTipo = ZZTipo
            WWFragancia = ZZFragancia
            WWCalidad = ZZCalidad
            WWTamano = ZZTamano
            
            ControlPrecioII = 0
            Call Calcula_Costo
            If WPorceIva <> 0 Then
                WWPrecio = WWPrecio * 1.21
            End If
            
            If WWMoneda = 0 Then
                Rem WVector1.Col = 4
                Rem WVector1.Text = "$"
                WWPrecioII = WWPrecio / WWParidad
                    Else
                Rem WVector1.Col = 4
                Rem WVector1.Text = "U$S"
                WWPrecioII = WWPrecio
                WWPrecio = WWPrecio * WWParidad
                WWPrecioSB = WWPrecioSB * WWParidad
            End If
            
            Call Redondeo(WWPrecio)
            Call Redondeo(WWPrecioII)
            Call Redondeo(WWPrecioSB)
                
            WWImporte = WWPrecio * WWCanti
            WWImporteII = WWImporte / WWParidad
            
            WVector1.Row = Ciclo
            WVector1.CellBackColor = &HFFFFC0
            WVector1.Col = 9
            WVector1.Text = Str$(ControlPrecioII)
            
            
            If ControlPrecioII = 0 Then
                WVector1.Row = Ciclo
                WVector1.Col = 4
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWDto))
                WVector1.Col = 5
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecio))
                WVector1.Col = 6
                WVector1.CellBackColor = &HFFFFC0
                WVector1.Text = Pusing("###,###.##", Str$(WWImporte))
                WVector1.Col = 10
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecioSB))
                WVector1.Col = 11
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Str$(WWMoneda)
                    Else
                WVector1.Row = Ciclo
                WVector1.Col = 4
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWDto))
                WVector1.Col = 5
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecio))
                WVector1.Col = 6
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWImporte))
                WVector1.Col = 10
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Pusing("###,###.##", Str$(WWPrecioSB))
                WVector1.Col = 11
                WVector1.CellBackColor = &HFF00FF
                WVector1.Text = Str$(WWMoneda)
            End If
        End If
        
    Next Ciclo
    
    If ControlPrecioI <> 0 Then
        m$ = "Existen precios NO Activos"
        aaaaaa% = MsgBox(m$, 0, "Precios")
        Rem WWPrecio = 0
        Rem WWDto = 0
        Rem Exit Sub
    End If
    
    Call Calcula_Click

End Sub




Private Sub Calcula_Costo()


    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
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
        ControlPrecioI = 1
        ControlPrecioII = 1
        Rem m$ = "Precio No Activo"
        Rem aaaaaa% = MsgBox(m$, 0, "Precios")
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

    WWPrecioSB = WWPrecio
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
        WWPrecioSB = WWPrecio
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

End Sub




Private Sub Lee_CtaCte()

    WSaldo = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.Cliente = " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " and CtaCte.Saldo <> 0 "
    ZSql = ZSql + " Order by CtaCte.Cliente,CtaCte.OrdFecha,CtaCte.Impre,CtaCte.Numero"
        
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
    
        With rstCtaCte
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WSaldo = WSaldo + rstCtaCte!Saldo
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCte.Close
    End If
    
    
    txtUserName = "SA"
    txtPassword = "Sw58125812"
    txtOdbc = "FraganciasII"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.Cliente = " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " and CtaCte.Saldo <> 0 "
    ZSql = ZSql + " Order by CtaCte.Cliente,CtaCte.OrdFecha,CtaCte.Impre,CtaCte.Numero"
        
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
    
        With rstCtaCte
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WSaldo = WSaldo + rstCtaCte!Saldo
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCte.Close
    End If
                
                
    txtUserName = "SA"
    txtPassword = "Sw58125812"
    txtOdbc = "Fragancias"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    Saldo.Caption = Pusing("###,###,###.##", Str$(WSaldo))


End Sub



