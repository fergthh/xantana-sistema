VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactura 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturas"
   ClientHeight    =   9090
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
   ScaleHeight     =   9090
   ScaleWidth      =   13950
   Visible         =   0   'False
   Begin VB.TextBox VtoCae 
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
      Left            =   12360
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   64
      Text            =   " "
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox Cae 
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
      Left            =   10560
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   63
      Text            =   " "
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox Dolar 
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
      Left            =   11040
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   60
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
      Index           =   9
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton BorraRenglon 
      Caption         =   "Borra Renglon"
      Height          =   615
      Left            =   11160
      TabIndex        =   58
      Top             =   7560
      Width           =   1575
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
      Left            =   10440
      TabIndex        =   55
      Top             =   960
      Width           =   1815
   End
   Begin VB.CheckBox Contado 
      Caption         =   "Contado"
      Height          =   255
      Left            =   8760
      TabIndex        =   54
      Top             =   960
      Width           =   1575
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
      ItemData        =   "factura.frx":0000
      Left            =   2520
      List            =   "factura.frx":0007
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
      Left            =   9000
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
      Left            =   5640
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
      Left            =   9480
      MouseIcon       =   "factura.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":031F
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Elimina el Registro"
      Top             =   7560
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
      MouseIcon       =   "factura.frx":0B61
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":0E6B
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Pedidos de Clientes"
      Top             =   5640
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
      MouseIcon       =   "factura.frx":1735
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":1A3F
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Impresion"
      Top             =   6720
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
      MouseIcon       =   "factura.frx":2281
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":258B
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1440
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
      Left            =   12960
      MouseIcon       =   "factura.frx":2DCD
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":30D7
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Elimina el Registro"
      Top             =   2520
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
      MouseIcon       =   "factura.frx":3919
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":3C23
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3480
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
      MouseIcon       =   "factura.frx":4465
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":476F
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Consulta de Datos"
      Top             =   4560
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
      MouseIcon       =   "factura.frx":4FB1
      MousePointer    =   99  'Custom
      Picture         =   "factura.frx":52BB
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Menu Principal"
      Top             =   7800
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
         Left            =   4680
         TabIndex        =   28
         Top             =   240
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
      Left            =   8160
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
      Left            =   7440
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
      Left            =   6960
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
      Left            =   12480
      TabIndex        =   67
      Top             =   600
      Width           =   1335
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
      Left            =   12480
      TabIndex        =   66
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "CAE"
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
      Left            =   9840
      TabIndex        =   65
      Top             =   120
      Width           =   855
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
      TabIndex        =   62
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
      Left            =   10320
      TabIndex        =   61
      Top             =   480
      Width           =   1455
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
      Left            =   8280
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
      Left            =   4680
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
      Left            =   6120
      TabIndex        =   20
      Top             =   120
      Width           =   1215
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
Attribute VB_Name = "PrgFactura"
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
Private WImpoDto1 As Double
Private WImpoDto2 As Double
Private WImpoDto3 As Double
Private WDescuento As Double
Private WCodIva As String
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
Dim ZZPasaDatos(100, 10) As String

Dim ControlPrecioI As Integer
Dim ControlPrecioII As Integer
Dim PrecioRedondeo As Double
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

Dim ZZCantidad As String
Dim ZZCantidadII As String
Dim WParcial As Double

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
Dim ZZImpreNumero As String
Dim ZZImpreTipo As String
Dim ZZImpreComprobante As String




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
Dim WFecha As String











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

Private Sub Anula_Click()

    T$ = "Anulacion de Comprobantes"
    m$ = "Desea Anular el Comprobante "
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then

        T$ = "Anulacion de Comprobantes"
        m$ = "Esta Seguro que Desea Anular el Comprobante "
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
        
                    
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "01"
                
            Claveven$ = Letra.Text + WTipo + WPunto + Auxi + "01"
               
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
                
                If ZSaldo <> ZTotal And Contado.Value = 0 Then
                
                    m$ = "El comprobante se encuentra total o parcialmente cancelado"
                    aaaaaa% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
                    
                    Exit Sub
                    
                End If
                
            End If
            
            Erase WVector
            Erase ZZVector
        
            For WRenglon = 1 To 50
            
                ZZZLetra = Letra.Text
                
                ZZZPunto = Punto.Text
                Call Ceros(ZZZPunto, 1)
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "01" + ZZZLetra + ZZZPunto + Auxi + Auxi1
                    
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
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
            
            If Letra.Text = "A" Then
                ZSql = ""
                ZSql = ZSql + "DELETE MovimientoInsumo"
                ZSql = ZSql + " Where MovimientoInsumo.Tipo = " + "'" + "4" + "'"
                ZSql = ZSql + " and MovimientoInsumo.Numero = " + "'" + Numero.Text + "'"
                spMovimientoInsumo = ZSql
                Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                ZSql = ""
                ZSql = ZSql + "DELETE MovimientoInsumo"
                ZSql = ZSql + " Where MovimientoInsumo.Tipo = " + "'" + "5" + "'"
                ZSql = ZSql + " and MovimientoInsumo.Numero = " + "'" + Numero.Text + "'"
                spMovimientoInsumo = ZSql
                Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            
            
            
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
            
            Erase WVector
        
            For WRenglon = 1 To 50
            
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
        ZZDireccionII = rstCliente!DireccionII
        ZZLocalidadII = IIf(IsNull(rstCliente!LocalidadII), "", rstCliente!LocalidadII)
        ZZProvinciaII = IIf(IsNull(rstCliente!ProvinciaII), "0", rstCliente!ProvinciaII)
        ZZPostalII = IIf(IsNull(rstCliente!PostalII), "", rstCliente!PostalII)
        rstCliente.Close
    End If
    
    ZZCuitII = ""
    
    Call Calcula_Barra
    
    
    ZZLetra = "X"
    ZZTipo = "01"
    ZZPunto = "0001"
    Auxi1 = Numero.Text
    Call Ceros(Auxi1, 8)
    ZZFactura = Auxi1
    ZZImpreNumero = ZZFactura
    ZZfecha = Fecha.Text
    ZZCliente = Cliente.Text
    ZZNombre = Trim(ZZRazon)
    ZZDireccion = Trim(ZZDireccionII) + " - " + Trim(ZZLocalidadII) + " - " + Trim(Provincia(ZZProvinciaII))
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
    ZZPostal = WPostalII
    
    ZZLugarFactura = 0
    
    Call Numtolet
    
    For A = 1 To 30
    
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
            ZZImpre4 = ""
            
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
            ZSql = ZSql + "Impre2 ,"
            ZSql = ZSql + "Impre3 ,"
            ZSql = ZSql + "Impre4 ,"
            ZSql = ZSql + "Cae ,"
            ZSql = ZSql + "VtoCae ,"
            ZSql = ZSql + "ImpreBarra ,"
            ZSql = ZSql + "ImpreBarraII ,"
            ZSql = ZSql + "ImpreTipo ,"
            ZSql = ZSql + "ImpreComprobante ,"
            ZSql = ZSql + "ImpreNumero ,"
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
            ZSql = ZSql + "'" + ZZImpre2 + "',"
            ZSql = ZSql + "'" + ZZImpre3 + "',"
            ZSql = ZSql + "'" + ZZImpreComprobante + "',"
            ZSql = ZSql + "'" + Cae.Text + "',"
            ZSql = ZSql + "'" + VtoCae.Text + "',"
            ZSql = ZSql + "'" + ZZImpreBarra + "',"
            ZSql = ZSql + "'" + ZZImpreBarraII + "',"
            ZSql = ZSql + "'" + ZZImpreTipo + "',"
            ZSql = ZSql + "'" + ZZImpreComprobante + "',"
            ZSql = ZSql + "'" + ZZImpreNumero + "',"
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
        ZZImpre4 = ""
        
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
        ZSql = ZSql + "Impre2 ,"
        ZSql = ZSql + "Impre3 ,"
        ZSql = ZSql + "Impre4 ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "VtoCae ,"
        ZSql = ZSql + "ImpreBarra ,"
        ZSql = ZSql + "ImpreBarraII ,"
        ZSql = ZSql + "ImpreTipo ,"
        ZSql = ZSql + "ImpreComprobante ,"
        ZSql = ZSql + "ImpreNumero ,"
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
        ZSql = ZSql + "'" + ZZImpre2 + "',"
        ZSql = ZSql + "'" + ZZImpre3 + "',"
        ZSql = ZSql + "'" + ZZImpreComprobante + "',"
        ZSql = ZSql + "'" + Cae.Text + "',"
        ZSql = ZSql + "'" + VtoCae.Text + "',"
        ZSql = ZSql + "'" + ZZImpreBarra + "',"
        ZSql = ZSql + "'" + ZZImpreBarraII + "',"
        ZSql = ZSql + "'" + ZZImpreTipo + "',"
        ZSql = ZSql + "'" + ZZImpreComprobante + "',"
        ZSql = ZSql + "'" + ZZImpreNumero + "',"
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
        
    Listado.SQLQuery = "SELECT Factura.Factura, Factura.Renglon, Factura.Fecha, Factura.Cliente, Factura.Nombre, Factura.Direccion, Factura.Cuit, Factura.Remito, Factura.Descripcion, Factura.Neto, Factura.Neto1, Factura.Iva1, Factura.Iva2, Factura.Total, Factura.Imprepago, Factura.Impre1, Factura.Impre2, Factura.CondIva, Factura.Item, Factura.Articulo, Factura.Cantidad, Factura.Precio, Factura.Impre3, Factura.Impre4, Factura.PorceIva, Factura.Postal, Factura.Cae, Factura.VtoCae, Factura.ImpreBarra, Factura.ImpreBarraII, Factura.ImpreTipo, Factura.ImpreNumero " _
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
    
    If Letra.Text = "A" Then
        Listado.ReportFileName = "facturaFEA.rpt"
            Else
        Listado.ReportFileName = "facturaFEb.rpt"
    End If
    
    Listado.Action = 1

End Sub



Private Sub Impresion_Caratula()
    
    Listado.WindowTitle = "Impresion de Caratula"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Cliente.Cliente, Cliente.Razon, Cliente.Direccion, Cliente.Localidad, Cliente.Provincia, Cliente.Expreso, Cliente.Fantasia, " _
            + "Provincia.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Cliente Cliente, " _
            + DSQ + ".dbo.Provincia Provincia " _
            + "Where " _
            + "Cliente.Provincia = Provincia.Codigo AND " _
            + "Cliente.Cliente >= '" + Cliente.Text + "' AND " _
            + "Cliente.Cliente <= '" + Cliente.Text + "'"
    
    Listado.Connect = Connect()
    
    Uno = "{Cliente.Cliente} in " + Chr$(34) + Cliente.Text + Chr$(34) + " to " + Chr$(34) + Cliente.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    Listado.CopiesToPrinter = 1
    Rem Listado.Destination = 0
    
    Listado.ReportFileName = "caratula.rpt"
    
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
        
        For CicloII = 1 To 50
            
            ZZArticulo = WVector1.TextMatrix(CicloII, 1)
            ZZDesArticulo = WVector1.TextMatrix(CicloII, 2)
            ZZCantidad = WVector1.TextMatrix(CicloII, 3)
            ZZPrecio = WVector1.TextMatrix(CicloII, 4)
            ZZImporte = WVector1.TextMatrix(CicloII, 5)
            ZZStock = WVector1.TextMatrix(CicloII, 6)
            ZZClavePedido = WVector1.TextMatrix(CicloII, 8)
            
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
                
            End If
        
        Next CicloII

    
        ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
        
        For WRenglon = 1 To 50
                
            Articulo = ZZPasaDatos(WRenglon, 1)
            Cantidad = Val(ZZPasaDatos(WRenglon, 3))
            ClavePedido = ZZPasaDatos(WRenglon, 8)
                
            If Cantidad <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Pedido SET "
                ZSql = ZSql + " Entregado = Entregado + " + "'" + Str$(Cantidad) + "',"
                ZSql = ZSql + " Marca = " + "'" + "" + "'"
                ZSql = ZSql + " Where ClavePedido = " + "'" + ClavePedido + "'"
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
            
        If ZZNivel = 1 Then
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
    
        If ZZNivel = 1 Then
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
    
    If Val(WEmpresa) = 1 Then
    
        T$ = "Emision de Facturas"
        m$ = "Desea reimprimir el remito"
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
            ZZPasaImpre = 1
            Call Impresion_remito
        End If
        
        T$ = "Emision de Facturas"
        m$ = "Desea reimprimir la factura"
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
            ZZPasaImpre = 1
            Call Impresion_FacturaFe
        End If
        
        T$ = "Emision de Facturas"
        m$ = "Desea reimprimir la caratula de Transporte"
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
            ZZPasaImpre = 1
            Call Impresion_Caratula
        End If
    
            Else
        
        T$ = "Emision de Facturas"
        m$ = "Desea reimprimir la factura"
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
            ZZPasaImpre = 1
            Call Impresion_FacturaRemito
        End If
        
    End If
            

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
    
    For A = 1 To 50
    
        WCantidad = Val(WVector1.TextMatrix(A, 3))
        WPrecio = Val(WVector1.TextMatrix(A, 5))
        
        Rem If Letra.Text = "B" Then
        Rem     WWImpre = WPrecio * (1 + (ConfigIva1) / 100)
        Rem     Call Redondeo(WWImpre)
        Rem     WPrecio = WWImpre
        Rem End If
        
        WParcial = (WPrecio * WCantidad)
        Call Redondeo(WParcial)
        WNeto = WNeto + WParcial
        
    Next A
    
    Call Calcula_Importe
    
End Sub

Private Sub Calcula_Importe()

    WIva1 = 0
    WIva2 = 0
    
    If Letra.Text = "A" Then
        Select Case Val(WCodIva)
            Case 2
                WIva1 = WNeto * ((ConfigIva1) / 100)
                WIva2 = WNeto * ((ConfigIva2) / 100)
                Call Redondeo(WIva1)
                Call Redondeo(WIva2)
            Case Else
                WIva1 = WNeto * ((ConfigIva1) / 100)
                Call Redondeo(WIva1)
        End Select
    End If
    
    WWIva(1) = WIva1
    WWIva(2) = WIva2
    
    WTotal = WNeto + WIva1 + WIva2
    
    SubTotal.Caption = Str$(WNeto + WImpoDto)
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
    PrgFactura.Hide
    Unload Me
    If ZZPasaProcesoFactura = 0 Then
        MenuVen.Show
            Else
        PrgControlPedido.Show
    End If
End Sub

Private Sub Graba_Click()



    T$ = "Grabacion de Facturas"
    m$ = "Desea emitir la factura "
    Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
    
        ZZLineas = 0
        Erase ZZPasaDatos
        
        For CicloII = 1 To 50
            
            ZZArticulo = WVector1.TextMatrix(CicloII, 1)
            ZZDesArticulo = WVector1.TextMatrix(CicloII, 2)
            ZZCantidad = WVector1.TextMatrix(CicloII, 3)
            ZZPrecio = WVector1.TextMatrix(CicloII, 5)
            ZZImporte = WVector1.TextMatrix(CicloII, 6)
            ZZPedido = WVector1.TextMatrix(CicloII, 7)
            ZZClavePedido = WVector1.TextMatrix(CicloII, 8)
            
            If Val(ZZCantidad) <> 0 Then
                
                    
                ZZLineas = ZZLineas + 1
                ZZPasaDatos(ZZLineas, 1) = ZZArticulo
                ZZPasaDatos(ZZLineas, 2) = ZZDesArticulo
                ZZPasaDatos(ZZLineas, 3) = ZZCantidad
                ZZPasaDatos(ZZLineas, 4) = ZZPrecio
                ZZPasaDatos(ZZLineas, 5) = ZZImporte
                ZZPasaDatos(ZZLineas, 7) = ZZPedido
                ZZPasaDatos(ZZLineas, 8) = ZZClavePedido
                
            End If
        
        Next CicloII
    
        If ZZLineas > 30 Then
            m$ = "La factura a emitor supera los 30 renglones"
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
            WPlazo1 = Str$(rstCondPago!Dias)
            rstCondPago.Close
        End If
        
        WFecha = Fecha.Text
        Call Calcula_vencimiento(WFecha, WPlazo1, WVencimiento)
        
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
        
        If ZZNivel = 1 Then
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
            
            If ZZNivel = 1 Then
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
            
        If Trim(Cae.Text) <> "" Then
            Exit Sub
        End If
        
        If Trim(Cae.Text) = "" Then
            Call Calcula_Cae
            If Trim(Cae.Text) = "" Then
                Exit Sub
            End If
        End If
            
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
        ZZCae = Cae.Text
        ZZVtoCae = VtoCae.Text
        
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
            
        For A = 1 To 50
            
            WRenglon = WRenglon + 1
            
            Articulo = UCase(ZZPasaDatos(WRenglon, 1))
            DesArticulo = ZZPasaDatos(WRenglon, 2)
            Cantidad = Val(ZZPasaDatos(WRenglon, 3))
            Precio = Val(ZZPasaDatos(WRenglon, 4))
            Preciosalva = Val(ZZPasaDatos(WRenglon, 4))
            ZZZPedido = ZZPasaDatos(WRenglon, 7)
            ZZZClavePedido = ZZPasaDatos(WRenglon, 8)
                
            If Cantidad <> 0 Then
                        
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                            
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
                
                
                Auxi2 = Str$(Cantidad)
                Call Ceros(Auxi2, 8)
    
                ZZTipo = "01"
                ZZNumero = Numero.Text
                ZZRenglon = Renglon
                ZZArticulo = Articulo
                ZZDescripcion = DesArticulo
                ZZCantidad = Auxi2
                ZZCantidadII = Auxi2
                ZZPrecio = Str$(Precio)
                ZZPrecioSalva = Str$(Preciosalva)
                ZZPrecioUs = Str$(XXPrecio)
                ZZImporte = Str$(Precio * Cantidad)
                ZZImporteUs = Str$(XXPrecio * Cantidad)
                ZZCliente = Cliente.Text
                ZZParidad = "0"
                ZZVendedor = "0"
                ZZRubro = "0"
                ZZLinea = "0"
                ZZCosto2 = "0"
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
                ZZDescuento = ""
                ZZPartida = ""
                
                ZZCantidadII = ZZCantidad
                
                ZZPrecioII = Str$(XXPrecio)
                
                ZZCosto1 = ""
                ZZCosto2 = ""
                ZZTipoComision = ""
                ZZMarca = ""
                ZZTipoII = ""
                
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
        
        If ZZNivel = 1 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        
        ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
        
        For WRenglon = 1 To 50
                
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
                                                    
                    If Trim(WInsumoII) = "" Then
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
                        
                        If Letra.Text = "A" Then
                            ZZTipoMov = "04"
                                Else
                            ZZTipoMov = "05"
                        End If
                                
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
        m$ = "Desea reimprimir el remito"
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
            ZZPasaImpre = 1
            Call Impresion_remito
        End If
        
        T$ = "Emision de Facturas"
        m$ = "Desea reimprimir la factura"
        Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
            ZZPasaImpre = 1
            Call Impresion_FacturaFe
        End If
    
        Rem T$ = "Emision de Facturas"
        Rem m$ = "Desea Imprimir la Factura"
        Rem Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
        Rem If Respuestaaaaaa% = 6 Then
        Rem     Call WImpresion
        Rem End If
        
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Fcaturas")
        
            
        Call Limpia_Click
        
        If ZZPasaProcesoFactura = 0 Then
            Cliente.SetFocus
                Else
            Call CmdClose_Click
        End If
        
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
                    Exit Sub
                    
                End If
            End If
        
            Erase WVector
            Erase ZZVector
        
            For WRenglon = 1 To 50
            
                ZZZLetra = Letra.Text
                
                ZZZPunto = Punto.Text
                Call Ceros(ZZZPunto, 1)
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 6)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "01" + ZZZLetra + ZZZPunto + Auxi + Auxi1
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
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
            
            If Letra.Text = "A" Then
                ZSql = ""
                ZSql = ZSql + "DELETE MovimientoInsumo"
                ZSql = ZSql + " Where MovimientoInsumo.Tipo = " + "'" + "4" + "'"
                ZSql = ZSql + " and MovimientoInsumo.Numero = " + "'" + Numero.Text + "'"
                spMovimientoInsumo = ZSql
                Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                ZSql = ""
                ZSql = ZSql + "DELETE MovimientoInsumo"
                ZSql = ZSql + " Where MovimientoInsumo.Tipo = " + "'" + "5" + "'"
                ZSql = ZSql + " and MovimientoInsumo.Numero = " + "'" + Numero.Text + "'"
                spMovimientoInsumo = ZSql
                Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql + "DELETE CtaCte"
            ZSql = ZSql + " Where Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            For WRenglon = 1 To 50
            
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
        Rem ConfigPunto = rstConfiguracion!Punto
        ConfigPunto = 9
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
    
    If ZZNivel = 1 Then
        txtUserName = "SA"
        txtPassword = "Sw58125812"
        txtOdbc = "FraganciasII"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If
    
    For WRenglon = 1 To 50
        
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
            WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Cantidad))
                
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###", Str$(rstEstadistica!Descuento))
            
            PrecioRedondeo = rstEstadistica!Precio
            Call Redondeo(PrecioRedondeo)
                
            WVector1.Col = 5
            WVector1.Text = Pusing("###,###.##", Str$(PrecioRedondeo))
            
            WVector1.Col = 6
            WVector1.Text = Pusing("###,###.##", Str$(PrecioRedondeo * rstEstadistica!CantidadII))
            
            ZZPedido = IIf(IsNull(rstEstadistica!Pedido), "", rstEstadistica!Pedido)
            ZZClavePedido = IIf(IsNull(rstEstadistica!ClavePedido), "", rstEstadistica!ClavePedido)
            
            WVector1.Col = 7
            WVector1.Text = ZZPedido
            
            WVector1.Col = 8
            WVector1.Text = ZZClavePedido
            
            rstEstadistica.Close
                
        End If
    
    Next WRenglon
    
    If ZZNivel = 1 Then
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
            If ZZNivel = 1 Then
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
    
            If ZZNivel = 1 Then
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
                                If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 3 Then
                                Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 3 Then
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
            
            If ZZNivel = 0 Then
                
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
    
            If ZZNivel = 1 Then
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
            
            ZZImpreHistorial = ""
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM HistorialCliente"
            ZSql = ZSql + " Where HistorialCliente.Cliente = " + "'" + Cliente.Text + "'"
            spHistorialCliente = ZSql
            Set rstHistorialCliente = db.OpenRecordset(spHistorialCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstHistorialCliente.RecordCount > 0 Then
                rstHistorialCliente.Close
                ZZPasaCliente = Cliente.Text
                ZZPasaProceso = 4
                PrgHistorialClienteConsulta.Show
                ZZImpreHistorial = "S"
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
    
        If ZZNivel = 1 Then
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
                            If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                            Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 3 Then
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
    
        If ZZNivel = 1 Then
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
    
        If ZZNivel = 1 Then
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
            Dolar.Text = Str$(rstCtaCte!Paridad)
            Cae.Text = IIf(IsNull(rstCtaCte!Cae), "", rstCtaCte!Cae)
            VtoCae.Text = IIf(IsNull(rstCtaCte!VtoCae), "", rstCtaCte!VtoCae)

            rstCtaCte.Close
    
            If ZZNivel = 1 Then
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
                WRazon = rstCliente!Fantasia
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                rstCliente.Close
            End If
            
            Call Lee_CtaCte
            
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
    
            If ZZNivel = 1 Then
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
                WFormato(Ciclo) = "###,###.##"
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
    
    For Ciclo = 1 To 50
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
    
    For A = 1 To 30
    
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
            ZZImpre4 = ""
            
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
        ZZImpre4 = ""
        
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
    
    Listado.CopiesToPrinter = 3
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
        
    Listado.ReportFileName = "remito.rpt"
    
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
                            WVector1.Text = Str$(rstPedido!Numero)
                            
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
    
    For Ciclo = 1 To 50
    
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
            If Letra.Text <> "A" Then
                WWPrecio = WWPrecio * 1.21
            End If
            Call Redondeo(WWPrecio)
            
            If WWMoneda = 0 Then
                Rem WVector1.Col = 4
                Rem WVector1.Text = "$"
                WWPrecioII = WWPrecio / WWParidad
                    Else
                Rem WVector1.Col = 4
                Rem WVector1.Text = "U$S"
                WWPrecioII = WWPrecio
                WWPrecio = WWPrecio * WWParidad
            End If
            
            Call Redondeo(WWPrecio)
            Call Redondeo(WWPrecioII)
                
            
                
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
        WWPrecio = WWPrecio - WImpoDto
        Call Redondeo(WWPrecio)
    End If

End Sub



Private Sub Impresion_FacturaRemito()


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
            Rem ZZZMoneda = Val(WVector1.TextMatrix(a, 9))
            
            ZZCantidad = Str$(ZZZCantidad)
            ZZPrecio = Str$(ZZZPrecio)
            ZZImporte = Str$(ZZZImporte)
            ZZDto = Str$(ZZZDto)
            ZZPorceIva = Str$(WWParidad)
            ZZDias = 0
            Rem ZZDias = ZZZMoneda
            
            If Trim(ZZArticulo) = "" Then
                ZZItem = ""
                ZZArticulo = ""
                ZZDescripcion = ""
                ZZCantidad = ""
                ZZPrecio = ""
                ZZImporte = ""
                ZZDto = ""
                ZZPorceIva = ""
                ZZDias = ""
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
            ZSql = ZSql + "'" + Str$(ZZDias) + "',"
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



Private Sub Calcula_Cae()
    
    Dim WSAA As Object, WSFEv1 As Object
    
    On Error GoTo ManejoError
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    ' Crear objeto interface Web Service Autenticaci?n y Autorizaci?n
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    'Debug.Print WSAA.InstallDir
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
    tra = WSAA.CreateTRA("wsfe")
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
        
    ZPath = "c:\salva\"
    ZNombre = "Mc"
    ZCuit = "30708403020"
    punto_vta = 9
    
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Rem Certificado = "..\..\reingart.crt" ' certificado de prueba
    Rem ClavePrivada = "..\..\reingart.key" ' clave privada de prueba
    
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    ' Llamar al web service para autenticar:
    proxy = "" '"usuario:clave@localhost:8000"
    Rem ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms", proxy) ' Homologaci?n
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms", proxy) ' Homologaci?n

    ' Imprimir el ticket de acceso, ToKen y Sign de autorizaci?n
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 24 horas
    ' (este per?odo se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electr?nica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    'Debug.Print WSFEv1.InstallDir
    
    ' Setear tocken y sing de autorizaci?n (pasos previos)
    WSFEv1.Token = WSAA.Token
    WSFEv1.Sign = WSAA.Sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = ZCuit
    
    ' Conectar al Servicio Web de Facturaci?n
    proxy = "" ' "usuario:clave@localhost:8000"
    wsdl = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
    cache = ""    'Rem Path
        
    oK = WSFEv1.Conectar(cache, wsdl, proxy, "") ' homologaci?n
    Debug.Print WSFEv1.Version
    
    ' mostrar bit?cora de depuraci?n:
    Debug.Print WSFEv1.DebugLog
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus
       
    ' Establezco los valores de la factura a autorizar:
    If Letra.Text = "A" Then
        tipo_cbte = 1
            Else
        tipo_cbte = 6
    End If
    
    cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
    If cbte_nro = "" Then
        cbte_nro = 0                ' no hay comprobantes emitidos
            Else
        cbte_nro = CLng(cbte_nro)   ' convertir a entero largo
    End If
    
    If cbte_nro + 1 <> Val(Numero.Text) Then
        m$ = "El numero del comprobante no coincide con el informado por la afip (" + Str$(cbte_nro) + ")"
        A% = MsgBox(m$, 0, "Ingreso de Facturas")
        Exit Sub
    End If
    
    Rem dada
    Rem dada
    Rem dada
    Rem dada

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WCuit = IIf(IsNull(rstCliente!Cuit), "", rstCliente!Cuit)
        Call Eval
        rstCliente.Close
    End If
    
    Rem Fecha = Format(Date, "yyyymmdd")
    
    Rem CONCEPTO   1-PRODUCTO    2-SERVICIOS     3-PRODUCTOS Y SERVICIOS
    Concepto = 1
    
    Rem TIPO DE DOCUMENTO
    If Len(WCuit) = 11 Then
        tipo_doc = 80
            Else
        tipo_doc = 96
    End If
    
    Rem NUMERO DE DOCUMENTO
    nro_doc = Left$(WCuit + Space$(11), 11)
    
    Rem NUMERO DE DOCUMENTO
    cbte_nro = cbte_nro + 1
    cbt_desde = cbte_nro
    cbt_hasta = cbte_nro
    
    Rem IMPORTE TOTAL
    imp_total = Val(Total.Caption)
    
    Rem IMPORTE DE CONCEPTOS NO GRAVADOS POR EL IVA
    imp_tot_conc = 0
    
    Rem IMPORTE NETO
    imp_neto = Val(Neto.Caption)
    
    Rem IMPORTE IVA
    imp_iva = Val(Iva1.Caption)
    
    Rem suma de importes de otros impuestos
    imp_trib = 0
    
    Rem IMPORTE EXENTO DE IVA
    imp_op_ex = 0
    
    Rem FECHA
    ZZfecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    fecha_cbte = ZZfecha
    
    Rem VENCIMIENTO
    fecha_venc_pago = ""
    
    Rem FECHAS DE SERVICIOS PARA SERVICIOS
    ' Fechas del per?odo del servicio facturado (solo si concepto = 1?)
    fecha_serv_desde = ""
    fecha_serv_hasta = ""
    
    Rem MONEDA
    moneda_id = "PES"
    
    Rem COTIZACION
    moneda_ctz = 1

    oK = WSFEv1.CrearFactura(Concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz)
    
    ' Agrego los comprobantes asociados:
    Rem If False Then ' solo nc/nd
    Rem     tipo = 19
    Rem     pto_vta = 2
    Rem     nro = 1234
    Rem     ok = WSFEv1.AgregarCmpAsoc(tipo, pto_vta, nro)
    Rem End If
        
    ' Agrego impuestos varios
    Rem id = 99
    Rem Desc = "Impuesto Municipal Matanza'"
    Rem base_imp = "100.00"
    Rem alic = "1.00"
    Rem importe = "1.00"
    Rem ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)

    ' Agrego tasas de IVA
    If Val(Iva1.Caption) = 0 Then
        id = 3 ' 0%
        base_imp = Val(Neto.Caption)
        Importe = Val(Iva1.Caption)
        oK = WSFEv1.AgregarIva(id, base_imp, Importe)
            Else
        id = 5 ' 21%
        base_imp = Val(Neto.Caption)
        Importe = Val(Iva1.Caption)
        oK = WSFEv1.AgregarIva(id, base_imp, Importe)
    End If
    
    
    
    ' Habilito reprocesamiento autom?tico (predeterminado):
    WSFEv1.Reprocesar = True

    ' Solicito CAE:
    Cae = WSFEv1.CAESolicitar()
    
    Debug.Print "Resultado", WSFEv1.resultado
    Debug.Print "CAE", WSFEv1.Cae

    Debug.Print "Numero de comprobante:", WSFEv1.CbteNro
    
    ' Imprimo pedido y respuesta XML para depuraci?n (errores de formato)
    Debug.Print WSFEv1.XmlRequest
    Debug.Print WSFEv1.XmlResponse
    
    Debug.Print "Reprocesar:", WSFEv1.Reprocesar
    Debug.Print "Reproceso:", WSFEv1.Reproceso
    Debug.Print "CAE:", WSFEv1.Cae
    Debug.Print "EmisionTipo:", WSFEv1.EmisionTipo

    MsgBox "Resultado:" & WSFEv1.resultado & " CAE: " & Cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
    ' Muestro los errores
    If WSFEv1.ErrMsg <> "" Then
        MsgBox WSFEv1.ErrMsg, vbExclamation, "Error"
    End If
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEv1.eventos:
        MsgBox evento, vbInformation, "Evento"
    Next
    
    ' Buscar la factura
    cae2 = WSFEv1.CompConsultar(tipo_cbte, punto_vta, cbte_nro)
    
    Debug.Print "Fecha Comprobante:", WSFEv1.FechaCbte
    Debug.Print "Fecha Vencimiento CAE", WSFEv1.Vencimiento
    Debug.Print "Importe Total:", WSFEv1.ImpTotal
    Debug.Print "Resultado:", WSFEv1.resultado
    
    If Cae <> cae2 Then
        MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!: " & Cae & " vs " & cae2
    Else
        MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
    End If
        
        
    If WSFEv1.resultado = "A" Then
        Cae.Text = Cae
        If Len(Trim(WSFEv1.Vencimiento)) = 8 Then
            VtoCae.Text = Right$(WSFEv1.Vencimiento, 2) + "/" + Mid$(WSFEv1.Vencimiento, 5, 2) + "/" + Left$(WSFEv1.Vencimiento, 4)
                Else
            VtoCae.Text = WSFEv1.Vencimiento
        End If
    End If

    Exit Sub
ManejoError:
    ' Si hubo error:
    Rem Debug.Print WSFEv1.Excepcion
    Debug.Print Err.Description            ' descripci?n error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Print WSFEv1.XmlRequest
            Debug.Print WSFEv1.XmlResponse
            Debug.Print WSFEv1.traceback
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    Debug.Print WSFEv1.XmlRequest
    Debug.Assert False
    Debug.Print WSFEv1.traceback
End Sub

Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub

Private Sub Calcula_Barra()
    
    Dim ZZCara(1000) As String
    
    ZZNumero = "30708403020"
    
    If Letra.Text = "A" Then
        ZZNumero = ZZNumero + "01"
        ZZImpreTipo = "01"
        ZZImpreComprobante = "FACTURA"
            Else
        ZZNumero = ZZNumero + "06"
        ZZImpreTipo = "06"
        ZZImpreComprobante = "FACTURA"
    End If
            
    Auxi1 = Punto.Text
    Call Ceros(Auxi1, 4)
    ZZNumero = ZZNumero + Auxi1
    
    ZZNumero = ZZNumero + Trim(Cae.Text)
    
    ZZFechaCae = VtoCae.Text
    ZZOrdFechaCae = Right$(ZZFechaCae, 4) + Mid$(ZZFechaCae, 4, 2) + Left$(ZZFechaCae, 2)
    ZZNumero = ZZNumero + ZZOrdFechaCae
    
    ZZCara(0) = "!"
    ZZCara(1) = Chr$(34)
    ZZCara(2) = "#"
    ZZCara(3) = "$"
    ZZCara(4) = "%"
    ZZCara(5) = "&"
    ZZCara(6) = "?"
    ZZCara(7) = "("
    ZZCara(8) = ")"
    ZZCara(9) = "*"
    ZZCara(10) = "+"
    ZZCara(11) = ","
    ZZCara(12) = "-"
    ZZCara(13) = "."
    ZZCara(14) = "/"
    ZZCara(15) = "0"
    ZZCara(16) = "1"
    ZZCara(17) = "2"
    ZZCara(18) = "3"
    ZZCara(19) = "4"
    ZZCara(20) = "5"
    ZZCara(21) = "6"
    ZZCara(22) = "7"
    ZZCara(23) = "8"
    ZZCara(24) = "9"
    ZZCara(25) = ":"
    ZZCara(26) = ";"
    ZZCara(27) = "<"
    ZZCara(28) = "="
    ZZCara(29) = ">"
    ZZCara(30) = "?"
    ZZCara(31) = "@"
    ZZCara(32) = "A"
    ZZCara(33) = "B"
    ZZCara(34) = "C"
    ZZCara(35) = "D"
    ZZCara(36) = "E"
    ZZCara(37) = "F"
    ZZCara(38) = "G"
    ZZCara(39) = "H"
    ZZCara(40) = "I"
    ZZCara(41) = "J"
    ZZCara(42) = "K"
    ZZCara(43) = "L"
    ZZCara(44) = "M"
    ZZCara(45) = "N"
    ZZCara(46) = "O"
    ZZCara(47) = "P"
    ZZCara(48) = "Q"
    ZZCara(49) = "R"
    ZZCara(50) = "S"
    ZZCara(51) = "T"
    ZZCara(52) = "U"
    ZZCara(53) = "V"
    ZZCara(54) = "W"
    ZZCara(55) = "X"
    ZZCara(56) = "Y"
    ZZCara(57) = "Z"
    ZZCara(58) = "["
    ZZCara(59) = "\"
    ZZCara(60) = "]"
    ZZCara(61) = "^"
    ZZCara(62) = "_"
    ZZCara(63) = "`"
    ZZCara(64) = "a"
    ZZCara(65) = "b"
    ZZCara(66) = "c"
    ZZCara(67) = "d"
    ZZCara(68) = "e"
    ZZCara(69) = "f"
    ZZCara(70) = "g"
    ZZCara(71) = "h"
    ZZCara(72) = "i"
    ZZCara(73) = "j"
    ZZCara(74) = "k"
    ZZCara(75) = "l"
    ZZCara(76) = "m"
    ZZCara(77) = "n"
    ZZCara(78) = "o"
    ZZCara(79) = "p"
    ZZCara(80) = "q"
    ZZCara(81) = "r"
    ZZCara(82) = "s"
    ZZCara(83) = "t"
    ZZCara(84) = "u"
    ZZCara(85) = "v"
    ZZCara(86) = "w"
    ZZCara(87) = "x"
    ZZCara(88) = "y"
    ZZCara(89) = "z"
    ZZCara(90) = "�"
    ZZCara(91) = "�"
    ZZCara(92) = "�"
    ZZCara(93) = "�"
    ZZCara(94) = "�"
    ZZCara(95) = "�"
    ZZCara(96) = "�"
    ZZCara(97) = "�"
    ZZCara(98) = "�"
    ZZCara(99) = "�"
    
    Rem ZZNumero = "3070306062119000260321213344273201008198"
    Rem ZZNumero = "000102030405060708091011121314151617181920"
    Rem ZZNumero = "2122232425262728293031323334353637383940"
    Rem ZZNumero = "4142434445464748495051525354555657585960"
    Rem ZZNumero = "6162636465666768697071727374757677787980"
    Rem ZZNumero = "81828384858687888990919293949596979899"
    Rem ZZNumero = "307030606211900026032121334427320100819"
    
    ZZSumaI = 0
    ZZSumaII = 0
    
    For Ciclo = 1 To 39 Step 2
        ZZSumaI = ZZSumaI + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    ZZSumaI = ZZSumaI * 3
    
    For Ciclo = 2 To 39 Step 2
        ZZSumaII = ZZSumaII + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    
    ZZSuma = ZZSumaI + ZZSumaII
    ZZVerifica = ZZSuma
    ZZDigi = 0
    
    Do
    
        ZZVerifi = Int(ZZVerifica / 10) * 10
        
        If ZZVerifi = ZZVerifica Then
            Exit Do
        End If
        
        ZZDigi = ZZDigi + 1
        
        ZZVerifica = ZZSuma + ZZDigi
        
    Loop
    
    ZZNumero = ZZNumero + Trim(Str$(ZZDigi))
    
    lccar = ""
    barralargo = ZZNumero
    
    For lni = 1 To Len(barralargo) Step 2
        ZZLugar = Val(Mid(barralargo, lni, 2))
        lccar = lccar + ZZCara(ZZLugar)
    Next
    
    Rem barralargo = "{" + lccar + "}"
    barralargo = "(" + lccar + ")"
    
    ZZImpreBarra = barralargo
    ZZImpreBarraII = ZZNumero

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


