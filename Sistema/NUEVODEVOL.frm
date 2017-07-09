VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgNuevoDevol 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Devoluciones de Productos"
   ClientHeight    =   7995
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   11565
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
   ScaleHeight     =   7995
   ScaleWidth      =   11565
   Visible         =   0   'False
   Begin VB.TextBox RemitoII 
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
      TabIndex        =   70
      Text            =   " "
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox Cae 
      Height          =   360
      Left            =   9240
      TabIndex        =   67
      Top             =   6120
      Width           =   1935
   End
   Begin VB.TextBox VtoCae 
      Height          =   360
      Left            =   9240
      TabIndex        =   66
      Top             =   6600
      Width           =   1575
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
      Left            =   9840
      MaxLength       =   8
      TabIndex        =   64
      Text            =   " "
      Top             =   480
      Width           =   1215
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
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.CheckBox Comision 
      Caption         =   "Comision 50%"
      Height          =   255
      Left            =   8640
      TabIndex        =   63
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox Partida 
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
      MaxLength       =   6
      TabIndex        =   62
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame PantallaConfirma 
      Height          =   1335
      Left            =   2280
      TabIndex        =   59
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
         TabIndex        =   61
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
         TabIndex        =   60
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
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   4200
      Width           =   375
   End
   Begin VB.ComboBox TipoIva 
      Height          =   360
      Left            =   6480
      TabIndex        =   57
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Expreso 
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
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   53
      Text            =   " "
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
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
      Height          =   975
      Left            =   10560
      MouseIcon       =   "NUEVODEVOL.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Elimina el Registro"
      Top             =   2760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Vendedor 
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
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   49
      Text            =   " "
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Pago 
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
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   44
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
      Left            =   10560
      MouseIcon       =   "NUEVODEVOL.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Pedidos de Clientes"
      Top             =   1680
      Visible         =   0   'False
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
      Left            =   10560
      MouseIcon       =   "NUEVODEVOL.frx":1720
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":1A2A
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Impresion"
      Top             =   3840
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
      Left            =   9600
      MouseIcon       =   "NUEVODEVOL.frx":226C
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":2576
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1680
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
      Left            =   9600
      MouseIcon       =   "NUEVODEVOL.frx":2DB8
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":30C2
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Elimina el Registro"
      Top             =   2760
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
      Left            =   9600
      MouseIcon       =   "NUEVODEVOL.frx":3904
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":3C0E
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3840
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
      Left            =   9600
      MouseIcon       =   "NUEVODEVOL.frx":4450
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":475A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Consulta de Datos"
      Top             =   4920
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
      Left            =   10560
      MouseIcon       =   "NUEVODEVOL.frx":4F9C
      MousePointer    =   99  'Custom
      Picture         =   "NUEVODEVOL.frx":52A6
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Menu Principal"
      Top             =   4920
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
      Left            =   3960
      MaxLength       =   8
      TabIndex        =   34
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   4800
      TabIndex        =   23
      Top             =   5760
      Width           =   2895
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
         Left            =   1440
         TabIndex        =   48
         Top             =   720
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
         Left            =   120
         TabIndex        =   47
         Top             =   720
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
         TabIndex        =   33
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
         Left            =   120
         TabIndex        =   32
         Top             =   1080
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
         Left            =   120
         TabIndex        =   31
         Top             =   1320
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
         Left            =   120
         TabIndex        =   30
         Top             =   1680
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
         Left            =   1440
         TabIndex        =   29
         Top             =   240
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
         Left            =   1440
         TabIndex        =   28
         Top             =   1080
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
         Left            =   1440
         TabIndex        =   27
         Top             =   1320
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
         Left            =   1440
         TabIndex        =   26
         Top             =   1680
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
         Left            =   1440
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Descuento"
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
         TabIndex        =   24
         Top             =   480
         Width           =   1095
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   17
      Top             =   3240
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   360
      Left            =   3480
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      Index           =   4
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3720
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
      Left            =   6480
      MaxLength       =   6
      TabIndex        =   9
      Text            =   " "
      Top             =   480
      Width           =   855
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
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   4455
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
      Height          =   1560
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.ListBox Pantalla 
      Height          =   1740
      ItemData        =   "NUEVODEVOL.frx":5AE8
      Left            =   120
      List            =   "NUEVODEVOL.frx":5AEF
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   18
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
      Height          =   4095
      Left            =   120
      TabIndex        =   36
      Top             =   1560
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7223
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label16 
      Caption         =   "Vto.Cae"
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
      Left            =   7920
      TabIndex        =   69
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label14 
      Caption         =   "Cae"
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
      Left            =   8040
      TabIndex        =   68
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label Label15 
      Caption         =   "Factura"
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
      Left            =   8040
      TabIndex        =   65
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo Iva"
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
      Left            =   5160
      TabIndex        =   56
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label DesExpreso 
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
      TabIndex        =   55
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
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
      Left            =   120
      TabIndex        =   54
      Top             =   1200
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label DesVendedor 
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
      Left            =   7320
      TabIndex        =   51
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label18 
      Caption         =   "Vendedor"
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
      Left            =   5160
      TabIndex        =   50
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
      TabIndex        =   46
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
      TabIndex        =   45
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
      Left            =   3120
      TabIndex        =   35
      Top             =   480
      Visible         =   0   'False
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
      TabIndex        =   22
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Descuento"
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
      Left            =   5160
      TabIndex        =   10
      Top             =   480
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
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   3375
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
Attribute VB_Name = "PrgNuevoDevol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
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
Dim WPuntoII As String


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

Dim ZZCantidad As String
Dim ZZCantidadII As String

Dim WVector(100, 10) As String
Dim ZZPasaDatos(100, 10) As String

Dim WWArticulo As String
Dim WWDescripcion As String
Dim WWCantidad As String
Dim WWPrecio As String
Dim WWImpre As Double
Dim WWDto(10) As Double
Dim WWIva(10) As Double

Dim XXPrecio As Double
Dim XXDto As Double
Dim XXComision As Double
Dim XXImpoComision As Double

Dim ZZImpreBarra As String
Dim ZZImpreBarraII As String
Dim ZZZLetra As String
Dim ZZZPunto As String
Dim ZZZNumeroFac As String

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
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then

        T$ = "Anulacion de Comprobantes"
        m$ = "Esta seguro que Desea Anular el Comprobante "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "02"
                
            ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
               
            Rem ZSql = ""
            Rem ZSql = ZSql + "Select *"
            Rem ZSql = ZSql + " FROM Ctacte"
            Rem ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
            Rem spCtaCte = ZSql
            Rem Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstCtaCte.RecordCount > 0 Then
            Rem     ZSaldo = rstCtaCte!Saldo
            Rem     ZTotal = rstCtaCte!Total
            Rem     rstCtaCte.Close
            Rem     If ZSaldo <> ZTotal Then
            Rem         m$ = "El comprobante se encuentra total o parcialmente cancelado"
            Rem         a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
            Rem         Exit Sub
            Rem     End If
            Rem End If
            
            Erase WVector
        
            For WRenglon = 1 To 50
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "02" + Auxi + Auxi1
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                
                    Articulo = rstEstadistica!Articulo
                    Cantidad = rstEstadistica!Cantidad
                    CantidadII = rstEstadistica!CantidadII
                        
                    rstEstadistica.Close
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Salidas = Salidas - " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " StockI = StockI + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                
            Next WRenglon
            
            ZSql = ""
            ZSql = ZSql + "DELETE Estadistica"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "02" + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
            If Val(Remito.Text) <> 0 Then
                
                WPunto = Punto.Text
                Call Ceros(WPunto, 4)
                    
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                    
                WTipo = "02"
                    
                WClave = Letra.Text + WTipo + WPunto + Auxi + "01"
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Ctacte"
                ZSql = ZSql + " Where Ctacte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                    ZZBajaSaldo = rstCtaCte!Total * -1
                    ZZBajaSaldoUs = rstCtaCte!Totalus * -1
                    rstCtaCte.Close
                End If
            
                WLetra = Letra.Text
                WTipo = "01"
                WPuntoII = RemitoII.Text
                Call Ceros(WPuntoII, 4)
                WNumero = Remito.Text
                
                Auxi$ = Remito.Text
                Call Ceros(Auxi$, 8)
                WClave = WLetra + WTipo + WPuntoII + Auxi$ + "01"
                
                ZSql = ""
                ZSql = ZSql + "UPDATE CtaCte SET "
                ZSql = ZSql + " Saldo = Saldo + " + "'" + Str$(ZZBajaSaldo) + "',"
                ZSql = ZSql + " SaldoUs = SaldoUs + " + "'" + Str$(ZZBajaSaldoUs) + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "02"
                
            WClave = Letra.Text + WTipo + WPunto + Auxi + "01"
            
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
            ZSql = ZSql + " and Tipo = " + "'" + "02" + "'"
            ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            Call Limpia_Click
            
            Cliente.SetFocus
            
        End If

    End If

End Sub


Private Sub Consulta_Click()

    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem ""
    Opcion.AddItem "Condicion de Pago"
    Opcion.AddItem "Articulos"

    Opcion.Visible = True
     
 End Sub

Private Sub Impresion_Click()

    Call Impresion_FacturaFe
    
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
            
        Case 1
            
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
    
    For a = 1 To 50
    
        WCantidad = Val(WVector1.TextMatrix(a, 3))
        WPrecio = Val(WVector1.TextMatrix(a, 4))
        
        If Letra.Text = "B" Then
            If TipoIva.ListIndex = 0 Then
                WWImpre = WPrecio * (1 + (ConfigIva1) / 100)
                    Else
                WWImpre = WPrecio * (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(WWImpre)
            WPrecio = WWImpre
        End If
        
        If Val(WCantidad) <> 0 Then
            Select Case Partida.Text
                Case "/"
                    WCantidad = WCantidad / 2
                Case "?"
                    WCantidad = WCantidad / 12
                Case Else
            End Select
        End If
        
        WNeto = WNeto + (WPrecio * WCantidad)
        
    Next a
    
    Call Calcula_Importe
    
End Sub

Private Sub CalculaReal_Click()

    WNeto = 0
    
    For a = 1 To 50
    
        WCantidad = Val(WVector1.TextMatrix(a, 3))
        WPrecio = Val(WVector1.TextMatrix(a, 4))
        
        If Letra.Text = "B" Then
            WWImpre = WPrecio * 1.21
            Call Redondeo(WWImpre)
            WPrecio = WWImpre
        End If
        
        If Val(WCantidad) <> 0 Then
            Select Case Partida.Text
                Case "/"
                    WCantidad = WCantidad / 2
                Case "?"
                    WCantidad = WCantidad / 12
                Case Else
            End Select
        End If
        
        WNeto = WNeto + (WPrecio * WCantidad)
        
    Next a
    
    Call Calcula_Importe
    
End Sub


Private Sub Calcula_Importe()

    WImpoDto = 0
    WDescuento = Val(Descuento.Text)
    
    WDescuento = WDescuento
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
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
                If TipoIva.ListIndex = 0 Then
                    WIva1 = WNeto * ((ConfigIva1) / 100)
                    Call Redondeo(WIva1)
                        Else
                    WIva1 = WNeto * ((ConfigIva2) / 100)
                    Call Redondeo(WIva1)
                End If
        End Select
    End If
    
    WWIva(1) = WIva1
    WWIva(2) = WIva2
    
    WTotal = WNeto + WIva1 + WIva2
    
    SubTotal.Caption = Str$(WNeto + WImpoDto)
    Dto.Caption = Str$(WImpoDto)
    Neto.Caption = Str$(WNeto)
    Iva1.Caption = Str$(WIva1)
    Iva2.Caption = Str$(WIva2)
    Total.Caption = Str$(WTotal)
    
    SubTotal.Caption = Pusing("###,###.##", SubTotal.Caption)
    Dto.Caption = Pusing("###,###.##", Dto.Caption)
    Neto.Caption = Pusing("###,###.##", Neto.Caption)
    Iva1.Caption = Pusing("###,###.##", Iva1.Caption)
    Iva2.Caption = Pusing("###,###.##", Iva2.Caption)
    Total.Caption = Pusing("###,###.##", Total.Caption)

End Sub

Private Sub cmdClose_Click()
    PrgNuevoDevol.Hide
    Unload Me
    Menu4.Show
End Sub

Private Sub Graba_Click()

    Call Calcula_Click
    
    WNeto = Val(Neto.Caption)
    WIva1 = Val(Iva1.Caption)
    WIva2 = Val(Iva2.Caption)
    WTotal = Val(Total.Caption)
    
    WPuntoII = RemitoII.Text
    Call Ceros(WPuntoII, 4)
            
    Auxi = Remito.Text
    Call Ceros(Auxi, 8)
            
    WTipo = "01"
            
    ClaveVen$ = Letra.Text + WTipo + WPuntoII + Auxi + "01"
        
    ZSaldo = 0
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ctacte"
    ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        ZSaldo = rstCtaCte!Saldo
        rstCtaCte.Close
    End If
    
    If Val(Total.Caption) > ZSaldo Then
        m$ = "El monto de la nota de credito supera la factura"
        a% = MsgBox(m$, 0, "Ingreso de Facturas")
        Exit Sub
    End If
    
    
    
    
    
    
        
    
    ZZLineas = 0
    Erase ZZPasaDatos
        
    For ZZCiclo = 1 To 50
        
        ZZArticulo = WVector1.TextMatrix(ZZCiclo, 1)
        ZZDesArticulo = WVector1.TextMatrix(ZZCiclo, 2)
        ZZCantidad = WVector1.TextMatrix(ZZCiclo, 3)
        ZZPrecio = WVector1.TextMatrix(ZZCiclo, 4)
        
        If Val(ZZCantidad) <> 0 Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Exibidores"
            ZSql = ZSql + " Where Exibidores.Codigo = " + "'" + ZZArticulo + "'"
            spExibidores = ZSql
            Set rstExibidores = db.OpenRecordset(spExibidores, dbOpenSnapshot, dbSQLPassThrough)
            If rstExibidores.RecordCount > 0 Then
            
                ZZLineas = ZZLineas + 1
                ZZPasaDatos(ZZLineas, 1) = ZZArticulo
                ZZPasaDatos(ZZLineas, 2) = ZZDesArticulo
                ZZPasaDatos(ZZLineas, 3) = ZZCantidad
                ZZPasaDatos(ZZLineas, 4) = ZZPrecio
                ZZPasaDatos(ZZLineas, 7) = "1"
            
                rstExibidores.Close
                
                For WRenglon = 1 To 100
                
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    
                    WClave = Trim(ZZArticulo) + Auxi1
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Exibidores"
                    ZSql = ZSql + " Where Exibidores.Clave = " + "'" + WClave + "'"
                    spExibidores = ZSql
                    Set rstExibidores = db.OpenRecordset(spExibidores, dbOpenSnapshot, dbSQLPassThrough)
                    If rstExibidores.RecordCount > 0 Then
                        ZZLineas = ZZLineas + 1
                        ZZPasaDatos(ZZLineas, 1) = rstExibidores!Articulo
                        ZZPasaDatos(ZZLineas, 3) = Str$(Val(ZZCantidad) * rstExibidores!Cantidad)
                        ZZPasaDatos(ZZLineas, 7) = "2"
                        rstExibidores.Close
                    End If
                
                Next WRenglon
                    
                    Else
                
                ZZLineas = ZZLineas + 1
                ZZPasaDatos(ZZLineas, 1) = ZZArticulo
                ZZPasaDatos(ZZLineas, 2) = ZZDesArticulo
                ZZPasaDatos(ZZLineas, 3) = ZZCantidad
                ZZPasaDatos(ZZLineas, 4) = ZZPrecio
                ZZPasaDatos(ZZLineas, 7) = "0"
                
            End If
            
        End If
        
    Next ZZCiclo
        
    For CicloII = 1 To 50
    
        If Val(ZZPasaDatos(CicloII, 7)) = 2 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZPasaDatos(CicloII, 1) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZPasaDatos(CicloII, 2) = rstArticulo!Descripcion
                ZZPasaDatos(CicloII, 4) = Str$(rstArticulo!Precio)
                rstArticulo.Close
            End If
            
        End If
        
    Next CicloII
        
        
        
        
        
        
        
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
            
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
            
    WTipo = "02"
            
    ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ctacte"
    ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        rstCtaCte.Close
        m$ = "Factura ya emitida"
        a% = MsgBox(m$, 0, "Ingreso de Facturas")
        Exit Sub
    End If
    
    
    
    
    
    
    Pasa = "S"
    
    For a = 1 To 50
        
        Articulo = WVector1.TextMatrix(a, 1)
        ZDescripcion = WVector1.TextMatrix(a, 2)
        Cantidad = Val(WVector1.TextMatrix(a, 3))
        
        If Val(Cantidad) <> 0 Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZIva = rstArticulo!Iva
                rstArticulo.Close
            End If
            
            If ZZIva <> TipoIva.ListIndex Then
                m$ = "La condicion de Iva del Articulo " + ZDescripcion + " no coincide con el informado en la factura"
                a% = MsgBox(m$, 0, "Emision de Facturas")
                Exit Sub
            End If
            
            Select Case Partida.Text
                Case "/"
                    MiResultado = Val(Cantidad) Mod 2
                    If MiResultado <> 0 Then
                        m$ = "Las cantidades no son concordantes con el tipo de facturacion en el articulo " + ZDescripcion
                        a% = MsgBox(m$, 0, "Emision de Facturas")
                        Exit Sub
                    End If
                Case "?"
                    MiResultado = Val(Cantidad) Mod 12
                    If MiResultado <> 0 Then
                        m$ = "Las cantidades no son concordantes con el tipo de facturacion"
                        a% = MsgBox(m$, 0, "Emision de Facturas")
                        Exit Sub
                    End If
                Case Else
            End Select
            
        End If
        
                                        
    Next a
    
    If Letra.Text = "B" Then
        If TipoIva.ListIndex = 0 Then
            WNeto = Val(Total.Caption) / (1 + ((ConfigIva1) / 100))
                Else
            WNeto = Val(Total.Caption) / (1 + ((ConfigIva2) / 100))
        End If
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
    
    ZZTipo = "02"
    ZZImpre = "NC"
            
    ZZPunto = WPunto
    ZZLetra = Letra.Text
    ZZNumero = Auxi
    ZZRenglon = "01"
    ZZCliente = Cliente.Text
    ZZfecha = Fecha.Text
    ZZEstado = "0"
    ZZVencimiento = Fecha.Text
    ZZTotal = Str$(WTotal * -1)
    ZZSaldo = Str$(WTotal * -1)
    
    If Letra.Text = "B" Then
        If TipoIva.ListIndex = 0 Then
            WNeto = WTotal / (1 + ((ConfigIva1) / 100))
                Else
            WNeto = WTotal / (1 + ((ConfigIva2) / 100))
        End If
        Call Redondeo(WNeto)
        WIva1 = WTotal - WNeto
        WIva2 = 0
        ZZNeto = Str$(WNeto * -1)
        ZZIva1 = Str$(WIva1 * -1)
        ZZIva2 = Str$(WIva2 * -1)
            Else
        ZZNeto = Str$(WNeto * -1)
        ZZIva1 = Str$(WIva1 * -1)
        ZZIva2 = Str$(WIva2 * -1)
    End If
    
    Select Case Partida.Text
        Case "/"
            ZZTotalUs = Str$(WTotal + WNeto)
            ZZSaldoUs = Str$(WTotal + WNeto)
        Case "?"
            ZZTotalUs = Str$(WTotal + (WNeto * 11))
            ZZSaldoUs = Str$(WTotal + (WNeto * 11))
        Case Else
            ZZTotalUs = Str$(WTotal)
            ZZSaldoUs = Str$(WTotal)
    End Select
    
    ZZTotalUs = Str$(Val(ZZTotalUs) * -1)
    ZZSaldoUs = Str$(Val(ZZSaldoUs) * -1)
    
    ZZExento = "0"
    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZZOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZZPedido = Pedido.Text
    ZZRemito = ""
    ZZOrden = ""
    ZZProvincia = WProvincia
    ZZVendedor = Vendedor.Text
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
    ZZParidad = "0"
    ZZRemito1 = ""
    ZZRemito2 = ""
    ZZBusqueda = ZZLetra + WPunto + Auxi
    
    ZZDescuento = Descuento.Text
    ZZPago = Pago.Text
    ZZPartida = Partida.Text
    ZZExpreso = Expreso.Text
    ZZTipoIva = Str$(TipoIva.ListIndex)
    ZZComision = Str$(Comision.Value)
    ZZRemito = Remito.Text
    
    ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
    
    ZZCae = Cae.Text
    ZZVtoCae = VtoCae.Text
    
    ZZLinea = ""
    
    ZZNetoTotal = ZZNeto
    If ZZPartida = "/" Then
        ZZNetoTotal = Str$(WNeto * 2 * -1)
    End If
    If ZZPartida = "?" Then
        ZZNetoTotal = Str$(WNeto * 12 * -1)
    End If
    
    If Val(Remito.Text) <> 0 Then
        ZZBajaSaldo = ZZSaldo
        ZZBajaSaldoUs = ZZSaldoUs
        ZZSaldo = "0"
        ZZSaldoUs = "0"
            Else
        ZZBajaSaldo = "0"
        ZZBajaSaldoUs = "0"
    End If
    
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
    ZSql = ZSql + "'" + ZZBusqueda + "')"
                            
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    Renglon = 0
    WRenglon = 0
        
    For a = 1 To 50
        
        WRenglon = WRenglon + 1
        
        Articulo = UCase(ZZPasaDatos(WRenglon, 1))
        DesArticulo = ZZPasaDatos(WRenglon, 2)
        Cantidad = Val(ZZPasaDatos(WRenglon, 3)) * -1
        Precio = Val(ZZPasaDatos(WRenglon, 4))
        TipoII = Val(ZZPasaDatos(WRenglon, 7))
        Preciosalva = Val(ZZPasaDatos(WRenglon, 4))
            
        If Cantidad <> 0 Then
        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Numero.Text)
            Call Ceros(Auxi1, 8)
            
            If TipoII = 1 Then
                Precio = 0
            End If
            
            ZZCosto1 = "0"
            ZZComision = "0"
            ZZTipoComision = Str$(Comision.Value)
            
            If TipoII <> 1 Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZCosto1 = Str$(rstArticulo!Costo)
                    ZZComision = Str$(rstArticulo!Comision)
                    rstArticulo.Close
                End If
            End If
            
            XXPrecio = Precio
            If Val(Descuento.Text) <> 0 Then
                XXDto = XXPrecio * (Val(Descuento.Text) / 100)
                Call Redondeo(XXDto)
                XXPrecio = XXPrecio - XXDto
                Call Redondeo(XXPrecio)
            End If
            
            If Comision.Value = 0 Then
                XXComision = ZZComision
                    Else
                XXComision = ZZComision * 0.5
            End If
            
            XXImpoComision = XXPrecio * (XXComision / 100)
            Call Redondeo(XXImpoComision)
            
            XXPrecio = XXPrecio - XXImpoComision
            Call Redondeo(XXPrecio)
            
            ZZTipo = "02"
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
            ZZParidad = "0"
            ZZVendedor = "0"
            ZZRubro = "0"
            ZZLinea = "0"
            ZZCosto2 = "0"
            ZZCoeficiente = "0"
            ZZPedido = "0"
            ZZfecha = Fecha.Text
            ZZImporte1 = "0"
            ZZImporte2 = "0"
            ZZImporte3 = "0"
            ZZImporte4 = "0"
            ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZZWArticulo = ""
            ZZRemito = ""
            ZZTipoII = Str$(TipoII)
            
            ZZZLetra = Letra.Text
            ZZLetra = Letra.Text
            
            ZZZPunto = Punto.Text
            Call Ceros(ZZZPunto, 1)
            
            ZZZNumeroFac = Numero.Text
            Call Ceros(ZZZNumeroFac, 6)
            
            ZZClave = "02" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi
            
            ZZWDate = Date$
            ZZClaveCtacte = "02" + Auxi1 + "01"
            
            ZZImprefactura = "N/C"
            ZZNroFactura = Auxi1
            ZZTalle = Talle
            ZZColor = XXColor
            ZZCuenta = WCuenta
            ZZDescuento = Descuento.Text
            ZZPartida = Partida.Text
            
            ZZCantidadII = ZZCantidad
            If ZZPartida = "/" Then
                ZZCantidadII = Str$(Val(ZZCantidad) / 2)
            End If
            If ZZPartida = "?" Then
                ZZCantidadII = Str$(Val(ZZCantidad) / 12)
            End If
            
            ZZPrecioII = Str$(XXPrecio)
            
            ZZSumaComision = ZZSumaComision + (XXImpoComision * Val(ZZCantidadII))
            
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
            ZSql = ZSql + "Coeficiente ,"
            ZSql = ZSql + "Pedido ,"
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
            ZSql = ZSql + "'" + ZZCoeficiente + "',"
            ZSql = ZSql + "'" + ZZPedido + "',"
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
                        
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Salidas = Salidas + " + "'" + Str$(Cantidad) + "',"
            ZSql = ZSql + " Stock = Stock - " + "'" + Str$(Cantidad) + "',"
            ZSql = ZSql + " StockI = StockI - " + "'" + Str$(Cantidad) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
        End If
                                        
    Next a
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    
    ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Costo = " + "'" + Str$(ZZSumaComision) + "'"
    ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Impresion_FacturaFe

    Rem T$ = "Emision de Facturas"
    Rem m$ = "Desea Imprimir la Factura"
    Rem Respuesta% = MsgBox(m$, 32 + 4, T$)
    Rem If Respuesta% = 6 Then
    Rem     Call WImpresion
    Rem End If
    
    If Val(Remito.Text) <> 0 Then
    
        WLetra = Letra.Text
        WTipo = "01"
        WPuntoII = RemitoII.Text
        Call Ceros(WPuntoII, 4)
        WNumero = Remito.Text
        
        Auxi$ = Remito.Text
        Call Ceros(Auxi$, 8)
        WClave = WLetra + WTipo + WPuntoII + Auxi$ + "01"
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCte SET "
        ZSql = ZSql + " Saldo = Saldo + " + "'" + ZZBajaSaldo + "',"
        ZSql = ZSql + " SaldoUs = SaldoUs + " + "'" + ZZBajaSaldoUs + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    End If
        
    Call Limpia_Click
        
    Cliente.SetFocus
        
End Sub










Private Sub CmdDelete_Click()

    T$ = "Baja de Comprobantes"
    m$ = "Desea Borrar el Comprobante "
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then

        T$ = "Baja de Comprobantes"
        m$ = "Esta seguro que Desea Borrar el Comprobante "
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "02"
                
            ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
               
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                ZSaldo = rstCtaCte!Saldo
                ZTotal = rstCtaCte!Total
                rstCtaCte.Close
                If ZSaldo <> ZTotal Then
                    m$ = "El comprobante se encuentra total o parcialmente cancelado"
                    a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
                    Exit Sub
                End If
            End If
        
            Erase WVector
        
            For WRenglon = 1 To 50
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "02" + Auxi + Auxi1
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                
                    Articulo = rstEstadistica!Articulo
                    Cantidad = rstEstadistica!Cantidad
                    CantidadII = rstEstadistica!CantidadII
                        
                    rstEstadistica.Close
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Salidas = Salidas - " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " StockI = StockI + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                
            Next WRenglon
            
            ZSql = ""
            ZSql = ZSql + "DELETE Estadistica"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "02" + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            If Val(Remito.Text) <> 0 Then
                
                WPunto = Punto.Text
                Call Ceros(WPunto, 4)
                    
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                    
                WTipo = "02"
                    
                WClave = Letra.Text + WTipo + WPunto + Auxi + "01"
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Ctacte"
                ZSql = ZSql + " Where Ctacte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                    ZZBajaSaldo = rstCtaCte!Total * -1
                    ZZBajaSaldoUs = rstCtaCte!Totalus * -1
                    rstCtaCte.Close
                End If
            
                WLetra = Letra.Text
                WTipo = "01"
                WPuntoII = RemitoII.Text
                Call Ceros(WPuntoII, 4)
                WNumero = Remito.Text
                
                Auxi$ = Remito.Text
                Call Ceros(Auxi$, 8)
                WClave = WLetra + WTipo + WPuntoII + Auxi$ + "01"
                
                ZSql = ""
                ZSql = ZSql + "UPDATE CtaCte SET "
                ZSql = ZSql + " Saldo = Saldo + " + "'" + Str$(ZZBajaSaldo) + "',"
                ZSql = ZSql + " SaldoUs = SaldoUs + " + "'" + Str$(ZZBajaSaldoUs) + "'"
                ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            
            ZSql = ""
            ZSql = ZSql + "DELETE CtaCte"
            ZSql = ZSql + " Where Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Tipo = " + "'" + "02" + "'"
            ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
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
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Descuento.Text = ""
    Pedido.Text = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Vendedor.Text = ""
    Expreso.Text = ""
    DesVendedor.Caption = ""
    DesExpreso.Caption = ""
    Comision.Value = 0
    Partida.Text = ""
    Remito.Text = ""
    RemitoII.Text = "3"
    Cae.Text = ""
    VtoCae.Text = ""
    
    Renglon = 0
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
    TipoIva.ListIndex = 0

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
            
        Case 2
            Indice = Pantalla.ListIndex
            Pago.Text = WIndice.List(Indice)
            Call Pago_Keypress(13)
            
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
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Descuento.Text = ""
    Pedido.Text = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Vendedor.Text = ""
    Expreso.Text = ""
    DesVendedor.Caption = ""
    DesExpreso.Caption = ""
    Comision.Value = 0
    Partida.Text = ""
    Remito.Text = ""
    RemitoII.Text = "3"
    Cae.Text = ""
    VtoCae.Text = ""
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
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
        ConfigPunto = 3
        CantiFac = rstConfiguracion!CantiFac
        CantiRem = rstConfiguracion!CantiRem
        CantiArti = rstConfiguracion!CantiArti
        rstConfiguracion.Close
    End If
    
    TipoIva.Clear
    
    TipoIva.AddItem "21 %"
    TipoIva.AddItem "10.5 %"
    
    TipoIva.ListIndex = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 50
            
        If Val(Punto.Text) = 1 Then
    
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "02" + Auxi + Auxi1
            
                Else
            
            ZZZLetra = Letra.Text
            
            ZZZPunto = Punto.Text
            Call Ceros(ZZZPunto, 1)
            
            ZZZNumeroFac = Numero.Text
            Call Ceros(ZZZNumeroFac, 6)
                
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "02" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi1
                
        End If
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Estadistica"
        ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
        
            ZZTipoII = IIf(IsNull(rstEstadistica!TipoII), "o", rstEstadistica!TipoII)
            If ZZTipoII = 0 Or ZZTipoII = 1 Then
            
                Canti = rstEstadistica!Cantidad
                
                Renglon = Renglon + 1
                        
                WVector1.Row = Renglon
                        
                WVector1.Col = 1
                WVector1.Text = rstEstadistica!Articulo
                Auxi1 = rstEstadistica!Articulo
                    
                WVector1.Col = 2
                WVector1.Text = IIf(IsNull(rstEstadistica!Descripcion), "", rstEstadistica!Descripcion)
                    
                WVector1.Col = 3
                WVector1.Text = Pusing("###,###", Str$(rstEstadistica!Cantidad * -1))
                    
                If ZZTipoII = 0 Then
                
                    WVector1.Col = 4
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Precio))
                    
                    WVector1.Col = 5
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Precio * rstEstadistica!CantidadII * -1))
                    
                       Else
                
                    WVector1.Col = 4
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Preciosalva))
                    
                    WVector1.Col = 5
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Preciosalva * rstEstadistica!CantidadII * -1))
                    
                End If
                
                rstEstadistica.Close
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
            End If
                
        End If
    
    Next WRenglon

    Call Calcula_Click
    
    Graba.Enabled = True

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Trim(Cliente.Text) <> "" Then
            Auxi = UCase(Left$(Cliente.Text, 1))
            Auxi1 = Mid$(Cliente.Text, 2, 5)
            Call Ceros(Auxi1, 3)
            Cliente.Text = Auxi + "-" + Auxi1
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            Descuento.Text = Str$(rstCliente!Descuento)
            Descuento.Text = Pusing("###,###.##", Descuento.Text)
            Vendedor.Text = rstCliente!Vendedor
            Pago.Text = rstCliente!Condicion
            Expreso.Text = rstCliente!Expreso
            WProvincia = rstCliente!Provincia
            WCodIva = rstCliente!Iva
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            Select Case Val(WCodIva)
                Case 1, 2
                    Letra.Text = "A"
                Case Else
                    Letra.Text = "B"
            End Select
            ZMarca = IIf(IsNull(rstCliente!Marca), "0", rstCliente!Marca)
            
            Rem If Letra.Text = "B" Then
            Rem     m$ = "COLOQUE EL FORMULARIO B"
            Rem     a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Rem End If
            
            rstCliente.Close
                
            Rem ZSql = ""
            Rem ZSql = ZSql + "Select *"
            Rem ZSql = ZSql + " FROM ClienteAdicional"
            Rem ZSql = ZSql + " Where ClienteAdicional.Cliente = " + "'" + Cliente.Text + "'"
            Rem ZSql = ZSql + " and ClienteAdicional.Linea = " + "'" + Linea.Text + "'"
            Rem spClienteAdicional = ZSql
            Rem Set rstClienteAdicional = db.OpenRecordset(spClienteAdicional, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstClienteAdicional.RecordCount > 0 Then
            Rem     Descuento.Text = Str$(rstClienteAdicional!Descuento)
            Rem     Descuento.Text = Pusing("###,###.##", Descuento.Text)
            Rem     Vendedor.Text = rstClienteAdicional!Vendedor
            Rem     rstClienteAdicional.Close
            Rem End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Codigo = " + "'" + Vendedor.Text + "'"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                DesVendedor.Caption = rstVendedor!Nombre
                rstVendedor.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Expreso"
            ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
            spExpreso = ZSql
            Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
            If rstExpreso.RecordCount > 0 Then
                DesExpreso.Caption = rstExpreso!Nombre
                rstExpreso.Close
                    Else
                DesExpreso.Caption = ""
            End If

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
            WTipo = "02"
            
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
                                If Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 5 Then
                                    Numero.Text = Str$(Val(rstCtaCte!Numero) + 1)
                                    Rem Remito.Text = Str$(Val(rstCtaCte!Remito) + 1)
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
            
            Confirma.Text = Partida.Text
            Remito.SetFocus
            Rem PantallaConfirma.Visible = True
            Rem Confirma.SetFocus
            
                Else
                
            Cliente.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub RemitoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Remito.SetFocus
    End If
    If KeyAscii = 27 Then
        RemitoII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Remito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WPuntoII = RemitoII.Text
        Call Ceros(WPuntoII, 4)
            
        Auxi = Remito.Text
        Call Ceros(Auxi, 8)
            
        WTipo = "01"
            
        ClaveVen$ = Letra.Text + WTipo + WPuntoII + Auxi + "01"
           
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            ZCliente = rstCtaCte!Cliente
            rstCtaCte.Close
            If Trim(ZCliente) <> Trim(Cliente.Text) Then
                m$ = "El Cliente de la factura no coincide con el de la nota de credito"
                a% = MsgBox(m$, 0, "Nota de Credito")
                    Else
                PantallaConfirma.Visible = True
                Confirma.SetFocus
            End If
                
                Else
                    
            m$ = "Factura Inexistente"
            a% = MsgBox(m$, 0, "Nota de Credito")
                
        End If
    End If
    If KeyAscii = 27 Then
        Remito.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub



Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Numero.Text
        Call Ceros(WPunto, 4)
        
        Numero.Text = "1"
        WTipo = "02"
        
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
                            If Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 5 Then
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
            
        WTipo = "02"
            
        ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
           
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            
            Fecha.Text = rstCtaCte!Fecha
            Cliente.Text = rstCtaCte!Cliente
            Pedido.Text = Str$(Val(rstCtaCte!Pedido))
            Descuento.Text = Str$(rstCtaCte!Descuento)
            Descuento.Text = Pusing("###,###.##", Descuento.Text)
            Pago.Text = rstCtaCte!Pago
            Partida.Text = rstCtaCte!Partida
            Comision.Value = rstCtaCte!Comision
            Expreso.Text = rstCtaCte!Expreso
            TipoIva.ListIndex = rstCtaCte!TipoIva
            Comision.Value = rstCtaCte!Comision
            Remito.Text = Trim(rstCtaCte!Remito)
            Cae.Text = IIf(IsNull(rstCtaCte!Cae), "", rstCtaCte!Cae)
            VtoCae.Text = IIf(IsNull(rstCtaCte!VtoCae), "", rstCtaCte!VtoCae)
            
            rstCtaCte.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                Descuento.Text = Str$(rstCliente!Descuento)
                Descuento.Text = Pusing("###,###.##", Descuento.Text)
                WProvincia = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
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
            Rem Pedido.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Descuento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento.Text = Pusing("###,###.##", Descuento.Text)
        Pago.SetFocus
    End If
    If KeyAscii = 27 Then
        Descuento.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            Rem Expreso.SetFocus
                Else
            Pago.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Pago.Text = ""
        DesPago.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub EXPRESO_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Expreso"
        ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
        spExpreso = ZSql
        Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
        If rstExpreso.RecordCount > 0 Then
            DesExpreso.Caption = rstExpreso!Nombre
            rstExpreso.Close
            Confirma.Text = Partida.Text
            PantallaConfirma.Visible = True
            Confirma.SetFocus
                Else
            Expreso.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Expreso.Text = ""
        DesExpreso.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Confirma_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Confirma.Text = Trim(UCase(Confirma.Text))
        If Confirma.Text = "S" Or Confirma.Text = "N" Or Confirma.Text = "/" Or Confirma.Text = "?" Then
            PantallaConfirma.Visible = False
            If Confirma.Text <> "N" Then
                Partida.Text = Confirma.Text
                Call Lee_Pedido
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
        Case 4
            Rem If WVector1.Row < WVector1.Rows - 1 Then
            If WVector1.Row < 19 Then
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
            If Trim(WVector1.Text) <> "" Then
                Auxi = UCase(Left$(WVector1.Text, 1))
                Auxi1 = Mid$(WVector1.Text, 2, 5)
                Call Ceros(Auxi1, 5)
                WVector1.Text = Auxi + Auxi1
            End If
            
            ZPasa = "S"
            If ZPasa = "S" Then
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZIva = rstArticulo!Iva
                    If ZZIva <> TipoIva.ListIndex Then
                        m$ = "La condicion de Iva del Articulo " + ZDescripcion + " no coincide con el informado en la factura"
                        a% = MsgBox(m$, 0, "Emision de Facturas")
                        WControl = "N"
                            Else
                        WVector1.Col = 2
                        WVector1.Text = rstArticulo!Descripcion
                        WVector1.Col = 4
                        WVector1.Text = Str$(rstArticulo!Precio)
                        WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        WVector1.Col = 2
                    End If
                    rstArticulo.Close
                        Else
                    WControl = "N"
                End If
        
            End If
            
        Case 3
            Select Case Partida.Text
                Case "/"
                    MiResultado = Val(WVector1.TextMatrix(WVector1.Row, 3)) Mod 2
                    If MiResultado <> 0 Then
                        m$ = "Las cantidades no son concordantes con el tipo de facturacion en el articulo " + ZDescripcion
                        a% = MsgBox(m$, 0, "Emision de Facturas")
                        WVector1.TextMatrix(WVector1.Row, 3) = ""
                        WControl = "N"
                    End If
                Case "?"
                    MiResultado = Val(WVector1.TextMatrix(WVector1.Row, 3)) Mod 12
                    If MiResultado <> 0 Then
                        m$ = "Las cantidades no son concordantes con el tipo de facturacion"
                        a% = MsgBox(m$, 0, "Emision de Facturas")
                        WVector1.TextMatrix(WVector1.Row, 3) = ""
                        WControl = "N"
                    End If
                Case Else
            End Select
            
            WCantidad = Val(WVector1.Text)
            If Val(WCantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        WCantidad = WCantidad / 2
                    Case "?"
                        WCantidad = WCantidad / 12
                    Case Else
                End Select
            End If
        
            WVector1.TextMatrix(WVector1.Row, 5) = Str$(WCantidad * Val(WVector1.TextMatrix(WVector1.Row, 4)))
            WVector1.TextMatrix(WVector1.Row, 5) = Pusing("###,###.##", WVector1.TextMatrix(WVector1.Row, 5))
            
        Case 4
            WCantidad = Val(WVector1.TextMatrix(WVector1.Row, 3))
            If Val(WCantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        WCantidad = WCantidad / 2
                    Case "?"
                        WCantidad = WCantidad / 12
                    Case Else
                End Select
            End If
        
            WVector1.TextMatrix(WVector1.Row, 5) = Str$(WCantidad * Val(WVector1.Text))
            WVector1.TextMatrix(WVector1.Row, 5) = Pusing("###,###.##", WVector1.TextMatrix(WVector1.Row, 5))
            
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
    WVector1.Cols = 6
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
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
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
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
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

    'Convertir en letras el número en Text1
    
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

Private Sub Partida_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call CmdDelete_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call PedidoAyuda_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub


Sub Impresion_Factura()


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WProvincia = rstCliente!Provincia
        WCodIva = rstCliente!Iva
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZCuitII = rstExpreso!Cuit
        rstExpreso.Close
    End If
    

    If Letra.Text = "A" Then
    

        Open "lpt3" For Output As #1
        Rem Open "dada.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(40); "NOTA DE CREDITO";
        Print #1, Tab(66); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(10); Trim(WRazon); " "; Cliente.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE"
        End Select
        
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35);
        If Val(Remito.Text) <> 0 Then
            Print #1, Tab(55); "Afecta Fc. :"; Remito.Text
                Else
            Print #1, Tab(55); ""
        End If
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            If Val(Cantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        Cantidad = Cantidad / 2
                    Case "?"
                        Cantidad = Cantidad / 12
                    Case Else
                End Select
            End If
            Precio = Val(WVector1.TextMatrix(a, 4))
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(22); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial))
                    Else
                Print #1, ""
            End If
            
        Next a
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(17); "Dto."; Alinea("###.##", Descuento.Text);
        End If
        If TipoIva.ListIndex = 0 Then
            Print #1, Tab(47); "21";
                Else
            Print #1, Tab(47); "10.5";
        End If
        Print #1, ""
        
        Print #1, Tab(2); Alinea("###,###.##", SubTotal.Caption);
        Print #1, Tab(15); Alinea("###,###.##", Dto.Caption);
        Print #1, Tab(27); Alinea("###,###.##", Neto.Caption);
        Print #1, Tab(38); Alinea("###,###.##", Iva1.Caption);
        Print #1, Tab(51); Alinea("###,###.##", Iva2.Caption);
        Print #1, Tab(63); Alinea("###,###.##", Total.Caption)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
            Else
    

        Open "lpt1" For Output As #1
        Rem Open "dada.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(40); "NOTA DE CREDITO";
        Print #1, Tab(66); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(10); Trim(WRazon); " "; Cliente.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE"
        End Select
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35);
        If Val(Remito.Text) <> 0 Then
            Print #1, Tab(55); "Afecta Fc. :"; Remito.Text
                Else
            Print #1, Tab(55); ""
        End If
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            If Val(Cantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        Cantidad = Cantidad / 2
                    Case "?"
                        Cantidad = Cantidad / 12
                    Case Else
                End Select
            End If
            
            Precio = Val(WVector1.TextMatrix(a, 4))
            If TipoIva.ListIndex = 0 Then
                WWImpre = Precio * (1 + (ConfigIva1) / 100)
                    Else
                WWImpre = Precio * (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(WWImpre)
            Precio = WWImpre
            
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(21); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial))
                    Else
                Print #1, ""
            End If
            
        Next a
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(7); "SubTotal:"; Alinea("###,###.##", SubTotal.Caption);
            Print #1, Tab(31); "Descuento:"; Alinea("###.##", Descuento.Text); "% "; Alinea("###,###.##", Dto.Caption);
        End If
        
        Print #1, Tab(65); Alinea("###,###.##", Total.Caption);
        Print #1, Tab(90); ZZCuitII
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
    End If
            
            

End Sub

Private Sub Lee_Pedido()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 99
    
        Auxi = Pedido.Text
        Call Ceros(Auxi, 8)
            
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Clave = " + "'" + WClave + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(rstPedido!Articulo) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZIva = rstArticulo!Iva
                rstArticulo.Close
            End If
            
            If ZZIva = TipoIva.ListIndex And rstPedido!facturado = 0 Then
        
                Canti = rstPedido!Cantidad
                
                Renglon = Renglon + 1
                        
                WVector1.Row = Renglon
                        
                WVector1.Col = 1
                WVector1.Text = Trim(rstPedido!Articulo)
                Auxi1 = rstPedido!Articulo
                    
                WVector1.Col = 3
                WVector1.Text = Pusing("###,###", Str$(rstPedido!Cantidad))
                    
                WVector1.Col = 4
                WVector1.Text = Pusing("###,###.##", Str$(rstPedido!Precio))
                
                WCantidad = rstPedido!Cantidad
                If Val(WCantidad) <> 0 Then
                    Select Case Partida.Text
                        Case "/"
                            WCantidad = WCantidad / 2
                        Case "?"
                            WCantidad = WCantidad / 12
                        Case Else
                    End Select
                End If
                
                
                WVector1.Col = 5
                WVector1.Text = Pusing("###,###.##", Str$(rstPedido!Precio * WCantidad))
                    
                rstPedido.Close
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                    Else
                    
                rstPedido.Close
                
            End If
            
        End If
    
    Next WRenglon
    Call Calcula_Click
    
End Sub

Private Sub Calcula_Cae()
    
    Dim WSAA As Object, WSFEv1 As Object
    
    On Error GoTo ManejoError
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    Rem Cae.Text = "12345678901234"
    Rem VtoCae.Text = "14/04/2011"
    Rem Exit Sub
    
    ' Crear objeto interface Web Service Autenticaci?n y Autorizaci?n
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    'Debug.Print WSAA.InstallDir
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
    tra = WSAA.CreateTRA("wsfe")
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
        
    ZPath = "c:\salva\"
    ZNombre = "celugama"
    ZCuit = "30637671622"
    punto_vta = 3
    
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
        
    ok = WSFEv1.Conectar(cache, wsdl, proxy, "") ' homologaci?n
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
        tipo_cbte = 3
            Else
        tipo_cbte = 8
    End If
    
    cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
    If cbte_nro = "" Then
        cbte_nro = 0                ' no hay comprobantes emitidos
            Else
        cbte_nro = CLng(cbte_nro)   ' convertir a entero largo
    End If
    
    If cbte_nro + 1 <> Val(Numero.Text) Then
        m$ = "Numero de comprobante no coincide con el de la afip"
        a% = MsgBox(m$, 0, "Ingreso de Facturas")
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WCuit = rstCliente!Cuit
        rstCliente.Close
        Call Eval
    End If
    
    Rem Fecha = Format(Date, "yyyymmdd")
    
    Rem CONCEPTO   1-PRODUCTO    2-SERVICIOS     3-PRODUCTOS Y SERVICIOS
    concepto = 1
    
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

    ok = WSFEv1.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
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
    If TipoIva.ListIndex = 0 Then
        Id = 5 ' 21%
            Else
        Id = 4 ' 10.50%
    End If
    base_imp = Val(Neto.Caption)
    Importe = Val(Iva1.Caption)
    ok = WSFEv1.AgregarIva(Id, base_imp, Importe)
    
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

    MsgBox "Resultado:" & WSFEv1.resultado & " CAE: " & Cae & " Venc: " & WSFEv1.vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
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
    Debug.Print "Fecha Vencimiento CAE", WSFEv1.vencimiento
    Debug.Print "Importe Total:", WSFEv1.ImpTotal
    Debug.Print "Resultado:", WSFEv1.resultado
    
    If Cae <> cae2 Then
        MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!: " & Cae & " vs " & cae2
    Else
        MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
    End If
        
    If WSFEv1.resultado = "A" Then
        Cae.Text = Trim(Cae)
        If Len(Trim(WSFEv1.vencimiento)) = 8 Then
            VtoCae.Text = Right$(WSFEv1.vencimiento, 2) + "/" + Mid$(WSFEv1.vencimiento, 5, 2) + "/" + Left$(WSFEv1.vencimiento, 4)
                Else
            VtoCae.Text = WSFEv1.vencimiento
        End If
    End If

    Exit Sub
ManejoError:
    ' Si hubo error:
    Debug.Print WSFEv1.Excepcion
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

    X = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               X = X + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               X = X + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               X = X + Y: MinusOk = 0

        End If

    Next

    WCuit = X

End Sub

Private Sub Calcula_Barra()
    
    Dim ZZCara(1000) As String
    
    ZZNumero = "30637671622"
    
    If Letra.Text = "A" Then
        ZZNumero = ZZNumero + "03"
            Else
        ZZNumero = ZZNumero + "08"
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
    ZZCara(90) = "¡"
    ZZCara(91) = "¢"
    ZZCara(92) = "£"
    ZZCara(93) = "¤"
    ZZCara(94) = "¥"
    ZZCara(95) = "¦"
    ZZCara(96) = "§"
    ZZCara(97) = "¨"
    ZZCara(98) = "©"
    ZZCara(99) = "ª"
    
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
        ZZRazon = rstCliente!Razon
        ZZDireccion = rstCliente!Direccion
        ZZLocalidad = rstCliente!Localidad
        ZZPostal = rstCliente!Postal
        ZZCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZCuitII = rstExpreso!Cuit
        rstExpreso.Close
    End If
    
    ZZLetra = "X"
    ZZTipo = "01"
    ZZPunto = "0001"
    Auxi1 = Numero.Text
    Call Ceros(Auxi1, 8)
    ZZFactura = Auxi1
    ZZfecha = Fecha.Text
    ZZCliente = Cliente.Text
    ZZNombre = Trim(ZZRazon)
    ZZDireccion = Trim(ZZDireccion)
    ZZLocalidad = Trim(ZZLocalidad)
    ZZPartida = Partida.Text
    ZZNeto = Neto.Caption
    ZZDto = Dto.Caption
    ZZNeto1 = SubTotal.Caption
    ZZIva1 = Iva1.Caption
    ZZIva2 = Iva2.Caption
    ZZTotal = Total.Caption
    ZZImprepago = Left$(DesPago.Caption, 35)
    ZZImpreIva = Iva(Val(ZZCodIva))
    If TipoIva.ListIndex = 0 Then
        ZZPorceIva = "21"
            Else
        ZZPorceIva = "10.5"
    End If
    ZZPorceDto = Descuento.Text
    Select Case Partida.Text
        Case "/"
            ZZPostal = "CP:B" + WPostal + "BIE"
        Case "?"
            ZZPostal = "CP%B" + WPostal + "BIE"
        Case Else
            ZZPostal = "CP B" + WPostal + "BIE"
    End Select
    
    Call Calcula_Barra
    
    For a = 1 To 35
        
        ZZRenglon = Str$(a)
        Auxi1 = ZZRenglon
        Call Ceros(Auxi1, 2)
        ZZRenglon = Auxi1
        
        ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
        
        ZZItem = ZZRenglon
        
        ZZArticulo = WVector1.TextMatrix(a, 1)
        ZZDescripcion = WVector1.TextMatrix(a, 2)
        ZZZCantidad = Val(WVector1.TextMatrix(a, 3))
        If Val(ZZZCantidad) <> 0 Then
            Select Case Partida.Text
                Case "/"
                    ZZZCantidad = ZZZCantidad / 2
                Case "?"
                    ZZZCantidad = ZZZCantidad / 12
                Case Else
            End Select
        End If
        ZZZPrecio = Val(WVector1.TextMatrix(a, 4))
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
        
        ZZImpre1 = ""
        ZZImpre2 = ""
        If Trim(Remito.Text) <> "" Then
            ZZImpre2 = "Afecta a Factura : " + Remito.Text
        End If
        Auxi2 = Numero.Text
        Call Ceros(Auxi2, 8)
        ZZImpre3 = Auxi2
        ZZImpre4 = "NOTA DE CREDITO"
        
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
        ZSql = ZSql + "Impre1 ,"
        ZSql = ZSql + "Impre2 ,"
        ZSql = ZSql + "Impre3 ,"
        ZSql = ZSql + "Impre4 ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "VtoCae ,"
        ZSql = ZSql + "ImpreBarra ,"
        ZSql = ZSql + "ImpreBarraII ,"
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
        ZSql = ZSql + "'" + ZZDireccion + "',"
        ZSql = ZSql + "'" + ZZLocalidad + "',"
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
        ZSql = ZSql + "'" + ZZImpre1 + "',"
        ZSql = ZSql + "'" + ZZImpre2 + "',"
        ZSql = ZSql + "'" + ZZImpre3 + "',"
        ZSql = ZSql + "'" + ZZImpre4 + "',"
        ZSql = ZSql + "'" + Cae.Text + "',"
        ZSql = ZSql + "'" + VtoCae.Text + "',"
        ZSql = ZSql + "'" + ZZImpreBarra + "',"
        ZSql = ZSql + "'" + ZZImpreBarraII + "',"
        ZSql = ZSql + "'" + ZZPorceIva + "',"
        ZSql = ZSql + "'" + ZZPorceDto + "')"
                                
        spFactura = ZSql
        Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
    
    Next a
    
    Listado.WindowTitle = "Emision de Comprobantes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Factura.Factura, Factura.Renglon, Factura.Fecha, Factura.Cliente, Factura.Nombre, Factura.Direccion, Factura.Localidad, Factura.Cuit, Factura.Descripcion, Factura.Neto, Factura.Dto, Factura.Neto1, Factura.Iva1, Factura.Iva2, Factura.Total, Factura.Imprepago, Factura.CondIva, Factura.Item, Factura.Articulo, Factura.Cantidad, Factura.Precio, Factura.PordeDto, Factura.Postal, Factura.Impre1 " _
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
    If Val(Vendedor.Text) <> 1 And Val(Vendedor.Text) <> 7 Then
        Listado.CopiesToPrinter = 4
            Else
        Listado.CopiesToPrinter = 3
    End If
    Rem Listado.Destination = 0
    
    If Letra.Text = "A" Then
        Listado.ReportFileName = "ImpreFacturaElectronicaA.rpt"
            Else
        Listado.ReportFileName = "ImpreFacturaElectronicaB.rpt"
    End If
    
    Listado.Action = 1

End Sub



