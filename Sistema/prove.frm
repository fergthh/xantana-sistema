VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form PrgProve 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Proveedores"
   ClientHeight    =   8370
   ClientLeft      =   1140
   ClientTop       =   300
   ClientWidth     =   9855
   LinkTopic       =   "Form2"
   ScaleHeight     =   8370
   ScaleWidth      =   9855
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
      Left            =   6840
      MouseIcon       =   "prove.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Salida"
      Top             =   3840
      Width           =   855
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
      Left            =   5880
      MouseIcon       =   "prove.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Registro Siguiente"
      Top             =   3840
      Width           =   855
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
      Left            =   4920
      MouseIcon       =   "prove.frx":0E98
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Registro Anterior"
      Top             =   3840
      Width           =   855
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
      Left            =   3960
      MouseIcon       =   "prove.frx":15E4
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":18EE
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Primer Registro"
      Top             =   3840
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
      Left            =   8760
      MouseIcon       =   "prove.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":203A
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Salida"
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Lista 
      Caption         =   "Listado F9"
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
      Left            =   7800
      MouseIcon       =   "prove.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":2B86
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Impresion "
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
      Left            =   3000
      MouseIcon       =   "prove.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":36D2
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Consulta de Datos"
      Top             =   3840
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
      Left            =   2040
      MouseIcon       =   "prove.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":421E
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3840
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
      Left            =   1080
      MouseIcon       =   "prove.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":4D6A
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Elimina el Registro"
      Top             =   3840
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
      Left            =   120
      MouseIcon       =   "prove.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "prove.frx":58B6
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1440
      TabIndex        =   18
      Top             =   5040
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Acepta 
         Caption         =   "Confirma F11"
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
         Left            =   3960
         MouseIcon       =   "prove.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "prove.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela F12"
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
         Left            =   5040
         MouseIcon       =   "prove.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "prove.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Hasta 
         BeginProperty Font 
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
         MaxLength       =   8
         TabIndex        =   24
         Text            =   " "
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Desde 
         BeginProperty Font 
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
         MaxLength       =   8
         TabIndex        =   23
         Text            =   " "
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   2280
         TabIndex        =   22
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
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
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
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
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
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
      Left            =   120
      TabIndex        =   38
      Top             =   5040
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.TextBox NombreCheque 
      BeginProperty Font 
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
      MaxLength       =   50
      TabIndex        =   13
      Text            =   " "
      Top             =   3360
      Width           =   6975
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
      Height          =   1980
      Left            =   720
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox Iva 
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
      Left            =   6600
      TabIndex        =   12
      Text            =   " "
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox Ganancia 
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
      Left            =   2280
      TabIndex        =   11
      Text            =   " "
      Top             =   3000
      Width           =   2415
   End
   Begin VB.ComboBox Provincia 
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
      Left            =   2280
      TabIndex        =   4
      Text            =   " "
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Dias 
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
      Left            =   6600
      MaxLength       =   4
      TabIndex        =   7
      Text            =   " "
      Top             =   1920
      Width           =   855
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   10
      Text            =   " "
      Top             =   2640
      Width           =   5535
   End
   Begin VB.TextBox EMail 
      BeginProperty Font 
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
      MaxLength       =   30
      TabIndex        =   8
      Text            =   " "
      Top             =   2280
      Width           =   2655
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
      Height          =   285
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   6
      Text            =   " "
      Top             =   1920
      Width           =   2655
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
      Left            =   6600
      MaxLength       =   15
      TabIndex        =   9
      Text            =   " "
      Top             =   2280
      Width           =   2655
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
      Left            =   6600
      MaxLength       =   20
      TabIndex        =   5
      Text            =   " "
      Top             =   1560
      Width           =   2655
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Text            =   " "
      Top             =   1200
      Width           =   5535
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   840
      Width           =   5535
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
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1455
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8400
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "prove.rpt"
      Destination     =   1
      WindowTitle     =   "Listados de Proveedores"
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
      Left            =   8040
      TabIndex        =   17
      Top             =   1200
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
      Height          =   2220
      ItemData        =   "prove.frx":6F90
      Left            =   120
      List            =   "prove.frx":6F97
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   9135
   End
   Begin VB.TextBox Nombre 
      BeginProperty Font 
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
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label15 
      Caption         =   "Cheque a nombre de "
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
      TabIndex        =   37
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label Label13 
      Caption         =   "Codigo de Iva"
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
      Left            =   5160
      TabIndex        =   35
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label12 
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
      Left            =   120
      TabIndex        =   34
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label11 
      Caption         =   "Dias de Plazo"
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
      Left            =   5160
      TabIndex        =   33
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Categoria Ganancia"
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
      TabIndex        =   32
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label9 
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
      TabIndex        =   31
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "E Mail"
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
      TabIndex        =   30
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label7 
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
      Left            =   120
      TabIndex        =   29
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Cuit"
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
      Left            =   5160
      TabIndex        =   28
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   5160
      TabIndex        =   27
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   840
      Width           =   1815
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
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Proveedor"
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
      TabIndex        =   14
      Top             =   180
      Width           =   2055
   End
End
Attribute VB_Name = "PrgProve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Dim ResultadoCuit As String

Sub Verifica_datos()
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        Nombre.Text = Trim(rstProveedor!Nombre)
        Direccion.Text = Trim(rstProveedor!Direccion)
        Localidad.Text = Trim(rstProveedor!Localidad)
        Postal.Text = Trim(rstProveedor!Postal)
        Cuit.Text = Trim(rstProveedor!Cuit)
        Telefono.Text = Trim(rstProveedor!Telefono)
        email.Text = Trim(rstProveedor!email)
        Observaciones.Text = Trim(rstProveedor!Observaciones)
        Dias.Text = rstProveedor!Dias
        Iva.ListIndex = rstProveedor!Iva
        Ganancia.ListIndex = rstProveedor!Ganancia
        Provincia.ListIndex = rstProveedor!Provincia
        NombreCheque.Text = Trim(rstProveedor!NombreCheque)
        rstProveedor.Close
        Call Format_datos
    End If
    
End Sub

Sub Imprime_Descripcion()
End Sub

Private Sub Acepta_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Proveedor SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + "1" + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)

    Listado.WindowTitle = "Listado de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Trim(Desde.Text) = "" Then
         Desde.Text = ""
    End If
    If Trim(Hasta.Text) = "" Then
         Hasta.Text = "ZZZZ"
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Proveedor.Proveedor, Proveedor.Nombre, Proveedor.Direccion, Proveedor.Telefono, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Proveedor Proveedor, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Proveedor.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Proveedor.Proveedor >= '" + Desde.Text + "' AND " _
                + "Proveedor.Proveedor <= '" + Hasta.Text + "'"
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Proveedor.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Listado.SelectionFormula = "{Proveedor.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Proveedor.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Proveedor.Text <> "" Then
    
    
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            rstProveedor.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "',"
            ZSql = ZSql + " Direccion = " + "'" + Direccion.Text + "',"
            ZSql = ZSql + " Localidad = " + "'" + Localidad.Text + "',"
            ZSql = ZSql + " Postal = " + "'" + Postal.Text + "',"
            ZSql = ZSql + " Cuit = " + "'" + Cuit.Text + "',"
            ZSql = ZSql + " Telefono = " + "'" + Telefono.Text + "',"
            ZSql = ZSql + " EMail = " + "'" + email.Text + "',"
            ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
            ZSql = ZSql + " Dias = " + "'" + Dias.Text + "',"
            ZSql = ZSql + " Ganancia = " + "'" + Str$(Ganancia.ListIndex) + "',"
            ZSql = ZSql + " Iva = " + "'" + Str$(Iva.ListIndex) + "',"
            ZSql = ZSql + " Provincia = " + "'" + Str$(Provincia.ListIndex) + "',"
            ZSql = ZSql + " NombreCheque = " + "'" + NombreCheque.Text + "'"
            ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Proveedor ("
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Postal ,"
            ZSql = ZSql + "Cuit ,"
            ZSql = ZSql + "Telefono ,"
            ZSql = ZSql + "Email ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Dias ,"
            ZSql = ZSql + "Ganancia ,"
            ZSql = ZSql + "Iva ,"
            ZSql = ZSql + "Provincia ,"
            ZSql = ZSql + "NombreCheque )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Proveedor.Text + "',"
            ZSql = ZSql + "'" + Nombre.Text + "',"
            ZSql = ZSql + "'" + Direccion.Text + "',"
            ZSql = ZSql + "'" + Localidad.Text + "',"
            ZSql = ZSql + "'" + Postal.Text + "',"
            ZSql = ZSql + "'" + Cuit.Text + "',"
            ZSql = ZSql + "'" + Telefono.Text + "',"
            ZSql = ZSql + "'" + email.Text + "',"
            ZSql = ZSql + "'" + Observaciones.Text + "',"
            ZSql = ZSql + "'" + Dias.Text + "',"
            ZSql = ZSql + "'" + Str$(Ganancia.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Iva.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Provincia.ListIndex) + "',"
            ZSql = ZSql + "'" + NombreCheque.Text + "')"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
    
        
        Rem Call CmdLimpiar_Click
    
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Proveedores")
    
        
        Proveedor.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Proveedor.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            rstProveedor.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE Proveedor"
                ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
                
            End If
        End If
        
    
        If ZZNivel = 0 Then
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            rstProveedor.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE Proveedor"
                ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
                
            End If
        End If
        
    
        If ZZNivel = 0 Then
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        
        
        
    End If
    Proveedor.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Proveedor.Text = ""
    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Cuit.Text = ""
    Telefono.Text = ""
    email.Text = ""
    Observaciones.Text = ""
    Dias.Text = ""
    NombreCheque = ""
    Iva.ListIndex = 0
    Ganancia.ListIndex = 0
    Provincia.ListIndex = 0
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Proveedor) as [ProveedorMayor]"
    Rem ZSql = ZSql + " FROM Proveedor"
    Rem spProveedor = ZSql
    Rem Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstProveedor.RecordCount > 0 Then
    Rem     rstProveedor.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstProveedor!ProveedorMayor), "0", rstProveedor!ProveedorMayor)
    Rem     Proveedor.Text = ZUltimo + 1
    Rem     rstProveedor.Close
    Rem End If
    
    Proveedor.SetFocus
    
End Sub

Private Sub CmdClose_Click()
    PrgProve.Hide
    Unload Me
    If ZZSistema = 1 Then
        MenuAdminis.Show
            Else
        MenuVen.Show
    End If
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Direccion.SetFocus
    End If
    If KeyAscii = 27 Then
        Nombre.Text = ""
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
        Provincia.SetFocus
    End If
    If KeyAscii = 27 Then
        Localidad.Text = ""
    End If
End Sub

Private Sub Provincia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Postal.SetFocus
    End If
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Postal_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Telefono.SetFocus
    End If
    If KeyAscii = 27 Then
        Postal.Text = ""
    End If
End Sub

Private Sub Telefono_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dias.SetFocus
    End If
    If KeyAscii = 27 Then
        Telefono.Text = ""
    End If
End Sub

Private Sub Dias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        email.SetFocus
    End If
    If KeyAscii = 27 Then
        Dias.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub EMail_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuit.SetFocus
    End If
    If KeyAscii = 27 Then
        email.Text = ""
    End If
End Sub

Private Sub Cuit_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuit.Text <> "" Then
            Rem Call verifica_cuit(Cuit.Text, ResultadoCuit)
            ResultadoCuit = "S"
            If ResultadoCuit = "S" Then
                Observaciones.SetFocus
                    Else
                Cuit.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Cuit.Text = ""
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ganancia.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Ganancia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Iva.SetFocus
    End If
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        NombreCheque.SetFocus
    End If
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NombreCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        NombreCheque.Text = ""
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Proveedor.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                rstProveedor.Close
                Call Imprime_Datos
                    Else
                WProveedor = Proveedor.Text
                CmdLimpiar_Click
                Proveedor.Text = WProveedor
            End If
        End If
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    Opcion.Clear
    
    Opcion.AddItem "Proveedores"

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
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
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
            Indice = Pantalla.ListIndex
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
        
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    Proveedor.Text = ""
    Nombre.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Cuit.Text = ""
    Telefono.Text = ""
    email.Text = ""
    Observaciones.Text = ""
    Dias.Text = ""
    NombreCheque.Text = ""
    
    Iva.Clear
    
    Iva.AddItem ""
    Iva.AddItem "No Inscripto"
    Iva.AddItem "Consumidor Final"
    Iva.AddItem "Resp.Inscripto"
    Iva.AddItem "Exento"
    Iva.AddItem "No Responsable"
    Iva.AddItem "Monotributo"
    Iva.AddItem "No Catalogado"
    
    Provincia.Clear
    
    Provincia.AddItem ""
    Provincia.AddItem "Capital Federal"
    Provincia.AddItem "Buenos Aires"
    Provincia.AddItem "Catamarca"
    Provincia.AddItem "Cordoba"
    Provincia.AddItem "Corrientes"
    Provincia.AddItem "Chaco"
    Provincia.AddItem "Chubut"
    Provincia.AddItem "Entre Rios"
    Provincia.AddItem "Formosa"
    Provincia.AddItem "Jujuy"
    Provincia.AddItem "La Pampa"
    Provincia.AddItem "La Rioja"
    Provincia.AddItem "Mendoza"
    Provincia.AddItem "Misiones"
    Provincia.AddItem "Neuquen"
    Provincia.AddItem "Rio Negro"
    Provincia.AddItem "Salta"
    Provincia.AddItem "San Juan"
    Provincia.AddItem "San Luis"
    Provincia.AddItem "Santa Cruz"
    Provincia.AddItem "Santa Fe"
    Provincia.AddItem "Santiago del Estero"
    Provincia.AddItem "Tucuman"
    Provincia.AddItem "Tierra del Fuego"
    Provincia.AddItem "Exterior"
    
    Ganancia.Clear
    
    Ganancia.AddItem ""
    Ganancia.AddItem "Bienes"
    Ganancia.AddItem "Honorarios"
    Ganancia.AddItem "Alquileres"
    Ganancia.AddItem "Exento"
    Ganancia.AddItem "Servicios"
    
    Iva.ListIndex = 0
    Ganancia.ListIndex = 0
    Provincia.ListIndex = 0
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Proveedor) as [ProveedorMayor]"
    Rem ZSql = ZSql + " FROM Proveedor"
    Rem spProveedor = ZSql
    Rem Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstProveedor.RecordCount > 0 Then
    Rem     rstProveedor.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstProveedor!ProveedorMayor), "0", rstProveedor!ProveedorMayor)
    Rem     Proveedor.Text = ZUltimo + 1
    Rem     rstProveedor.Close
    Rem End If

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
            
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedor"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Direccion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Localidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Provincia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Postal_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Telefono_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Dias_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Email_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cuit_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Ganancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Iva_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub NombreCheque_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 120
            Call Lista_Click
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
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor < " + "'" + Proveedor.Text + "'"
    ZSql = ZSql + " Order by Proveedor.Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        With rstProveedor
            .MoveLast
            Proveedor.Text = Trim(rstProveedor!Proveedor)
        End With
        rstProveedor.Close
        Call Imprime_Datos
        Proveedor.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Proveedores")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Proveedor) as [ProveedorMenor]"
    ZSql = ZSql + " FROM Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        rstProveedor.MoveFirst
        ZUltimo = IIf(IsNull(rstProveedor!ProveedorMenor), "", rstProveedor!ProveedorMenor)
        Proveedor.Text = Trim(ZUltimo)
        rstProveedor.Close
        Call Imprime_Datos
        Proveedor.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Proveedor) as [ProveedorMayor]"
    ZSql = ZSql + " FROM Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        rstProveedor.MoveLast
        ZUltimo = IIf(IsNull(rstProveedor!ProveedorMayor), "", rstProveedor!ProveedorMayor)
        Proveedor.Text = Trim(ZUltimo)
        rstProveedor.Close
        Call Imprime_Datos
        Proveedor.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor > " + "'" + Proveedor.Text + "'"
    ZSql = ZSql + " Order by Proveedor.Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        With rstProveedor
            .MoveFirst
            Proveedor.Text = Trim(rstProveedor!Proveedor)
        End With
        rstProveedor.Close
        Call Imprime_Datos
        Proveedor.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Proveedores")
    End If
End Sub









