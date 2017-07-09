VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgDespacho 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Despacho"
   ClientHeight    =   7380
   ClientLeft      =   1125
   ClientTop       =   750
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   9750
   Begin VB.TextBox Importador 
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
      MaxLength       =   50
      TabIndex        =   37
      Text            =   " "
      Top             =   2640
      Width           =   4935
   End
   Begin VB.TextBox Puerto 
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
      MaxLength       =   50
      TabIndex        =   35
      Text            =   " "
      Top             =   2280
      Width           =   4935
   End
   Begin VB.TextBox Aduana 
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
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   1920
      Width           =   4935
   End
   Begin VB.TextBox Posicion 
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
      MaxLength       =   50
      TabIndex        =   31
      Text            =   " "
      Top             =   1560
      Width           =   4935
   End
   Begin VB.TextBox Origen 
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
      MaxLength       =   50
      TabIndex        =   29
      Text            =   " "
      Top             =   1200
      Width           =   4935
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
      MouseIcon       =   "Despacho.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Elimina el Registro"
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Consulta de Datos"
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Impresion "
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Salida"
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":46D2
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Primer Registro"
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":4B14
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":4E1E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Registro Anterior"
      Top             =   3120
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
      MouseIcon       =   "Despacho.frx":5260
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":556A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Registro Siguiente"
      Top             =   3120
      Width           =   855
   End
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
      MouseIcon       =   "Despacho.frx":59AC
      MousePointer    =   99  'Custom
      Picture         =   "Despacho.frx":5CB6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Salida"
      Top             =   3120
      Width           =   855
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
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Descripcion 
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
      MaxLength       =   50
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox Codigo 
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
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1920
      TabIndex        =   7
      Top             =   4200
      Visible         =   0   'False
      Width           =   5535
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
         Left            =   3360
         MouseIcon       =   "Despacho.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "Despacho.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   4440
         MouseIcon       =   "Despacho.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "Despacho.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Hasta 
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   13
         Text            =   " "
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Desde 
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   360
         Width           =   855
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
         Left            =   1800
         TabIndex        =   11
         Top             =   1320
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
         Left            =   240
         TabIndex        =   10
         Top             =   1320
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
         TabIndex        =   9
         Top             =   720
         Width           =   1455
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
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Despacho.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Bancos"
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
      Left            =   5760
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Numero 
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
      MaxLength       =   20
      TabIndex        =   1
      Top             =   840
      Width           =   3615
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
      Height          =   2460
      ItemData        =   "Despacho.frx":6F90
      Left            =   120
      List            =   "Despacho.frx":6F97
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   8175
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
      Left            =   1320
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label8 
      Caption         =   "Importador"
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
      TabIndex        =   38
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Puerto / Aeropuerto"
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
      TabIndex        =   36
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Aduana"
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
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Posicion"
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
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Origen"
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
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Descripcion"
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
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro. de Despacho"
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
      TabIndex        =   4
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Despacho"
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
      Width           =   2295
   End
End
Attribute VB_Name = "PrgDespacho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Imprime_Nombre()
End Sub

Sub Verifica_datos()
    Rem If Val(descripcion.text) = 0 Then
    Rem      descripcion.text = "0"
    Rem End If
End Sub

Sub Format_datos()
    Rem descripcion.text = Pusing("###,###.##", descripcion.text)
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Despacho"
    ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Codigo.Text + "'"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        Numero.Text = Trim(rstDespacho!Numero)
        Descripcion.Text = Trim(rstDespacho!Descripcion)
        Origen.Text = Trim(rstDespacho!Origen)
        Posicion.Text = Trim(rstDespacho!Posicion)
        Aduana.Text = Trim(rstDespacho!Aduana)
        Puerto.Text = Trim(rstDespacho!Puerto)
        Importador.Text = Trim(rstDespacho!Importador)
        rstDespacho.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
End Sub

Private Sub Acepta_Click()

    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Despacho SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Despachos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Despacho.Codigo, Despacho.Numero, Despacho.Descripcion, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Despacho Despacho, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Despacho.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Despacho.Codigo >= " + Desde.Text + " AND " _
                + "Despacho.Codigo <= " + Hasta.Text
    
    Listado.GroupSelectionFormula = "{Despacho.Codigo} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{Despacho.Codigo} in " + Desde.Text + " to " + Hasta.Text
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
     Codigo.SetFocus
End Sub

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Despacho"
        ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Codigo.Text + "'"
        spDespacho = ZSql
        Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
        If rstDespacho.RecordCount > 0 Then
            rstDespacho.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Despacho SET "
            ZSql = ZSql + " Numero = " + "'" + Numero.Text + "',"
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
            ZSql = ZSql + " Origen = " + "'" + Origen.Text + "',"
            ZSql = ZSql + " Posicion = " + "'" + Posicion.Text + "',"
            ZSql = ZSql + " Aduana = " + "'" + Aduana.Text + "',"
            ZSql = ZSql + " Puerto = " + "'" + Puerto.Text + "',"
            ZSql = ZSql + " Importador = " + "'" + Importador.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Despacho ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Posicion ,"
            ZSql = ZSql + "Aduana ,"
            ZSql = ZSql + "Puerto ,"
            ZSql = ZSql + "Importador )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + Origen.Text + "',"
            ZSql = ZSql + "'" + Posicion.Text + "',"
            ZSql = ZSql + "'" + Aduana.Text + "',"
            ZSql = ZSql + "'" + Puerto.Text + "',"
            ZSql = ZSql + "'" + Importador.Text + "')"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
    End If
End Sub

Private Sub CmdDelete_Click()
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Despacho"
        ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Codigo.Text + "'"
        spDespacho = ZSql
        Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
        If rstDespacho.RecordCount > 0 Then
            rstDespacho.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
    
                ZSql = ""
                ZSql = ZSql + "DELETE Despacho"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spDespacho = ZSql
                Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Numero.Text = ""
    Descripcion.Text = ""
    Origen.Text = ""
    Posicion.Text = ""
    Aduana.Text = ""
    Puerto.Text = ""
    Importador.Text = ""

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Despacho"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        rstDespacho.MoveLast
        ZUltimo = IIf(IsNull(rstDespacho!CodigoMayor), "0", rstDespacho!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstDespacho.Close
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgDespacho.Hide
    Unload Me
    Menu3.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Origen.SetFocus
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
End Sub

Private Sub Origen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Posicion.SetFocus
    End If
    If KeyAscii = 27 Then
        Origen.Text = ""
    End If
End Sub

Private Sub Posicion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Aduana.SetFocus
    End If
    If KeyAscii = 27 Then
        Posicion.Text = ""
    End If
End Sub

Private Sub Aduana_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Puerto.SetFocus
    End If
    If KeyAscii = 27 Then
        Aduana.Text = ""
    End If
End Sub

Private Sub Puerto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Importador.SetFocus
    End If
    If KeyAscii = 27 Then
        Puerto.Text = ""
    End If
End Sub

Private Sub Importador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Importador.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Despacho"
            ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Codigo.Text + "'"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
            If rstDespacho.RecordCount > 0 Then
                rstDespacho.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Despachos"

     Rem Opcion.Visible = True
     Opcion.ListIndex = 0
     Call Opcion_Click
     
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
            ZSql = ZSql + " FROM Despacho"
            ZSql = ZSql + " Order by Despacho.Codigo"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
            If rstDespacho.RecordCount > 0 Then
                With rstDespacho
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstDespacho.Close
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
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    Codigo.Text = ""
    Numero.Text = ""
    Descripcion.Text = ""
    Origen.Text = ""
    Posicion.Text = ""
    Aduana.Text = ""
    Puerto.Text = ""
    Importador.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Despacho"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        rstDespacho.MoveLast
        ZUltimo = IIf(IsNull(rstDespacho!CodigoMayor), "0", rstDespacho!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstDespacho.Close
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
            ZSql = ZSql + " FROM Despacho"
            ZSql = ZSql + " Where Despacho.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Despacho.Codigo"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
            If rstDespacho.RecordCount > 0 Then
                With rstDespacho
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstDespacho.Close
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

Private Sub Codigo_DblClick()

    Opcion.Clear
    Opcion.AddItem "Despachos"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Numer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Origen_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Posicion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Aduana_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Puerto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Importador_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call CmdDelete_Click
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
            Call cmdClose_Click
        Case 122
            Call Acepta_Click
        Case 123
            Call Cancela_click
        Case Else
    End Select
End Sub



Private Sub Anterior_Click()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Despacho"
    ZSql = ZSql + " Where Despacho.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Despacho.Codigo"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        With rstDespacho
            .MoveLast
            Codigo.Text = rstDespacho!Codigo
        End With
        rstDespacho.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de Despachos")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM Despacho"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        rstDespacho.MoveFirst
        ZUltimo = IIf(IsNull(rstDespacho!CodigoMenor), "0", rstDespacho!CodigoMenor)
        Codigo.Text = ZUltimo
        rstDespacho.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
        ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Despacho"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        rstDespacho.MoveLast
        ZUltimo = IIf(IsNull(rstDespacho!CodigoMayor), "0", rstDespacho!CodigoMayor)
        Codigo.Text = ZUltimo
        rstDespacho.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Despacho"
    ZSql = ZSql + " Where Despacho.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Despacho.Codigo"
    spDespacho = ZSql
    Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
    If rstDespacho.RecordCount > 0 Then
        With rstDespacho
            .MoveFirst
            Codigo.Text = rstDespacho!Codigo
        End With
        rstDespacho.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Despachos")
    End If
End Sub

















