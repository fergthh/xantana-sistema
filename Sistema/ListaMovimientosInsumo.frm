VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaMovimientosInsumos 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Movimientos de Insumos"
   ClientHeight    =   3120
   ClientLeft      =   1725
   ClientTop       =   1455
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3120
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2655
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton Cancela 
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
         Left            =   4080
         MouseIcon       =   "ListaMovimientosInsumo.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaMovimientosInsumo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salida"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Impre 
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
         Left            =   2280
         MouseIcon       =   "ListaMovimientosInsumo.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaMovimientosInsumo.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1440
         Width           =   855
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
         Left            =   480
         MouseIcon       =   "ListaMovimientosInsumo.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaMovimientosInsumo.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1440
         Width           =   855
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
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
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
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
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListPedCli.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
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
      Left            =   6960
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaMovimientosInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WSaldo As Double

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    Rem On Error GoTo WError
    
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Varios = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
        
    Listado.WindowTitle = "Listado de Movimientos de Insumos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
            
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT MovimientoInsumo.Tipo, MovimientoInsumo.Numero, MovimientoInsumo.Insumo, MovimientoInsumo.Cantidad, MovimientoInsumo.Fecha, MovimientoInsumo.OrdFecha, MovimientoInsumo.Deposito, MovimientoInsumo.DepositoII, " _
            + "Insumo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.MovimientoInsumo MovimientoInsumo, " _
            + DSQ + ".dbo.Insumo Insumo " _
            + "Where " _
            + "MovimientoInsumo.Insumo = Insumo.Codigo AND " _
            + "MovimientoInsumo.Insumo >= ' ' AND " _
            + "MovimientoInsumo.Insumo <= 'ZZZZZZZZZ' AND " _
            + "MovimientoInsumo.OrdFecha >= '" + WDesde + "' AND " _
            + "MovimientoInsumo.OrdFecha <= '" + WHasta + "'"
    
    Listado.Connect = Connect()
            
    Uno = "{MovimientoInsumo.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {MovimientoInsumo.Insumo} in " + Chr$(34) + "" + Chr$(34) + " to " + Chr$(34) + "ZZZZZ" + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.ReportFileName = "ListaMovimientosInsumo.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaMovimientosInsumos.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFec.Text = "  /  /    "
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Frame2.Visible = True
End Sub


Private Sub DesdeFec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Panta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub













