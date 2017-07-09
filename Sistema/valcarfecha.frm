VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgValcarFecha 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cheques por Fecha de Ingreso"
   ClientHeight    =   3165
   ClientLeft      =   3210
   ClientTop       =   720
   ClientWidth     =   6030
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   6030
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   5295
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
         Left            =   1440
         MouseIcon       =   "valcarfecha.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "valcarfecha.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Consulta de Datos"
         Top             =   1320
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
         Left            =   240
         MouseIcon       =   "valcarfecha.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "valcarfecha.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1320
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
         Left            =   2760
         MouseIcon       =   "valcarfecha.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "valcarfecha.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1320
         Width           =   855
      End
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
         MouseIcon       =   "valcarfecha.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "valcarfecha.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salida"
         Top             =   1320
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin MSMask.MaskEdBox Desdefec 
         Height          =   255
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5520
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "carteraFecha.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Valores en Cartera"
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
      Left            =   5400
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgValcarFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxiliar As String
Private WLinea As Single
Private WRecibo As String
Private WCheque As String
Private WBanco As String
Private WFecha As String
Private WImporte As Double
Private WFechaImpre As String
Private WVencimiento As String
Private Baja As Integer
Private WW As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim XParam As String
Dim XProveedor As String
Dim WVenci As String

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

    WDesde = Right$(Desdefec.Text, 4) + Mid$(Desdefec.Text, 4, 2) + Left$(Desdefec.Text, 2)
    WHasta = Right$(Hastafec.Text, 4) + Mid$(Hastafec.Text, 4, 2) + Left$(Hastafec.Text, 2)
    
    WTitulo = "Del " + Desdefec.Text + " al " + Hastafec.Text
    If Val(WEmpresa) = 1 Then
        ZVarios = "Cheques por Fecha de Ingreso"
            Else
        ZVarios = "*****   Cheques por Fecha de Ingreso   ********"
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Actividad = " + "'" + WTitulo + "',"
    ZSql = ZSql + " Varios = " + "'" + ZVarios + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Recibos SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Recibos.Clave, Recibos.Recibo, Recibos.Cliente, Recibos.FechaOrd, Recibos.TipoReg, Recibos.Tipo2, Recibos.Numero2, Recibos.Fecha2, Recibos.banco2, Recibos.Importe2, Recibos.Estado2, Recibos.FechaOrd2, Recibos.Orden, Recibos.Deposito, Recibos.SucursalCheque, Recibos.TipoCheque, Recibos.ClaseCheque, Recibos.ProveedorSalida, Recibos.BancoSalida, " _
            + "Auxiliar.Nombre, Auxiliar.Actividad, Auxiliar.Varios, " _
            + "Cliente.Razon " _
            + "From " _
            + DSQ + ".dbo.recibos Recibos, " _
            + DSQ + ".dbo.Auxiliar Auxiliar, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Recibos.CodigoEmpresa = Auxiliar.Empresa AND " _
            + "Recibos.Cliente = Cliente.Cliente AND " _
            + "Recibos.TipoReg = '2' AND " _
            + "Recibos.Tipo2 = '02' AND " _
            + "Recibos.Estado2 >= ' ' AND " _
            + "Recibos.Estado2 <= 'Z' AND " _
            + "Recibos.FechaOrd >= '" + WDesde + "' AND " _
            + "Recibos.FechaOrd <= '" + WHasta + "' AND " _
            + "Recibos.Orden <> 999999"

    Listado.Connect = Connect()
    
    Uno = "{Recibos.FechaOrd} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgValcarFecha.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desdefec.Text, Auxi)
        If Auxi = "S" Then
            Hastafec.SetFocus
                Else
            Desdefec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Desdefec.Text = "  /  /    "
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hastafec.Text, Auxi)
        If Auxi = "S" Then
            Desdefec.SetFocus
                Else
            Hastafec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hastafec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Desdefec.Text = "  /  /    "
    Hastafec.Text = "  /  /    "
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

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Panta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_Click
        Case Else
    End Select
End Sub

