VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgValcarNumero 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cheques por Numero de Recibo"
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
      TabIndex        =   3
      Top             =   120
      Width           =   5295
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
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   1
         Text            =   " "
         Top             =   840
         Width           =   975
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
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   975
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
         Left            =   1440
         MouseIcon       =   "ValcarNumero.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ValcarNumero.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
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
         MouseIcon       =   "ValcarNumero.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ValcarNumero.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   6
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
         MouseIcon       =   "ValcarNumero.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ValcarNumero.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   5
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
         MouseIcon       =   "ValcarNumero.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ValcarNumero.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salida"
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Hasta Recibo"
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
         Left            =   960
         TabIndex        =   9
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Desde Recibo"
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
         Left            =   960
         TabIndex        =   8
         Top             =   420
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5520
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "carteraNumero.rpt"
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
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgValcarNumero"
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
Dim ZDesde As String
Dim ZHasta As String

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
    
    WTitulo = ""
    If Val(WEmpresa) = 1 Then
        ZVarios = "Cheques por Numero de Recibo"
            Else
        ZVarios = "*****   Cheques por Numero de Recibo   ********"
    End If
    
    ZDesde = Desde.Text
    ZHasta = Hasta.Text
    Call Ceros(ZDesde, 6)
    Call Ceros(ZHasta, 6)
    
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
            + "Recibos.Recibo >= '" + ZDesde + "' AND " _
            + "Recibos.Recibo <= '" + ZHasta + "'"

    Listado.Connect = Connect()
    
    Uno = "{Recibos.Recibo} in " + Chr$(34) + ZDesde + Chr$(34) + " to " + Chr$(34) + ZHasta + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgValcarNumero.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Frame2.Visible = True
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
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

