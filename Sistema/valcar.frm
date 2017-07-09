VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgValcar 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cheques en Cartera"
   ClientHeight    =   3885
   ClientLeft      =   3315
   ClientTop       =   1575
   ClientWidth     =   6615
   LinkTopic       =   "Form2"
   ScaleHeight     =   3885
   ScaleWidth      =   6615
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
      Height          =   3615
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Tipo 
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
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
         Width           =   2415
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
         Left            =   120
         MouseIcon       =   "valcar.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "valcar.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2160
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
         Left            =   2040
         MouseIcon       =   "valcar.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "valcar.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2160
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
         Left            =   3840
         MouseIcon       =   "valcar.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "valcar.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salida"
         Top             =   2160
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
      Begin VB.Label Label1 
         Caption         =   "Tipo Listado"
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
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
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
      Left            =   5880
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "cartera.rpt"
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
      Left            =   6000
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgValcar"
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
Dim XCliente As String
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
        ZVarios = "Valores por Fecha"
            Else
        ZVarios = "*****   Valores por Fecha  ********"
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
    ZSql = ZSql + " CodigoEmpresa = " + "'" + "1" + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
    Select Case Tipo.ListIndex
        Case 0
            WDesde1 = "0"
            WHasta1 = "0"
            
            WDesde2 = "0"
            WHasta2 = "0"
        Case 1
            WDesde1 = "1"
            WHasta1 = "999999"
            
            WDesde2 = "0"
            WHasta2 = "0"
        Case Else
            WDesde1 = "0"
            WHasta1 = "0"
            
            WDesde2 = "1"
            WHasta2 = "999999"
    End Select
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Cliente, Recibos.Fecha, Recibos.TipoReg, Recibos.Tipo2, Recibos.Numero2, Recibos.Fecha2, Recibos.banco2, Recibos.Importe2, Recibos.Estado2, Recibos.FechaOrd2, Recibos.Orden, Recibos.Deposito, Recibos.TipoCheque, Recibos.ClaseCheque, Recibos.Periodo, Recibos.Destino, " _
            + "Cliente.Razon, " _
            + "Auxiliar.Nombre, Auxiliar.Actividad, Auxiliar.Varios " _
            + "From " _
            + DSQ + ".dbo.Recibos Recibos, " _
            + DSQ + ".dbo.Cliente Cliente, " _
            + DSQ + ".dbo.Auxiliar Auxiliar " _
            + "Where " _
            + "Recibos.Cliente = Cliente.Cliente AND " _
            + "Recibos.CodigoEmpresa = Auxiliar.Empresa AND " _
            + "Recibos.TipoReg = '2' AND " _
            + "Recibos.Tipo2 = '02' AND " _
            + "Recibos.FechaOrd2 >= '" + WDesde + "' AND " _
            + "Recibos.FechaOrd2 <= '" + WHasta + "' AND " _
            + "Recibos.Orden >= " + WDesde1 + " AND " _
            + "Recibos.Orden <= " + WHasta1 + " AND " _
            + "Recibos.Deposito >= " + WDesde2 + " AND " _
            + "Recibos.Deposito <= " + WHasta2
    
    Listado.Connect = Connect()
    
    If Val(WEmpresa) = 1 Then
        Listado.ReportFileName = "Cartera.rpt"
            Else
        Listado.ReportFileName = "CarteraR.rpt"
    End If
    
    Uno = "{Recibos.FechaOrd2} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Recibos.Orden} in " + WDesde1 + " to " + WHasta1
    Tres = " and {Recibos.Deposito} in " + WDesde2 + " to " + WHasta2
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgValcar.Hide
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

    Tipo.Clear
    
    Tipo.AddItem "En cartera"
    Tipo.AddItem "Entregado"
    Tipo.AddItem "Depositado"
    
    Tipo.ListIndex = 0

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








