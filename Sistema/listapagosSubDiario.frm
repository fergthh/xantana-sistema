VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaPagosSubDiario 
   AutoRedraw      =   -1  'True
   Caption         =   "SubDiario de Pagos"
   ClientHeight    =   5865
   ClientLeft      =   2655
   ClientTop       =   855
   ClientWidth     =   6900
   LinkTopic       =   "Form2"
   ScaleHeight     =   5865
   ScaleWidth      =   6900
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
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   5295
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
         MouseIcon       =   "listapagosSubDiario.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "listapagosSubDiario.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1200
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
         Left            =   2880
         MouseIcon       =   "listapagosSubDiario.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "listapagosSubDiario.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consulta de Datos"
         Top             =   1200
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
         Left            =   1560
         MouseIcon       =   "listapagosSubDiario.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "listapagosSubDiario.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1200
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
         Left            =   4200
         MouseIcon       =   "listapagosSubDiario.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "listapagosSubDiario.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salida"
         Top             =   1200
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin VB.Label Label2 
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
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaPagosSubDiario.rpt"
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
      Left            =   5760
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      ItemData        =   "listapagosSubDiario.frx":2D30
      Left            =   360
      List            =   "listapagosSubDiario.frx":2D37
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   6375
   End
End
Attribute VB_Name = "PrgListaPagosSubDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZMes(100) As String

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    WTitulo = ZMes(Val(Mid$(Hasta.Text, 4, 2))) + " de " + Right$(Hasta.Text, 4)
    ZVarios = "S U B D I A R I O    D E    P A G O S"
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Actividad = " + "'" + WTitulo + "',"
    ZSql = ZSql + " Varios = " + "'" + ZVarios + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Pagos SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Pagos por Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Pagos.Orden, Pagos.Proveedor, Pagos.Fecha, Pagos.FechaOrd, Pagos.Tipo2, Pagos.Fecha2, Pagos.Importe2, " _
            + "Auxiliar.Nombre, Auxiliar.Actividad, Auxiliar.Varios, " _
            + "Proveedor.Nombre, Proveedor.Cuit " _
            + "From " _
            + DSQ + ".dbo.Pagos Pagos, " _
            + DSQ + ".dbo.Auxiliar Auxiliar, " _
            + DSQ + ".dbo.Proveedor Proveedor " _
            + "Where " _
            + "Pagos.CodigoEmpresa = Auxiliar.Empresa AND " _
            + "Pagos.Proveedor = Proveedor.Proveedor AND " _
            + "Pagos.FechaOrd >= '" + WDesde + "' AND " _
            + "Pagos.FechaOrd <= '" + WHasta + "' AND " _
            + "Pagos.Importe2 <> 0"
    
    Listado.Connect = Connect()
    
    Uno = "{Pagos.FechaOrd} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34) + " and "
    Dos = "{Pagos.Importe2} <> 0"
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Action = 1
    
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgListaPagosSubDiario.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = "  /  /    "
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    ZMes(1) = "Enero"
    ZMes(2) = "Febrero"
    ZMes(3) = "Marzo"
    ZMes(4) = "Abril"
    ZMes(5) = "Mayo"
    ZMes(6) = "Junio"
    ZMes(7) = "Julio"
    ZMes(8) = "Agosto"
    ZMes(9) = "Septiembre"
    ZMes(10) = "Octubre"
    ZMes(11) = "Noviembre"
    ZMes(12) = "Diciembre"


    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
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

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
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




