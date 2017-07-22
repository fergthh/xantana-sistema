VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaSolicitudInsumo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Solicitud de Insumos"
   ClientHeight    =   4140
   ClientLeft      =   1725
   ClientTop       =   1455
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4140
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   4815
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
         MouseIcon       =   "ListaSolicitudInsumos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaSolicitudInsumos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salida"
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
         MouseIcon       =   "ListaSolicitudInsumos.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaSolicitudInsumos.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2160
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
         MouseIcon       =   "ListaSolicitudInsumos.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaSolicitudInsumos.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2160
         Width           =   855
      End
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
         Left            =   2400
         TabIndex        =   7
         Top             =   1560
         Width           =   2055
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   960
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
         Left            =   2400
         TabIndex        =   4
         Top             =   480
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
      Begin VB.Label Label5 
         Caption         =   "Tipo"
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
         Left            =   840
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
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
         Left            =   840
         TabIndex        =   3
         Top             =   960
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
         Left            =   840
         TabIndex        =   2
         Top             =   480
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
Attribute VB_Name = "PrgListaSolicitudInsumo"
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
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Solicitud SET "
    ZSql = ZSql + " Saldo = Cantidad - Pedido - Ajuste"
    spSolicitud = ZSql
    Set rstSolicitud = db.OpenRecordset(spSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        
    Listado.WindowTitle = "Listado de Solicitud de Ordenes de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
            
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If Tipo.ListIndex = 0 Then
        
        Listado.SQLQuery = "SELECT Solicitud.Numero, Solicitud.Renglon, Solicitud.Fecha, Solicitud.OrdFecha, Solicitud.Insumo, Solicitud.Cantidad, Solicitud.Observaciones, Solicitud.Pedido, " _
                + "Insumo.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Solicitud Solicitud, " _
                + DSQ + ".dbo.Insumo Insumo " _
                + "Where " _
                + "Solicitud.Insumo = Insumo.Codigo AND " _
                + "Solicitud.OrdFecha >= '" + WDesde + "' AND " _
                + "Solicitud.OrdFecha <= '" + WHasta + "' AND " _
                + "Solicitud.Saldo > 0"
                
        Listado.Connect = Connect()
        
        Uno = "{Solicitud.Saldo} > 0"
        Dos = " and {Solicitud.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
        
        Listado.ReportFileName = "ListaSolicitudInsumoPend.rpt"
    
            Else
            
        Listado.SQLQuery = "SELECT Solicitud.Numero, Solicitud.Renglon, Solicitud.Fecha, Solicitud.OrdFecha, Solicitud.Insumo, Solicitud.Cantidad, Solicitud.Observaciones, Solicitud.Pedido, " _
                + "Insumo.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Solicitud Solicitud, " _
                + DSQ + ".dbo.Insumo Insumo " _
                + "Where " _
                + "Solicitud.Insumo = Insumo.Codigo AND " _
                + "Solicitud.OrdFecha >= '" + WDesde + "' AND " _
                + "Solicitud.OrdFecha <= '" + WHasta + "'"
                
        Listado.Connect = Connect()
        
        Uno = "{Solicitud.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        
        Listado.GroupSelectionFormula = Uno
        Listado.SelectionFormula = Uno
        
        Listado.ReportFileName = "ListaSolicitudInsumo.rpt"
            
    End If

    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaSolicitudInsumo.Hide
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
            DesdeFec.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Pendiente"
    Tipo.AddItem "Completo"
    
    Tipo.ListIndex = 0

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












