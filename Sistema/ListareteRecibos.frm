VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListareterecibos 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Retenciones Recibidas"
   ClientHeight    =   4785
   ClientLeft      =   2790
   ClientTop       =   1395
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   4815
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
         Left            =   2160
         TabIndex        =   10
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
         Left            =   600
         MouseIcon       =   "ListareteRecibos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListareteRecibos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2040
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
         Left            =   1920
         MouseIcon       =   "ListareteRecibos.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListareteRecibos.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2040
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
         Left            =   3240
         MouseIcon       =   "ListareteRecibos.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListareteRecibos.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salida"
         Top             =   2040
         Width           =   855
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   2160
         TabIndex        =   3
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
      Begin MSMask.MaskEdBox Desdefecha 
         Height          =   300
         Left            =   2160
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
         Left            =   600
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta fecha"
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
         Left            =   600
         TabIndex        =   5
         Top             =   720
         Width           =   1695
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
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaReteRecibos.rpt"
      Destination     =   1
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
      Left            =   4680
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListareterecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Recibos SET "
    ZSql = ZSql + " Impo1List = 0"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
    Select Case Tipo.ListIndex
        Case 0
            ZVarios = "Retenciones de Ganancias"
            ZSql = ""
            ZSql = ZSql + "UPDATE Recibos SET "
            ZSql = ZSql + " ImpoList = NroRetGanancias " + ","
            ZSql = ZSql + " Impo1List = RetGanancias"
            ZSql = ZSql + " Where Renglon = 1"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 1
            ZVarios = "Retenciones de Iva"
            ZSql = ""
            ZSql = ZSql + "UPDATE Recibos SET "
            ZSql = ZSql + " ImpoList = NroRetIva" + ","
            ZSql = ZSql + " Impo1List = RetIva"
            ZSql = ZSql + " Where Renglon = 1"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 2
            ZVarios = "Retenciones de Ingresos Brutos"
            ZSql = ""
            ZSql = ZSql + "UPDATE Recibos SET "
            ZSql = ZSql + " ImpoList = NroRetOtra" + ","
            ZSql = ZSql + " Impo1List = RetOtra"
            ZSql = ZSql + " Where Renglon = 1"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 3
            ZVarios = "Retenciones de Suss"
            ZSql = ""
            ZSql = ZSql + "UPDATE Recibos SET "
            ZSql = ZSql + " ImpoList = NroRetSuss" + ","
            ZSql = ZSql + " Impo1List = RetSuss"
            ZSql = ZSql + " Where Renglon = 1"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
        Case Else
    End Select
    
    WTitulo = "Del " + Desde.Text + " al " + Hasta.Text
    
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
    
    Listado.WindowTitle = "Listado de Retenciones Recibidas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Cliente, Recibos.Fecha, Recibos.Fechaord, Recibos.Impolist, Recibos.Impo1list, " _
                + "Cliente.Razon, Cliente.Cuit, " _
                + "Auxiliar.Nombre, Auxiliar.Actividad, Auxiliar.Varios " _
                + "From " _
                + DSQ + ".dbo.Recibos Recibos, " _
                + DSQ + ".dbo.Cliente Cliente, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Recibos.Cliente = Cliente.Cliente AND " _
                + "Recibos.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Recibos.Fechaord >= '" + WDesde + "' AND " _
                + "Recibos.Fechaord <= '" + WHasta + "' AND " _
                + "Recibos.Impo1list <> 0"
    
    Listado.Connect = Connect()
    
    Uno = "{Recibos.FechaOrd} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Recibos.Impo1list} <> 0.00"
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Action = 1
    

    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgListareterecibos.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Desdefecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefecha.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Ganancias"
    Tipo.AddItem "Iva"
    Tipo.AddItem "Ingresos Brutos"
    Tipo.AddItem "Suss"
    
    Tipo.ListIndex = 0

    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub DesdeFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
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





