VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaStockGrupo 
   AutoRedraw      =   -1  'True
   Caption         =   "Stock por Grupo"
   ClientHeight    =   3585
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3585
   ScaleWidth      =   8145
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
      Height          =   3135
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   6135
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
         Left            =   2520
         TabIndex        =   9
         Top             =   1320
         Width           =   2895
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
         Left            =   4440
         MouseIcon       =   "listastockgrupo.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "listastockgrupo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salida"
         Top             =   1920
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
         Left            =   2640
         MouseIcon       =   "listastockgrupo.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "listastockgrupo.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1920
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
         Left            =   720
         MouseIcon       =   "listastockgrupo.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "listastockgrupo.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1920
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   4
         Text            =   " "
         Top             =   840
         Width           =   1335
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Grupo"
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
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Grupo"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7920
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PrecioFam.rpt"
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
      Left            =   7440
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaStockGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Producto As String
Private Costo As Double


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
    If Val(Desde.Text) = 0 Then
        Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
        Hasta.Text = "0"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Stock por Grupo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Select Case Tipo.ListIndex
        Case 0
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Color, Articulo.Grupo, Articulo.Minimo, Articulo.Entradas, Articulo.Salidas, Articulo.Stock, Articulo.StockAnterior, " _
                + "Auxiliar.Nombre, " _
                + "Familia.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Familia Familia " _
                + "Where " _
                + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Grupo = Familia.Codigo AND " _
                + "Articulo.Grupo >= " + Desde.Text + " AND " _
                + "Articulo.Grupo <= " + Hasta.Text + " AND " _
                + "Articulo.Stock <> 0"
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Stock} <> 0"
            Dos = " and {Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockGrupo.rpt"
            Listado.Action = 1
            
        Case 1
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Color, Articulo.Grupo, Articulo.Minimo, Articulo.Entradas, Articulo.Salidas, Articulo.Stock, Articulo.StockAnterior, " _
                + "Auxiliar.Nombre, " _
                + "Familia.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Familia Familia " _
                + "Where " _
                + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Grupo = Familia.Codigo AND " _
                + "Articulo.Grupo >= " + Desde.Text + " AND " _
                + "Articulo.Grupo <= " + Hasta.Text
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
            Dos = ""
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockGrupoII.rpt"
            Listado.Action = 1
            
        Case 2
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Color, Articulo.Grupo, Articulo.Minimo, Articulo.Entradas, Articulo.Salidas, Articulo.Stock, Articulo.StockAnterior, " _
                + "Auxiliar.Nombre, " _
                + "Familia.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Familia Familia " _
                + "Where " _
                + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Grupo = Familia.Codigo AND " _
                + "Articulo.Grupo >= " + Desde.Text + " AND " _
                + "Articulo.Grupo <= " + Hasta.Text + " AND " _
                + "Articulo.Stock <> 0"
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Stock} <> 0"
            Dos = " and {Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockGrupoIII.rpt"
            Listado.Action = 1
            
        Case 3
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Color, Articulo.Grupo, Articulo.Minimo, Articulo.Entradas, Articulo.Salidas, Articulo.Stock, Articulo.StockAnterior, " _
                + "Auxiliar.Nombre, " _
                + "Familia.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Familia Familia " _
                + "Where " _
                + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Grupo = Familia.Codigo AND " _
                + "Articulo.Grupo >= " + Desde.Text + " AND " _
                + "Articulo.Grupo <= " + Hasta.Text
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
            Dos = ""
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockGrupoIV.rpt"
            Listado.Action = 1
            
        Case Else
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Stock, Articulo.StockI, Articulo.StockII, " _
                + "Auxiliar.Nombre, " _
                + "Familia.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Familia Familia " _
                + "Where " _
                + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Grupo = Familia.Codigo AND " _
                + "Articulo.Grupo >= " + Desde.Text + " AND " _
                + "Articulo.Grupo <= " + Hasta.Text + " AND " _
                + "Articulo.Stock <> 0"
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Stock} <> 0"
            Dos = " and {Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockGrupoV"
            ..rpt"
            Listado.Action = 1
            
    End Select
            
            
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaStockGrupo.Hide
    Unload Me
    Menu231.Show
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Stock"
    Tipo.AddItem "Completo"
    Tipo.AddItem "Cliente Stock"
    Tipo.AddItem "Cliente Completo"
    Tipo.AddItem "Por Deposito"
    
    Tipo.ListIndex = 1

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
            Call Cancela_click
        Case Else
    End Select
End Sub
















