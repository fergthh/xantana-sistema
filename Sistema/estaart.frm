VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEstaart 
   AutoRedraw      =   -1  'True
   Caption         =   "Estadistica de Ventas por Articulo"
   ClientHeight    =   6945
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   8145
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
      Left            =   840
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   6255
   End
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
      Height          =   4215
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   4935
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
         MouseIcon       =   "estaart.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "estaart.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salida"
         Top             =   3000
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
         Left            =   1440
         MouseIcon       =   "estaart.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "estaart.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Impresion x Impresora"
         Top             =   3000
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
         Left            =   2640
         MouseIcon       =   "estaart.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "estaart.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Consulta de Datos"
         Top             =   3000
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
         MouseIcon       =   "estaart.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "estaart.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   3000
         Width           =   855
      End
      Begin VB.ComboBox Tipolist 
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
         Left            =   360
         TabIndex        =   11
         Top             =   2400
         Width           =   2295
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   1800
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
         Left            =   1920
         TabIndex        =   9
         Top             =   1320
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
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   6
         Text            =   " "
         Top             =   720
         Width           =   2055
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
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   2055
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
         Left            =   360
         TabIndex        =   8
         Top             =   1800
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
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
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
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
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
         Top             =   360
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Estaart.rpt"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "estaart.frx":2D30
      Left            =   840
      List            =   "estaart.frx":2D37
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "PrgEstaart"
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
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia

    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Varios = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Estadistica SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Estadistica de Ventas por Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Select Case Tipolist.ListIndex
        Case 0
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Listado.SQLQuery = "SELECT Estadistica.Tipo, Estadistica.Numero, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Precio, Estadistica.Fecha, Estadistica.OrdFecha, " _
                    + "Auxiliar.Nombre, Auxiliar.Varios, " _
                    + "Articulo.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Estadistica Estadistica, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Articulo Articulo " _
                    + "Where " _
                    + "Estadistica.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Estadistica.Articulo = Articulo.Codigo AND " _
                    + "Estadistica.Articulo >= '" + Desde.Text + "' AND " _
                    + "Estadistica.Articulo <= '" + Hasta.Text + "' AND " _
                    + "Estadistica.OrdFecha >= '" + WDesde + "' AND " _
                    + "Estadistica.OrdFecha <= '" + WHasta + "'"
                
            Listado.Connect = Connect()
                    
            Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Dos = " and {Estadistica.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "EstaArt1.rpt"
            
        Case 1
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Listado.SQLQuery = "SELECT Estadistica.Tipo, Estadistica.Numero, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Precio, Estadistica.Fecha, Estadistica.OrdFecha, " _
                    + "Auxiliar.Nombre, Auxiliar.Varios, " _
                    + "Articulo.Descripcion, Articulo.Sector, " _
                    + "Sector.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Estadistica Estadistica, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Sector Sector " _
                    + "Where " _
                    + "Estadistica.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Estadistica.Articulo = Articulo.Codigo AND " _
                    + "Articulo.Sector = Sector.Codigo AND " _
                    + "Estadistica.Articulo >= '" + Desde.Text + "' AND " _
                    + "Estadistica.Articulo <= '" + Hasta.Text + "' AND " _
                    + "Estadistica.OrdFecha >= '" + WDesde + "' AND " _
                    + "Estadistica.OrdFecha <= '" + WHasta + "'"
        
            Listado.Connect = Connect()
                    
            Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Dos = " and {Estadistica.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
            
            Listado.ReportFileName = "EstaArtSECTOR.rpt"
            
        Case Else
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Listado.SQLQuery = "SELECT Estadistica.Tipo, Estadistica.Numero, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Precio, Estadistica.Cliente, Estadistica.Fecha, Estadistica.OrdFecha, " _
                    + "Auxiliar.Nombre, Auxiliar.Varios, " _
                    + "Articulo.Descripcion, " _
                    + "Cliente.Razon " _
                    + "From " _
                    + DSQ + ".dbo.Estadistica Estadistica, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Estadistica.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Estadistica.Articulo = Articulo.Codigo AND " _
                    + "Estadistica.Cliente = Cliente.Cliente AND " _
                    + "Estadistica.Articulo >= '" + Desde.Text + "' AND " _
                    + "Estadistica.Articulo <= '" + Hasta.Text + "' AND " _
                    + "Estadistica.OrdFecha >= '" + WDesde + "' AND " _
                    + "Estadistica.OrdFecha <= '" + WHasta + "'"
        
            Listado.Connect = Connect()
                    
            Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Dos = " and {Estadistica.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
            
            Listado.ReportFileName = "EstaArtCliente.rpt"
            
    End Select
   
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgEstaart.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFec.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
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
            Desde.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Tipolist.Clear

    Tipolist.AddItem "Resumido"
    Tipolist.AddItem "x sector"
    Tipolist.AddItem "x Cliente"
    
    Tipolist.ListIndex = 0

    Desde.Text = ""
    Hasta.Text = ""
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Ayuda.Visible = True
    Ayuda.Text = ""
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Indice = Pantalla.ListIndex
    Desde.Text = WIndice.List(Indice)
    Hasta.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    Desde.SetFocus
    
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
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
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

Private Sub TipoList_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 115
            Call Consulta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub








