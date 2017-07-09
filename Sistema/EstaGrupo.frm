VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEstaGrupo 
   AutoRedraw      =   -1  'True
   Caption         =   "Estadistica de Ventas por Grupo"
   ClientHeight    =   7020
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   7020
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
      Top             =   4560
      Visible         =   0   'False
      Width           =   6735
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
      Width           =   5175
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
         Left            =   3960
         MouseIcon       =   "estagrupo.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "estagrupo.frx":030A
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
         Left            =   1560
         MouseIcon       =   "estagrupo.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "estagrupo.frx":0E56
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
         Left            =   2760
         MouseIcon       =   "estagrupo.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "estagrupo.frx":19A2
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
         Left            =   360
         MouseIcon       =   "estagrupo.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "estagrupo.frx":24EE
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
         Left            =   720
         TabIndex        =   11
         Top             =   2400
         Width           =   2295
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   2280
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
         Left            =   2280
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
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   6
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
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   855
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
         Left            =   720
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
         Left            =   720
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
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
         TabIndex        =   5
         Top             =   720
         Width           =   1935
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
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Estacli.rpt"
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
      Left            =   6600
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
      ItemData        =   "estagrupo.frx":2D30
      Left            =   840
      List            =   "estagrupo.frx":2D37
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   6735
   End
End
Attribute VB_Name = "PrgEstaGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WImpo(10000) As Double
Dim ZNeto As Double

Private Producto As String
Private Costo As Double
Private WAuxi As String
Private Auxi As String

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

    WTitulo = "Del " + DesdeFec.Text + " al " + HastaFec.Text
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Estadistica SET "
    ZSql = ZSql + " Estadistica.Vendedor = Cliente.Vendedor"
    ZSql = ZSql + " From Estadistica, Cliente"
    ZSql = ZSql + " Where Estadistica.Cliente = Cliente.Cliente"
    ZSql = ZSql + " and Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    ZTotal = 0
    Erase WImpo
    
    ZSql = ""
    ZSql = ZSql + "Select *, Articulo.Grupo as [WGrupo]"
    ZSql = ZSql + " FROM Estadistica, Articulo"
    ZSql = ZSql + " Where Estadistica.Articulo = Articulo.Codigo"
    ZSql = ZSql + " and Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WGrupo = rstEstadistica!WGrupo
                    
                    If Tipolist.ListIndex = 0 Then
                    
                        ZImpo = rstEstadistica!Cantidad * rstEstadistica!Precio
                        ZDto = ZImpo * (rstEstadistica!Descuento / 100)
                        ZNeto = ZImpo - ZDto
                        
                        If rstEstadistica!Vendedor <> 1 Then
                            ZNeto = ZNeto * 0.9
                            Call Redondeo(ZNeto)
                        End If
                    
                        WImpo(WGrupo) = WImpo(WGrupo) + ZNeto
                        ZTotal = ZTotal + ZNeto
                        
                            Else
                            
                        WImpo(WGrupo) = WImpo(WGrupo) + rstEstadistica!Cantidad
                        ZTotal = ZTotal + rstEstadistica!Cantidad
                        
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Dolar = " + "'" + Str$(ZTotal) + "',"
    ZSql = ZSql + " Varios = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Estadistica SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    
    For Ciclo = 0 To 1000
    
        If WImpo(Ciclo) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Familia SET "
            ZSql = ZSql + " Ordena = " + "'" + Str$(WImpo(Ciclo)) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Str$(Ciclo) + "'"
            spFamilia = ZSql
            Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
    
    Next Ciclo
    
    
    
    
    
    Listado.WindowTitle = "Estadistica de Ventas por Grupo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Val(Desde.Text) = 0 Then
        Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
        Hasta.Text = "0"
    End If
    
    Listado.SQLQuery = "SELECT Estadistica.Tipo, Estadistica.Numero, Estadistica.Cantidad, Estadistica.Precio, Estadistica.Fecha, Estadistica.OrdFecha, Estadistica.Descuento, Estadistica.CantidadII, " _
            + "Articulo.Grupo, Articulo.Costo, " _
            + "Auxiliar.Nombre, Auxiliar.Varios, " _
            + "Familia.Descripcion, Familia.Ordena " _
            + "From " _
            + DSQ + ".dbo.Estadistica Estadistica, " _
            + DSQ + ".dbo.Articulo Articulo, " _
            + DSQ + ".dbo.Auxiliar Auxiliar, " _
            + DSQ + ".dbo.Familia Familia " _
            + "Where " _
            + "Estadistica.Articulo = Articulo.Codigo AND " _
            + "Estadistica.CodigoEmpresa = Auxiliar.Empresa AND " _
            + "Articulo.Grupo = Familia.Codigo AND " _
            + "Estadistica.OrdFecha >= '" + WDesde + "' AND " _
            + "Estadistica.OrdFecha <= '" + WHasta + "' AND " _
            + "Articulo.Grupo >= " + Desde.Text + " AND " _
            + "Articulo.Grupo <= " + Hasta.Text
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
            
    Uno = "{Estadistica.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    
    Select Case Tipolist.ListIndex
        Case 0, 1
            Listado.ReportFileName = "EstaGrupo.rpt"
        Case 2
            Listado.SQLQuery = "SELECT Estadistica.Tipo, Estadistica.Numero, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Precio, Estadistica.Vendedor, Estadistica.Fecha, Estadistica.OrdFecha, Estadistica.Descuento, Estadistica.CantidadII, " _
                + "Articulo.Descripcion, Articulo.Grupo, Articulo.Costo, " _
                + "Auxiliar.Nombre, Auxiliar.Varios, " _
                + "Familia.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Estadistica Estadistica, " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Familia Familia " _
                + "Where " _
                + "Estadistica.Articulo = Articulo.Codigo AND " _
                + "Estadistica.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Grupo = Familia.Codigo AND " _
                + "Estadistica.OrdFecha >= '" + WDesde + "' AND " _
                + "Estadistica.OrdFecha <= '" + WHasta + "' AND " _
                + "Articulo.Grupo >= " + Desde.Text + " AND " _
                + "Articulo.Grupo <= " + Hasta.Text
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
        
            Listado.ReportFileName = "EstaGrupoarti.rpt"
            
        Case Else
            Listado.SQLQuery = "SELECT Estadistica.Tipo, Estadistica.Numero, Estadistica.Articulo, Estadistica.Cantidad, Estadistica.Precio, Estadistica.Cliente, Estadistica.Vendedor, Estadistica.Fecha, Estadistica.OrdFecha, Estadistica.Descuento, Estadistica.CantidadII, " _
                + "Articulo.Grupo, Articulo.Costo, " _
                + "Auxiliar.Nombre, Auxiliar.Varios, " _
                + "Cliente.Razon, " _
                + "Familia.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Estadistica Estadistica, " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Cliente Cliente, " _
                + DSQ + ".dbo.Familia Familia " _
                + "Where " _
                + "Estadistica.Articulo = Articulo.Codigo AND " _
                + "Estadistica.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Estadistica.Cliente = Cliente.Cliente AND " _
                + "Articulo.Grupo = Familia.Codigo AND " _
                + "Estadistica.OrdFecha >= '" + WDesde + "' AND " _
                + "Estadistica.OrdFecha <= '" + WHasta + "' AND " _
                + "Articulo.Grupo >= " + Desde.Text + " AND " _
                + "Articulo.Grupo <= " + Hasta.Text
    
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Listado.ReportFileName = "EstaGrupoClie.rpt"
    End Select
   
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgEstaGrupo.Hide
    Unload Me
    Menu42.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
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

    Tipolist.AddItem "Por Importe"
    Tipolist.AddItem "Por Unidad"
    Tipolist.AddItem "Por Producto"
    Tipolist.AddItem "Por Cliente"
    
    Tipolist.ListIndex = 0

    Desde.Text = ""
    Hasta.Text = ""
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Ayuda.Visible = True
    Ayuda.Text = ""

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Familia"
    ZSql = ZSql + " Order by Familia.Codigo"
    spFamilia = ZSql
    Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
    If rstFamilia.RecordCount > 0 Then
        With rstFamilia
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
        rstFamilia.Close
    End If
            
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
            ZSql = ZSql + " FROM Familia"
            ZSql = ZSql + " Where Familia.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Familia.Codigo"
            spFamilia = ZSql
            Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
            If rstFamilia.RecordCount > 0 Then
                With rstFamilia
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
                rstFamilia.Close
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












