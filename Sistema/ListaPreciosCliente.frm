VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaPreciosCliente 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Precios (Cliente)"
   ClientHeight    =   3180
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3180
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
      Height          =   2775
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Cliente 
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
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   2
         Text            =   " "
         Top             =   1200
         Width           =   1095
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
         Left            =   4560
         MouseIcon       =   "ListaPreciosCliente.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaPreciosCliente.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salida"
         Top             =   1560
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
         MouseIcon       =   "ListaPreciosCliente.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaPreciosCliente.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1560
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
         Left            =   840
         MouseIcon       =   "ListaPreciosCliente.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaPreciosCliente.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1560
         Width           =   855
      End
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   720
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
         TabIndex        =   11
         Top             =   240
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
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label DesCliente 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   1200
         Width           =   3375
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
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaPreciosCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WVector(5000) As String
Dim ZVector(5000, 10) As String
Dim ZImpre(60, 10) As String

Dim ZZCodigo As String
Dim ZZDescripcion As String
Dim ZZPrecio As String
Dim ZZDesGrupo As String

Dim XCodigo As String
Dim XDescripcion As String
Dim XPrecio As String
Dim XDesGrupo As String
Dim XCodigoII As String
Dim XDescripcionII As String
Dim XPrecioII As String
Dim XDesGrupoII As String


Dim ZZEntra As Integer
Dim ZZPagina As Integer
Dim ZZRenglon As Integer
Dim ZZProceso As Integer
Dim XLugar As Integer

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
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Imprime = " + "'" + "N" + "',"
    ZSql = ZSql + " ListaCodigo = Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    Erase ZVector
    ZLugar = 0
    Pasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Estadistica.Cliente = " + "'" + Cliente.Text + "'"
    ZSql = ZSql + " order by Estadistica.Cliente, Estadistica.Articulo"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WArticulo = rstEstadistica!Articulo
                    
                    If Pasa = 0 Then
                        Pasa = 1
                        Corte = WArticulo
                    End If
                    
                    If WArticulo <> Corte Then
                        ZLugar = ZLugar + 1
                        WVector(ZLugar) = Corte
                        Corte = WArticulo
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEstadistica.Close
    End If
    
    If Pasa <> 0 Then
        ZLugar = ZLugar + 1
        WVector(ZLugar) = Corte
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZArticulo = WVector(Ciclo)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Imprime = " + "'" + "S" + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
        
    ZSql = ""
    ZSql = ZSql + "DELETE ListaPrecios"
    spListaPrecios = ZSql
    Set rstListaPrecios = db.OpenRecordset(spListaPrecios, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Lista de Precios por Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *, Familia.Descripcion as [WDesGrupo], Familia.Estado as [WEstado]"
    ZSql = ZSql + " FROM Articulo, Familia"
    ZSql = ZSql + " Where Articulo.Grupo = Familia.Codigo"
    Rem ZSql = ZSql + " and Articulo.WEstado = 0"
    ZSql = ZSql + " Order by Articulo.WDesGrupo, Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstArticulo!WEstado = 0 Then
                    
                        If rstArticulo!Imprime = "S" Then
                    
                            ZLugar = ZLugar + 1
                            ZVector(ZLugar, 1) = rstArticulo!Grupo
                            ZVector(ZLugar, 2) = rstArticulo!WDesGrupo
                            ZVector(ZLugar, 3) = rstArticulo!Codigo
                            ZVector(ZLugar, 4) = rstArticulo!Descripcion
                            ZVector(ZLugar, 5) = Str$(rstArticulo!Precio)
                            
                        End If
                    
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
    
    ZZPagina = 0
    ZZRenglon = 0
    
    ZZCodigo = ""
    ZZDescripcion = ""
    ZZPrecio = ""
    ZZDesGrupo = ""
    
    Pasa = 0
    XLugar = 0
    Erase ZImpre
    
    For Ciclo = 1 To ZLugar
    
        ZGrupo = ZVector(Ciclo, 1)
        ZDesGrupo = ZVector(Ciclo, 2)
        ZCodigo = ZVector(Ciclo, 3)
        ZDescripcion = ZVector(Ciclo, 4)
        ZPrecio = ZVector(Ciclo, 5)
    
        If Pasa = 0 Then
        
            ZZCodigo = ""
            ZZDescripcion = ""
            ZZPrecio = ""
            ZZDesGrupo = ZDesGrupo
            Call Graba_Registro
            
            ZCorte = ZDesGrupo
            Pasa = 1
            
        End If
        
        If ZCorte <> ZDesGrupo Then
            
            ZZCodigo = ""
            ZZDescripcion = ""
            ZZPrecio = ""
            ZZDesGrupo = ""
            Call Graba_Registro
            
            ZZCodigo = ""
            ZZDescripcion = ""
            ZZPrecio = ""
            ZZDesGrupo = ZDesGrupo
            Call Graba_Registro
            
            ZCorte = ZDesGrupo
            
        End If
    
        ZZCodigo = ZCodigo
        ZZDescripcion = ZDescripcion
        ZZPrecio = ZPrecio
        ZZDesGrupo = ""
        Call Graba_Registro
        
    Next Ciclo
    
    Call Graba

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ListaPrecios.Pagina, ListaPrecios.Renglon, ListaPrecios.DesGrupo, ListaPrecios.Codigo, ListaPrecios.Descripcion, ListaPrecios.Precio, ListaPrecios.DesGrupoII, ListaPrecios.CodigoII, ListaPrecios.DescripcionII, ListaPrecios.PrecioII, ListaPrecios.Titulo " _
            + "From " _
            + DSQ + ".dbo.ListaPrecios ListaPrecios " _
            + "Where " _
            + "ListaPrecios.Pagina >= 0 AND " _
            + "ListaPrecios.Pagina <= 9999 "
            
    Uno = "{ListaPrecios.Pagina} in 0 to 9999"
        
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.ReportFileName = "ListaPrecios.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next

    
End Sub

Private Sub Cancela_click()
    PrgListaPreciosCliente.Hide
    Unload Me
    Menu21.Show
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
            Cliente.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub
  
Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Trim(Cliente.Text) <> "" Then
            Auxi = UCase(Left$(Cliente.Text, 1))
            Auxi1 = Mid$(Cliente.Text, 2, 5)
            Call Ceros(Auxi1, 3)
            Cliente.Text = Auxi + "-" + Auxi1
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            rstCliente.Close
            DesdeFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Cliente.Text = ""
    DesCliente.Caption = ""
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

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
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














Private Sub Graba_Registro()


    XLugar = XLugar + 1
    If XLugar = 121 Then
        Call Graba
        XLugar = 1
        Erase ZImpre
    End If
    
    If XLugar <= 60 Then
        ZImpre(XLugar, 1) = ZZCodigo
        ZImpre(XLugar, 2) = ZZDescripcion
        ZImpre(XLugar, 3) = ZZPrecio
        ZImpre(XLugar, 4) = ZZDesGrupo
            Else
        XXLugar = XLugar - 60
        ZImpre(XXLugar, 5) = ZZCodigo
        ZImpre(XXLugar, 6) = ZZDescripcion
        ZImpre(XXLugar, 7) = ZZPrecio
        ZImpre(XXLugar, 8) = ZZDesGrupo
    End If
        
End Sub

Private Sub Graba()

    ZZPagina = ZZPagina + 1
    
    For Ciclo = 1 To 60
    
        XCodigo = ZImpre(Ciclo, 1)
        XDescripcion = ZImpre(Ciclo, 2)
        XPrecio = ZImpre(Ciclo, 3)
        XDesGrupo = ZImpre(Ciclo, 4)
        
        XCodigoII = ZImpre(Ciclo, 5)
        XDescripcionII = ZImpre(Ciclo, 6)
        XPrecioII = ZImpre(Ciclo, 7)
        XDesGrupoII = ZImpre(Ciclo, 8)
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ListaPrecios ("
        ZSql = ZSql + "Pagina ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "DesGrupo ,"
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "DesGrupoII ,"
        ZSql = ZSql + "CodigoII ,"
        ZSql = ZSql + "DescripcionII ,"
        ZSql = ZSql + "PrecioII ,"
        ZSql = ZSql + "Titulo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Str$(ZZPagina) + "',"
        ZSql = ZSql + "'" + Str$(Ciclo) + "',"
        ZSql = ZSql + "'" + XDesGrupo + "',"
        ZSql = ZSql + "'" + XCodigo + "',"
        ZSql = ZSql + "'" + XDescripcion + "',"
        ZSql = ZSql + "'" + XPrecio + "',"
        ZSql = ZSql + "'" + XDesGrupoII + "',"
        ZSql = ZSql + "'" + XCodigoII + "',"
        ZSql = ZSql + "'" + XDescripcionII + "',"
        ZSql = ZSql + "'" + XPrecioII + "',"
        ZSql = ZSql + "'" + DesCliente.Caption + "')"
        
        spListaPrecios = ZSql
        Set rstListaPrecios = db.OpenRecordset(spListaPrecios, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo

End Sub


