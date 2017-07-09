VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaPreciosFuturo 
   AutoRedraw      =   -1  'True
   Caption         =   "Lista de Precios Futuro"
   ClientHeight    =   2565
   ClientLeft      =   1935
   ClientTop       =   750
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2565
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Titulo 
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
         MaxLength       =   50
         TabIndex        =   5
         Text            =   " "
         Top             =   360
         Width           =   4095
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
         MouseIcon       =   "ListaPreciosFuturo.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaPreciosFuturo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salida"
         Top             =   960
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
         Left            =   2880
         MouseIcon       =   "ListaPreciosFuturo.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaPreciosFuturo.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Impresion x Impresora"
         Top             =   960
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
         Left            =   960
         MouseIcon       =   "ListaPreciosFuturo.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaPreciosFuturo.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Titulo"
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
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Listcomi.rpt"
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
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaPreciosFuturo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(5000, 10) As String
Dim ZImpre(60, 10) As String

Dim ZZCodigo As String
Dim ZZDescripcion As String
Dim ZZPrecio As String
Dim ZZPrecioFuturo As String
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
    ZSql = ZSql + "DELETE ListaPrecios"
    spListaPrecios = ZSql
    Set rstListaPrecios = db.OpenRecordset(spListaPrecios, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Lista de Precios"
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
                    
                    If rstArticulo!CostoFuturo <> 0 Then
                
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar, 1) = rstArticulo!Grupo
                        ZVector(ZLugar, 2) = rstArticulo!WDesGrupo
                        ZVector(ZLugar, 3) = rstArticulo!Codigo
                        ZVector(ZLugar, 4) = rstArticulo!Descripcion
                        ZVector(ZLugar, 5) = Str$(rstArticulo!Precio)
                        ZVector(ZLugar, 6) = Str$(rstArticulo!PrecioFuturo)
                        
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
    ZZPrecioFuturo = ""
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
        ZPrecioFuturo = ZVector(Ciclo, 6)
    
        If Pasa = 0 Then
        
            ZZCodigo = ""
            ZZDescripcion = ""
            ZZPrecio = ""
            ZZPrecioFuturo = ""
            ZZDesGrupo = ZDesGrupo
            Call Graba_Registro
            
            ZCorte = ZDesGrupo
            Pasa = 1
            
        End If
        
        If ZCorte <> ZDesGrupo Then
            
            ZZCodigo = ""
            ZZDescripcion = ""
            ZZPrecio = ""
            ZZPrecioFuturo = ""
            ZZDesGrupo = ""
            Call Graba_Registro
            
            ZZCodigo = ""
            ZZDescripcion = ""
            ZZPrecio = ""
            ZZPrecioFuturo = ""
            ZZDesGrupo = ZDesGrupo
            Call Graba_Registro
            
            ZCorte = ZDesGrupo
            
        End If
    
        ZZCodigo = ZCodigo
        ZZDescripcion = ZDescripcion
        ZZPrecio = ZPrecio
        ZZPrecioFuturo = ZPrecioFuturo
        ZZDesGrupo = ""
        Call Graba_Registro
        
    Next Ciclo
    
    Call Graba

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT ListaPrecios.Pagina, ListaPrecios.Renglon, ListaPrecios.DesGrupo, ListaPrecios.Codigo, ListaPrecios.Descripcion, ListaPrecios.Precio, ListaPrecios.DesGrupoII, ListaPrecios.CodigoII, ListaPrecios.DescripcionII, ListaPrecios.PrecioII, ListaPrecios.Titulo, ListaPrecios.PrecioIII, ListaPrecios.PrecioIV " _
            + "From " _
            + DSQ + ".dbo.ListaPrecios ListaPrecios " _
            + "Where " _
            + "ListaPrecios.Pagina >= 0 AND " _
            + "ListaPrecios.Pagina <= 9999 "
            
    Uno = "{ListaPrecios.Pagina} in 0 to 9999"
        
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.ReportFileName = "ListaPreciosFuturo.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
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
        ZImpre(XLugar, 4) = ZZPrecioFuturo
        ZImpre(XLugar, 5) = ZZDesGrupo
            Else
        XXLugar = XLugar - 60
        ZImpre(XXLugar, 6) = ZZCodigo
        ZImpre(XXLugar, 7) = ZZDescripcion
        ZImpre(XXLugar, 8) = ZZPrecio
        ZImpre(XXLugar, 9) = ZZPrecioFuturo
        ZImpre(XXLugar, 10) = ZZDesGrupo
    End If
        
End Sub

Private Sub Graba()

    ZZPagina = ZZPagina + 1
    
    For Ciclo = 1 To 60
    
        XCodigo = ZImpre(Ciclo, 1)
        XDescripcion = ZImpre(Ciclo, 2)
        XPrecio = ZImpre(Ciclo, 3)
        XPrecioIII = ZImpre(Ciclo, 4)
        XDesGrupo = ZImpre(Ciclo, 5)
        
        XCodigoII = ZImpre(Ciclo, 6)
        XDescripcionII = ZImpre(Ciclo, 7)
        XPrecioII = ZImpre(Ciclo, 8)
        XPrecioIV = ZImpre(Ciclo, 9)
        XDesGrupoII = ZImpre(Ciclo, 10)
            
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
        ZSql = ZSql + "PrecioIII ,"
        ZSql = ZSql + "PrecioIV ,"
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
        ZSql = ZSql + "'" + XPrecioIII + "',"
        ZSql = ZSql + "'" + XPrecioIV + "',"
        ZSql = ZSql + "'" + Titulo.Text + "')"
        
        spListaPrecios = ZSql
        Set rstListaPrecios = db.OpenRecordset(spListaPrecios, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo

End Sub

Private Sub Cancela_click()
    PrgListaPreciosFuturo.Hide
    Unload Me
    Menu21.Show
End Sub

Sub Form_Load()
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













