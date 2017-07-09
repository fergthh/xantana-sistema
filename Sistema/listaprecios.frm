VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaPrecios 
   AutoRedraw      =   -1  'True
   Caption         =   "Lista de Precios"
   ClientHeight    =   6945
   ClientLeft      =   1935
   ClientTop       =   750
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   8145
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
      Height          =   2460
      ItemData        =   "listaprecios.frx":0000
      Left            =   720
      List            =   "listaprecios.frx":0007
      TabIndex        =   10
      Top             =   4320
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.ListBox Opcion 
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
      Left            =   840
      TabIndex        =   16
      Top             =   4320
      Visible         =   0   'False
      Width           =   3615
   End
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
      Left            =   720
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      Begin VB.ComboBox Activo 
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
         Left            =   4080
         TabIndex        =   17
         Top             =   720
         Width           =   2055
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
         Left            =   2160
         MouseIcon       =   "listaprecios.frx":0015
         MousePointer    =   99  'Custom
         Picture         =   "listaprecios.frx":031F
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Consulta de Datos"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Linea 
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
         MaxLength       =   10
         TabIndex        =   12
         Text            =   " "
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Tipo 
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
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   11
         Text            =   " "
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox Abierto 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Lista 
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
         MaxLength       =   8
         TabIndex        =   5
         Top             =   240
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
         Left            =   5400
         MouseIcon       =   "listaprecios.frx":0B61
         MousePointer    =   99  'Custom
         Picture         =   "listaprecios.frx":0E6B
         Style           =   1  'Graphical
         TabIndex        =   2
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
         Left            =   3840
         MouseIcon       =   "listaprecios.frx":16AD
         MousePointer    =   99  'Custom
         Picture         =   "listaprecios.frx":19B7
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   600
         MouseIcon       =   "listaprecios.frx":21F9
         MousePointer    =   99  'Custom
         Picture         =   "listaprecios.frx":2503
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codigo "
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
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1815
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
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label DesLista 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Lista"
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
         TabIndex        =   6
         Top             =   240
         Width           =   1215
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
Attribute VB_Name = "PrgListaPrecios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(20000) As String

Dim ZZCodigo As String
Dim ZZDescripcion As String
Dim ZZPrecio As String

Dim XCodigo As String
Dim XDescripcion As String
Dim XPrecio As String

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

    WTitulo = ""

    If Trim(Linea.Text) <> "" Then
        WTitulo = "Linea:" + Linea.Text
    End If
    If Trim(Tipo.Text) <> "" Then
        WTitulo = WTitulo + " Producto:" + Tipo.Text
    End If

    ZSql = ""
    ZSql = ZSql + "UPDATE Precios SET "
    ZSql = ZSql + " Marca = " + "'" + "" + "',"
    ZSql = ZSql + " Titulo = " + "'" + WTitulo + "'"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)

    Listado.WindowTitle = "Lista de Precios"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    ZLugar = 0
    Erase ZVector
    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WAno = Right$(WFecha, 4)
    WMes = Mid$(WFecha, 4, 2)
    WDia = Left$(WFecha, 2)
    WOrdFecha = WAno + WMes + WDia
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Lista = " + "'" + Lista.Text + "'"
    If Trim(Linea.Text) <> "" Then
        ZSql = ZSql + " and Precios.Linea = " + "'" + Linea.Text + "'"
    End If
    If Trim(Tipo.Text) <> "" Then
        ZSql = ZSql + " and Precios.Tipo = " + "'" + Tipo.Text + "'"
    End If
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Abierto.ListIndex = 1 Or WOrdFecha <= rstPrecios!OrdHasta Then
                
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = rstPrecios!Codigo
                    
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZVector(Ciclo)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZCodigo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZZActivo = rstArticulo!Activo
            rstArticulo.Close
        End If
    
        If Activo.ListIndex = 1 Or ZZActivo = 0 Then
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Precios SET "
            ZSql = ZSql + " Marca = " + "'" + "X" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + ZCodigo + "'"
            ZSql = ZSql + " and Lista = " + "'" + Lista.Text + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Precios.Codigo, Precios.Linea, Precios.Tipo, Precios.Lista, Precios.Desde, Precios.Hasta, Precios.Tope1, Precios.Valor1, Precios.Tope2, Precios.Valor2, Precios.Tope3, Precios.Valor3, Precios.Tope4, Precios.Valor4, Precios.Moneda, Precios.Marca, Precios.Titulo, " _
            + "Articulo.DescripcionII,  " _
            + "Lista.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Precios Precios, " _
            + DSQ + ".dbo.Articulo Articulo, " _
            + DSQ + ".dbo.Lista Lista " _
            + "Where " _
            + "Precios.Codigo = Articulo.Codigo AND " _
            + "Precios.Lista = Lista.Codigo AND " _
            + "Precios.Marca = 'X'"
            
    Listado.Connect = Connect()
    
    Uno = "{Precios.Marca} = " + Chr$(34) + "X" + Chr$(34)
        
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.ReportFileName = "ListaPrecios.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub


Private Sub Cancela_click()
    PrgListaPrecios.Hide
    Unload Me
    MenuVen.Show
End Sub

Sub Form_Load()
    Lista.Text = ""
    DesLista.Caption = ""
    Linea.Text = ""
    Tipo.Text = ""
    
    Abierto.Clear
    
    Abierto.AddItem "Abiertas"
    Abierto.AddItem "Todas"
    
    Abierto.ListIndex = 0
        
    Activo.Clear
    
    Activo.AddItem "Activos"
    Activo.AddItem "Todas"
    
    Activo.ListIndex = 0
    

End Sub

Private Sub Titulo_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Lista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM lista"
        ZSql = ZSql + " Where lista.Codigo = " + "'" + Lista.Text + "'"
        spLista = ZSql
        Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstLista.RecordCount > 0 Then
            DesLista.Caption = rstLista!Descripcion
            rstLista.Close
        End If
    End If
    
    If KeyAscii = 27 Then
        Lista.Text = ""
        DesLista.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub LInea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Linea.Text = UCase(Linea.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Lineas"
        ZSql = ZSql + " Where Lineas.Codigo = " + "'" + Linea.Text + "'"
        spLinea = ZSql
        Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstLinea.RecordCount > 0 Then
            rstLinea.Close
            Tipo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Linea.Text = ""
    End If
End Sub

Private Sub Tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tipo.Text = UCase(Tipo.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoPro"
        ZSql = ZSql + " Where TipoPro.Codigo = " + "'" + Tipo.Text + "'"
        spTipoPro = ZSql
        Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoPro.RecordCount > 0 Then
            rstTipoPro.Close
            Lista.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
    End If
End Sub









Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Listas"
     Opcion.AddItem "Linea"
     Opcion.AddItem "Tipo"

     Rem Opcion.Visible = True
     
     Opcion.ListIndex = 0
     Call Opcion_Click
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lista"
            ZSql = ZSql + " Order by Lista.Descripcion"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
            If rstLista.RecordCount > 0 Then
                With rstLista
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
                rstLista.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Order by Lineas.Descripcion"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                With rstLinea
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
                rstLinea.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoPro"
            ZSql = ZSql + " Order by TipoPro.Descripcion"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoPro.RecordCount > 0 Then
                With rstTipoPro
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
                rstTipoPro.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Lista.Text = WIndice.List(Indice)
            Call Lista_KeyPress(13)
        Case 1
            Indice = Pantalla.ListIndex
            Linea.Text = WIndice.List(Indice)
            Call LInea_KeyPress(13)
        Case 2
            Indice = Pantalla.ListIndex
            Tipo.Text = WIndice.List(Indice)
            Call Tipo_KeyPress(13)
            
        Case Else
    End Select
    
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
            ZSql = ZSql + " FROM Lista"
            ZSql = ZSql + " Where Lista.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Lista.Descripcion"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
            If rstLista.RecordCount > 0 Then
                With rstLista
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
                rstLista.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Where Lineas.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Lineas.Descripcion"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                With rstLinea
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
                rstLinea.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoPro"
            ZSql = ZSql + " Where TipoPro.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by TipoPro.Descripcion"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoPro.RecordCount > 0 Then
                With rstTipoPro
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
                rstTipoPro.Close
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

Private Sub Lista_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Listas"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub LInea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Listas"
    Opcion.AddItem "Linea"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Tipo_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Listas"
    Opcion.AddItem "Linea"
    Opcion.AddItem "Tipo"

    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub


