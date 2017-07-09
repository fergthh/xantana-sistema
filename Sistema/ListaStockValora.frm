VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaStockValora 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Valorizacion de stock de Articulo"
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
         MouseIcon       =   "ListaStockValora.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaStockValora.frx":030A
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
         MouseIcon       =   "ListaStockValora.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaStockValora.frx":0E56
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
         MouseIcon       =   "ListaStockValora.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaStockValora.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1920
         Width           =   855
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   4
         Text            =   " "
         Top             =   840
         Width           =   1335
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Linea"
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
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Linea"
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
Attribute VB_Name = "PrgListaStockValora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZProducto(10000, 10) As String
Dim ZTerminado(100, 2) As String
Dim ZCosto(100, 2) As String

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

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + "1" + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        WWParidad = rstDolar!Paridad
    End If

    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Costo = " + "'" + "0" + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)



    Erase ZProducto
    ZRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Stock > 0"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
    
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZRenglon = ZRenglon + 1
                    ZProducto(ZRenglon, 1) = rstArticulo!Codigo
                    ZProducto(ZRenglon, 2) = rstArticulo!Linea
                    ZProducto(ZRenglon, 3) = rstArticulo!Tipo
                    ZProducto(ZRenglon, 4) = rstArticulo!Fragancia
                    ZProducto(ZRenglon, 5) = rstArticulo!Calidad
                    ZProducto(ZRenglon, 6) = rstArticulo!Tamano
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstArticulo.Close
    End If
    

    For Ciclo = 1 To ZRenglon
    
        Erase ZTerminado
        Erase ZCosto
        ZLugar = 1
        ZLugarII = 0
        
        ZZCostoI = 0
        ZZCostoII = 0
        
        ZTerminado(ZLugar, 1) = ZProducto(Ciclo, 1)
        ZTerminado(ZLugar, 2) = "1"
        
        For CicloII = 1 To ZLugar
        
            ZZCombo = ""
            ZZArticulo = Trim(UCase(ZTerminado(CicloII, 1)))
            ZZProduccion = Val(ZTerminado(CicloII, 2))
            
            
            For ZZRenglon = 1 To 100
                
                Auxi1 = ZZRenglon
                Call Ceros(Auxi1, 2)
                
                ZZCodigo = ZZArticulo
                ZZClave = ZZCodigo + Auxi1
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Formula"
                ZSql = ZSql + " Where Formula.Clave = " + "'" + ZZClave + "'"
                spFormula = ZSql
                Set rstFormula = db.OpenRecordset(spFormula, dbOpenSnapshot, dbSQLPassThrough)
                If rstFormula.RecordCount > 0 Then
                    
                    ZZZInsumo = Trim(rstFormula!Insumo)
                    ZZZTerminado = Trim(rstFormula!terminado)
                    ZZZCantidad = rstFormula!Cantidad
                    
                    ZZCombo = Trim(rstFormula!Combo)
                    ZZCanti = ZZZCantidad * ZZProduccion
                    
                    rstFormula.Close
                        
                    If Trim(ZZZInsumo) <> "" Then
                        ZLugarII = ZLugarII + 1
                        ZCosto(ZLugarII, 1) = ZZZInsumo
                        ZCosto(ZLugarII, 2) = Str$(ZZCanti)
                    End If
                        
                    If Trim(ZZZTerminado) <> "" Then
                        ZLugar = ZLugar + 1
                        ZTerminado(ZLugar, 1) = ZZZTerminado
                        ZTerminado(ZLugar, 2) = Str$(ZZCanti)
                    End If
                        
                End If
                                                
            Next ZZRenglon
                                            
            If Trim(ZZCombo) <> "" Then
                
                For ZZRenglon = 1 To 100
                
                    Auxi1 = ZZRenglon
                    Call Ceros(Auxi1, 2)
                    
                    ZZCodigo = ZZCombo
                    ZZClave = ZZCodigo + Auxi1
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Combo"
                    ZSql = ZSql + " Where Combo.Clave = " + "'" + ZZClave + "'"
                    spCombo = ZSql
                    Set rstCombo = db.OpenRecordset(spCombo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCombo.RecordCount > 0 Then
                        
                        ZZInsumo = Trim(rstCombo!Insumo)
                        ZZCanti = rstCombo!Cantidad * ZZProduccion
                        
                        rstCombo.Close
                        
                        ZLugarII = ZLugarII + 1
                        ZCosto(ZLugarII, 1) = ZZInsumo
                        ZCosto(ZLugarII, 2) = Str$(ZZCanti)
                
                    End If
                
                Next ZZRenglon
            
            End If
            
        Next CicloII
        
        For CicloII = 1 To ZLugarII
        
            ZZCodigo = ZCosto(CicloII, 1)
            ZZCantidad = Val(ZCosto(CicloII, 2))
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Insumo"
            ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZZCodigo + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstInsumo.RecordCount > 0 Then
                ZZCosto = rstInsumo!Costo
                ZZMoneda = rstInsumo!Moneda
                rstInsumo.Close
            End If
            
            If ZZMoneda = 2 Then
                ZZCostoII = ZZCostoII + (ZZCantidad * ZZCosto)
                    Else
                ZZCostoI = ZZCostoI + (ZZCantidad * ZZCosto)
            End If
            
        Next CicloII
        
        
        If ZZCostoI <> 0 Then
            ZZCostoI = ZZCostoI / WWParidad
        End If
        
        ZZCosto = ZZCostoI + ZZCostoII
        
        ZZListaFecha = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + "0" + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + UCase(Trim(ZProducto(Ciclo, 2))) + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + UCase(Trim(ZProducto(Ciclo, 3))) + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + UCase(Trim(ZProducto(Ciclo, 4))) + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + UCase(Trim(ZProducto(Ciclo, 5))) + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + UCase(Trim(ZProducto(Ciclo, 6))) + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            ZZListaFecha = rstPrecios!Hasta
            rstPrecios.Close
        End If
        
        ZZfecha = Right$(Date$, 4) + Left$(Date$, 2) + Mid$(Date$, 4, 2)
        zzFechaCierre = Right$(ZZListaFecha, 4) + Mid$(ZZListaFecha, 4, 2) + Left$(ZZListaFecha, 2)
        
        If ZZfecha < zzFechaCierre Then
            ZZListaFecha = ""
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Costo = " + "'" + Str$(ZZCosto) + "',"
        ZSql = ZSql + " ListaFecha = " + "'" + ZZListaFecha + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZProducto(Ciclo, 1) + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
                                       



    Listado.WindowTitle = "Listado de Articulos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    ZZDesde = "0.01"
    ZZHasta = "99999999"
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Linea, Articulo.Descripcion, Articulo.Stock, Articulo.Costo, Articulo.Fecha " _
            + "From " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "Articulo.Linea >= '" + Desde.Text + "' AND " _
            + "Articulo.Linea <= '" + Hasta.Text + "' AND " _
            + "articulo.Stock >= " + ZZDesde + " AND " _
            + "Articulo.Stock <= " + ZZHasta
    
    Listado.Connect = Connect()
    
    Uno = "{Articulo.Linea} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Dos = " and {Articulo.Stock} in " + ZZDesde + " to " + ZZHasta
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.ReportFileName = "ListaStockVALORA.rpt"
    
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaMinimoInsumos.Hide
    Unload Me
    MenuVen.Show
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
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Sub Form_Load()
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
















