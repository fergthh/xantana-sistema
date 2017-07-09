VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaDisponible 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de stock de Insumos"
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
      Begin VB.CheckBox Sele 
         Caption         =   "Solo Stock"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
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
         MouseIcon       =   "ListaDisponible.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaDisponible.frx":030A
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
         MouseIcon       =   "ListaDisponible.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaDisponible.frx":0E56
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
         MouseIcon       =   "ListaDisponible.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaDisponible.frx":19A2
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
         MaxLength       =   50
         TabIndex        =   4
         Text            =   " "
         Top             =   840
         Width           =   1935
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
         MaxLength       =   50
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta"
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
         Caption         =   "Desde "
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
Attribute VB_Name = "PrgListaDisponible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Producto As String
Private Costo As Double


Dim WTrabajo(10000, 10) As String
Dim WLugar As Integer

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
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Insumo SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + "1" + "',"
    ZSql = ZSql + " Venta = " + "'" + "0" + "',"
    ZSql = ZSql + " Compra = " + "'" + "0" + "',"
    ZSql = ZSql + " Disponible = " + "'" + "0" + "'"
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    Erase WTrabajo
    WLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Order by Pedido.Clave"
        
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                
                    Impo1 = 0
                    Impo2 = 0
                    Impo3 = 0
                    Impo4 = 0
                    
                    ZZArticulo = rstPedido!Articulo
                    ZZCantidad = rstPedido!Cantidad
                    ZZFacturado = rstPedido!facturado
                    ZZAjuste = rstPedido!Ajuste
                    ZZImpo = ZZCantidad - ZZFacturado - ZZAjuste
                    If ZZImpo < 0 Then
                        ZZImpo = 0
                    End If
                        
                    If ZZImpo > 0 Then
                        WLugar = WLugar + 1
                        WTrabajo(WLugar, 1) = ZZArticulo
                        WTrabajo(WLugar, 2) = Str$(ZZImpo)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
            
        End With
        
        rstPedido.Close
    End If
    
    
    For Ciclo = 1 To WLugar
    
                    
        ZZArticulo = Trim(WTrabajo(Ciclo, 1))
        ZZCantidad = Val(WTrabajo(Ciclo, 2))
        
        WWInsumoII = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WWInsumoII = IIf(IsNull(rstArticulo!InsumoII), "", rstArticulo!InsumoII)
            rstArticulo.Close
        End If
                                        
        If Trim(WInsumoII) = "" Then
    
            ZZCombo = ""
            ZZProduccion = ZZCantidad
            
            For ZZRenglon = 1 To 100
                
                Auxi1 = ZZRenglon
                Call Ceros(Auxi1, 2)
                
                ZZCodigo = Trim(ZZArticulo)
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
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Insumo SET "
                        ZSql = ZSql + " Venta = Venta + " + "'" + Str$(ZZCanti) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZZZInsumo + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                        
                    If Trim(ZZZTerminado) <> "" Then
                        WLugar = WLugar + 1
                        WTrabajo(WLugar, 1) = ZZZTerminado
                        WTrabajo(WLugar, 2) = Str$(ZZCanti)
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
                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Insumo SET "
                        ZSql = ZSql + " Venta = Venta + " + "'" + Str$(ZZCanti) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZZInsumo + "'"
                        spInsumo = ZSql
                        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                
                    End If
                
                Next ZZRenglon
            
            End If
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "UPDATE Insumo SET "
            ZSql = ZSql + " Venta = Venta - " + "'" + Str$(Cantidad) + "',"
            ZSql = ZSql + " StockII = StockII - " + "'" + Str$(Cantidad) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WInsumoII + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
    Next Ciclo
                    
    
    
    Erase WTrabajo
    WLugar = 0
            
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Orden"
    ZSql = ZSql + " Where Orden.Cantidad > Orden.Pedida"
    ZSql = ZSql + " Order by Orden.Clave"
    spOrden = ZSql
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        With rstOrden
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZCantidad = rstOrden!Cantidad - rstOrden!pedida - rstOrden!Ajuste
                    If ZZCantidad > 0 Then
                        WLugar = WLugar + 1
                        WTrabajo(WLugar, 1) = rstOrden!Insumo
                        WTrabajo(WLugar, 2) = Str$(ZZCantidad)
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOrden.Close
    End If
    
    
    
    
    For Ciclo = 1 To WLugar
    
                    
        ZZInsumo = Trim(WTrabajo(Ciclo, 1))
        ZZCantidad = Val(WTrabajo(Ciclo, 2))
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Insumo SET "
        ZSql = ZSql + " Compra = Compra + " + "'" + Str$(ZZCantidad) + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZZInsumo + "'"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
    Next Ciclo
                    
    
    
    
    
    
    
    
    
    
    Listado.WindowTitle = "Listado de Insumo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Insumo.Codigo, Insumo.Descripcion, Insumo.Stock, Insumo.Venta, Insumo.Compra " _
            + "From " _
            + DSQ + ".dbo.Insumo Insumo " _
            + "Where " _
            + "Insumo.Codigo >= '" + Desde.Text + "' AND " _
            + "Insumo.Codigo <= '" + Hasta.Text + "'"

    Listado.Connect = Connect()
    
    Uno = "{Insumo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Rem Dos = " and ({Insumo.Stock} <> 0 or {Insumo.Venta} <> 0 or {Insumo.Compra} <> 0)"
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.ReportFileName = "ListaDisponible.rpt"
    
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgListaDisponible.Hide
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
            Call Cancela_Click
        Case Else
    End Select
End Sub
















