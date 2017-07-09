VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaStockProveedor 
   AutoRedraw      =   -1  'True
   Caption         =   "Stock por Provedor"
   ClientHeight    =   4005
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4005
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
      Height          =   3615
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Minimo 
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
         TabIndex        =   11
         Text            =   " "
         Top             =   1800
         Width           =   1335
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
         Left            =   2520
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
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
         MouseIcon       =   "listastockProveedor.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "listastockProveedor.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salida"
         Top             =   2400
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
         MouseIcon       =   "listastockProveedor.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "listastockProveedor.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2400
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
         MouseIcon       =   "listastockProveedor.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "listastockProveedor.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2400
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
      Begin VB.Label Label4 
         Caption         =   "Minimo Ranking"
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
         TabIndex        =   12
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Desde Proveedor"
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
         Caption         =   "Desde Proveedor"
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
Attribute VB_Name = "PrgListaStockProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Producto As String
Private Costo As Double

Dim ZVector(10000) As String
Dim ZImpre(100, 2) As String
Dim ZCanti(100) As Double

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
    
    Auxi = Str$(Val(Right$(Date$, 4)) - 1)
    Call Ceros(Auxi, 4)
    
    ZHasta = "31" + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Mes = Left$(Date$, 2)
    Ano = Right$(Date$, 4)
    
    ZImpre(1, 1) = Mes
    ZImpre(1, 2) = Ano
    
    For Ciclo = 1 To 11
    
        Auxi = Str$(Val(Mes) - 1)
        Call Ceros(Auxi, 2)
        Mes = Auxi
        If Val(Mes) = 0 Then
            Mes = "12"
            Auxi = Str$(Val(Ano) - 1)
            Call Ceros(Auxi, 4)
            Ano = Auxi
        End If
        ZImpre(Ciclo + 1, 1) = Mes
        ZImpre(Ciclo + 1, 2) = Ano
        
        If Ciclo = 11 Then
            ZDesde = "01" + "/" + Mes + "/" + Ano
        End If
        
    Next Ciclo
    
    ZLugar = 0
    Erase ZVector

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Proveedor >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Articulo.Proveedor <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    If Tipo.ListIndex = 1 Or !Stock <> 0 Then
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = !Codigo
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZArticulo = ZVector(Ciclo)
        Erase ZCanti
        
        ZEmbarque = 0
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenImportacion"
        ZSql = ZSql + " Where OrdenImportacion.Articulo = " + "'" + ZArticulo + "'"
        ZSql = ZSql + " and OrdenImportacion.Estado = 0"
        spOrdenImportacion = ZSql
        Set rstOrdenImportacion = db.OpenRecordset(spOrdenImportacion, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenImportacion.RecordCount > 0 Then
            With rstOrdenImportacion
                .MoveFirst
                Do
                    If .EOF = False Then
                        ZEmbarque = ZEmbarque + rstOrdenImportacion!Cantidad
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrdenImportacion.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Estadistica"
        ZSql = ZSql + " Where Estadistica.OrdFecha >= " + "'" + ZDesde + "'"
        ZSql = ZSql + " and Estadistica.OrdFecha <= " + "'" + ZHasta + "'"
        ZSql = ZSql + " and Estadistica.Articulo = " + "'" + ZArticulo + "'"
        ZSql = ZSql + " Order by Estadistica.Clave"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
            With rstEstadistica
                .MoveFirst
                Do
                    If .EOF = False Then
                        For CicloII = 1 To 12
                            If ZImpre(CicloII, 1) = Mid$(!Fecha, 4, 2) And ZImpre(CicloII, 2) = Mid$(!Fecha, 7, 4) Then
                                ZCanti(CicloII) = ZCanti(CicloII) + !Cantidad
                                Exit For
                            End If
                        Next CicloII
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEstadistica.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Embarque = " + "'" + Str$(ZEmbarque) + "',"
        ZSql = ZSql + " Venta1 = " + "'" + Str$(ZCanti(1)) + "',"
        ZSql = ZSql + " Venta2 = " + "'" + Str$(ZCanti(2)) + "',"
        ZSql = ZSql + " Venta3 = " + "'" + Str$(ZCanti(3)) + "',"
        ZSql = ZSql + " Venta4 = " + "'" + Str$(ZCanti(4)) + "',"
        ZSql = ZSql + " Venta5 = " + "'" + Str$(ZCanti(5)) + "',"
        ZSql = ZSql + " Venta6 = " + "'" + Str$(ZCanti(6)) + "',"
        ZSql = ZSql + " Venta7 = " + "'" + Str$(ZCanti(7)) + "',"
        ZSql = ZSql + " Venta8 = " + "'" + Str$(ZCanti(8)) + "',"
        ZSql = ZSql + " Venta9 = " + "'" + Str$(ZCanti(9)) + "',"
        ZSql = ZSql + " Venta10 = " + "'" + Str$(ZCanti(10)) + "',"
        ZSql = ZSql + " Venta11 = " + "'" + Str$(ZCanti(11)) + "',"
        ZSql = ZSql + " Venta12 = " + "'" + Str$(ZCanti(12)) + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Impre1 = " + "'" + ZImpre(1, 1) + "/" + ZImpre(1, 2) + "',"
    ZSql = ZSql + " Impre2 = " + "'" + ZImpre(2, 1) + "/" + ZImpre(2, 2) + "',"
    ZSql = ZSql + " Impre3 = " + "'" + ZImpre(3, 1) + "/" + ZImpre(3, 2) + "',"
    ZSql = ZSql + " Impre4 = " + "'" + ZImpre(4, 1) + "/" + ZImpre(4, 2) + "',"
    ZSql = ZSql + " Impre5 = " + "'" + ZImpre(5, 1) + "/" + ZImpre(5, 2) + "',"
    ZSql = ZSql + " Impre6 = " + "'" + ZImpre(6, 1) + "/" + ZImpre(6, 2) + "',"
    ZSql = ZSql + " Impre7 = " + "'" + ZImpre(7, 1) + "/" + ZImpre(7, 2) + "',"
    ZSql = ZSql + " Impre8 = " + "'" + ZImpre(8, 1) + "/" + ZImpre(8, 2) + "',"
    ZSql = ZSql + " Impre9 = " + "'" + ZImpre(9, 1) + "/" + ZImpre(9, 2) + "',"
    ZSql = ZSql + " Impre10 = " + "'" + ZImpre(10, 1) + "/" + ZImpre(10, 2) + "',"
    ZSql = ZSql + " Impre11 = " + "'" + ZImpre(11, 1) + "/" + ZImpre(11, 2) + "',"
    ZSql = ZSql + " Impre12 = " + "'" + ZImpre(12, 1) + "/" + ZImpre(12, 2) + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    Listado.WindowTitle = "Stock por Proveedor"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Select Case Tipo.ListIndex
        Case 0
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Minimo, Articulo.Stock, Articulo.Venta1, Articulo.Venta2, Articulo.Venta3, Articulo.Venta4, Articulo.Venta5, Articulo.Venta6, Articulo.Venta7, Articulo.Venta8, Articulo.Venta9, Articulo.Venta10, Articulo.Venta11, Articulo.Venta12, Articulo.Impre1, Articulo.Impre2, Articulo.Impre3, Articulo.Impre4, Articulo.Impre5, Articulo.Impre6, Articulo.Impre7, Articulo.Impre8, Articulo.Impre9, Articulo.Impre10, Articulo.Impre11, Articulo.Impre12, " _
                    + "Auxiliar.Nombre, " _
                    + "Familia.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Familia Familia, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Articulo.Grupo = Familia.Codigo AND " _
                    + "Articulo.Proveedor = Proveedor.Proveedor AND " _
                    + "Articulo.Proveedor >= " + Desde.Text + " AND " _
                    + "Articulo.Proveedor <= " + Hasta.Text + " AND " _
                    + "Articulo.Stock > 0"
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Stock} > 0"
            Dos = " and {Articulo.Proveedor} in " + Desde.Text + " to " + Hasta.Text
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockProveedor.rpt"
            Listado.Action = 1
            
        Case 1
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Minimo, Articulo.Stock, Articulo.Venta1, Articulo.Venta2, Articulo.Venta3, Articulo.Venta4, Articulo.Venta5, Articulo.Venta6, Articulo.Venta7, Articulo.Venta8, Articulo.Venta9, Articulo.Venta10, Articulo.Venta11, Articulo.Venta12, Articulo.Impre1, Articulo.Impre2, Articulo.Impre3, Articulo.Impre4, Articulo.Impre5, Articulo.Impre6, Articulo.Impre7, Articulo.Impre8, Articulo.Impre9, Articulo.Impre10, Articulo.Impre11, Articulo.Impre12, Articulo.Embarque, " _
                    + "Auxiliar.Nombre, " _
                    + "Familia.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Familia Familia, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Articulo.Grupo = Familia.Codigo AND " _
                    + "Articulo.Proveedor = Proveedor.Proveedor AND " _
                    + "Articulo.Proveedor >= " + Desde.Text + " AND " _
                    + "Articulo.Proveedor <= " + Hasta.Text
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Proveedor} in " + Desde.Text + " to " + Hasta.Text
            Dos = ""
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockProveedorII.rpt"
            Listado.Action = 1
            
        Case 2
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Minimo, Articulo.Stock, Articulo.Venta1, Articulo.Venta2, Articulo.Venta3, Articulo.Venta4, Articulo.Venta5, Articulo.Venta6, Articulo.Venta7, Articulo.Venta8, Articulo.Venta9, Articulo.Venta10, Articulo.Venta11, Articulo.Venta12, Articulo.Impre1, Articulo.Impre2, Articulo.Impre3, Articulo.Impre4, Articulo.Impre5, Articulo.Impre6, Articulo.Impre7, Articulo.Impre8, Articulo.Impre9, Articulo.Impre10, Articulo.Impre11, Articulo.Impre12, " _
                    + "Auxiliar.Nombre, " _
                    + "Familia.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Familia Familia, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Articulo.Grupo = Familia.Codigo AND " _
                    + "Articulo.Proveedor = Proveedor.Proveedor AND " _
                    + "Articulo.Proveedor >= " + Desde.Text + " AND " _
                    + "Articulo.Proveedor <= " + Hasta.Text + " AND " _
                    + "Articulo.Stock <= 0"
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
                    
            Uno = "{Articulo.Stock} <= 0"
            Dos = " and {Articulo.Proveedor} in " + Desde.Text + " to " + Hasta.Text
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockProveedorSdt.rpt"
            Listado.Action = 1
            
        Case Else
            Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Minimo, Articulo.Stock, Articulo.Venta1, Articulo.Venta2, Articulo.Venta3, Articulo.Venta4, Articulo.Venta5, Articulo.Venta6, Articulo.Venta7, Articulo.Venta8, Articulo.Venta9, Articulo.Venta10, Articulo.Venta11, Articulo.Venta12, Articulo.Impre1, Articulo.Impre2, Articulo.Impre3, Articulo.Impre4, Articulo.Impre5, Articulo.Impre6, Articulo.Impre7, Articulo.Impre8, Articulo.Impre9, Articulo.Impre10, Articulo.Impre11, Articulo.Impre12, Articulo.Embarque, " _
                    + "Auxiliar.Nombre, " _
                    + "Familia.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Familia Familia, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Articulo.Grupo = Familia.Codigo AND " _
                    + "Articulo.Proveedor = Proveedor.Proveedor AND " _
                    + "Articulo.Proveedor >= " + Desde.Text + " AND " _
                    + "Articulo.Proveedor <= " + Hasta.Text
            
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
            
            Minimo.Text = Str$(Val(Minimo.Text))
                    
            Uno = "{Articulo.Proveedor} in " + Desde.Text + " to " + Hasta.Text
            Dos = " and {@Venta} >= " + Minimo.Text
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
                
            Listado.ReportFileName = "ListaStockProveedorRanking.rpt"
            Listado.Action = 1
        
    End Select
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaStockProveedor.Hide
    Unload Me
    Menu231.Show
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

    Tipo.Clear
    
    Tipo.AddItem "Stock"
    Tipo.AddItem "Completo"
    Tipo.AddItem "SDT"
    Tipo.AddItem "Ranking"
    
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
















