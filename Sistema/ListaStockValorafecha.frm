VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaStockValoraFecha 
   AutoRedraw      =   -1  'True
   Caption         =   "Valorizacion de Stock"
   ClientHeight    =   4275
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4275
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
      Height          =   3855
      Left            =   960
      TabIndex        =   3
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
         MouseIcon       =   "ListaStockValorafecha.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaStockValorafecha.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salida"
         Top             =   2640
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
         MouseIcon       =   "ListaStockValorafecha.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaStockValorafecha.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2640
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
         MouseIcon       =   "ListaStockValorafecha.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaStockValorafecha.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2640
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
         TabIndex        =   5
         Text            =   " "
         Top             =   1560
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
         TabIndex        =   1
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   2520
         TabIndex        =   0
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
      Begin VB.Label Label4 
         Caption         =   "Fecha"
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
         Top             =   600
         Width           =   1695
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
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
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
         TabIndex        =   4
         Top             =   1080
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
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaStockValoraFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZPasa(10000, 2) As String

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
    ZSql = ZSql + " StockTrabajo = Stock" + ","
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFecha = WAno + WMes + WDia
    
    
    
    
    Erase ZZPasa
    ZLugar = 0
    ZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Estadistica"
    ZSql = ZSql + " Where Estadistica.OrdFecha > " + "'" + WFecha + "'"
    ZSql = ZSql + " Order by Estadistica.Articulo"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZPasa = 0 Then
                        ZPasa = 1
                        ZCorte = rstEstadistica!Articulo
                        ZSuma = 0
                    End If
                    
                    If ZCorte <> rstEstadistica!Articulo Then
                        ZLugar = ZLugar + 1
                        ZZPasa(ZLugar, 1) = ZCorte
                        ZZPasa(ZLugar, 2) = Str$(ZSuma)
                        ZCorte = rstEstadistica!Articulo
                        ZSuma = 0
                    End If
                        
                    ZSuma = ZSuma + rstEstadistica!Cantidad
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstEstadistica.Close
    End If
    
    If ZPasa <> 0 Then
        ZLugar = ZLugar + 1
        ZZPasa(ZLugar, 1) = ZCorte
        ZZPasa(ZLugar, 2) = Str$(ZSuma)
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZArticulo = ZZPasa(Ciclo, 1)
        ZCantidad = ZZPasa(Ciclo, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " StockTrabajo = StockTrabajo + " + "'" + ZCantidad + "'" + ""
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase ZZPasa
    ZLugar = 0
    ZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM MovStk"
    ZSql = ZSql + " Where MovStk.OrdFecha > " + "'" + WFecha + "'"
    ZSql = ZSql + " Order by MovStk.Articulo"
    spMovStk = ZSql
    Set rstMovstk = db.OpenRecordset(spMovStk, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovstk.RecordCount > 0 Then
    
        With rstMovstk
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZPasa = 0 Then
                        ZPasa = 1
                        ZCorte = rstMovstk!Articulo
                        ZSuma = 0
                    End If
                    
                    If ZCorte <> rstMovstk!Articulo Then
                        ZLugar = ZLugar + 1
                        ZZPasa(ZLugar, 1) = ZCorte
                        ZZPasa(ZLugar, 2) = Str$(ZSuma)
                        ZCorte = rstMovstk!Articulo
                        ZSuma = 0
                    End If
                        
                    ZSuma = ZSuma + rstMovstk!Cantidad
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstMovstk.Close
    End If
    
    If ZPasa <> 0 Then
        ZLugar = ZLugar + 1
        ZZPasa(ZLugar, 1) = ZCorte
        ZZPasa(ZLugar, 2) = Str$(ZSuma)
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZArticulo = ZZPasa(Ciclo, 1)
        ZCantidad = ZZPasa(Ciclo, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " StockTrabajo = StockTrabajo - " + "'" + ZCantidad + "'" + ""
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Listado.WindowTitle = "Valorizacion del Stock a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Stock, Articulo.Costo, Articulo.StockTrabajo, " _
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
                    + "Articulo.StockTrabajo > 0"
            
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
                
    Uno = "{Articulo.Proveedor} in " + Desde.Text + " to " + Hasta.Text + ""
    Dos = " and {Articulo.StockTrabajo} > 0.00"
        
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
            
    Listado.ReportFileName = "ListaStockValorizaFecha.rpt"
        
        
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaStockValoraFecha.Hide
    Unload Me
    Menu231.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
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
        Fecha.SetFocus
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
















