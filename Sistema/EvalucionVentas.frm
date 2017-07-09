VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaEvolucionVentas 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Evolucion de Ventas"
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
         MouseIcon       =   "EvalucionVentas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "EvalucionVentas.frx":030A
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
         MouseIcon       =   "EvalucionVentas.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "EvalucionVentas.frx":0E56
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
         MouseIcon       =   "EvalucionVentas.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "EvalucionVentas.frx":19A2
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
      Begin VB.Label NombreII 
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
      Begin VB.Label NombreI 
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
Attribute VB_Name = "PrgListaEvolucionVentas"
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
    
    For Ciclo = 1 To 5
    
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
        
        If Ciclo = 5 Then
            ZDesde = "01" + "/" + Mes + "/" + Ano
        End If
        
    Next Ciclo
    
    ZLugar = 0
    Erase ZVector
    
    If Tipo.ListIndex = 0 Then

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
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = !Codigo
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
        End If
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Grupo >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Articulo.Grupo <= " + "'" + Hasta.Text + "'"
        ZSql = ZSql + " Order by Articulo.Codigo"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = !Codigo
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
        End If
            
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZArticulo = Trim(ZVector(Ciclo))
        Erase ZCanti
        
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
                        For CicloII = 1 To 6
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
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Embarque = " + "'" + Str$(ZEmbarque) + "',"
        ZSql = ZSql + " Venta6 = " + "'" + Str$(ZCanti(1)) + "',"
        ZSql = ZSql + " Venta5 = " + "'" + Str$(ZCanti(2)) + "',"
        ZSql = ZSql + " Venta4 = " + "'" + Str$(ZCanti(3)) + "',"
        ZSql = ZSql + " Venta3 = " + "'" + Str$(ZCanti(4)) + "',"
        ZSql = ZSql + " Venta2 = " + "'" + Str$(ZCanti(5)) + "',"
        ZSql = ZSql + " Venta1 = " + "'" + Str$(ZCanti(6)) + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Impre6 = " + "'" + ZImpre(1, 1) + "/" + ZImpre(1, 2) + "',"
    ZSql = ZSql + " Impre5 = " + "'" + ZImpre(2, 1) + "/" + ZImpre(2, 2) + "',"
    ZSql = ZSql + " Impre4 = " + "'" + ZImpre(3, 1) + "/" + ZImpre(3, 2) + "',"
    ZSql = ZSql + " Impre3 = " + "'" + ZImpre(4, 1) + "/" + ZImpre(4, 2) + "',"
    ZSql = ZSql + " Impre2 = " + "'" + ZImpre(5, 1) + "/" + ZImpre(5, 2) + "',"
    ZSql = ZSql + " Impre1 = " + "'" + ZImpre(6, 1) + "/" + ZImpre(6, 2) + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Auxi6 = " + "'" + ZImpre(1, 1) + "/" + ZImpre(1, 2) + "',"
    ZSql = ZSql + " Auxi5 = " + "'" + ZImpre(2, 1) + "/" + ZImpre(2, 2) + "',"
    ZSql = ZSql + " Auxi4 = " + "'" + ZImpre(3, 1) + "/" + ZImpre(3, 2) + "',"
    ZSql = ZSql + " Auxi3 = " + "'" + ZImpre(4, 1) + "/" + ZImpre(4, 2) + "',"
    ZSql = ZSql + " Auxi2 = " + "'" + ZImpre(5, 1) + "/" + ZImpre(5, 2) + "',"
    ZSql = ZSql + " Auxi1 = " + "'" + ZImpre(6, 1) + "/" + ZImpre(6, 2) + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " PrecioList = 1"
    Rem ZSql = ZSql + " PrecioList = Venta1 + Venta2 + Venta3 + Venta4 + Venta5 +  Venta6"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Reposicion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Tipo.ListIndex = 0 Then
    
        Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Stock, Articulo.Venta1, Articulo.Venta2, Articulo.Venta3, Articulo.Venta4, Articulo.Venta5, Articulo.Venta6, Articulo.PrecioList, Articulo.Embarque,  " _
                    + "Auxiliar.Nombre, Auxiliar.Auxi1, Auxiliar.Auxi2, Auxiliar.Auxi3, Auxiliar.Auxi4, Auxiliar.Auxi5, Auxiliar.Auxi6, " _
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
                    + "Articulo.PrecioList > 0"
            
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
                
        Uno = "{Articulo.PrecioList} > 0"
        Dos = " and {Articulo.Proveedor} in " + Desde.Text + " to " + Hasta.Text
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
            
        Listado.ReportFileName = "ListaVentaSemestralProveedor.rpt"
        Listado.Action = 1
            
                Else
    
        Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Stock, Articulo.Venta1, Articulo.Venta2, Articulo.Venta3, Articulo.Venta4, Articulo.Venta5, Articulo.Venta6, Articulo.PrecioList, Articulo.Embarque, " _
                    + "Auxiliar.Nombre, Auxiliar.Auxi1, Auxiliar.Auxi2, Auxiliar.Auxi3, Auxiliar.Auxi4, Auxiliar.Auxi5, Auxiliar.Auxi6, " _
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
                    + "Articulo.Grupo >= " + Desde.Text + " AND " _
                    + "Articulo.Grupo <= " + Hasta.Text + " AND " _
                    + "Articulo.PrecioList > 0"
            
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
                
        Uno = "{Articulo.PrecioList} > 0"
        Dos = " and {Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
            
        Listado.ReportFileName = "ListaVentaSemestralGrupo.rpt"
        Listado.Action = 1
                
    End If
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaEvolucionVentas.Hide
    Unload Me
    Menu2.Show
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
    
    Tipo.AddItem "Proveedor"
    Tipo.AddItem "Grupo"
    
    Tipo.ListIndex = 0
    
    NombreI.Caption = "Desde Proveedor"
    NombreII.Caption = "Hasta Proveedor"

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

Private Sub Tipo_click()
    If Tipo.ListIndex = 0 Then
        NombreI.Caption = "Desde Proveedor"
        NombreII.Caption = "Hasta Proveedor"
            Else
        NombreI.Caption = "Desde Grupo"
        NombreII.Caption = "Hasta Grupo"
    End If
End Sub
