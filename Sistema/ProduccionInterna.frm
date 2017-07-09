VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProduccionInterna 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Produccion Interna"
   ClientHeight    =   8175
   ClientLeft      =   195
   ClientTop       =   375
   ClientWidth     =   12285
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   12285
   Visible         =   0   'False
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo F8"
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
      Left            =   11400
      MouseIcon       =   "ProduccionInterna.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ProduccionInterna.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salida"
      Top             =   5400
      Width           =   855
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba F1"
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
      Left            =   11400
      MouseIcon       =   "ProduccionInterna.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "ProduccionInterna.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia F3"
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
      Left            =   11400
      MouseIcon       =   "ProduccionInterna.frx":1298
      MousePointer    =   99  'Custom
      Picture         =   "ProduccionInterna.frx":15A2
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2040
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
      Left            =   11400
      MouseIcon       =   "ProduccionInterna.frx":1DE4
      MousePointer    =   99  'Custom
      Picture         =   "ProduccionInterna.frx":20EE
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Consulta de Datos"
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
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
      Left            =   11400
      MouseIcon       =   "ProduccionInterna.frx":2930
      MousePointer    =   99  'Custom
      Picture         =   "ProduccionInterna.frx":2C3A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Menu Principal"
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   15
      Top             =   2400
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   14
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   4200
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2880
      Width           =   375
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   9
      Top             =   480
      Width           =   5295
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
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   6855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Height          =   1260
      Left            =   2280
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3720
      TabIndex        =   5
      Top             =   120
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
   Begin VB.TextBox Numero 
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
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9360
      TabIndex        =   2
      Top             =   7920
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
      Height          =   1740
      ItemData        =   "ProduccionInterna.frx":347C
      Left            =   120
      List            =   "ProduccionInterna.frx":3483
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   16
      Top             =   2400
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4815
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8493
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label3 
      Caption         =   "Observaciones"
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
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Movimiento"
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
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgProduccionInterna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private XIndice As Single
Dim Vector(100, 10) As String
Private Auxi As String
Private XColor As String
Private XArticulo As String
Private WTipopro As Integer

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Consulta_Click()

    Opcion.Clear
    Opcion.AddItem "Articulo"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    Call Opcion_Click
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
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
                            IngresaItem = Trim(!Codigo) + "       " + !Descripcion
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
            Ayuda.SetFocus
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub CmdClose_Click()
    PrgProduccionInterna.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Graba_Click()

    If Val(Numero.Text) = 0 Then
    
        ZSql = ""
        ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
        ZSql = ZSql + " FROM Produccion"
        spProduccion = ZSql
        Set rstProduccion = db.OpenRecordset(spProduccion, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduccion.RecordCount > 0 Then
            rstProduccion.MoveLast
            ZUltimo = IIf(IsNull(rstProduccion!NumeroMayor), "0", rstProduccion!NumeroMayor)
            Numero.Text = ZUltimo + 1
            rstProduccion.Close
        End If
    
    End If

    For WRenglon = 1 To 100
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Produccion"
        ZSql = ZSql + " Where Produccion.Clave = " + "'" + WClave + "'"
        spProduccion = ZSql
        Set rstProduccion = db.OpenRecordset(spProduccion, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduccion.RecordCount > 0 Then
            
            WArticulo = rstProduccion!Articulo
            WCantidad = rstProduccion!Cantidad
            
            WWInsumoII = ""
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WWInsumoII = IIf(IsNull(rstArticulo!InsumoII), "", rstArticulo!InsumoII)
                rstArticulo.Close
            End If
                                            
            If Trim(WInsumoII) = "" Then
                
                If WCantidad > 0 Then
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Stock = Stock - " + "'" + Str$(WCantidad) + "',"
                    ZSql = ZSql + " StockII = StockII - " + "'" + Str$(WCantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(WCantidad * -1) + "',"
                    ZSql = ZSql + " StockII = StockII + " + "'" + Str$(WCantidad * -1) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + WArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                ZZCombo = ""
                ZZArticulo = WArticulo
                ZZProduccion = WCantidad * -1
                
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
                        
                        ZZTipoProceso = IIf(IsNull(rstFormula!TipoProceso), "", rstFormula!TipoProceso)
                        
                        rstFormula.Close
                            
                        If Trim(UCase(ZZTipoProceso)) <> "T" Then
                                
                            If Trim(ZZZInsumo) <> "" Then
                            
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Insumo SET "
                                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(ZZCanti) + "',"
                                ZSql = ZSql + " StockII = StockII - " + "'" + Str$(ZZCanti) + "'"
                                ZSql = ZSql + " Where Codigo = " + "'" + ZZZInsumo + "'"
                                spInsumo = ZSql
                                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                                
                            End If
                                
                            If Trim(ZZZTerminado) <> "" Then
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Articulo SET "
                                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(ZZCanti) + "',"
                                ZSql = ZSql + " StockII = StockII - " + "'" + Str$(ZZCanti) + "'"
                                ZSql = ZSql + " Where Codigo = " + "'" + ZZZTerminado + "'"
                                spArticulo = ZSql
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
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
                            ZZTipoProceso = IIf(IsNull(rstCombo!TipoProceso), "", rstCombo!TipoProceso)
                            
                            rstCombo.Close
                            
                            If Trim(UCase(ZZTipoProceso)) <> "T" Then
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Insumo SET "
                                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(ZZCanti) + "',"
                                ZSql = ZSql + " StockII = StockII - " + "'" + Str$(ZZCanti) + "'"
                                ZSql = ZSql + " Where Codigo = " + "'" + ZZInsumo + "'"
                                spInsumo = ZSql
                                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
                        End If
                    
                    Next ZZRenglon
                
                End If
            
            End If
            
        End If
    
    Next WRenglon
    
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Produccion"
    ZSql = ZSql + " Where Produccion.Numero = " + "'" + Numero.Text + "'"
    spProduccion = ZSql
    Set rstProduccion = db.OpenRecordset(spProduccion, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "DELETE MovimientoInsumo"
    ZSql = ZSql + " Where MovimientoInsumo.Tipo = " + "'" + "3" + "'"
    ZSql = ZSql + " and MovimientoInsumo.Numero = " + "'" + Numero.Text + "'"
    spMovimientoInsumo = ZSql
    Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
    

    Renglon = 0
    WRenglon = 0
    ZZRenglonMov = 0
        
    For IRow = 1 To 100
            
        WVector1.Row = IRow
            
        WVector1.Col = 1
        Articulo = WVector1.Text
                    
        WVector1.Col = 3
        Cantidad = Val(WVector1.Text)
        
        If Cantidad <> 0 Then
                    
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Numero.Text)
            Call Ceros(Auxi1, 6)
                    
            ZZNumero = Numero.Text
            ZZRenglon = Auxi
            Rem ZZRenglon = Trim(ZZRenglon)
            ZZArticulo = Articulo
            ZZCantidad = Str$(Cantidad)
            ZZfecha = Fecha.Text
            ZZAuxiliar = "0"
            ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZZClave = Auxi1 + Auxi
            ZZObservaciones = Trim(Observaciones.Text)
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Produccion ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Auxiliar ,"
            ZSql = ZSql + "Observaciones )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZAuxiliar + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "')"
                            
            spProduccion = ZSql
            Set rstProduccion = db.OpenRecordset(spProduccion, dbOpenSnapshot, dbSQLPassThrough)
                
            
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
            
                If Val(ZZCantidad) > 0 Then
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Stock = Stock + " + "'" + ZZCantidad + "',"
                    ZSql = ZSql + " StockII = StockII + " + "'" + ZZCantidad + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + ZZArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Stock = Stock - " + "'" + Str$(Val(ZZCantidad) * -1) + "',"
                    ZSql = ZSql + " StockII = StockII - " + "'" + Str$(Val(ZZCantidad) * -1) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + ZZArticulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                
                ZZCombo = ""
                ZZArticulo = Articulo
                ZZProduccion = ZZCantidad
                
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
                        ZZCanti = ZZZCantidad * Val(ZZProduccion)
                        
                        ZZTipoProceso = IIf(IsNull(rstFormula!TipoProceso), "", rstFormula!TipoProceso)
                        
                        rstFormula.Close
                            
                        If Trim(UCase(ZZTipoProceso)) <> "T" Then
                                
                            If Trim(ZZZInsumo) <> "" Then
                            
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Insumo SET "
                                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(ZZCanti) + "',"
                                ZSql = ZSql + " StockII = StockII - " + "'" + Str$(ZZCanti) + "'"
                                ZSql = ZSql + " Where Codigo = " + "'" + ZZZInsumo + "'"
                                spInsumo = ZSql
                                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)

                                
                                Rem
                                Rem doty de alta el movimiento en listado
                                Rem
                                
                                ZZTipoMov = "03"
                                ZZNumeroMov = Numero.Text
                                ZZRenglonMov = ZZRenglonMov + 1
                                
                                Auxi1 = Numero.Text
                                Call Ceros(Auxi1, 6)
                                Auxi2 = Str$(ZZRenglonMov)
                                Call Ceros(Auxi2, 2)
                                ZZClaveMov = ZZTipoMov + Auxi1 + Auxi2
                                
                                ZSql = ""
                                ZSql = ZSql + "INSERT INTO MovimientoInsumo ("
                                ZSql = ZSql + "Clave ,"
                                ZSql = ZSql + "Tipo ,"
                                ZSql = ZSql + "Numero ,"
                                ZSql = ZSql + "Renglon ,"
                                ZSql = ZSql + "Insumo ,"
                                ZSql = ZSql + "Cantidad ,"
                                ZSql = ZSql + "Fecha ,"
                                ZSql = ZSql + "OrdFecha ,"
                                ZSql = ZSql + "Deposito ,"
                                ZSql = ZSql + "DepositoII ,"
                                ZSql = ZSql + "Concepto )"
                                ZSql = ZSql + "Values ("
                                ZSql = ZSql + "'" + ZZClaveMov + "',"
                                ZSql = ZSql + "'" + ZZTipoMov + "',"
                                ZSql = ZSql + "'" + ZZNumeroMov + "',"
                                ZSql = ZSql + "'" + Str$(ZZRenglonMov) + "',"
                                ZSql = ZSql + "'" + ZZZInsumo + "',"
                                ZSql = ZSql + "'" + Str$(ZZCanti * -1) + "',"
                                ZSql = ZSql + "'" + ZZfecha + "',"
                                ZSql = ZSql + "'" + ZZOrdFecha + "',"
                                ZSql = ZSql + "'" + "1" + "',"
                                ZSql = ZSql + "'" + "0" + "',"
                                ZSql = ZSql + "'" + "0" + "')"
                                                
                                spMovimientoInsumo = ZSql
                                Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
                                
                                
                            End If
                                
                            If Trim(ZZZTerminado) <> "" Then
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Articulo SET "
                                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(ZZCanti) + "',"
                                ZSql = ZSql + " StockII = StockII - " + "'" + Str$(ZZCanti) + "'"
                                ZSql = ZSql + " Where Codigo = " + "'" + ZZZTerminado + "'"
                                spArticulo = ZSql
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                            End If
                            
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
                            ZZTipoProceso = IIf(IsNull(rstCombo!TipoProceso), "", rstCombo!TipoProceso)
                            
                            rstCombo.Close
                            
                            If Trim(UCase(ZZTipoProceso)) <> "T" Then
                            
                                ZSql = ""
                                ZSql = ZSql + "UPDATE Insumo SET "
                                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(ZZCanti) + "',"
                                ZSql = ZSql + " StockII = StockII - " + "'" + Str$(ZZCanti) + "'"
                                ZSql = ZSql + " Where Codigo = " + "'" + ZZInsumo + "'"
                                spInsumo = ZSql
                                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                                

                                
                                Rem
                                Rem doty de alta el movimiento en listado
                                Rem
                                
                                ZZTipoMov = "03"
                                ZZNumeroMov = Numero.Text
                                ZZRenglonMov = ZZRenglonMov + 1
                                
                                Auxi1 = Numero.Text
                                Call Ceros(Auxi1, 6)
                                Auxi2 = Str$(ZZRenglonMov)
                                Call Ceros(Auxi2, 2)
                                ZZClaveMov = ZZTipoMov + Auxi1 + Auxi2
                                
                                ZSql = ""
                                ZSql = ZSql + "INSERT INTO MovimientoInsumo ("
                                ZSql = ZSql + "Clave ,"
                                ZSql = ZSql + "Tipo ,"
                                ZSql = ZSql + "Numero ,"
                                ZSql = ZSql + "Renglon ,"
                                ZSql = ZSql + "Insumo ,"
                                ZSql = ZSql + "Cantidad ,"
                                ZSql = ZSql + "Fecha ,"
                                ZSql = ZSql + "OrdFecha ,"
                                ZSql = ZSql + "Deposito ,"
                                ZSql = ZSql + "DepositoII ,"
                                ZSql = ZSql + "Concepto )"
                                ZSql = ZSql + "Values ("
                                ZSql = ZSql + "'" + ZZClaveMov + "',"
                                ZSql = ZSql + "'" + ZZTipoMov + "',"
                                ZSql = ZSql + "'" + ZZNumeroMov + "',"
                                ZSql = ZSql + "'" + Str$(ZZRenglonMov) + "',"
                                ZSql = ZSql + "'" + ZZInsumo + "',"
                                ZSql = ZSql + "'" + Str$(ZZCanti * -1) + "',"
                                ZSql = ZSql + "'" + ZZfecha + "',"
                                ZSql = ZSql + "'" + ZZOrdFecha + "',"
                                ZSql = ZSql + "'" + "1" + "',"
                                ZSql = ZSql + "'" + "0" + "',"
                                ZSql = ZSql + "'" + "0" + "')"
                                                
                                spMovimientoInsumo = ZSql
                                Set rstMovimientoInsumo = db.OpenRecordset(spMovimientoInsumo, dbOpenSnapshot, dbSQLPassThrough)
                                
                                
                            End If
                            
                        End If
                    
                    Next ZZRenglon
                
                End If
            
            End If
            
                
        End If
                                       
    Next IRow
    
    
    Rem T$ = "Movimiento de Stock"
    Rem m$ = "Desea Imprimir el Comprobante"
    Rem Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
    Rem If Respuestaaaaaa% = 6 Then
    Rem     Call Impresion
    Rem End If
                    
    Rem Call Limpia_Click
    
    m$ = "Grabacion realizada"
    aaaaaa% = MsgBox(m$, 0, "Archivo de Movimientos de Stock")
    
    
    Numero.SetFocus
        
End Sub

Private Sub Impresion()
                        
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Produccion.Numero, Produccion.Renglon, Produccion.fecha, Produccion.Articulo, Produccion.Cantidad, Produccion.Observaciones, Produccion.Stock, " _
            + "Articulo.Descripcion, Articulo.Stock " _
            + "From " _
            + DSQ + ".dbo.Produccion Produccion, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "Produccion.Articulo = Articulo.Codigo AND " _
            + "Produccion.Numero >= " + Numero.Text + " AND " _
            + "Produccion.Numero <= " + Numero.Text
    
    Listado.Connect = Connect()
    
    Uno = "{Produccion.Numero} in " + Numero.Text + " to " + Numero.Text
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
            
    Listado.ReportFileName = "ImpreProduccion.rpt"
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    
    Listado.Action = 1

End Sub

Private Sub Limpia_Click()

    Call Limpia_Vector

    Numero.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    
    Renglon = 0
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    Rem ZSql = ZSql + " FROM Produccion"
    Rem sppRODUCCION = ZSql
    Rem Set rstPRODUCCION = db.OpenRecordset(sppRODUCCION, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPRODUCCION.RecordCount > 0 Then
    Rem     rstPRODUCCION.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstPRODUCCION!NumeroMayor), "0", rstPRODUCCION!NumeroMayor)
    Rem     Numero.Text = ZUltimo + 1
    Rem     rstPRODUCCION.Close
    Rem End If
    
    Numero.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Indice = Pantalla.ListIndex
            Claveven$ = WIndice.List(Indice)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Claveven$ + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstArticulo!Codigo
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            Ayuda.Visible = False
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Numero.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Observaciones.Text = ""
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    Rem ZSql = ZSql + " FROM Produccion"
    Rem sppRODUCCION = ZSql
    Rem Set rstPRODUCCION = db.OpenRecordset(sppRODUCCION, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPRODUCCION.RecordCount > 0 Then
    Rem     rstPRODUCCION.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstPRODUCCION!NumeroMayor), "0", rstPRODUCCION!NumeroMayor)
    Rem     Numero.Text = ZUltimo + 1
    Rem     rstPRODUCCION.Close
    Rem End If
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 100
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Produccion"
        ZSql = ZSql + " Where Produccion.Clave = " + "'" + WClave + "'"
        spProduccion = ZSql
        Set rstProduccion = db.OpenRecordset(spProduccion, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduccion.RecordCount > 0 Then
            
            Renglon = Renglon + 1
                
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = rstProduccion!Articulo
            Auxi1 = rstProduccion!Articulo
                
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###", Str$(rstProduccion!Cantidad))
            
            rstProduccion.Close
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                rstArticulo.Close
            End If
                    
        End If
    
    Next WRenglon
    
    WVector1.Col = 1
    WVector1.Row = 1

End Sub

Private Sub Numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 6)
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Produccion"
        ZSql = ZSql + " Where Produccion.Numero = " + "'" + Auxi + "'"
        spProduccion = ZSql
        Set rstProduccion = db.OpenRecordset(spProduccion, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduccion.RecordCount > 0 Then
            
            Fecha.Text = rstProduccion!Fecha
            Observaciones.Text = rstProduccion!Observaciones
            rstProduccion.Close
            
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
            
                Else
                
            WNumero = Numero.Text
            Numero.Text = WNumero
            Fecha.SetFocus
            
        End If
        
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Observaciones.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            aaaaaa% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub


Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
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
    
    Rem If XIndice = 0 And KeyAscii <> 13 Then
    Rem     Exit Sub
    Rem End If
    
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
                            IngresaItem = Trim(!Codigo) + "      " + !Descripcion
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


Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub Ultimo_Click()
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    ZSql = ZSql + " FROM Produccion"
    spProduccion = ZSql
    Set rstProduccion = db.OpenRecordset(spProduccion, dbOpenSnapshot, dbSQLPassThrough)
    If rstProduccion.RecordCount > 0 Then
        rstProduccion.MoveLast
        Numero.Text = IIf(IsNull(rstProduccion!NumeroMayor), "0", rstProduccion!NumeroMayor)
        rstProduccion.Close
        Call Numero_Keypress(13)
    End If

End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 3
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZInsumoII = IIf(IsNull(rstArticulo!InsumoII), "", rstArticulo!InsumoII)
                If Trim(ZZInsumoII) <> "" Then
                    rstArticulo.Close
                    m$ = "No se puede generar un movimiento para este articulo ya que el stock se lleva en el insumo asociado (" + ZZInsumoII + ")"
                    aaaaaa% = MsgBox(m$, 0, "Archivo de Produccion Interna")
                    WControl = "N"
                        Else
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                        Else
                WControl = "N"
            End If
            
        Case 3
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        WVector1.Col = 3
        WAuxi2 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 4
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 25
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Produccion"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

        Opcion.Clear
    
        Opcion.AddItem "Articulo"

        Rem Opcion.Visible = False
    
        Opcion.ListIndex = 0
    
        Call Opcion_Click
    
    End If
    
End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Graba_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub






















