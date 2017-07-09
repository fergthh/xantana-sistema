VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form prgBusquedaArti 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Articulos"
   ClientHeight    =   8130
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   8130
   ScaleWidth      =   11790
   Visible         =   0   'False
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
      Left            =   9480
      MouseIcon       =   "busquedaarti.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "busquedaarti.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Salida"
      Top             =   6480
      Width           =   735
   End
   Begin VB.Frame PantaArticulo 
      Height          =   6615
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   7815
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
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   6255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   11033
         _Version        =   327680
         BackColor       =   16777152
      End
   End
   Begin VB.TextBox Tamano 
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
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   12
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Calidad 
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
      Left            =   4680
      MaxLength       =   10
      TabIndex        =   11
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Fragancia 
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
      Left            =   3600
      MaxLength       =   10
      TabIndex        =   10
      Text            =   " "
      Top             =   0
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
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   9
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox DescripcionII 
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   8
      Top             =   720
      Width           =   5535
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
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
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
      Height          =   1740
      Left            =   8640
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
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
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Articulo.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Clientes"
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
      Left            =   8400
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Descripcion 
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   2
      Top             =   360
      Width           =   5535
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
      Height          =   5820
      ItemData        =   "busquedaarti.frx":0B4C
      Left            =   8040
      List            =   "busquedaarti.frx":0B53
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   7680
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5040
      Width           =   1935
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
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "prgBusquedaArti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZPrecio  As Double
Dim ZZMargen As Double
Dim ZZFoto As Image
Dim ZZTextil As Integer
Dim ZZCodAnt As String

Dim WMovi(20000, 3) As String


Sub Imprime_Datos()
    
    PantaArticulo.Visible = False
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.LInea = " + "'" + Linea.Text + "'"
    ZSql = ZSql + " and Articulo.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and Articulo.fragancia = " + "'" + Fragancia.Text + "'"
    ZSql = ZSql + " and Articulo.Calidad = " + "'" + Calidad.Text + "'"
    ZSql = ZSql + " and Articulo.Tamano = " + "'" + Tamano.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Descripcion.Text = Trim(rstArticulo!Descripcion)
        DescripcionII.Text = Trim(rstArticulo!DescripcionII)
        rstArticulo.Close
    End If
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
    Linea.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    
    Linea.Text = ""
    Tipo.Text = ""
    Fragancia.Text = ""
    Calidad.Text = ""
    Tamano.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    
    Call Limpia_Vector
    
    PantaArticulo.Visible = True
    
    
    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click
    
    
    Linea.SetFocus
End Sub

Private Sub cmdClose_Click()
    ZZPasaLinea = Linea.Text
    ZZPasaTipo = Tipo.Text
    ZZPasaFragancia = Fragancia.Text
    ZZPasaCalidad = Calidad.Text
    ZZPasaTamaño = Tamano.Text
    prgBusquedaArti.Hide
    Unload Me
    Select Case ZZPasaProcesoII
        Case 1
            PrgClienteBonifica.Show
        Case 2
            PrgPedido.Show
        Case 3
            PrgPto.Show
        Case Else
            MenuVen.Show
    End Select
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
            Call Tipo_DblClick
            Call Busqueda
            Tipo.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Linea.Text = ""
        Call Busqueda
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
            Call Fragancia_DblClick
            Call Busqueda
            Fragancia.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        Call Busqueda
    End If
End Sub

Private Sub Fragancia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fragancia.Text = UCase(Fragancia.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Fragancia"
        ZSql = ZSql + " Where Fragancia.Codigo = " + "'" + Fragancia.Text + "'"
        spFragancia = ZSql
        Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
        If rstFragancia.RecordCount > 0 Then
            rstFragancia.Close
            Call Calidad_DblClick
            Call Busqueda
            Calidad.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fragancia.Text = ""
        Call Busqueda
    End If
End Sub

Private Sub Calidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Calidad.Text = UCase(Calidad.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM calidad"
        ZSql = ZSql + " Where calidad.Codigo = " + "'" + Calidad.Text + "'"
        spCalidad = ZSql
        Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
        If rstCalidad.RecordCount > 0 Then
            rstCalidad.Close
            Call Tamano_DblClick
            Call Busqueda
            Tamano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Calidad.Text = ""
        Call Busqueda
    End If
End Sub

Private Sub Tamano_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tamano.Text = UCase(Tamano.Text)
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Tamano"
        ZSql = ZSql + " Where Tamano.Codigo = " + "'" + Tamano.Text + "'"
        spTamano = ZSql
        Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
        If rstTamano.RecordCount > 0 Then
            rstTamano.Close
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.LInea = " + "'" + Linea.Text + "'"
            ZSql = ZSql + " and Articulo.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Articulo.fragancia = " + "'" + Fragancia.Text + "'"
            ZSql = ZSql + " and Articulo.Calidad = " + "'" + Calidad.Text + "'"
            ZSql = ZSql + " and Articulo.Tamano = " + "'" + Tamano.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
                Call Imprime_Datos
                Descripcion.SetFocus
                    Else
                WLinea = Linea.Text
                WTipo = Tipo.Text
                WFragancia = Fragancia.Text
                WCalidad = Calidad.Text
                WTamano = Tamano.Text
                CmdLimpiar_Click
                Linea.Text = WLinea
                Tipo.Text = WTipo
                Fragancia.Text = WFragancia
                Calidad.Text = WCalidad
                Tamano.Text = WTamano
                Call Busqueda
            End If
            Descripcion.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tamano.Text = ""
        Call Busqueda
    End If
End Sub
    
Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Lineaa"
     Opcion.AddItem "Tipos"
     Opcion.AddItem "Fragancias"
     Opcion.AddItem "Calidad"
     Opcion.AddItem "Tamano"

     Opcion.Visible = True
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
            
        Case 1
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
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Fragancia"
            ZSql = ZSql + " Order by Fragancia.Descripcion"
            spFragancia = ZSql
            Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
            If rstFragancia.RecordCount > 0 Then
                With rstFragancia
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
                rstFragancia.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Calidad"
            ZSql = ZSql + " Order by Calidad.Descripcion"
            spCalidad = ZSql
            Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
            If rstCalidad.RecordCount > 0 Then
                With rstCalidad
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
                rstCalidad.Close
            End If
            
        Case 4
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Tamano"
            ZSql = ZSql + " Order by Tamano.Descripcion"
            spTamano = ZSql
            Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
            If rstTamano.RecordCount > 0 Then
                With rstTamano
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
                rstTamano.Close
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
            Linea.Text = WIndice.List(Indice)
            Call LInea_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Tipo.Text = WIndice.List(Indice)
            Call Tipo_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            Fragancia.Text = WIndice.List(Indice)
            Call Fragancia_KeyPress(13)
            
        Case 3
            Indice = Pantalla.ListIndex
            Calidad.Text = WIndice.List(Indice)
            Call Calidad_KeyPress(13)
            
        Case 4
            Indice = Pantalla.ListIndex
            Tamano.Text = WIndice.List(Indice)
            Call Tamano_KeyPress(13)
            
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    Linea.Text = ""
    Tipo.Text = ""
    Fragancia.Text = ""
    Calidad.Text = ""
    Tamano.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    
    Call Limpia_Vector
    Call LInea_DblClick
    
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
            
        Case 1
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
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Fragancia"
            ZSql = ZSql + " Where Fragancia.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Fragancia.Descripcion"
            spFragancia = ZSql
            Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
            If rstFragancia.RecordCount > 0 Then
                With rstFragancia
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
                rstFragancia.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Calidad"
            ZSql = ZSql + " Where Calidad.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Calidad.Descripcion"
            spCalidad = ZSql
            Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
            If rstCalidad.RecordCount > 0 Then
                With rstCalidad
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
                rstCalidad.Close
            End If
            
        Case 4
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Tamano"
            ZSql = ZSql + " Where Tamano.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Tamano.Descripcion"
            spTamano = ZSql
            Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
            If rstTamano.RecordCount > 0 Then
                With rstTamano
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
                rstTamano.Close
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

Private Sub LInea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Tipo_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Fragancia_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub Calidad_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 3
    
    Call Opcion_Click

End Sub

Private Sub Tamano_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Lineas de Ventas"
    Opcion.AddItem "Despachos"

    Opcion.ListIndex = 4
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub LInea_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fragancia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Calidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tamano_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 3
    WVector1.FixedRows = 1
    WVector1.Rows = 10001
    
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
                WVector1.Text = "Codigo"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 5000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
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
    Rem modificar el Tamano de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Busqueda()

    Rem On Error GoTo WError
    
    Call Limpia_Vector
    PantaArticulo.Visible = True
    ZLugar = 0
    
    If Trim(Linea.Text) = "" And Trim(Tipo.Text) = "" And Trim(Fragancia.Text) = "" And Trim(Calidad.Text) = "" And Trim(Tamano.Text) = "" Then
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Descripcion <> ''"
    If Trim(Linea.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Linea = " + "'" + Linea.Text + "'"
    End If
    If Trim(Tipo.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Tipo = " + "'" + Tipo.Text + "'"
    End If
    If Trim(Fragancia.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Fragancia = " + "'" + Fragancia.Text + "'"
    End If
    If Trim(Calidad.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Calidad = " + "'" + Calidad.Text + "'"
    End If
    If Trim(Tamano.Text) <> "" Then
        ZSql = ZSql + " And Articulo.Tamano = " + "'" + Tamano.Text + "'"
    End If
    ZSql = ZSql + " Order by Articulo.Codigo"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    ZLugar = ZLugar + 1
                    WVector1.TextMatrix(ZLugar, 1) = !Codigo
                    WVector1.TextMatrix(ZLugar, 2) = !Descripcion
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If

End Sub


Private Sub WVector1_DblClick()

    WVector1.Col = 1
    ZZClave = WVector1.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZClave + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Linea.Text = rstArticulo!Linea
        Tipo.Text = rstArticulo!Tipo
        Fragancia.Text = rstArticulo!Fragancia
        Calidad.Text = rstArticulo!Calidad
        Tamano.Text = rstArticulo!Tamano
        rstArticulo.Close
        Call Tamano_KeyPress(13)
        Linea.SetFocus
    End If
    
End Sub

Private Sub WVector1_Click()

    WVector1.Col = 1
    ZZClave = WVector1.Text
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZClave + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Linea.Text = rstArticulo!Linea
        Tipo.Text = rstArticulo!Tipo
        Fragancia.Text = rstArticulo!Fragancia
        Calidad.Text = rstArticulo!Calidad
        Tamano.Text = rstArticulo!Tamano
        rstArticulo.Close
    End If
    
    Call cmdClose_Click
    
End Sub


