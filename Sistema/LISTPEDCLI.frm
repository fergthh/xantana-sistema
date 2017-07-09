VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListPedCli 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Pedidos por Cliente"
   ClientHeight    =   8220
   ClientLeft      =   1920
   ClientTop       =   555
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   8145
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
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   4815
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
         Left            =   3720
         MouseIcon       =   "LISTPEDCLI.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "LISTPEDCLI.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salida"
         Top             =   3120
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
         Left            =   1320
         MouseIcon       =   "LISTPEDCLI.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "LISTPEDCLI.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Impresion x Impresora"
         Top             =   3120
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
         Left            =   2520
         MouseIcon       =   "LISTPEDCLI.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "LISTPEDCLI.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Consulta de Datos"
         Top             =   3120
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
         Left            =   120
         MouseIcon       =   "LISTPEDCLI.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "LISTPEDCLI.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   3120
         Width           =   855
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
         Left            =   1560
         TabIndex        =   13
         Top             =   2520
         Width           =   3015
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
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
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   6
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1335
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
         Left            =   720
         TabIndex        =   12
         Top             =   2520
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
         Left            =   720
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
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
         Left            =   720
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Cliente"
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
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Cliente"
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
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListPedCli.rpt"
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
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "LISTPEDCLI.frx":2D30
      Left            =   120
      List            =   "LISTPEDCLI.frx":2D37
      TabIndex        =   1
      Top             =   5040
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListPedCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WSaldo As Double

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
    
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
    WWParidad = 0
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
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Varios = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    Select Case Tipo.ListIndex
        Case 1
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + " Paridad = " + "'" + Str$(WWParidad) + "',"
            ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "',"
            ZSql = ZSql + " Importe1 = Fabrica " + ","
            ZSql = ZSql + " Importe2 = Facturado " + ","
            ZSql = ZSql + " Importe3 = Fabrica - Facturado"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
        Case Else
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + " Paridad = " + "'" + Str$(WWParidad) + "',"
            ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "',"
            ZSql = ZSql + " Importe1 = Cantidad - Ajuste " + ","
            ZSql = ZSql + " Importe2 = Fabrica " + ","
            ZSql = ZSql + " Importe3 = Cantidad - Fabrica - Ajuste"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    End Select
        
    Listado.WindowTitle = "Listado de Pedidos Pendientes por Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
            
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case Tipo.ListIndex
        Case 2
            Listado.SQLQuery = "SELECT Pedido.Numero, Pedido.Articulo, Pedido.Cantidad, Pedido.Cliente, Pedido.Fecha, Pedido.OrdFecha, Pedido.Facturado, Pedido.Lista, Pedido.Saldo, " _
                    + "Articulo.Descripcion, " _
                    + "Cliente.Razon, " _
                    + "Auxiliar.Nombre, Auxiliar.Varios " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar " _
                    + "Where " _
                    + "Pedido.Articulo = Articulo.Codigo AND " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Pedido.Cliente >= '" + Desde.Text + "' AND " _
                    + "Pedido.Cliente <= '" + Hasta.Text + "' AND " _
                    + "Pedido.OrdFecha >= '" + WDesde + "' AND " _
                    + "Pedido.OrdFecha <= '" + WHasta + "'"
                    
            Listado.Connect = Connect()
            
            Uno = "{Pedido.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Dos = " and {Pedido.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos
            Listado.SelectionFormula = Uno + Dos
            
            Listado.ReportFileName = "ListPedCliII.rpt"
            
        Case Else
            Listado.SQLQuery = "SELECT Pedido.Numero, Pedido.Articulo, Pedido.Precio, Pedido.Cliente, Pedido.Fecha, Pedido.Importe1, Pedido.Importe2, Pedido.Importe3, Pedido.OrdFecha, Pedido.Descuento, Pedido.Facturado, Pedido.Lista, Pedido.Moneda, Pedido.PrecioII, Pedido.Paridad, " _
                    + "Articulo.Descripcion, " _
                    + "Cliente.Razon, " _
                    + "Auxiliar.Nombre, Auxiliar.Varios " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar " _
                    + "Where " _
                    + "Pedido.Articulo = Articulo.Codigo AND " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Pedido.Cliente >= '" + Desde.Text + "' AND " _
                    + "Pedido.Cliente <= '" + Hasta.Text + "' AND " _
                    + "Pedido.OrdFecha >= '" + WDesde + "' AND " _
                    + "Pedido.OrdFecha <= '" + WHasta + "' AND " _
                    + "Pedido.Importe3 > 0"
                    
            
            Listado.Connect = Connect()
            
            Uno = "{Pedido.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Dos = " and {Pedido.Importe3} > 0"
            Tres = " and {Pedido.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos + Tres
            Listado.SelectionFormula = Uno + Dos + Tres
            
            Listado.ReportFileName = "ListPedCliPendII.rpt"
    End Select

    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListPedCli.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFec.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            Desde.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Pendiente Fabricacion"
    Tipo.AddItem "Pendiente Facturacion"
    Tipo.AddItem "Completo"
    
    Tipo.ListIndex = 0

    Desde.Text = ""
    Hasta.Text = ""
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Ayuda.Visible = True
    Ayuda.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Order by Cliente.Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = !Cliente + " " + !Fantasia
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Cliente
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Indice = Pantalla.ListIndex
    Desde.Text = WIndice.List(Indice)
    Hasta.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    Desde.SetFocus
    
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
                ZSql = ZSql + " FROM Cliente"
                ZSql = ZSql + " Where Cliente.Fantasia LIKE " + "'" + "%" + ZAyuda + "%" + "'"
                ZSql = ZSql + " Order by Cliente.Cliente"
                spCliente = ZSql
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    With rstCliente
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                IngresaItem = rstCliente!Cliente + " " + rstCliente!Fantasia
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstCliente!Cliente
                                WIndice.AddItem IngresaItem
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCliente.Close
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

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 115
            Call Consulta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub













