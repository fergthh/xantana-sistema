VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListPedArt 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Pedidos por Sector"
   ClientHeight    =   8220
   ClientLeft      =   1935
   ClientTop       =   750
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   8220
   ScaleWidth      =   8145
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   960
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Cliente 
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
         MaxLength       =   10
         TabIndex        =   2
         Text            =   " "
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox Sector 
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
         MaxLength       =   4
         TabIndex        =   3
         Text            =   " "
         Top             =   1200
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
         TabIndex        =   4
         Top             =   2280
         Width           =   3015
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
         Left            =   240
         MouseIcon       =   "listpedart.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "listpedart.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   3000
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
         Left            =   2640
         MouseIcon       =   "listpedart.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "listpedart.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Consulta de Datos"
         Top             =   3000
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
         Left            =   1440
         MouseIcon       =   "listpedart.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "listpedart.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Impresion x Impresora"
         Top             =   3000
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
         Left            =   3840
         MouseIcon       =   "listpedart.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "listpedart.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salida"
         Top             =   3000
         Width           =   855
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   600
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
         Left            =   1560
         TabIndex        =   0
         Top             =   600
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
      Begin VB.Label DesSector 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3240
         TabIndex        =   18
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label DesCliente 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3240
         TabIndex        =   17
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label7 
         Caption         =   "Cliente"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Sector"
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
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1575
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
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1575
      End
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
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   7815
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListPedArt.rpt"
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
      TabIndex        =   8
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
      ItemData        =   "listpedart.frx":2D30
      Left            =   240
      List            =   "listpedart.frx":2D37
      TabIndex        =   6
      Top             =   4800
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListPedaRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WSaldo As Double

Dim ZZAyudaCli(10000) As String
Dim ZZLugarCli As Integer


Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    Cliente.Text = UCase(Cliente.Text)
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    If Trim(Cliente.Text) <> "" Then
        ZZDesdeCli = Cliente.Text
        ZZHastaCli = Cliente.Text
            Else
        ZZDesdeCli = ""
        ZZHastaCli = "ZZZZZZZZZZ"
    End If
    
    If Trim(Sector.Text) <> "" Then
        ZZDesdeSector = Sector.Text
        ZZHastaSector = Sector.Text
            Else
        ZZDesdeSector = "00"
        ZZHastaSector = "99"
    End If
    
    WTitulo = "entre el " + DesdeFec.Text + " hasta el " + HastaFec.Text
    
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
            ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "',"
            ZSql = ZSql + " Importe1 = Fabrica " + ","
            ZSql = ZSql + " Importe2 = Facturado " + ","
            ZSql = ZSql + " Importe3 = Fabrica - Facturado"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
        Case Else
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "',"
            ZSql = ZSql + " Importe1 = Cantidad - Ajuste " + ","
            ZSql = ZSql + " Importe2 = Fabrica " + ","
            ZSql = ZSql + " Importe3 = Cantidad - Fabrica - Ajuste"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    End Select
    
    Listado.WindowTitle = "Listado de Pedidos por Sector"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case Tipo.ListIndex
        Case 2
            Listado.SQLQuery = "SELECT Pedido.Numero, Pedido.Articulo, Pedido.Cantidad, Pedido.Cliente, Pedido.Fecha, Pedido.OrdFecha, Pedido.Facturado, Pedido.Saldo, " _
                    + "Articulo.Descripcion, Articulo.Sector, " _
                    + "Auxiliar.Nombre, Auxiliar.Varios, " _
                    + "Cliente.Razon, " _
                    + "Sector.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Sector Sector " _
                    + "Where " _
                    + "Pedido.Articulo = Articulo.Codigo AND " _
                    + "Pedido.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Articulo.Sector = Sector.Codigo AND " _
                    + "Pedido.Cliente >= '" + ZZDesdeCli + "' AND " _
                    + "Pedido.Cliente <= '" + ZZHastaCli + "' AND " _
                    + "Pedido.OrdFecha >= '" + WDesde + "' AND " _
                    + "Pedido.OrdFecha <= '" + WHasta + "' AND " _
                    + "Articulo.Sector >= '" + ZZDesdeSector + "' AND " _
                    + "Articulo.Sector <= '" + ZZHastaSector + "'"
                    
                    
            Listado.Connect = Connect()
    
            Uno = "{Pedido.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Dos = " and {Articulo.Sector} in " + Chr$(34) + ZZDesdeSector + Chr$(34) + " to " + Chr$(34) + ZZHastaSector + Chr$(34)
            Tres = " and {Pedido.Cliente} in " + Chr$(34) + ZZDesdeCli + Chr$(34) + " to " + Chr$(34) + ZZHastaCli + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos + Tres
            Listado.SelectionFormula = Uno + Dos + Tres
            
            Listado.ReportFileName = "ListPedArtII.rpt"
        
        Case Else
            Listado.SQLQuery = "SELECT Pedido.Numero, Pedido.Articulo, Pedido.Cantidad, Pedido.Cliente, Pedido.Fecha, Pedido.OrdFecha, Pedido.Facturado, Pedido.Saldo, Pedido.Importe1, Pedido.Importe2, Pedido.Importe3,  " _
                    + "Articulo.Descripcion, Articulo.Sector, " _
                    + "Auxiliar.Nombre, Auxiliar.Varios, " _
                    + "Cliente.Razon, " _
                    + "Sector.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Sector Sector " _
                    + "Where " _
                    + "Pedido.Articulo = Articulo.Codigo AND " _
                    + "Pedido.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Articulo.Sector = Sector.Codigo AND " _
                    + "Pedido.Cliente >= '" + ZZDesdeCli + "' AND " _
                    + "Pedido.Cliente <= '" + ZZHastaCli + "' AND " _
                    + "Pedido.OrdFecha >= '" + WDesde + "' AND " _
                    + "Pedido.OrdFecha <= '" + WHasta + "' AND " _
                    + "Pedido.Importe3 > 0 AND " _
                    + "Articulo.Sector >= '" + ZZDesdeSector + "' AND " _
                    + "Articulo.Sector <= '" + ZZHastaSector + "'"
    
            Listado.Connect = Connect()
            
            Uno = "{Pedido.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Dos = " and {Pedido.Importe3} > 0"
            Tres = " and {Articulo.Sector} in " + Chr$(34) + ZZDesdeSector + Chr$(34) + " to " + Chr$(34) + ZZHastaSector + Chr$(34)
            Cuatro = " and {Pedido.Cliente} in " + Chr$(34) + ZZDesdeCli + Chr$(34) + " to " + Chr$(34) + ZZHastaCli + Chr$(34)
            
            Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
            Listado.SelectionFormula = Uno + Dos + Tres + Cuatro
            
            Listado.ReportFileName = "ListPedArtPendII.rpt"
    
    End Select

    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListPedaRT.Hide
    Unload Me
    MenuVen.Show
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
            Sector.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Private Sub Sector_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Sector.Text) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Sector"
            ZSql = ZSql + " Where Sector.Codigo = " + "'" + Sector.Text + "'"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
            If rstSector.RecordCount > 0 Then
                DesSector.Caption = rstSector!Descripcion
                Cliente.SetFocus
                    Else
                DesSector.Caption = ""
                Sector.SetFocus
            End If
                Else
            DesSector.Caption = ""
            Sector.SetFocus
            Cliente.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Sector.Text = ""
        DesSector.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Trim(Cliente.Text) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Fantasia
                DesdeFec.SetFocus
                    Else
                Cliente.Text = ""
                DesCliente.Caption = ""
            End If
                Else
            Cliente.Text = ""
            DesCliente.Caption = ""
            DesdeFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub






Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Pendiente Fabricacion"
    Tipo.AddItem "Pendiente Facturacion"
    Tipo.AddItem "Completo"
    
    Tipo.ListIndex = 0

    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Sector.Text = ""
    Cliente.Text = ""
    DesSector.Caption = ""
    DesCliente.Caption = ""
    
    Frame2.Visible = True

End Sub

Private Sub Consulta_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    Opcion.Clear
    
    Opcion.AddItem "Sector"
    Opcion.AddItem "Clientes"

    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()


    On Error GoTo WError

    Opcion.Visible = False

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex

    Ayuda.Visible = True
    Ayuda.Text = ""
     
    Dim IngresaItem As String
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Sector"
            ZSql = ZSql + " Order by Sector.Codigo"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
            If rstSector.RecordCount > 0 Then
                With rstSector
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstSector!Codigo + " " + rstSector!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstSector!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSector.Close
            End If
            
        Case 1
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Ayuda.Visible = False
    Sector.SetFocus

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Sector.Text = WIndice.List(Indice)
            Call Sector_Keypress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
                    
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
            ZSql = ZSql + " FROM Sector"
            ZSql = ZSql + " Where Sector.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Sector.Codigo"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
            If rstSector.RecordCount > 0 Then
                With rstSector
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
                rstSector.Close
            End If
            
        Case 1
            Erase ZZAyudaCli
            ZZLugarCli = 0
        
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.fantasia LIKE " + "'" + "%" + ZAyuda + "%" + "'"
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
                            ZZLugarCli = ZZLugarCli + 1
                            ZZAyudaCli(ZZLugarCli) = rstCliente!Cliente
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.razon LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            ZZEntra = "S"
                            For Ciclo = 1 To ZZLugarCli
                                If UCase(ZZAyudaCli(Ciclo)) = UCase(rstCliente!Cliente) Then
                                    ZZEntra = "N"
                                    Exit For
                                End If
                            Next Ciclo
                            If ZZEntra = "S" Then
                                IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstCliente!Cliente
                                WIndice.AddItem IngresaItem
                            End If
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

Private Sub Sector_KeyDown(KeyCode As Integer, Shift As Integer)
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













