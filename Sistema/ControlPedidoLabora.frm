VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgControlPedido 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "load"
   ClientHeight    =   8835
   ClientLeft      =   165
   ClientTop       =   1245
   ClientWidth     =   14865
   FillStyle       =   2  'Horizontal Line
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   14865
   Begin VB.TextBox NroPedido 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   30
      Text            =   " "
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton Entrega 
      Caption         =   "ENTREGADO"
      Height          =   615
      Left            =   13200
      TabIndex        =   29
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton Fabricado 
      Caption         =   "FABRICADO"
      Height          =   615
      Left            =   13200
      TabIndex        =   28
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton FacturaAuto 
      Caption         =   "FACTURA"
      Height          =   615
      Left            =   13200
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   4320
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3960
      Width           =   375
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   13560
      TabIndex        =   22
      Text            =   "Text2"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   495
      Left            =   13560
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton PtoAuto 
      Caption         =   "PRESUPUESTO"
      Height          =   615
      Left            =   13200
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton RecAuto 
      Caption         =   "RECIBOS"
      Height          =   615
      Left            =   13200
      TabIndex        =   19
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Bonifica 
      Caption         =   "PEDIDOS"
      Height          =   615
      Left            =   13200
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox Color 
      Height          =   315
      Left            =   1320
      TabIndex        =   16
      Text            =   " "
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu F10"
      Height          =   975
      Left            =   11640
      MouseIcon       =   "ControlPedidoLabora.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ControlPedidoLabora.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salida"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3840
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   6375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   11245
      _Version        =   327680
      BackColor       =   -2147483633
      FillStyle       =   1
      GridLines       =   3
      GridLinesFixed  =   1
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   6960
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.TextBox Cliente 
      Height          =   300
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1575
   End
   Begin VB.ListBox Pantalla 
      Height          =   1620
      ItemData        =   "ControlPedidoLabora.frx":0B4C
      Left            =   6960
      List            =   "ControlPedidoLabora.frx":0B53
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.ListBox WIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSMask.MaskEdBox HastaFecha 
      Height          =   285
      Left            =   4320
      TabIndex        =   12
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSMask.MaskEdBox DesdeFecha 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pedido"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   31
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label9 
      Caption         =   "Color"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta Fecha"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   3000
      TabIndex        =   15
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Desde Fecha"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "PrgControlPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Importe1 As Double
Private Importe2 As Double
Private Importe3 As Double
Private WTipo As Integer
Private WSaldo As Double

Dim ZZFacturado As Double
Dim ZZCantidad As Double
Dim ZZFabricado As Double
Dim ZZEntregado As Double

Dim ZZPedidos(10000, 10) As String
Dim ZZLugarPedido As Integer

Private Sub Bonifica_Click()
    ZZPasaProcesoPedido = 1
    WVector1.Col = 1
    ZZPasaPedido = ""
    PrgPedido.Show
End Sub

Private Sub cmdClose_Click()
    PrgControlPedido.Hide
    Unload Me
    MenuVen.Show
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
            
    Rem Pantalla.Visible = True
    Rem Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Color_click()
    Call Proceso_Click
    WVector1.TopRow = 1
    WVector1.Col = 1
    WVector1.Row = 1
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Entrega_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE Pedido SET "
    ZSql = ZSql + " Entregado = Fabrica"
    ZSql = ZSql + " Where Numero = " + "'" + WVector1.TextMatrix(WVector1.Row, 1) + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Proceso_Click

End Sub

Private Sub Fabricado_Click()
    ZZPasaProcesoFabrica = 1
    ZZPasaPedido = WVector1.TextMatrix(WVector1.Row, 1)
    PrgProduccionPedido.Show
End Sub

Private Sub FacturaAuto_Click()
    ZZPasaProcesoFactura = 1
    ZZPasaPedido = WVector1.TextMatrix(WVector1.Row, 1)
    PrgFactura.Show
End Sub

Private Sub Form_Activate()
    Call Cliente_KeyPress(13)
End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Cliente.Text = WIndice.List(Indice)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Fantasia
        rstCliente.Close
        Call Proceso_Click
    End If
    
    Cliente.SetFocus
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
 
    Color.Clear
    
    Color.AddItem ""
    Color.AddItem "S/Imprimir (Celeste)"
    Color.AddItem "Precio (Violeta)"
    Color.AddItem "Fabricado (Rosa)"
    Color.AddItem "Facturado (verde)"
    Color.AddItem "Entregado (Amarillo)"
    
    Color.ListIndex = 0
 
    Cliente.Text = ""
    DesCliente.Caption = ""
    
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    
    NroPedido.Text = ""

    WVector1.Col = 1
    WVector1.Row = 1
    
    Call Consulta_Click
    Rem Cliente.Text = PCliente
    Rem Cliente.SetFocus
    
    Call Proceso_Click
    
End Sub

Private Sub Proceso_Click()

    WSalida = "N"
    
    Call Limpia_Vector
    
    Renglon = 0
    WSaldo = 0
    ZZFechaActual = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    If DesdeFecha.Text <> "  /  /    " Then
        WAno = Right$(DesdeFecha.Text, 4)
        WMes = Mid$(DesdeFecha.Text, 4, 2)
        WDia = Left$(DesdeFecha.Text, 2)
        WDesde = WAno + WMes + WDia
            Else
        WDesde = "00000000"
    End If
        
    If HastaFecha.Text <> "  /  /    " Then
        WAno = Right$(HastaFecha.Text, 4)
        WMes = Mid$(HastaFecha.Text, 4, 2)
        WDia = Left$(HastaFecha.Text, 2)
        WHasta = WAno + WMes + WDia
            Else
        WHasta = "99999999"
    End If
        
    
    Erase ZZPedidos
    ZZLugarPedido = 0
    
    ZZZPasa = 0
    ZZCorte = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Pedido.OrdFecha <= " + "'" + WHasta + "'"
    
    If Trim(Cliente.Text) <> "" Then
        ZSql = ZSql + " and Pedido.Cliente = " + "'" + Cliente.Text + "'"
    End If

    ZSql = ZSql + " Order by Pedido.Clave"
    
    
        
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZZZPasa = 0 Then
                        ZZZPasa = 1
                        ZZCorte = rstPedido!Numero
                        ZZCliente = rstPedido!Cliente
                        ZZfecha = rstPedido!Fecha
                        Rem precio = 0
                        ZZEstado1 = 0
                        Rem fabricado
                        ZZEstado2 = 0
                        Rem pediente de fab.
                        ZZEstado3 = 0
                        Rem sin imprimir
                        ZZEstado4 = 0
                        Rem impreso
                        ZZEstado5 = 0
                        Rem Anulado
                        ZZEstado6 = 0
                        Rem entragado
                        ZZEstado7 = 0
                        Rem pendiente
                        ZZEstado8 = 0
                    End If
                    
                    If ZZCorte <> rstPedido!Numero Then
                        
                        ZZPasa = "N"
                        
                        If Color.ListIndex = 0 Then
                            ZZPasa = "S"
                        End If
                        
                        If Color.ListIndex = 1 And ZZEstado5 = 1 Then
                            ZZPasa = "S"
                        End If
                        
                        If Color.ListIndex = 2 And ZZEstado1 = 1 Then
                            ZZPasa = "S"
                        End If
                        
                        If Color.ListIndex = 3 And ZZEstado2 = 1 Then
                            ZZPasa = "S"
                        End If
                        
                        If Color.ListIndex = 4 And ZZEstado7 = 1 Then
                            ZZPasa = "S"
                        End If
                        
                        If Color.ListIndex = 5 And ZZEstado6 = 1 Then
                            ZZPasa = "S"
                        End If
                        
                        If Val(NroPedido.Text) <> 0 Then
                            ZZPasa = "N"
                            If Val(NroPedido.Text) = Val(ZZCorte) Then
                                ZZPasa = "S"
                            End If
                        End If
                        
                        If ZZPasa = "S" Then
                            
                            If ZZEstado1 <> 0 Or ZZEstado2 <> 0 Or ZZEstado6 <> 0 Or ZZEstado7 <> 0 Then
                                
                                Renglon = Renglon + 1
                                WVector1.Row = Renglon
                                
                                WVector1.Col = 1
                                WVector1.CellBackColor = &H8000000F
                                WVector1.Text = ZZCorte
                                
                                WVector1.Col = 2
                                WVector1.CellBackColor = &H8000000F
                                WVector1.Text = ZZfecha
                                
                                WVector1.Col = 3
                                WVector1.CellBackColor = &H8000000F
                                WVector1.Text = ZZCliente
                                
                                If ZZEstado5 = 1 Then
                                    WVector1.Col = 4
                                    WVector1.CellBackColor = &HFFFF80
                                    WVector1.Text = Space(100)
                                        Else
                                    WVector1.Col = 4
                                    WVector1.CellBackColor = &H8000000F
                                    WVector1.Text = Space(100)
                                End If
                                
                                If ZZEstado1 = 1 Then
                                    WVector1.Col = 5
                                    WVector1.CellBackColor = &HFF00FF
                                    WVector1.Text = Space(100)
                                        Else
                                    WVector1.Col = 5
                                    WVector1.CellBackColor = &H8000000F
                                    WVector1.Text = Space(100)
                                End If
                                
                                If ZZEstado2 = 1 Then
                                    WVector1.Col = 6
                                    Rem WVector1.CellBackColor = &HFFC0FF
                                    WVector1.CellBackColor = &HC0C0FF
                                    WVector1.Text = Space(100)
                                        Else
                                    WVector1.Col = 6
                                    WVector1.CellBackColor = &H8000000F
                                    WVector1.Text = Space(100)
                                End If
                                
                                If ZZEstado7 = 1 Then
                                    WVector1.Col = 7
                                    WVector1.CellBackColor = &HFF00&
                                    WVector1.Text = Space(100)
                                        Else
                                    WVector1.Col = 7
                                    WVector1.CellBackColor = &H8000000F
                                    WVector1.Text = Space(100)
                                End If
                                
                                If ZZEstado6 = 1 Then
                                    WVector1.Col = 8
                                    WVector1.CellBackColor = &HFFFF&
                                    WVector1.Text = Space(100)
                                        Else
                                    WVector1.Col = 8
                                    WVector1.CellBackColor = &H8000000F
                                    WVector1.Text = Space(100)
                                End If
                            End If
                            
                        End If
                        
                        ZZCorte = rstPedido!Numero
                        ZZCliente = rstPedido!Cliente
                        ZZfecha = rstPedido!Fecha
                        
                        Rem precio = 0
                        ZZEstado1 = 0
                        Rem fabricado
                        ZZEstado2 = 0
                        Rem pediente de fab.
                        ZZEstado3 = 0
                        Rem sin imprimir
                        ZZEstado4 = 0
                        Rem impreso
                        ZZEstado5 = 0
                        Rem Anulado
                        ZZEstado6 = 0
                        Rem entragado
                        ZZEstado7 = 0
                        Rem pendiente
                        ZZEstado8 = 0
                        
                    End If
                    ZZCantidad = rstPedido!Cantidad - rstPedido!Ajuste

                    ZZFacturado = rstPedido!facturado
                    ZZFabricado = rstPedido!fabrica
                    ZZEntregado = rstPedido!Entregado
                         
                    Call Redondeo(ZZFacturado)
                    Call Redondeo(ZZFabricado)
                    Call Redondeo(ZZEntregado)
                        
                    If rstPedido!Precio = 0 And ZZCantidad <> 0 And ZZFacturado = 0 Then
                        ZZEstado1 = 1
                    End If
                        
                    If ZZFabricado < ZZCantidad Then
                        ZZEstado2 = 1
                    End If
                            
                    If ZZFacturado < ZZFabricado And ZZFabricado > 0 Then
                        ZZEstado7 = 1
                    End If
                            
                    If ZZEntregado < ZZFacturado And ZZFacturado > 0 Then
                        ZZEstado6 = 1
                    End If
                        
                    If rstPedido!MarcaII = "S" Then
                        Rem ZZEstado6 = 1
                            Else
                        If rstPedido!Marca = "X" Then
                            ZZEstado4 = 1
                                Else
                            ZZEstado5 = 1
                        End If
                    End If
                    
                    
                    
                    
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstPedido.Close
    End If
    
    If ZZZPasa <> 0 Then
                        
        ZZPasa = "N"
        
        If Color.ListIndex = 0 Then
            ZZPasa = "S"
        End If
        
        If Color.ListIndex = 1 And ZZEstado5 = 1 Then
            ZZPasa = "S"
        End If
        
        If Color.ListIndex = 2 And ZZEstado1 = 1 Then
            ZZPasa = "S"
        End If
        
        If Color.ListIndex = 3 And ZZEstado2 = 1 Then
            ZZPasa = "S"
        End If
        
        If Color.ListIndex = 4 And ZZEstado7 = 1 Then
            ZZPasa = "S"
        End If
        
        If Color.ListIndex = 5 And ZZEstado6 = 1 Then
            ZZPasa = "S"
        End If
        
        If Val(NroPedido.Text) <> 0 Then
            ZZPasa = "N"
            If Val(NroPedido.Text) = Val(ZZCorte) Then
                ZZPasa = "S"
            End If
        End If
        
        
        If ZZPasa = "S" Then
            
            If ZZEstado1 <> 0 Or ZZEstado2 <> 0 Or ZZEstado6 <> 0 Or ZZEstado7 <> 0 Then
                
                Renglon = Renglon + 1
                WVector1.Row = Renglon
                
                WVector1.Col = 1
                WVector1.CellBackColor = &H8000000F
                WVector1.Text = ZZCorte
                
                WVector1.Col = 2
                WVector1.CellBackColor = &H8000000F
                WVector1.Text = ZZfecha
                
                WVector1.Col = 3
                WVector1.CellBackColor = &H8000000F
                WVector1.Text = ZZCliente
                
                If ZZEstado5 = 1 Then
                    WVector1.Col = 4
                    WVector1.CellBackColor = &HFFFF80
                    WVector1.Text = Space(100)
                        Else
                    WVector1.Col = 4
                    WVector1.CellBackColor = &H8000000F
                    WVector1.Text = Space(100)
                End If
                
                If ZZEstado1 = 1 Then
                    WVector1.Col = 5
                    WVector1.CellBackColor = &HFF00FF
                    WVector1.Text = Space(100)
                        Else
                    WVector1.Col = 5
                    WVector1.CellBackColor = &H8000000F
                    WVector1.Text = Space(100)
                End If
                
                If ZZEstado2 = 1 Then
                    WVector1.Col = 6
                    Rem WVector1.CellBackColor = &HFFC0FF
                    WVector1.CellBackColor = &HC0C0FF
                    WVector1.Text = Space(100)
                        Else
                    WVector1.Col = 6
                    WVector1.CellBackColor = &H8000000F
                    WVector1.Text = Space(100)
                End If
                
                If ZZEstado7 = 1 Then
                    WVector1.Col = 7
                    WVector1.CellBackColor = &HFF00&
                    WVector1.Text = Space(100)
                        Else
                    WVector1.Col = 7
                    WVector1.CellBackColor = &H8000000F
                    WVector1.Text = Space(100)
                End If
                
                If ZZEstado6 = 1 Then
                    WVector1.Col = 8
                    WVector1.CellBackColor = &HFFFF&
                    WVector1.Text = Space(100)
                        Else
                    WVector1.Col = 8
                    WVector1.CellBackColor = &H8000000F
                    WVector1.Text = Space(100)
                End If
        
            End If
        
        End If
        
    End If
    
    
    For Ciclo = 1 To Renglon
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WVector1.TextMatrix(Ciclo, 3) + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WVector1.TextMatrix(Ciclo, 3) = rstCliente!Fantasia
            rstCliente.Close
        End If
    Next Ciclo
    
    
    WVector1.Rows = Renglon + 1
    
    WVector1.Col = 0
    WVector1.Row = 0

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WCliente = Cliente.Text
        Cliente.Text = WCliente
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Fantasia
            rstCliente.Close
            Call Proceso_Click
            If WVector1.Rows > 1 Then
                WVector1.TopRow = 1
                WVector1.Col = 1
                WVector1.Row = 1
            End If
            Cliente.SetFocus
                Else
            Call Proceso_Click
            Rem WVector1.TopRow = 1
            Rem WVector1.Col = 1
            Rem WVector1.Row = 1
            Cliente.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
        Call Proceso_Click
    End If
End Sub


Private Sub NroPedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Proceso_Click
        Cliente.SetFocus
    End If
    If KeyAscii = 27 Then
        NroPedido.Text = ""
        Call Proceso_Click
    End If
End Sub



Private Sub Limpia_Vector()

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 9
    WVector1.Rows = 10000
    WVector1.FixedRows = 1
    
    WVector1.ColWidth(0) = 200
    
    WVector1.Row = 0
    
    WVector1.Col = 1
    WVector1.Text = "Pedido"
    WVector1.ColWidth(1) = 1300
    WVector1.ColAlignment(1) = flexAlignRightCenter
    
    WVector1.Col = 2
    WVector1.Text = "Fecha"
    WVector1.ColWidth(2) = 1300
    WVector1.ColAlignment(2) = flexAlignRightCenter
    
    WVector1.Col = 3
    WVector1.Text = "Cliente"
    WVector1.ColWidth(3) = 3500
    WVector1.ColAlignment(3) = flexAlignLeftCenter
    
    WVector1.Col = 4
    WVector1.Text = "S/Imprimir"
    WVector1.ColWidth(5) = 1300
    WVector1.ColAlignment(5) = flexAlignRightCenter
    
    WVector1.Col = 5
    WVector1.Text = "S/Precio"
    WVector1.ColWidth(4) = 1300
    WVector1.ColAlignment(4) = flexAlignRightCenter
    
    WVector1.Col = 6
    WVector1.Text = "F/Fabricar"
    WVector1.ColWidth(6) = 1300
    WVector1.ColAlignment(6) = flexAlignRightCenter
    
    WVector1.Col = 7
    WVector1.Text = "F/Facturar"
    WVector1.ColWidth(7) = 1300
    WVector1.ColAlignment(7) = flexAlignRightCenter
    
    WVector1.Col = 8
    WVector1.Text = "F/Entregar"
    WVector1.ColWidth(7) = 1300
    WVector1.ColAlignment(7) = flexAlignRightCenter
    
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
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

Private Sub PtoAuto_Click()
    ZZPasaProcesoFactura = 1
    ZZNivelFactura = 1
    ZZPasaPedido = WVector1.TextMatrix(WVector1.Row, 1)
    PrgFacturaRemito.Show
End Sub

Private Sub WVector1_DblClick()

    ZZPasaProcesoPedido = 1
    ZZPedidoControles = 1
    WVector1.Col = 1
    ZZPasaPedido = WVector1.Text
    PrgPedido.Show
    
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

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub WVector1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub










Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            Call Proceso_Click
            HastaFecha.SetFocus
                Else
            DesdeFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFecha.Text = "  /  /    "
        Call Proceso_Click
    End If
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Call Proceso_Click
            DesdeFecha.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
        Call Proceso_Click
    End If
End Sub


