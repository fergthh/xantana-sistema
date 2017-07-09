VERSION 5.00
Begin VB.Form Menu41 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Control de Gestion - Clientes"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   375
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "menu41factura.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "menu41factura.frx":0442
   ScaleHeight     =   8310
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.TextBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5880
      TabIndex        =   1
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "4) Listado de Pedidos por Vendedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   2880
      Width           =   6015
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "3) Listado de Pedidos por Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   2400
      Width           =   6015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "2) Listado de Pedido por Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "1) Subdiario de Iva Ventas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      Caption         =   "Ingrese su Opcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   7080
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "LISTADOS DE OPERACIONES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
End
Attribute VB_Name = "Menu41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cash_Click()
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgCash.Show
End Sub

Private Sub Banco_Click()
    PrgBanco.Show
End Sub

Private Sub Concepto_Click()
    PrgConcepto.Show
End Sub

Private Sub Cuentas_Click()
    PrgCuenta.Show
End Sub

Private Sub Agenda_Click()
    OPEN_FILE_Agenda
    PrgAgenda.Show
End Sub

Private Sub Aplica_Click()
    PrgAplica.Show
End Sub

Private Sub ActalizaCodigo_Click()
    prgActualizacionIndividual.Show
End Sub

Private Sub ActualizaCostoFuturo_Click()
    prgActualizacionCostoFuturo.Show
End Sub

Private Sub ActualizaGeneral_Click()
    prgActualizacionGeneral.Show
End Sub

Private Sub Calculadora_Click()
    Calculator.Show
End Sub

Private Sub Cotiza_Click()
    OPEN_FILE_Cotiza
    OPEN_FILE_Configuracion
    PrgCotiza.Show
End Sub

Private Sub CondPago_Click()
    PrgCondPago.Show
End Sub

Private Sub ctactefecha_Click()
    PrgCtaCtefecha.Show
End Sub

Private Sub CtaCteCampana_Click()
    PrgCtaCteCampana.Show
End Sub

Private Sub ConsultaPrecios_Click()
    prgConsultaArticulo.Show
End Sub

Private Sub Despacho_Click()
    PrgDespacho.Show
End Sub

Private Sub Empre_Click()
    Empresa.Show
End Sub

Private Sub Escalas_Click()
    PrgEscalas.Show
End Sub

Private Sub Cliente_Click()
    prgcliente.Show
End Sub

Private Sub Comiven_Click()
    PrgVentVend.Show
End Sub

Private Sub ComparaVen_Click()
    PrgComparaVen.Show
End Sub

Private Sub CompArt_Click()
    PrgCompArt.Show
End Sub

Private Sub CtaCte_Click()
    PrgCtaCte.Show
End Sub

Private Sub ctacte2_Click()
    PrgCtaCte1.Show
End Sub

Private Sub CtacteVen_Click()
    PrgCtaCteVen.Show
End Sub

Private Sub devol_Click()
    PrgDevol.Show
End Sub

Private Sub Expreso_Click()
    PrgExpreso.Show
End Sub

Private Sub factu_Click()
    PrgFactura.Show
End Sub

Private Sub FactuExpo_Click()
    OPEN_FILE_Configuracion
    WVarios = 1
    PrgFactuExpo.Show
End Sub

Private Sub Fallados_Click()
    PrgFallados.Show
End Sub

Private Sub Factura_Click()

End Sub

Private Sub Fin_Click()
    Rem Menu.WindowState = 1
    Rem a = Shell("Sistema.exe", 1)
    Close
    End
End Sub

Private Sub Form_Load()
    Menu1.Caption = "Sistema de Ventas : " + WNombreEmpresa
End Sub

Private Sub HistorialCliente_Click()
    PrgHistorialCliente.Show
End Sub

Private Sub ListaClienteProvincia_Click()
    PrgListaClieProvincia.Show
End Sub

Private Sub ListaClienteVendedor_Click()
    PrgListaClieVende.Show
End Sub

Private Sub ListaClienteZona_Click()
    PrgListaClieZona.Show
End Sub

Private Sub ListaCostoProveedor_Click()
    PrgListaCostoProveedor.Show
End Sub

Private Sub listageneral_Click()
    PrgListaNovedadesPrecio.Show
End Sub

Private Sub ListaStockMinimo_Click()
    PrgListaStockMinimo.Show
End Sub

Private Sub ListaStockProve_Click()
    PrgListaStockProveedor.Show
End Sub

Private Sub ListaStockValora_Click()
    PrgListaStockValora.Show
End Sub

Private Sub ListaStocvGrupo_Click()
    PrgListaStockGrupo.Show
End Sub

Private Sub ListaventasSemestre_Click()
    PrgListaEvolucionVentas.Show
End Sub

Private Sub PasajecostoFuturo_Click()
    prgPasajeCostoFuturo.Show
End Sub

Private Sub Prove_Click()
    PrgProve.Show
End Sub

Private Sub IngrePtoVend_Click()
    PrgIngrePtoVend.Show
End Sub

Private Sub IvaCompo_Click()
    PrgIvaCompo.Show
End Sub

Private Sub Ivaven_Click()
    PrgIvaven.Show
End Sub

Private Sub Lineas_Click()
    PrgZona.Show
End Sub

Private Sub ListaClieVende_Click()
    PrgListaClieVende.Show
End Sub

Private Sub Listapre_Click()
    PrgPrecioFam.Show
End Sub

Private Sub ListCoti_Click()
    PrgListCoti.Show
End Sub

Private Sub Listmov_Click()
    PrgListMov.Show
End Sub

Private Sub ListPedArt_Click()
    PrgListPedaRT.Show
End Sub

Private Sub ListPedCli_Click()
    PrgListPedCli.Show
End Sub

Private Sub Minimo_Click()
    PrgMinimo.Show
End Sub

Private Sub MovStk_Click()
    PrgMovStk.Show
End Sub

Private Sub lledatos_Click()
    PrgLeeDatos.Show
End Sub

Private Sub MOvvar_Click()
    PrgMovStk.Show
End Sub

Private Sub OrdenFactura_Click()
    PrgOrdenFactura.Show
End Sub

Private Sub parametro_Click()
    PrgParametro.Show
End Sub

Private Sub Pedido_Click()
    PrgPedido.Show
End Sub

Private Sub Plantilla_Click()
    PrgPlantilla.Show
End Sub

Private Sub Prod_Click()
    prgArticulo.Show
End Sub

Private Sub Proyec_Click()
    PrgProyCta.Show
End Sub

Private Sub Proyecto_Click()
    PrgProyecto.Show
End Sub

Private Sub Tipopro_Click()
    PrgTipoPro.Show
End Sub

Private Sub Produccion_Click()
    PrgProduccion.Show
End Sub

Private Sub recibos_Click()
    PrgRecibos.Show
End Sub

Private Sub SaldosCta_Click()
    PrgSaldoCta.Show
End Sub

Private Sub Valua_Click()
    PrgValua.Show
End Sub

Private Sub SubLinea_Click()
    PrgFamilia.Show
End Sub

Private Sub TotalArticulo_Click()
    PrgTotalArticulo.Show
End Sub

Private Sub Varios1_Click()
    WVarios = 1
    PrgVarios.Show
End Sub

Private Sub Varios2_Click()
    WVarios = 2
    PrgVarios.Show
End Sub

Private Sub Varios3_Click()
    WVarios = 3
    PrgVarios.Show
End Sub

Private Sub Vende_Click()
    PrgVendedor.Show
End Sub

Private Sub VentaCampana_Click()
    PrgVentaCampana.Show
End Sub

Private Sub VentArt_Click()
    PrgEstaart.Show
End Sub

Private Sub VentasCli_Click()
    PrgVentClie.Show
End Sub

Private Sub Ventasproy_Click()
    PrgVentProy.Show
End Sub

Private Sub ventclie_Click()
    PrgEstacli.Show
End Sub

Private Sub Form_Activate()
    Menu.Caption = "Sistema de Ventas : " + WNombreEmpresa
    If WEmpresa = "" Then
        Menu.Hide
        Empresa.Show
    End If
End Sub

Private Sub VentPcia_Click()
    PrgVentPcia.Show
End Sub

Private Sub Opcion_Keypress(KeyAscii As Integer)
    If Val(Chr$(KeyAscii)) = 1 Then
        Menu41.Hide
        Unload Me
        PrgIvaven.Show
    End If
    If Val(Chr$(KeyAscii)) = 2 Then
        Menu41.Hide
        Unload Me
        PrgListPedCli.Show
    End If
    If Val(Chr$(KeyAscii)) = 3 Then
        Menu41.Hide
        Unload Me
        PrgListPedaRT.Show
    End If
    If KeyAscii = 27 Then
        Menu41.Hide
        Unload Me
        Menu4.Show
    End If
End Sub


