VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Control de Gestion - Ventas y Stock : "
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   660
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
   Icon            =   "MenuFactura.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "MenuFactura.frx":0442
   ScaleHeight     =   8310
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Logo 
      Height          =   1005
      Left            =   8520
      Picture         =   "MenuFactura.frx":0884
      Top             =   6360
      Width           =   2025
   End
   Begin VB.Label VentasII 
      BackColor       =   &H8000000C&
      Caption         =   "Ventas y Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "MenuFactura.frx":1202
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Ventas y Stock"
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   960
      Picture         =   "MenuFactura.frx":150C
      Top             =   120
      Width           =   9600
   End
   Begin VB.Menu Maestros 
      Caption         =   "Clientes"
      Begin VB.Menu Cliente 
         Caption         =   "Mantenimiento de Archivo de Clientes"
      End
      Begin VB.Menu ListaClienteVendedor 
         Caption         =   "Listado por Clientes por Vendedor"
      End
      Begin VB.Menu ListaClienteZona 
         Caption         =   "Listado por Clientes por Zona"
      End
      Begin VB.Menu ListaClienteProvincia 
         Caption         =   "Listado por Clientes por Provincia"
      End
      Begin VB.Menu HistorialCliente 
         Caption         =   "Historial del Cliente"
      End
   End
   Begin VB.Menu fgnhghj 
      Caption         =   "Articulos"
      Begin VB.Menu Prod 
         Caption         =   "Ingreso de Articulos"
      End
      Begin VB.Menu Prove 
         Caption         =   "Ingreso de Proveedores"
      End
      Begin VB.Menu sdfds 
         Caption         =   "Listado de Precios"
         Begin VB.Menu ListaCostoProveedor 
            Caption         =   "Listado de Costos por Proveedor"
         End
         Begin VB.Menu ConsultaPrecios 
            Caption         =   "Consulta de Precios"
         End
         Begin VB.Menu ListaNovedadesPrecio 
            Caption         =   "Listado de Novedades de Precios"
            Begin VB.Menu listageneral 
               Caption         =   "Listado General de Novedades de Precios"
            End
            Begin VB.Menu GrabaNovedades 
               Caption         =   "Grabacion de Nonedades es Diskette"
               Enabled         =   0   'False
            End
         End
         Begin VB.Menu ActualizafechaListaPrecio 
            Caption         =   "Actualizacion de Fechas de Lista de Precios  en Clientes"
            Enabled         =   0   'False
         End
         Begin VB.Menu ListaPreciosGrupo 
            Caption         =   "Lista de Precios por Grupo"
            Enabled         =   0   'False
         End
         Begin VB.Menu ListaPrecioProveedor 
            Caption         =   "Lista de Precios por Proveedor"
            Enabled         =   0   'False
         End
         Begin VB.Menu ListaPreciosDiskette 
            Caption         =   "Lista de Precios en Diskette"
            Enabled         =   0   'False
         End
         Begin VB.Menu ListaPreciosCliente 
            Caption         =   "Lista de Precios por Cliente"
            Enabled         =   0   'False
         End
         Begin VB.Menu ListaCostosImportacion 
            Caption         =   "Listado de Costos por Importacion"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu sdasd 
         Caption         =   "Actulizacion de Precios"
         Begin VB.Menu ActalizaCodigo 
            Caption         =   "Actualizacion en Forma Individual"
         End
         Begin VB.Menu ActualizaGeneral 
            Caption         =   "Actualizacion General"
         End
         Begin VB.Menu ActualizaCostoFuturo 
            Caption         =   "Actualizacion de Costo Futuro"
         End
         Begin VB.Menu PasajecostoFuturo 
            Caption         =   "Pasaje de Costo Futuro a Actual"
         End
      End
      Begin VB.Menu movstk 
         Caption         =   "Ingreso y Egreso de Stock"
      End
      Begin VB.Menu ListaStockProve 
         Caption         =   "Stock por Proveedor"
      End
      Begin VB.Menu ListaStockValora 
         Caption         =   "Valorizacion de Stock"
      End
      Begin VB.Menu ListaStocvGrupo 
         Caption         =   "Stock por Grupo"
      End
      Begin VB.Menu ListaStockMinimo 
         Caption         =   "Pedido de Reposicion"
      End
      Begin VB.Menu ListaventasSemestre 
         Caption         =   "Evolucion Semestral de Ventas"
      End
   End
   Begin VB.Menu sadfdsf 
      Caption         =   "Parametros"
      Begin VB.Menu CondPago 
         Caption         =   "Condiciones de Pago"
      End
      Begin VB.Menu SubLinea 
         Caption         =   "Grupos"
      End
      Begin VB.Menu Expreso 
         Caption         =   "Expreso"
      End
      Begin VB.Menu Despacho 
         Caption         =   "Despacho"
      End
      Begin VB.Menu Lineas 
         Caption         =   "Zonas"
      End
      Begin VB.Menu Vende 
         Caption         =   "Vendedores"
      End
   End
   Begin VB.Menu Nov 
      Caption         =   "Novedades"
      Begin VB.Menu Pedido 
         Caption         =   "Ingreso de Pedidos"
      End
      Begin VB.Menu Factu 
         Caption         =   "Ingreso de Facturas"
      End
      Begin VB.Menu GuiaTRansporte 
         Caption         =   "Ingreso de Guias de Transporte"
      End
      Begin VB.Menu NotasEnvio 
         Caption         =   "Ingreso de Notas de Envio"
      End
      Begin VB.Menu devol 
         Caption         =   "Ingreso de Devolucion de productos"
         Enabled         =   0   'False
      End
      Begin VB.Menu Varios1 
         Caption         =   "Emision de Facturas"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu Varios2 
         Caption         =   "Emision de Notas de Debito"
         Enabled         =   0   'False
      End
      Begin VB.Menu Varios3 
         Caption         =   "Emision de Notas de Credito"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu recibos 
         Caption         =   "Ingreso de Cobranas"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu fgh 
      Caption         =   "Procesos"
      Begin VB.Menu Escalas 
         Caption         =   "Ingreso de Configuracion del Sistema"
      End
      Begin VB.Menu parametro 
         Caption         =   "Ingreso de Datos de la Empresa"
      End
      Begin VB.Menu lledatos 
         Caption         =   "Traspaso de Datos"
      End
      Begin VB.Menu Calculadora 
         Caption         =   "Calculadora"
         Visible         =   0   'False
      End
      Begin VB.Menu Empre 
         Caption         =   "Cambio de Empresa"
         Visible         =   0   'False
      End
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menu"
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
    Menu.Caption = "Sistema de Ventas : " + WNombreEmpresa
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
