Attribute VB_Name = "SISTEMA"
'--------------------------------------------------------
' DEFINICIONES GLOBALES
'--------------------------------------------------------

Rem Global Const FILENAME = "LOCALIDAD.MDB"

Global PATH_PROG As String
Global coderr As Integer
Global Ds(30) As Integer
Global Const FILE_TYPE = ""
Global Lote As String
Global TipoImpre As String
Global XIndice As Integer
Global Text As String
Global Auxi As String
Global Auxi1 As String
Global Auxi2 As String
Global Validate As String
Global Cicla As Integer
Global WAuxi As Integer
Global XCol As Integer
Global XRow As Integer
Global Existe As String
Global Renglon As Integer
Global WProveedor As String
Global WTipo As String
Global WLetra As String
Global WPunto As String
Global WNumero As String
Global XProveedor As String
Global XTipo As String
Global XLetra As String
Global XPunto As String
Global XNumero As String
Global WImpo As Double
Global WCtaConcepto As String
Global Inicial As Double
Global WEmpresa As String
Global XEmpresa As String
Global PCliente As String
Global PTipo As String
Global PTerminado As String
Global PLote As String
Global WRecibo As String
Global WArti As String
Global ConfigIva1 As Double
Global ConfigIva2 As Double
Global ConfigPercepcion As Double
Global ConfigPunto As Integer
Global WVarios As Integer
Global WPosi As Integer
Global WPeriodo As String
Global WEmpresaConta As String
Global WLicencia As String
Global WFuncion As Integer
Global txtLetraDolares As String
Global ZSql As String
Global ZEmpresa As String
Global ZUsuario As String
Global ZAyuda As String
Global ZZPasaArticulo As String
Global ZZPasaCliente As String
Global ZZPasaNumero As String
Global ZZNivel As Integer
Global ZZNivelFactura As Integer
Global ZZUsuario As Integer
Global ZZClaveProceso As Integer
Global ZZPasaProceso As Integer
Global ZZProcesoPedido As Integer
Global ZZPasoArchivo As String
Global ZZPasaLinea As String
Global ZZPasaTipo As String
Global ZZPasaFragancia As String
Global ZZPasaTamaño As String
Global ZZPasaCalidad As String
Global ZZPasaProcesoII As Integer
Global ZZPasaProcesoCtaCte As Integer
Global ZZPasaPedido As String
Global ZZSistema As Integer
Global ZZPasaProcesoPedido As Integer
Global ZZPasaProcesoFactura As Integer
Global ZZPasaProcesoFabrica As Integer
Global ZZPedidoControles As Integer
Global ZZOperador As String
Global ZZImpreHistorial As String


Global DbConnect$
Global DSN$
Global UID$
Global PWD$
Global DSQ$

Global WNombreBase As String
Global WNombreEmpresa As String
Global WDireccionEmpresa As String
Global WLocalidadEmpresa As String
Global WCuitEmpresa As String
Global WCtaProveedor As String
Global WCtaEfectivo As String
Global WCtaCheques As String
Global WCtaChequeRecha As String
Global WCtaRetGan As String
Global WCtaRetIva As String
Global WCtaretOtra As String
Global WCtaRetSuss As String
Global WCtaDeudores As String
Global WCtaDocumentos As String


Global rstEmpresa As Recordset
Global spEmpresa As String
Global rstAuxiliar As Recordset
Global spAuxiliar As String
Global rstCuenta As Recordset
Global spCuenta As String
Global rstTipoPro As Recordset
Global spTipoPro As String
Global rstTipoClie As Recordset
Global spTipoClie As String
Global rstBanco As Recordset
Global spBanco As String
Global rstConceptos As Recordset
Global spConceptos As String
Global rstProyecto As Recordset
Global spProyecto As String
Global rstProveedor As Recordset
Global spProveedor As String
Global rstCierre As Recordset
Global spCierre As String
Global rstIvaComp As Recordset
Global spIvaComp As String
Global rstHoja As Recordset
Global spHoja As String
Global rstCtaCtePrv As Recordset
Global spCtaCtePrv As String
Global rstImpCyb As Recordset
Global spImpCyb As String
Global rstImpProy As Recordset
Global spImpProy As String
Global rstCompras As Recordset
Global spCompras As String
Global rstArticulo As Recordset
Global spArticulo As String
Global rstConfiguracion As Recordset
Global spConfiguracion As String
Global rstPago As Recordset
Global spPago As String
Global rstChequera As Recordset
Global spChequera As String
Global rstNroRet As Recordset
Global spNroRet As String
Global rstRecibos As Recordset
Global spRecibos As String
Global rstRetencion As Recordset
Global spRetencion As String
Global rstImpreOrd As Recordset
Global spImpreOrd As String
Global rstDepositos As Recordset
Global spDepositos As String
Global rstClientes As Recordset
Global spClientes As String
Global rstCtaCte As Recordset
Global spCtaCte As String
Global rstPosicion As Recordset
Global spPosicion As String
Global rstParametro As Recordset
Global spParametro As String
Global rstClave As Recordset
Global spClave As String
Global rstPagos As Recordset
Global spPagos As String
Global rstControl As Recordset
Global spControl As String
Global rstMovBan As Recordset
Global spMovBan As String
Global rstCash As Recordset
Global spCash As String
Global rstImpCtaCte As Recordset
Global spImpCtaCte As String
Global rstVendedor As Recordset
Global spVendedor As String
Global rstLinea As Recordset
Global spLinea As String
Global rstCliente As Recordset
Global spCliente As String
Global rstPlantilla As Recordset
Global spPlantilla As String
Global rstPedido As Recordset
Global spPedido As String
Global rstFactura As Recordset
Global spFactura As String
Global rstEstadistica As Recordset
Global spEstadistica As String
Global rstDesccomp As Recordset
Global spDesccomp As String
Global rstExpreso As Recordset
Global spExpreso As String
Global rstCondPago As Recordset
Global spCondPago As String
Global rstDespacho As Recordset
Global spDespacho As String
Global rstProduccion As Recordset
Global spProduccion As String
Global rstSubLineas As Recordset
Global spSubLineas As String
Global rstFamilia As Recordset
Global spFamilia As String
Global rstZona As Recordset
Global spZona As String
Global rstHistorialCliente As Recordset
Global spHistorialCliente As String
Global rstBcra As Recordset
Global spBcra As String
Global rstImpreRecibo As Recordset
Global spImpreRecibo As String
Global rstGuiaTransporte As Recordset
Global spGuiaTransporte As String
Global rstGastosBancarios As Recordset
Global spGastosBancarios As String
Global rstGastosCaja As Recordset
Global spGastosCaja As String
Global rstOrden As Recordset
Global spOrden As String
Global rstOrdenImportacion As Recordset
Global spOrdenImportacion As String
Global rstListaPrecios As Recordset
Global spListaPrecios As String
Global rstPrueba As Recordset
Global spPrueba As String
Global rstLineaInsumo As Recordset
Global spLineaInsumo As String
Global rstUbicacion As Recordset
Global spUbicacion As String
Global rstInsumo As Recordset
Global spInsumo As String
Global rstFormula As Recordset
Global spFormula As String
Global rstFormulaII As Recordset
Global spFormulaII As String
Global rstRequisicion As Recordset
Global spRequisicion As String
Global rstOrdenCompraInsumos As Recordset
Global spOrdenCompraInsumos As String
Global rstMovVarInsumos As Recordset
Global spMovVarInsumos As String
Global rstPedidoAuxi As Recordset
Global spPedidoAuxi As String
Global rstAperturaColor As Recordset
Global spAperturaColor As String
Global rstExibidores As Recordset
Global spExibidores As String
Global rstFragancia As Recordset
Global spFragancia As String
Global rstCalidad As Recordset
Global spCalidad As String
Global rstTamaño As Recordset
Global spTamaño As String
Global rstSector As Recordset
Global spSector As String
Global rstEsencia As Recordset
Global spEsencia As String
Global rstInsumoHistorial As Recordset
Global spInsumoHistorial As String
Global rstLista As Recordset
Global spLista As String
Global rstPrecios As Recordset
Global spPrecios As String
Global rstRemito As Recordset
Global spRemito As String
Global rstSolicitud As Recordset
Global spSolicitud As String
Global rstAnalizaPedido As Recordset
Global spAnalizaPedido As String
Global rstOperadorIngreso As Recordset
Global spOperadorIngreso As String
Global rstTransferencia As Recordset
Global spTransferencia As String
Global rstCombo As Recordset
Global spCombo As String
Global rstMovStkInsumo As Recordset
Global spMovStkInsumo As String
Global rstMovStkArticulo As Recordset
Global spMovStkArticulo As String

'--------------------------------------------------------
' VARIABLES OBJETO DEL TIPO "BASE DE DATOS" Y "DYNASETS"
'--------------------------------------------------------

Global DbsAdminis As Database
Global DbsAdminis1 As Database
Global DbsAdminis2 As Database

'definicion de tablas de base de datos de empresa


'definicion de tablas de base de datos  de administracion

Global rstMesas As Recordset
Global rstMozos As Recordset
Global rstRubros As Recordset
Global rstAutor As Recordset
Global rstDistribuidor As Recordset
Global rstColor As Recordset
Global rstEnvio As Recordset
Global rstPrecio As Recordset
Global rstEvolucion As Recordset
Global rstHistorico As Recordset
Global rstRecepcion As Recordset
Global rstEnvase As Recordset
Global rstComponente As Recordset
Global rstClieExpo As Recordset
Global rstArtiExpo As Recordset
Global rstOrdenImpo As Recordset
Global rstListaCompo As Recordset
Global rstLugar As Recordset
Global rstTransporte As Recordset
Global rstMovmes As Recordset
Global rstPosdat As Recordset
Global rstTarjeta As Recordset
Global rstCaja As Recordset
Global rstCupones As Recordset
Global rstVenta As Recordset
Global rstGastos As Recordset
Global rstMovstk As Recordset
Global rstMovi As Recordset
Global rstPasaClientes As Recordset
Global rstPasaCuentas As Recordset
Global rstPasaProve As Recordset
Global rstConceptoII As Recordset
Global rstPto As Recordset
Global rstPtoCue As Recordset
Global rstPtoVend As Recordset
Global rstAgenda As Recordset
Global rstLineas As Recordset
Global rstCotiza As Recordset
Global rstCuentaCon As Recordset
Global rstCompara As Recordset
Global rstEmpreCon As Recordset
Global rstAsiento As Recordset


'--------------------------------------------------------
' NOMBRE DE LAS TABLAS QUE COMPONEN LA BASE DE DATOS
'--------------------------------------------------------

Global Const TABLA_Mesas = "Mesas"
Global Const TABLA_Mozos = "Mozos"
Global Const TABLA_Rubros = "Rubros"
Global Const TABLA_Autor = "Autor"
Global Const TABLA_Distribuidor = "Distribuidor"

Global Const TABLA_Pedido = "Pedido"
Global Const TABLA_Articulo = "Articulo"
Global Const TABLA_Color = "Color"
Global Const TABLA_Envio = "Envio"
Global Const TABLA_Precio = "Precio"
Global Const TABLA_Evolucion = "Evolucion"
Global Const TABLA_Configuracion = "Configuracion"
Global Const TABLA_Historico = "Historico"
Global Const TABLA_Recepcion = "Recepcion"
Global Const TABLA_Auxiliar = "Auxiliar"
Global Const TABLA_Banco = "Banco"
Global Const TABLA_Envase = "Envase"
Global Const TABLA_Componente = "Componente"
Global Const TABLA_Formula = "Formula"
Global Const TABLA_ClieExpo = "ClieExpo"
Global Const TABLA_ArtiExpo = "ArtiExpo"
Global Const TABLA_OrdenImpo = "OrdenImpo"
Global Const TABLA_ListaCompo = "ListaCompo"
Global Const TABLA_Despacho = "Despacho"
Global Const TABLA_Lugar = "Lugar"
Global Const TABLA_Transporte = "Transporte"
Global Const TABLA_TipoPro = "TipoPro"
Global Const TABLA_Chequera = "Chequera"
Global Const TABLA_Control = "Control"
Global Const TABLA_Parametro = "Parametro"
Global Const TABLA_Proyecto = "Proyecto"
Global Const TABLA_Movmes = "Movmes"
Global Const TABLA_Posdat = "Posdat"
Global Const TABLA_Movban = "Movban"
Global Const TABLA_Depositos = "Depositos"
Global Const TABLA_Tarjeta = "Tarjeta"
Global Const TABLA_Caja = "Caja"
Global Const TABLA_Cupones = "Cupones"
Global Const TABLA_Venta = "Venta"
Global Const TABLA_Gastos = "Gastos"
Global Const TABLA_Movstk = "Movstk"
Global Const TABLA_Movi = "Movi"
Global Const TABLA_Compras = "Compras"
Global Const TABLA_PasaClientes = "PasaClientes"
Global Const TABLA_PasaCuentas = "PasaCuentas"
Global Const TABLA_PasaProve = "PasaProve"
Global Const TABLA_Conceptos = "Conceptos"
Global Const TABLA_ConceptoII = "ConceptoII"
Global Const TABLA_Pto = "Pto"
Global Const TABLA_PtoCue = "PtoCue"
Global Const TABLA_PtoVend = "PtoVend"
Global Const TABLA_Clientes = "Clientes"
Global Const TABLA_Agenda = "Agenda"
Global Const TABLA_Plantilla = "Plantilla"
Global Const TABLA_Lineas = "Lineas"
Global Const TABLA_Estadistica = "Estadistica"
Global Const TABLA_Cotiza = "Cotiza"
Global Const TABLA_CtaCtePrv = "CtaCtePrv"
Global Const TABLA_CtaCte = "CtaCte"
Global Const TABLA_CtaCte1 = "CtaCte1"
Global Const TABLA_CtaCte2 = "CtaCte2"
Global Const TABLA_Iva = "Iva"
Global Const TABLA_IvaComp = "Ivacomp"
Global Const TABLA_Pagos = "Pagos"
Global Const TABLA_Solicitud = "Solicitud"
Global Const TABLA_Ranking = "Ranking"
Global Const TABLA_Posicion = "Posicion"
Global Const TABLA_Etiqueta = "Etiqueta"
Global Const TABLA_Cash = "Cash"
Global Const TABLA_GastosProy = "GastosProy"
Global Const TABLA_Proceso1 = "Proceso1"
Global Const TABLA_Recibos = "Recibos"
Global Const TABLA_CheqCartera = "CheqCartera"
Global Const TABLA_Proveedor = "Proveedor"
Global Const TABLA_Cuenta = "Cuenta"
Global Const TABLA_CuentaCon = "CuentaCon"
Global Const TABLA_Compara = "Compara"
Global Const TABLA_Retencion = "Retencion"
Global Const TABLA_Numero = "Numero"
Global Const TABLA_Desccomp = "Desccomp"
Global Const TABLA_Factura = "Factura"
Global Const TABLA_ImpreOrd = "ImpreOrd"
Global Const TABLA_NroRet = "NroRet"
Global Const TABLA_Impctacte = "Impctacte"

Global Const TABLA_Empresa = "Empresa"
Global Const TABLA_EmpreCon = "EmpreCon"
Global Const TABLA_Auxi = "Auxi"
Global Const TABLA_Listado1 = "Listado1"
Global Const TABLA_Impcyb = "Impcyb"
Global Const TABLA_Imputac = "Imputac"
Global Const TABLA_Impproy = "Impproy"
Global Const TABLA_Asiento = "Asiento"



'--------------------------------------------------------
' CAMPOS CORRESPONDIENTES AL ARCHIVO DE Vendedor
'--------------------------------------------------------
 
 Global Const Codigo = "CODIGO"
 Global Const Descripcion = "DESCRIPCION"

Sub OPEN_FILE_Mesas()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstMesas = DbsAdminis.OpenRecordset("Mesas")
End Sub

Sub OPEN_FILE_Mozos()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstMozos = DbsAdminis.OpenRecordset("Mozos")
End Sub

Sub OPEN_FILE_Autor()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstAutor = DbsAdminis.OpenRecordset("Autor")
End Sub

Sub OPEN_FILE_Distribuidor()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstDistribuidor = DbsAdminis.OpenRecordset("Distribuidor")
End Sub

Sub OPEN_FILE_Rubros()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstRubros = DbsAdminis.OpenRecordset("Rubros")
End Sub

Sub OPEN_FILE_Articulo()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstArticulo = DbsAdminis.OpenRecordset("Articulo")
End Sub
 
Sub OPEN_FILE_Color()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstColor = DbsAdminis.OpenRecordset("Color")
End Sub
 
Sub OPEN_FILE_Envio()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstEnvio = DbsAdminis.OpenRecordset("Envio")
End Sub
 
Sub OPEN_FILE_Precio()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPrecio = DbsAdminis.OpenRecordset("Precio")
End Sub
 
Sub OPEN_FILE_Evolucion()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstEvolucion = DbsAdminis.OpenRecordset("Evolucion")
End Sub
 
Sub OPEN_FILE_Configuracion()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstConfiguracion = DbsAdminis.OpenRecordset("Configuracion")
End Sub
 
Sub OPEN_FILE_Historico()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstHistorico = DbsAdminis.OpenRecordset("Historico")
End Sub
 
Sub OPEN_FILE_Recepcion()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstRecepcion = DbsAdminis.OpenRecordset("Recepcion")
End Sub

Sub OPEN_FILE_Auxiliar()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstAuxiliar = DbsAdminis.OpenRecordset("Auxiliar")
End Sub

Sub OPEN_FILE_Banco()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstBanco = DbsAdminis.OpenRecordset("Banco")
End Sub

Sub OPEN_FILE_Envase()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstEnvase = DbsAdminis.OpenRecordset("caja")
End Sub

Sub OPEN_FILE_Componente()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstComponente = DbsAdminis.OpenRecordset("Componente")
End Sub

Sub OPEN_FILE_Formula()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstFormula = DbsAdminis.OpenRecordset("Formula")
End Sub

Sub OPEN_FILE_ClieExpo()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstClieExpo = DbsAdminis.OpenRecordset("ClieExpo")
End Sub

Sub OPEN_FILE_ArtiExpo()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstArtiExpo = DbsAdminis.OpenRecordset("ArtiExpo")
End Sub

Sub OPEN_FILE_OrdenImpo()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstOrdenImpo = DbsAdminis.OpenRecordset("OrdenImpo")
End Sub

Sub OPEN_FILE_ListaCompo()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstListaCompo = DbsAdminis.OpenRecordset("ListaCompo")
End Sub

Sub OPEN_FILE_Despacho()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstDespacho = DbsAdminis.OpenRecordset("Despacho")
End Sub

Sub OPEN_FILE_Lugar()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstLugar = DbsAdminis.OpenRecordset("Lugar")
End Sub

Sub OPEN_FILE_Transporte()
    Set DbsAdminis = OpenDatabase("Impo.mdb", False, False, FILE_TYPE)
    Set rstTransporte = DbsAdminis.OpenRecordset("Transporte")
End Sub

Sub OPEN_FILE_Vendedor()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstVendedor = DbsAdminis.OpenRecordset("Vendedor")
End Sub

Sub OPEN_FILE_TipoPro()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstTipoPro = DbsAdminis.OpenRecordset("TipoPro")
End Sub

Sub OPEN_FILE_Chequera()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstChequera = DbsAdminis.OpenRecordset("Chequera")
End Sub

Sub OPEN_FILE_Control()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstControl = DbsAdminis.OpenRecordset("Control")
End Sub

Sub OPEN_FILE_Parametro()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstParametro = DbsAdminis.OpenRecordset("Parametro")
End Sub

Sub OPEN_FILE_Proyecto()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstProyecto = DbsAdminis.OpenRecordset("Proyecto")
End Sub

Sub OPEN_FILE_Movmes()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstMovmes = DbsAdminis.OpenRecordset("Movmes")
End Sub

Sub OPEN_FILE_Pedido()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPedido = DbsAdminis.OpenRecordset("Pedido")
End Sub

Sub OPEN_FILE_Posdat()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPosdat = DbsAdminis.OpenRecordset("Posdat")
End Sub

Sub OPEN_FILE_Movban()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstMovBan = DbsAdminis.OpenRecordset("Movban")
End Sub

Sub OPEN_FILE_Depositos()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstDepositos = DbsAdminis.OpenRecordset("Depositos")
End Sub

Sub OPEN_FILE_Tarjeta()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstTarjeta = DbsAdminis.OpenRecordset("Tarjeta")
End Sub

Sub OPEN_FILE_Caja()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCaja = DbsAdminis.OpenRecordset("Caja")
End Sub

Sub OPEN_FILE_Cupones()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCupones = DbsAdminis.OpenRecordset("Cupones")
End Sub

Sub OPEN_FILE_Venta()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstVenta = DbsAdminis.OpenRecordset("Venta")
End Sub

Sub OPEN_FILE_Gastos()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstGastos = DbsAdminis.OpenRecordset("Gastos")
End Sub

Sub OPEN_FILE_Movstk()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstMovstk = DbsAdminis.OpenRecordset("Movstk")
End Sub

Sub OPEN_FILE_Movi()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstMovi = DbsAdminis.OpenRecordset("Movi")
End Sub

Sub OPEN_FILE_Compras()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCompras = DbsAdminis.OpenRecordset("Compras")
End Sub

Sub OPEN_FILE_PasaClientes()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPasaClientes = DbsAdminis.OpenRecordset("PasaClientes")
End Sub

Sub OPEN_FILE_PasaCuentas()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPasaCuentas = DbsAdminis.OpenRecordset("PasaCuentas")
End Sub

Sub OPEN_FILE_PasaProve()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPasaProve = DbsAdminis.OpenRecordset("PasaProve")
End Sub

Sub OPEN_FILE_Conceptos()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstConceptos = DbsAdminis.OpenRecordset("Conceptos")
End Sub

Sub OPEN_FILE_ConceptoII()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstConceptoII = DbsAdminis.OpenRecordset("ConceptoII")
End Sub

Sub OPEN_FILE_Pto()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPto = DbsAdminis.OpenRecordset("Pto")
End Sub

Sub OPEN_FILE_PtoCue()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPtoCue = DbsAdminis.OpenRecordset("PtoCue")
End Sub

Sub OPEN_FILE_PtoVend()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPtoVend = DbsAdminis.OpenRecordset("PtoVend")
End Sub

Sub OPEN_FILE_Clientes()
    Set DbsVentas = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstClientes = DbsVentas.OpenRecordset("Cliente")
End Sub

Sub OPEN_FILE_Agenda()
    Set DbsVentas = OpenDatabase("Agenda.mdb", False, False, FILE_TYPE)
    Set rstAgenda = DbsVentas.OpenRecordset("Agenda")
End Sub

Sub OPEN_FILE_Plantilla()
    Set DbsVentas = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPlantilla = DbsVentas.OpenRecordset("Plantilla")
End Sub
 
Sub OPEN_FILE_Lineas()
    Set DbsVentas = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstLineas = DbsVentas.OpenRecordset("Lineas")
End Sub
 
Sub OPEN_FILE_Cotiza()
    Set DbsVentas = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCotiza = DbsVentas.OpenRecordset("Cotiza")
End Sub
 
Sub OPEN_FILE_Estadistica()
    Set DbsVentas = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstEstadistica = DbsVentas.OpenRecordset("Estadistica")
End Sub

Sub OPEN_FILE_Proveedor()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstProveedor = DbsAdminis.OpenRecordset("Proveedor")
End Sub

Sub OPEN_FILE_Empresa()
    Set DbsAdminis = OpenDatabase("Empresa.mdb", False, False, FILE_TYPE)
    Set rstEmpresa = DbsAdminis.OpenRecordset("Empresa")
End Sub

Sub OPEN_FILE_EmpreCon()
    Set DbsAdminis = OpenDatabase("EmpreCon.mdb", False, False, FILE_TYPE)
    Set rstEmpreCon = DbsAdminis.OpenRecordset("Empresa")
End Sub

Sub OPEN_FILE_Cuenta()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCuenta = DbsAdminis.OpenRecordset("Cuenta")
End Sub

Sub OPEN_FILE_CuentaCon()
    Set DbsAdminis = OpenDatabase(WEmpresaConta + "Cont.mdb", False, False, FILE_TYPE)
    Set rstCuentaCon = DbsAdminis.OpenRecordset("Cuenta")
End Sub

Sub OPEN_FILE_Compara()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCompara = DbsAdminis.OpenRecordset("Compara")
End Sub

Sub OPEN_FILE_CtaCtePrv()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCtaCtePrv = DbsAdminis.OpenRecordset("CtaCtePrv")
End Sub

Sub OPEN_FILE_Iva()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstIva = DbsAdminis.OpenRecordset("Iva")
End Sub

Sub OPEN_FILE_Ivacomp()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstIvaComp = DbsAdminis.OpenRecordset("IvaComp")
End Sub

Sub OPEN_FILE_Pagos()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPagos = DbsAdminis.OpenRecordset("Pagos")
End Sub

Sub OPEN_FILE_Solicitud()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstSolicitud = DbsAdminis.OpenRecordset("Solicitud")
End Sub

Sub OPEN_FILE_Ranking()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstRanking = DbsAdminis.OpenRecordset("Ranking")
End Sub

Sub OPEN_FILE_Posicion()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstPosicion = DbsAdminis.OpenRecordset("Posicion")
End Sub

Sub OPEN_FILE_Etiqueta()
    Set DbsAdminis = OpenDatabase("Lista.mdb", False, False, FILE_TYPE)
    Set rstEtiqueta = DbsAdminis.OpenRecordset("Etiqueta")
End Sub

Sub OPEN_FILE_Cash()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCash = DbsAdminis.OpenRecordset("Cash")
End Sub

Sub OPEN_FILE_GastosProy()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstGastosProy = DbsAdminis.OpenRecordset("GastosProy")
End Sub

Sub OPEN_FILE_Proceso1()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstProceso1 = DbsAdminis.OpenRecordset("Proceso1")
End Sub

Sub OPEN_FILE_Recibos()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstRecibos = DbsAdminis.OpenRecordset("Recibos")
End Sub

Sub OPEN_FILE_CheqCartera()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCheqCartera = DbsAdminis.OpenRecordset("CheqCartera")
End Sub

Sub OPEN_FILE_Auxi()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstAuxi = DbsAdminis.OpenRecordset("Auxiliar")
End Sub

Sub OPEN_FILE_Listado1()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstListado1 = DbsAdminis.OpenRecordset("Listado1")
End Sub

Sub OPEN_FILE_Ctacte()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstCtaCte = DbsAdminis.OpenRecordset("Ctacte")
End Sub

Sub OPEN_FILE_Ctacte1()
    Set DbsAdminis1 = OpenDatabase("\ventas1\Adminis.mdb", False, False, FILE_TYPE)
    Set rstCtaCte1 = DbsAdminis1.OpenRecordset("Ctacte")
End Sub

Sub OPEN_FILE_Ctacte2()
    Set DbsAdminis2 = OpenDatabase("\ventas2\Adminis.mdb", False, False, FILE_TYPE)
    Set rstCtaCte2 = DbsAdminis2.OpenRecordset("Ctacte")
End Sub

Sub OPEN_FILE_ImpCtacte()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstImpCtaCte = DbsAdminis.OpenRecordset("ImpCtacte")
End Sub

Sub OPEN_FILE_Numero()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstNumero = DbsAdminis.OpenRecordset("Numero")
End Sub

Sub OPEN_FILE_DescComp()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstDesccomp = DbsAdminis.OpenRecordset("DescComp")
End Sub

Sub OPEN_FILE_Factura()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstFactura = DbsAdminis.OpenRecordset("Factura")
End Sub

Sub OPEN_FILE_ImpreOrd()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstImpreOrd = DbsAdminis.OpenRecordset("ImpreOrd")
End Sub

Sub OPEN_FILE_NroRet()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstNroRet = DbsAdminis.OpenRecordset("NroRet")
End Sub

Sub OPEN_FILE_Retencion()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstRetencion = DbsAdminis.OpenRecordset("Retencion")
End Sub

Sub OPEN_FILE_Impcyb()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstImpCyb = DbsAdminis.OpenRecordset("Impcyb")
End Sub

Sub OPEN_FILE_Imputac()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstImputac = DbsAdminis.OpenRecordset("Imputac")
End Sub

Sub OPEN_FILE_Impproy()
    Set DbsAdminis = OpenDatabase(WEmpresa + "Admi.mdb", False, False, FILE_TYPE)
    Set rstImpProy = DbsAdminis.OpenRecordset("Impproy")
End Sub

Sub OPEN_FILE_Asiento()
    Set DbsAdminis = OpenDatabase(WEmpresaConta + "Cont.mdb", False, False, FILE_TYPE)
    Select Case WPosi
        Case 1
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento1")
        Case 2
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento2")
        Case 3
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento3")
        Case 4
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento4")
        Case 5
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento5")
        Case 6
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento6")
        Case 7
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento7")
        Case 8
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento8")
        Case 9
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento9")
        Case 10
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento10")
        Case 11
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento11")
        Case 12
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento12")
        Case Else
            Set rstAsiento = DbsAdminis.OpenRecordset("Asiento")
    End Select
End Sub

Sub NumbersOnly(T As Control, KeyAscii As Integer)
'This Sub allows only the digits 0 to 9, an initial minus sign and one period.
If KeyAscii < Asc(" ") Then     ' Is this Control char?
    Exit Sub                    ' Yes, let it pass
End If
If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
     'don't discard it
ElseIf KeyAscii = Asc(".") Then 'if its a period
     If InStr(1, T, ".") Then 'if there is already a period
          KeyAscii = 0   'discard it
     End If
ElseIf KeyAscii = Asc("-") And T.SelStart = 0 Then
     'keep it, it's an initial minus sign
Else
    KeyAscii = 0  ' Discard all other characters
End If
'Now prevent any characters in front of a minus sign
If Mid$(T.Text, T.SelStart + T.SelLength + 1, 1) = "-" Then
    KeyAscii = 0   ' Discard characters before -
End If
End Sub

Sub Errores(coderr As Integer, Archivo As String, Mensaje As String)

    e = coderr
    Select Case e
        Case 3021
            m$ = Mensaje$
            aaaaaa% = MsgBox(m$, 0, "Archivo de " + Archivo$)
        Case Else
            m$ = Mensaje$
            aaaaaa% = MsgBox(m$, 0, "Archivo de Vendedor")
    End Select
    
End Sub

Sub Ceros(Campo As String, largo As Integer)

    L% = 1
    cadena$ = ""
    While L% <= Len(Campo) And L% > 0
        If Mid$(Campo, L%, 1) <> Chr$(32) Then cadena$ = cadena$ + Mid$(Campo, L%, 1)
        L% = L% + 1
    Wend
    Campo = Right$(String$(40, "0") + cadena$, largo)
    
End Sub


