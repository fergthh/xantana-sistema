/*
Navicat SQL Server Data Transfer

Source Server         : SQLSERVER
Source Server Version : 100000
Source Host           : GASTON-PC\SQLEXPRESS:1433
Source Database       : Fragancia
Source Schema         : dbo

Target Server Type    : SQL Server
Target Server Version : 100000
File Encoding         : 65001

Date: 2017-09-04 19:49:02
*/


-- ----------------------------
-- Table structure for Articulo
-- ----------------------------
DROP TABLE [dbo].[Articulo]
GO
CREATE TABLE [dbo].[Articulo] (
[Codigo] char(25) NULL ,
[Linea] char(4) NULL ,
[Tipo] char(4) NULL ,
[Fragancia] char(4) NULL ,
[Calidad] char(4) NULL ,
[Tamano] char(4) NULL ,
[Descripcion] char(50) NULL ,
[DescripcionII] char(20) NULL ,
[Activo] int NULL ,
[FechaInactivo] char(10) NULL ,
[Facturable] int NULL ,
[Sector] char(4) NULL ,
[Etiqueta] int NULL ,
[Insumo] char(16) NULL ,
[Stock] float(53) NULL ,
[StockI] float(53) NULL ,
[StockII] float(53) NULL ,
[StockIII] float(53) NULL ,
[StockIV] float(53) NULL ,
[StockV] float(53) NULL ,
[StockVI] float(53) NULL ,
[InsumoII] char(25) NULL ,
[Rubro] int NULL DEFAULT ((0)) ,
[Costo] float(53) NULL DEFAULT ((0)) ,
[Observaciones] char(200) NULL ,
[Iva] char(1) NULL 
)


GO

-- ----------------------------
-- Table structure for Auxiliar
-- ----------------------------
DROP TABLE [dbo].[Auxiliar]
GO
CREATE TABLE [dbo].[Auxiliar] (
[Empresa] smallint NULL ,
[Nombre] char(50) NULL ,
[Direccion] char(50) NULL ,
[Cuit] char(15) NULL ,
[Actividad] char(50) NULL ,
[CtaRetgan] char(10) NULL ,
[CtaRetIva] char(10) NULL ,
[CtaRetotro] char(10) NULL ,
[Ctadeudores] char(10) NULL ,
[CtaEfectivo] char(10) NULL ,
[CtaCheque] char(10) NULL ,
[CtaDocumentos] char(10) NULL ,
[CtaProveedores] char(50) NULL ,
[CtaIva21] char(50) NULL ,
[CtaIva5] char(50) NULL ,
[CtaIva27] char(50) NULL ,
[CtaIb] char(50) NULL ,
[CtaTerceros] char(50) NULL ,
[Auxi1] char(10) NULL ,
[Auxi2] char(10) NULL ,
[Auxi3] char(10) NULL ,
[Auxi4] char(10) NULL ,
[Impre] char(50) NULL ,
[Varios] char(50) NULL ,
[Auxi5] char(10) NULL ,
[Auxi6] char(10) NULL ,
[Dolar] float(53) NULL ,
[Porce] float(53) NULL ,
[TipoCosto] int NULL ,
[Impo1] float(53) NULL ,
[Impo2] float(53) NULL ,
[Impo3] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Banco
-- ----------------------------
DROP TABLE [dbo].[Banco]
GO
CREATE TABLE [dbo].[Banco] (
[Banco] int NULL ,
[Nombre] char(50) NULL ,
[Cuenta] char(10) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for Bcra
-- ----------------------------
DROP TABLE [dbo].[Bcra]
GO
CREATE TABLE [dbo].[Bcra] (
[Codigo] int NULL ,
[Descripcion] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Calidad
-- ----------------------------
DROP TABLE [dbo].[Calidad]
GO
CREATE TABLE [dbo].[Calidad] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for Clave
-- ----------------------------
DROP TABLE [dbo].[Clave]
GO
CREATE TABLE [dbo].[Clave] (
[Clave] nvarchar(20) NULL ,
[Nivel] int NULL ,
[Vto] nvarchar(8) NULL 
)


GO

-- ----------------------------
-- Table structure for Cliente
-- ----------------------------
DROP TABLE [dbo].[Cliente]
GO
CREATE TABLE [dbo].[Cliente] (
[Cliente] char(10) NULL ,
[Razon] char(50) NULL ,
[Direccion] char(50) NULL ,
[Localidad] char(50) NULL ,
[Provincia] char(2) NULL ,
[Postal] char(15) NULL ,
[Email] char(100) NULL ,
[Fax] char(30) NULL ,
[Telefono] char(50) NULL ,
[Cuit] char(15) NULL ,
[Observaciones] char(100) NULL ,
[Iva] char(1) NULL ,
[Importe1] float(53) NULL ,
[Importe2] float(53) NULL ,
[Importe3] float(53) NULL ,
[Importe4] float(53) NULL ,
[Importe5] float(53) NULL ,
[Importe6] float(53) NULL ,
[Dias] int NULL ,
[Empresa] int NULL ,
[Vendedor] int NULL ,
[Descuento] float(53) NULL ,
[Cuenta] char(10) NULL ,
[CodigoEmpresa] int NULL ,
[Expreso] char(50) NULL ,
[Partida] char(2) NULL ,
[Descuento1] float(53) NULL ,
[Descuento2] float(53) NULL ,
[Descuento3] float(53) NULL ,
[EntregaII] char(50) NULL ,
[EntregaIII] char(50) NULL ,
[EntregaIV] char(50) NULL ,
[EntregaV] char(50) NULL ,
[UltimaCompra] char(10) NULL ,
[OrdUltimaCompra] char(8) NULL ,
[Zona] int NULL ,
[NroLista] int NULL ,
[Condicion] int NULL ,
[UltimaLista] char(10) NULL ,
[OrdUltimaLista] char(8) NULL ,
[Marca] int NULL ,
[Ordena] float(53) NULL ,
[ClienteII] char(6) NULL ,
[ImpreRazon] char(50) NULL ,
[ImpreDireccion] char(50) NULL ,
[ImpreLocalidad] char(50) NULL ,
[ImpreProvincia] char(50) NULL ,
[ImpreBultos] int NULL ,
[ImpreDespacho] char(2) NULL ,
[TipoClie] int NULL ,
[Fantasia] char(50) NULL ,
[DireccionII] char(50) NULL ,
[FechaAlta] char(10) NULL ,
[PorceIva] float(53) NULL ,
[NombreI] char(50) NULL ,
[TelefonoI] char(50) NULL ,
[EmailI] char(100) NULL ,
[NombreII] char(50) NULL ,
[TelefonoII] char(50) NULL ,
[EmailII] char(100) NULL ,
[NombreIII] char(50) NULL ,
[TelefonoIII] char(50) NULL ,
[EmailIII] char(100) NULL ,
[LocalidadII] char(50) NULL ,
[ProvinciaII] char(2) NULL ,
[PostalII] char(15) NULL ,
[ObservacionesII] char(100) NULL ,
[Responsable] char(30) NULL ,
[PaginaWeb] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for ClienteBonifica
-- ----------------------------
DROP TABLE [dbo].[ClienteBonifica]
GO
CREATE TABLE [dbo].[ClienteBonifica] (
[Clave] char(35) NULL ,
[Cliente] char(10) NULL ,
[Codigo] char(25) NULL ,
[Linea] char(4) NULL ,
[Tipo] char(4) NULL ,
[Fragancia] char(4) NULL ,
[Calidad] char(4) NULL ,
[Tamano] char(4) NULL ,
[Desde] char(10) NULL ,
[Hasta] char(10) NULL ,
[OrdDesde] char(8) NULL ,
[OrdHasta] char(8) NULL ,
[Tope1] float(53) NULL ,
[Valor1] float(53) NULL ,
[Tope2] float(53) NULL ,
[Valor2] float(53) NULL ,
[Tope3] float(53) NULL ,
[Valor3] float(53) NULL ,
[Tope4] float(53) NULL ,
[Valor4] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for ClienteLista
-- ----------------------------
DROP TABLE [dbo].[ClienteLista]
GO
CREATE TABLE [dbo].[ClienteLista] (
[Clave] char(18) NULL ,
[Cliente] char(10) NULL ,
[Linea] char(4) NULL ,
[Tipo] char(4) NULL ,
[Lista] char(3) NULL 
)


GO

-- ----------------------------
-- Table structure for Combo
-- ----------------------------
DROP TABLE [dbo].[Combo]
GO
CREATE TABLE [dbo].[Combo] (
[Clave] char(12) NULL ,
[Codigo] char(10) NULL ,
[Renglon] int NULL ,
[Descripcion] char(50) NULL ,
[TipoProceso] char(1) NULL ,
[Insumo] char(16) NULL ,
[Cantidad] float(53) NULL ,
[Costo] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Conceptos
-- ----------------------------
DROP TABLE [dbo].[Conceptos]
GO
CREATE TABLE [dbo].[Conceptos] (
[Concepto] int NULL ,
[Nombre] char(50) NULL ,
[Cuenta] char(10) NULL ,
[CodigoEmpresa] int NULL ,
[Rubro] int NULL ,
[Agrupa] int NULL ,
[DesAgrupa] char(50) NULL ,
[Importe] float(53) NULL ,
[TituloI] char(50) NULL ,
[TituloII] char(50) NULL ,
[Impo1] float(53) NULL ,
[Impo2] float(53) NULL ,
[Impo3] float(53) NULL ,
[Impo4] float(53) NULL ,
[Impo5] float(53) NULL ,
[Impo6] float(53) NULL ,
[Impo7] float(53) NULL ,
[Impo8] float(53) NULL ,
[Impo9] float(53) NULL ,
[Impo10] float(53) NULL ,
[Impo11] float(53) NULL ,
[Impo12] float(53) NULL ,
[Marca] char(1) NULL ,
[Tipo] int NULL ,
[ImporteII] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for ConceptoStock
-- ----------------------------
DROP TABLE [dbo].[ConceptoStock]
GO
CREATE TABLE [dbo].[ConceptoStock] (
[Concepto] int NULL ,
[Nombre] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for CondPago
-- ----------------------------
DROP TABLE [dbo].[CondPago]
GO
CREATE TABLE [dbo].[CondPago] (
[Codigo] char(4) NULL ,
[Nombre] char(50) NULL ,
[Dias] int NULL ,
[CodigoEmpresa] int NULL ,
[Observaciones] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Configuracion
-- ----------------------------
DROP TABLE [dbo].[Configuracion]
GO
CREATE TABLE [dbo].[Configuracion] (
[Clave] int NULL ,
[Iva1] float(53) NULL ,
[Iva2] float(53) NULL ,
[Percepcion] float(53) NULL ,
[Punto] int NULL ,
[CantiFac] int NULL ,
[CantiRem] int NULL ,
[CantiArti] int NULL ,
[IvaServicio] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for CtaCte
-- ----------------------------
DROP TABLE [dbo].[CtaCte]
GO
CREATE TABLE [dbo].[CtaCte] (
[Clave] char(17) NULL ,
[Letra] char(1) NULL ,
[Tipo] char(2) NOT NULL ,
[Punto] char(4) NULL ,
[Numero] char(8) NOT NULL ,
[Renglon] char(4) NOT NULL ,
[Cliente] char(10) NULL ,
[fecha] char(10) NULL ,
[Estado] char(1) NULL ,
[Vencimiento] char(10) NULL ,
[Total] float(53) NULL ,
[Saldo] float(53) NULL ,
[OrdFecha] char(8) NULL ,
[OrdVencimiento] char(8) NULL ,
[Impre] char(2) NULL ,
[Neto] float(53) NULL ,
[Iva1] float(53) NULL ,
[Iva2] float(53) NULL ,
[Pedido] int NULL ,
[Remito] char(20) NULL ,
[Orden] char(10) NULL ,
[Provincia] char(2) NULL ,
[Vendedor] int NULL ,
[Costo] float(53) NULL ,
[Importe1] float(53) NULL ,
[Importe2] float(53) NULL ,
[Importe3] float(53) NULL ,
[Importe4] float(53) NULL ,
[Importe5] float(53) NULL ,
[Importe6] float(53) NULL ,
[Importe7] float(53) NULL ,
[Tipoventa] int NULL ,
[Proyecto] char(10) NULL ,
[Paridad] float(53) NULL ,
[TotalUs] float(53) NULL ,
[SaldoUs] float(53) NULL ,
[Remito1] char(50) NULL ,
[Remito2] char(50) NULL ,
[Busqueda] char(13) NULL ,
[Descuento] float(53) NULL ,
[Partida] char(1) NULL ,
[Pago] int NULL ,
[Lista] int NULL ,
[CodigoEmpresa] int NULL ,
[Linea] int NULL ,
[Exento] float(53) NULL ,
[NetoTotal] float(53) NULL ,
[Imprime] char(1) NULL ,
[Comision] int NULL ,
[Expreso] int NULL ,
[TipoIva] int NULL ,
[NroRemito] int NULL ,
[ClienteII] char(6) NULL ,
[Cae] char(20) NULL ,
[VtoCae] char(10) NULL 
)


GO

-- ----------------------------
-- Table structure for CtaCtePrv
-- ----------------------------
DROP TABLE [dbo].[CtaCtePrv]
GO
CREATE TABLE [dbo].[CtaCtePrv] (
[Clave] char(21) NULL ,
[Proveedor] int NOT NULL ,
[Letra] char(1) NULL ,
[Tipo] char(2) NOT NULL ,
[Punto] char(4) NULL ,
[Numero] char(8) NOT NULL ,
[fecha] char(10) NULL ,
[Estado] char(1) NULL ,
[Vencimiento] char(10) NULL ,
[Vencimiento1] char(50) NULL ,
[Total] float(53) NULL ,
[Saldo] float(53) NULL ,
[OrdFecha] char(8) NULL ,
[OrdVencimiento] char(8) NULL ,
[Impre] char(2) NULL ,
[SaldoList] float(53) NULL ,
[NroInterno] int NULL ,
[Lista] char(1) NULL ,
[Acumulado] float(53) NULL ,
[Observaciones] char(50) NULL ,
[Empresa] char(50) NULL ,
[ImpreObservaciones] char(50) NULL ,
[ImpreOrdFecha] char(10) NULL ,
[CodigoEmpresa] int NULL ,
[Orden] int NULL ,
[Item] int NULL ,
[ClaveOrden] char(10) NULL ,
[SubItem] int NULL 
)


GO

-- ----------------------------
-- Table structure for Cuenta
-- ----------------------------
DROP TABLE [dbo].[Cuenta]
GO
CREATE TABLE [dbo].[Cuenta] (
[Cuenta] char(10) NULL ,
[Descripcion] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Depositos
-- ----------------------------
DROP TABLE [dbo].[Depositos]
GO
CREATE TABLE [dbo].[Depositos] (
[Clave] char(8) NULL ,
[Deposito] char(6) NULL ,
[Renglon] char(2) NULL ,
[Banco] int NULL ,
[Fecha] char(10) NULL ,
[FechaOrd] char(10) NULL ,
[Importe] float(53) NULL ,
[Acredita] char(10) NULL ,
[AcreditaOrd] char(8) NULL ,
[Tipo2] char(2) NULL ,
[Numero2] char(8) NULL ,
[Fecha2] char(10) NULL ,
[Importe2] real NULL ,
[Observaciones2] char(20) NULL ,
[Empresa] int NULL ,
[Impolista] float(53) NULL ,
[ClaveCheque] char(8) NULL ,
[ClaveLectora] char(31) NULL 
)


GO

-- ----------------------------
-- Table structure for Desccomp
-- ----------------------------
DROP TABLE [dbo].[Desccomp]
GO
CREATE TABLE [dbo].[Desccomp] (
[Clave] char(17) NULL ,
[Letra] char(1) NULL ,
[Tipo] char(2) NULL ,
[Punto] int NULL ,
[Numero] int NULL ,
[Renglon] int NULL ,
[Descripcion] char(50) NULL ,
[Importe] float(53) NULL ,
[Empresa] int NULL ,
[WDate] char(8) NULL ,
[Cuenta] char(10) NULL 
)


GO

-- ----------------------------
-- Table structure for Dolar
-- ----------------------------
DROP TABLE [dbo].[Dolar]
GO
CREATE TABLE [dbo].[Dolar] (
[Codigo] int NULL ,
[Paridad] float(53) NULL ,
[ParidadII] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for dtproperties
-- ----------------------------
DROP TABLE [dbo].[dtproperties]
GO
CREATE TABLE [dbo].[dtproperties] (
[id] int NOT NULL IDENTITY(1,1) ,
[objectid] int NULL ,
[property] varchar(64) NOT NULL ,
[value] varchar(255) NULL ,
[uvalue] nvarchar(255) NULL ,
[lvalue] image NULL ,
[version] int NOT NULL DEFAULT (0) 
)


GO

-- ----------------------------
-- Table structure for Empresa
-- ----------------------------
DROP TABLE [dbo].[Empresa]
GO
CREATE TABLE [dbo].[Empresa] (
[Empresa] smallint NULL ,
[Nombre] nvarchar(50) NULL ,
[Direccion] nvarchar(50) NULL ,
[Localidad] nvarchar(50) NULL ,
[Cuit] nvarchar(30) NULL ,
[Actividad] nvarchar(50) NULL ,
[CtaRetgan] nvarchar(10) NULL ,
[CtaRetIva] nvarchar(10) NULL ,
[CtaRetotro] nvarchar(10) NULL ,
[Ctadeudores] nvarchar(10) NULL ,
[CtaEfectivo] nvarchar(10) NULL ,
[CtaCheque] nvarchar(10) NULL ,
[CtaDocumentos] nvarchar(10) NULL ,
[CtaProveedores] nvarchar(10) NULL ,
[CtaIva21] nvarchar(10) NULL ,
[CtaIva5] nvarchar(10) NULL ,
[CtaIva27] nvarchar(10) NULL ,
[CtaIb] nvarchar(10) NULL ,
[CtaGanancia] nvarchar(10) NULL ,
[CtaChequeRecha] nvarchar(10) NULL ,
[CtaIvaVen] nvarchar(10) NULL ,
[Ctaventas] nvarchar(10) NULL ,
[Telefono] nvarchar(30) NULL ,
[Condiva] nvarchar(50) NULL ,
[IngBrutos] nvarchar(30) NULL ,
[InicioAct] nvarchar(10) NULL ,
[CtaIva105] nvarchar(10) NULL ,
[CtaFondoFijo] nvarchar(10) NULL ,
[CtaImpInterno] nvarchar(10) NULL ,
[CtaImpCombustible] nvarchar(10) NULL ,
[CtaRetSuss] nvarchar(10) NULL ,
[NombreBase] char(20) NULL 
)


GO

-- ----------------------------
-- Table structure for Esencia
-- ----------------------------
DROP TABLE [dbo].[Esencia]
GO
CREATE TABLE [dbo].[Esencia] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for Estadistica
-- ----------------------------
DROP TABLE [dbo].[Estadistica]
GO
CREATE TABLE [dbo].[Estadistica] (
[Clave] char(17) NULL ,
[Letra] char(1) NULL ,
[Tipo] int NULL ,
[Punto] int NULL ,
[Numero] int NULL ,
[Renglon] int NULL ,
[Articulo] char(25) NULL ,
[Cantidad] int NULL ,
[Precio] float(53) NULL ,
[PrecioUs] float(53) NULL ,
[Importe] float(53) NULL ,
[ImporteUs] float(53) NULL ,
[Cliente] char(10) NULL ,
[Paridad] float(53) NULL ,
[Vendedor] int NULL ,
[Rubro] int NULL ,
[Linea] int NULL ,
[Costo1] float(53) NULL ,
[Costo2] float(53) NULL ,
[Coeficiente] float(53) NULL ,
[Pedido] int NULL ,
[Fecha] char(10) NULL ,
[Importe1] float(53) NULL ,
[Importe2] float(53) NULL ,
[Importe3] float(53) NULL ,
[Importe4] float(53) NULL ,
[OrdFecha] char(8) NULL ,
[WArticulo] char(8) NULL ,
[Remito] char(10) NULL ,
[WDate] char(10) NULL ,
[Marca] char(1) NULL ,
[ClaveCtacte] char(12) NULL ,
[Imprefactura] char(20) NULL ,
[NroFactura] char(8) NULL ,
[Cuenta] char(10) NULL ,
[Partida] char(1) NULL ,
[Descuento] float(53) NULL ,
[CantidadII] int NULL ,
[CodigoEmpresa] int NULL ,
[PrecioII] float(53) NULL ,
[Descripcion] char(50) NULL ,
[Comision] float(53) NULL ,
[TipoComision] float(53) NULL ,
[Lista] char(1) NULL ,
[TipoII] int NULL ,
[PrecioSalva] float(53) NULL ,
[TipoIII] int NULL ,
[CantidadImpre] int NULL 
)


GO

-- ----------------------------
-- Table structure for Expreso
-- ----------------------------
DROP TABLE [dbo].[Expreso]
GO
CREATE TABLE [dbo].[Expreso] (
[Codigo] char(4) NULL ,
[Nombre] char(50) NULL ,
[Direccion] char(50) NULL ,
[Localidad] char(50) NULL ,
[Provincia] char(2) NULL ,
[Postal] char(15) NULL ,
[Email] char(20) NULL ,
[Fax] char(20) NULL ,
[Telefono] char(20) NULL ,
[Cuit] char(15) NULL ,
[Observaciones] char(100) NULL ,
[Iva] char(1) NULL ,
[CodigoEmpresa] int NULL ,
[Estado] int NULL 
)


GO

-- ----------------------------
-- Table structure for Factura
-- ----------------------------
DROP TABLE [dbo].[Factura]
GO
CREATE TABLE [dbo].[Factura] (
[Clave] char(20) NULL ,
[Letra] char(1) NULL ,
[Tipo] int NULL ,
[Punto] char(4) NULL ,
[Factura] char(8) NULL ,
[Renglon] int NULL ,
[Fecha] char(10) NULL ,
[Cliente] char(10) NULL ,
[Nombre] char(50) NULL ,
[Direccion] char(50) NULL ,
[Localidad] char(50) NULL ,
[Partida] char(1) NULL ,
[Cuit] char(20) NULL ,
[Remito] char(10) NULL ,
[Descripcion] char(50) NULL ,
[Importe] float(53) NULL ,
[Neto] float(53) NULL ,
[Dto] float(53) NULL ,
[Neto1] float(53) NULL ,
[Iva1] float(53) NULL ,
[Iva2] float(53) NULL ,
[Total] float(53) NULL ,
[Percepcion] float(53) NULL ,
[Imprepago] char(50) NULL ,
[Impre1] char(100) NULL ,
[Impre2] char(100) NULL ,
[CondIva] char(30) NULL ,
[Item] int NULL ,
[Articulo] char(25) NULL ,
[Cantidad] float(53) NULL ,
[Precio] float(53) NULL ,
[Dias] int NULL ,
[TipoIva1] char(1) NULL ,
[TipoIva2] char(1) NULL ,
[Pago1] char(1) NULL ,
[Pago2] char(1) NULL ,
[Dia] int NULL ,
[Mes] int NULL ,
[Ano] int NULL ,
[Impre3] char(50) NULL ,
[Impre4] char(50) NULL ,
[PorceIva] float(53) NULL ,
[PordeDto] float(53) NULL ,
[Postal] char(30) NULL ,
[Cae] char(20) NULL ,
[VtoCae] char(20) NULL ,
[ImpreBarra] char(50) NULL ,
[ImpreBarraII] char(50) NULL ,
[DescriII] char(50) NULL ,
[CantiII] int NULL ,
[PrecioII] float(53) NULL ,
[ImpreIva] char(20) NULL 
)


GO

-- ----------------------------
-- Table structure for Familia
-- ----------------------------
DROP TABLE [dbo].[Familia]
GO
CREATE TABLE [dbo].[Familia] (
[Codigo] int NULL ,
[Descripcion] char(50) NULL ,
[Margen] float(53) NULL ,
[PorceIva] float(53) NULL ,
[Ubicacion] int NULL ,
[Estado] int NULL ,
[CodigoEmpresa] int NULL ,
[Ordena] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Formula
-- ----------------------------
DROP TABLE [dbo].[Formula]
GO
CREATE TABLE [dbo].[Formula] (
[Clave] char(27) NULL ,
[Articulo] char(25) NULL ,
[Renglon] int NULL ,
[Combo] char(10) NULL ,
[TipoProceso] char(1) NULL ,
[Insumo] char(16) NULL ,
[Terminado] char(25) NULL ,
[Cantidad] float(53) NULL ,
[Costo] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Fragancia
-- ----------------------------
DROP TABLE [dbo].[Fragancia]
GO
CREATE TABLE [dbo].[Fragancia] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for HistorialCliente
-- ----------------------------
DROP TABLE [dbo].[HistorialCliente]
GO
CREATE TABLE [dbo].[HistorialCliente] (
[Clave] char(9) NULL ,
[Cliente] char(6) NULL ,
[Renglon] int NULL ,
[Fecha] char(10) NULL ,
[Ordfecha] char(8) NULL ,
[Observaciones] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Hoja
-- ----------------------------
DROP TABLE [dbo].[Hoja]
GO
CREATE TABLE [dbo].[Hoja] (
[Clave] char(8) NULL ,
[Numero] int NULL ,
[Renglon] int NULL ,
[Fecha] char(10) NULL ,
[OrdFecha] char(8) NULL ,
[Pedidio] int NULL ,
[Cliente] char(10) NULL ,
[Observaciones] char(50) NULL ,
[Tiporeg] int NULL ,
[Articulo] char(25) NULL ,
[Insumo] char(16) NULL ,
[SemiTerminado] char(25) NULL ,
[Descripcion] char(50) NULL ,
[Cantidad] float(53) NULL ,
[Envase] char(10) NULL ,
[DesEnvase] char(50) NULL ,
[Observa] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Impctacte
-- ----------------------------
DROP TABLE [dbo].[Impctacte]
GO
CREATE TABLE [dbo].[Impctacte] (
[Clave] char(17) NULL ,
[Letra] char(1) NULL ,
[Tipo] char(2) NOT NULL ,
[Punto] int NULL ,
[Numero] int NOT NULL ,
[Renglon] int NOT NULL ,
[Cliente] char(6) NULL ,
[fecha] char(10) NULL ,
[Estado] char(1) NULL ,
[Vencimiento] char(10) NULL ,
[Total] float(53) NULL ,
[Saldo] float(53) NULL ,
[OrdFecha] char(8) NULL ,
[OrdVencimiento] char(8) NULL ,
[Impre] char(8) NULL ,
[Importe1] float(53) NULL ,
[Importe2] float(53) NULL ,
[Importe3] float(53) NULL ,
[Importe4] float(53) NULL ,
[Importe5] float(53) NULL ,
[Importe6] float(53) NULL ,
[Importe7] float(53) NULL ,
[Periodo] char(50) NULL ,
[DesEmpresa] char(50) NULL ,
[Agrupa] char(17) NULL 
)


GO

-- ----------------------------
-- Table structure for ImpCyb
-- ----------------------------
DROP TABLE [dbo].[ImpCyb]
GO
CREATE TABLE [dbo].[ImpCyb] (
[Clave] char(23) NULL ,
[Proveedor] int NULL ,
[Tipo] int NULL ,
[Letra] char(1) NULL ,
[Punto] int NULL ,
[Numero] int NULL ,
[Renglon] int NULL ,
[Cuenta] char(10) NULL ,
[Debito] float(53) NULL ,
[Credito] float(53) NULL ,
[Fecha] char(10) NULL ,
[OrdFecha] char(8) NULL ,
[Observaciones] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for ImpProy
-- ----------------------------
DROP TABLE [dbo].[ImpProy]
GO
CREATE TABLE [dbo].[ImpProy] (
[Clave] char(23) NULL ,
[Proveedor] int NULL ,
[Tipo] int NULL ,
[Letra] char(1) NULL ,
[Punto] int NULL ,
[Numero] int NULL ,
[Renglon] int NULL ,
[Proyecto] char(10) NULL ,
[Importe] float(53) NULL ,
[Fecha] char(10) NULL ,
[OrdFecha] char(8) NULL ,
[ImpoList] float(53) NULL ,
[Concepto] int NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for ImpreOrd
-- ----------------------------
DROP TABLE [dbo].[ImpreOrd]
GO
CREATE TABLE [dbo].[ImpreOrd] (
[Orden] int NULL ,
[Renglon] int NULL ,
[Proveedor] int NULL ,
[Fecha] char(10) NULL ,
[TipoReg] int NULL ,
[Tipo] char(10) NULL ,
[Numero] char(8) NULL ,
[Fecha1] char(10) NULL ,
[Importe] float(53) NULL ,
[Descripcion] char(50) NULL ,
[Total] float(53) NULL ,
[Retencion] float(53) NULL ,
[Observaciones] char(50) NULL ,
[Dia] int NULL ,
[Mes] int NULL ,
[Ano] int NULL ,
[Nombre] char(50) NULL ,
[Cuenta] char(10) NULL ,
[RetIb] float(53) NULL ,
[NroRet] int NULL ,
[NroRet1] int NULL ,
[Tasa] float(53) NULL ,
[Impo1] float(53) NULL ,
[Impo2] float(53) NULL ,
[Exepcion] float(53) NULL ,
[Impo3] float(53) NULL ,
[Impo4] float(53) NULL ,
[NombreEmpresa] char(50) NULL ,
[DireccionEmpresa] char(50) NULL ,
[LocalidadEmpresa] char(50) NULL ,
[CuitEmpresa] char(50) NULL ,
[Empresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for ImpreRecibo
-- ----------------------------
DROP TABLE [dbo].[ImpreRecibo]
GO
CREATE TABLE [dbo].[ImpreRecibo] (
[Copia] int NULL ,
[Recibo] int NULL ,
[Renglon] int NULL ,
[Fecha] char(10) NULL ,
[Razon] char(50) NULL ,
[PesosI] char(100) NULL ,
[PesosII] char(100) NULL ,
[Total] float(53) NULL ,
[FechaI] char(10) NULL ,
[NumeroI] char(8) NULL ,
[ImporteI] float(53) NULL ,
[Banco] char(30) NULL ,
[Sucursal] char(10) NULL ,
[NumeroII] char(8) NULL ,
[FechaII] char(10) NULL ,
[ImporteII] float(53) NULL ,
[Estructura] char(30) NULL ,
[ImporteIII] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Imputac
-- ----------------------------
DROP TABLE [dbo].[Imputac]
GO
CREATE TABLE [dbo].[Imputac] (
[Clave] char(30) NULL ,
[TipoMovi] int NULL ,
[Comprobante] int NULL ,
[TipoComp] int NULL ,
[LetraComp] char(1) NULL ,
[PuntoComp] int NULL ,
[NroComp] int NULL ,
[Renglon] int NULL ,
[Fecha] char(10) NULL ,
[Observaciones] char(50) NULL ,
[Cuenta] char(10) NULL ,
[Debito] float(53) NULL ,
[Credito] float(53) NULL ,
[FechaOrd] char(8) NULL ,
[Titulo] char(50) NULL ,
[Nombre] char(50) NULL ,
[Titulolist] char(50) NULL ,
[Impre] char(8) NULL 
)


GO

-- ----------------------------
-- Table structure for Insumo
-- ----------------------------
DROP TABLE [dbo].[Insumo]
GO
CREATE TABLE [dbo].[Insumo] (
[Codigo] char(16) NULL ,
[Descripcion] char(50) NULL ,
[Color] char(30) NULL ,
[Linea] char(4) NULL ,
[Proveedor] char(10) NULL ,
[UnidadCaja] char(10) NULL ,
[Costo] float(53) NULL ,
[FechaCosto] char(10) NULL ,
[FechaCierre] char(10) NULL ,
[FechaUltimaEntrada] char(10) NULL ,
[FechaUltimaSalida] char(10) NULL ,
[Minimo] float(53) NULL ,
[Entradas] float(53) NULL ,
[Salidas] float(53) NULL ,
[Stock] float(53) NULL ,
[CodigoEmpresa] int NULL ,
[OrdFechaCosto] char(8) NULL ,
[CodigoProveedor] char(20) NULL ,
[Ubicacion] int NULL ,
[Faltante] float(53) NULL ,
[CostoAnterior] float(53) NULL ,
[FechaCostoAnterior] char(10) NULL ,
[Articulo] char(10) NULL ,
[Notas] ntext NULL ,
[Moneda] int NULL ,
[StockI] float(53) NULL ,
[StockII] float(53) NULL ,
[StockIII] float(53) NULL ,
[StockIV] float(53) NULL ,
[StockV] float(53) NULL ,
[StockVI] float(53) NULL ,
[Asociado] char(25) NULL ,
[InsumoAsociado] char(16) NULL 
)


GO

-- ----------------------------
-- Table structure for InsumoHistorial
-- ----------------------------
DROP TABLE [dbo].[InsumoHistorial]
GO
CREATE TABLE [dbo].[InsumoHistorial] (
[Codigo] char(16) NULL ,
[Fecha] char(10) NULL ,
[OrdFecha] char(8) NULL ,
[Costo] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for iva
-- ----------------------------
DROP TABLE [dbo].[iva]
GO
CREATE TABLE [dbo].[iva] (
[Clave] char(50) NULL ,
[Proveedor] int NULL ,
[Tipo] char(2) NULL ,
[Letra] char(1) NULL ,
[Punto] char(4) NULL ,
[Numero] char(8) NULL ,
[Fecha] char(10) NULL ,
[Vencimiento] char(10) NULL ,
[Periodo] char(10) NULL ,
[Neto] float(53) NULL ,
[Iva21] float(53) NULL ,
[Iva5] float(53) NULL ,
[Iva27] float(53) NULL ,
[Ib] float(53) NULL ,
[Exento] float(53) NULL ,
[Contado] char(1) NULL ,
[Concepto] int NULL ,
[Impre] char(2) NULL ,
[Ordfecha] char(8) NULL ,
[Empresa] int NULL ,
[Iva105] float(53) NULL ,
[ImpInterno] float(53) NULL ,
[ImpCombustible] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for IvaComp
-- ----------------------------
DROP TABLE [dbo].[IvaComp]
GO
CREATE TABLE [dbo].[IvaComp] (
[Clave] char(21) NULL ,
[Proveedor] int NULL ,
[Tipo] char(2) NULL ,
[Letra] char(1) NULL ,
[Punto] char(4) NULL ,
[Numero] char(8) NULL ,
[Fecha] char(10) NULL ,
[Vencimiento] char(10) NULL ,
[Periodo] char(10) NULL ,
[Neto] float(53) NULL ,
[Iva21] float(53) NULL ,
[Iva5] float(53) NULL ,
[Iva27] float(53) NULL ,
[Ib] float(53) NULL ,
[Exento] float(53) NULL ,
[Contado] char(1) NULL ,
[Impre] char(2) NULL ,
[Ordfecha] char(8) NULL ,
[Netolist] float(53) NULL ,
[ExentoList] float(53) NULL ,
[Concepto] int NULL ,
[Observaciones] char(50) NULL ,
[Iva105] float(53) NULL ,
[ProveedorIva] int NULL ,
[MovStk] int NULL ,
[Banco] int NULL ,
[ImpInterno] float(53) NULL ,
[ImpCombustible] float(53) NULL ,
[Centro] int NULL ,
[CodigoEmpresa] int NULL ,
[OrdPeriodo] char(8) NULL ,
[ImpreNumero] char(8) NULL ,
[Agrupa] int NULL ,
[DesAgrupa] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Lineas
-- ----------------------------
DROP TABLE [dbo].[Lineas]
GO
CREATE TABLE [dbo].[Lineas] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL ,
[Cliente] char(10) NULL 
)


GO

-- ----------------------------
-- Table structure for Lista
-- ----------------------------
DROP TABLE [dbo].[Lista]
GO
CREATE TABLE [dbo].[Lista] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for ListaArticulos
-- ----------------------------
DROP TABLE [dbo].[ListaArticulos]
GO
CREATE TABLE [dbo].[ListaArticulos] (
[Clave] char(21) NULL ,
[Articulo] char(15) NULL ,
[Lista] char(4) NULL ,
[Renglon] char(2) NULL ,
[Neto] float(53) NULL ,
[Precio] float(53) NULL ,
[Fecha] char(10) NULL ,
[FechaOrd] char(8) NULL 
)


GO

-- ----------------------------
-- Table structure for movban
-- ----------------------------
DROP TABLE [dbo].[movban]
GO
CREATE TABLE [dbo].[movban] (
[banco] int NULL ,
[fecha] char(10) NULL ,
[fechaord] char(8) NULL ,
[Acredita] char(10) NULL ,
[AcreditaOrd] char(8) NULL ,
[observaciones] char(30) NULL ,
[numero] char(10) NULL ,
[debito] float(53) NULL ,
[credito] float(53) NULL ,
[comprobante] char(8) NULL ,
[empresa] int NULL ,
[Tipocomp] char(10) NULL ,
[Saldo] float(53) NULL ,
[DesEmpresa] char(50) NULL ,
[Periodo] char(50) NULL ,
[Proveedor] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Movstk
-- ----------------------------
DROP TABLE [dbo].[Movstk]
GO
CREATE TABLE [dbo].[Movstk] (
[Clave] char(8) NULL ,
[Numero] int NOT NULL ,
[Renglon] char(2) NOT NULL ,
[fecha] char(10) NULL ,
[OrdFecha] char(8) NULL ,
[Articulo] char(10) NULL ,
[Cantidad] float(53) NULL ,
[Auxiliar] float(53) NULL ,
[Observaciones] char(50) NULL ,
[CantidadII] float(53) NULL ,
[Stock] float(53) NULL ,
[StkAnt] float(53) NULL ,
[StkAct] float(53) NULL ,
[Carpeta] int NULL ,
[Deposito] int NULL ,
[StkAntII] float(53) NULL ,
[StkAntIII] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for NroRet
-- ----------------------------
DROP TABLE [dbo].[NroRet]
GO
CREATE TABLE [dbo].[NroRet] (
[Clave] int NULL ,
[Numero] int NULL 
)


GO

-- ----------------------------
-- Table structure for Operador
-- ----------------------------
DROP TABLE [dbo].[Operador]
GO
CREATE TABLE [dbo].[Operador] (
[Operador] int NULL ,
[Clave] char(10) NULL ,
[ClaveII] char(10) NULL ,
[Nivel] int NULL ,
[Nombre] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Orden
-- ----------------------------
DROP TABLE [dbo].[Orden]
GO
CREATE TABLE [dbo].[Orden] (
[Clave] char(8) NULL ,
[Numero] int NULL ,
[Renglon] int NULL ,
[Proveedor] int NULL ,
[Fecha] char(10) NULL ,
[OrdFecha] char(8) NULL ,
[Articulo] char(10) NULL ,
[Cantidad] float(53) NULL ,
[Descripcion] char(50) NULL ,
[Observaciones] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Pagos
-- ----------------------------
DROP TABLE [dbo].[Pagos]
GO
CREATE TABLE [dbo].[Pagos] (
[Clave] char(8) NULL ,
[Orden] char(6) NULL ,
[Renglon] char(2) NULL ,
[Proveedor] int NULL ,
[Fecha] char(10) NULL ,
[FechaOrd] char(8) NULL ,
[TipoOrd] char(1) NULL ,
[RetGanancias] real NULL ,
[RetIva] real NULL ,
[RetOtra] real NULL ,
[Retencion] real NULL ,
[TipoReg] char(1) NULL ,
[Tipo1] char(2) NULL ,
[Letra1] char(1) NULL ,
[Punto1] char(4) NULL ,
[Numero1] char(8) NULL ,
[Importe1] real NULL ,
[Tipo2] char(2) NULL ,
[Numero2] char(8) NULL ,
[Fecha2] char(10) NULL ,
[banco2] int NULL ,
[Importe2] real NULL ,
[Observaciones2] char(30) NULL ,
[Concepto] int NULL ,
[Observaciones] char(50) NULL ,
[Importe] float(53) NULL ,
[FechaOrd2] char(10) NULL ,
[ImpoList] float(53) NULL ,
[Cuenta] char(10) NULL ,
[Solicitud] int NULL ,
[ClaveCheque] char(10) NULL ,
[Nroret] int NULL ,
[NroRet1] int NULL ,
[PorceIva] float(53) NULL ,
[PorceRIva] float(53) NULL ,
[Exepcion] float(53) NULL ,
[Impo1] float(53) NULL ,
[Impo2] float(53) NULL ,
[Impo3] float(53) NULL ,
[Impo4] float(53) NULL ,
[CodigoEmpresa] int NULL ,
[XBrutoIva] float(53) NULL ,
[XNetoIva] float(53) NULL ,
[XIvaIva] float(53) NULL ,
[ClaveLectora] char(31) NULL ,
[Agrupa] int NULL ,
[DesAgrupa] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Parametro
-- ----------------------------
DROP TABLE [dbo].[Parametro]
GO
CREATE TABLE [dbo].[Parametro] (
[Clave] int NULL ,
[Minimo1] float(53) NULL ,
[Minimo2] float(53) NULL ,
[Minimo3] float(53) NULL ,
[Escala1] float(53) NULL ,
[Escala2] float(53) NULL ,
[Escala3] float(53) NULL ,
[Escala4] float(53) NULL ,
[Escala5] float(53) NULL ,
[RetMinima] float(53) NULL ,
[PorceBienes] float(53) NULL ,
[PorceServicios] float(53) NULL ,
[PorceTranspo] float(53) NULL ,
[MinimoIva] float(53) NULL ,
[IvaInscripto] float(53) NULL ,
[IvaNoInscripto] float(53) NULL ,
[TasaGen] float(53) NULL ,
[TasaBienes] float(53) NULL ,
[Tasa1] float(53) NULL ,
[Tasa2] float(53) NULL ,
[Tasa3] float(53) NULL ,
[Tasa4] float(53) NULL ,
[Tasa5] float(53) NULL ,
[TasaNoInscripto] float(53) NULL ,
[Minimo4] float(53) NULL ,
[Mes1] char(10) NULL ,
[Mes2] char(10) NULL ,
[Mes3] char(10) NULL ,
[Mes4] char(10) NULL ,
[Mes5] char(10) NULL ,
[Mes6] char(10) NULL ,
[IvaServicio] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Pedido
-- ----------------------------
DROP TABLE [dbo].[Pedido]
GO
CREATE TABLE [dbo].[Pedido] (
[Clave] char(10) NULL ,
[Numero] int NULL ,
[Renglon] int NULL ,
[Articulo] char(25) NULL ,
[Cantidad] float(53) NULL ,
[Precio] float(53) NULL ,
[Importe] float(53) NULL ,
[Cliente] char(10) NULL ,
[Fecha] char(10) NULL ,
[Importe1] float(53) NULL ,
[Importe2] float(53) NULL ,
[Importe3] float(53) NULL ,
[Importe4] float(53) NULL ,
[OrdFecha] char(8) NULL ,
[Descuento] float(53) NULL ,
[Observaciones] char(50) NULL ,
[FecEntrega] char(10) NULL ,
[OrdFecEntrega] char(8) NULL ,
[Facturado] float(53) NULL ,
[Cotiza] int NULL ,
[Pago] int NULL ,
[Partida] char(1) NULL ,
[Talle1] char(4) NULL ,
[Talle2] char(4) NULL ,
[Talle3] char(4) NULL ,
[Talle4] char(4) NULL ,
[Talle5] char(4) NULL ,
[Talle6] char(4) NULL ,
[Talle7] char(4) NULL ,
[Talle8] char(4) NULL ,
[Talle9] char(4) NULL ,
[Talle10] char(4) NULL ,
[Ajuste] float(53) NULL ,
[Descuento1] float(53) NULL ,
[Descuento2] float(53) NULL ,
[Descuento3] float(53) NULL ,
[Lista] int NULL ,
[CodigoEmpresa] int NULL ,
[Saldo] float(53) NULL ,
[Marca] char(1) NULL ,
[MarcaII] char(1) NULL ,
[Descripcion] char(50) NULL ,
[Grupo] int NULL ,
[Alternativa] int NULL ,
[OCompra] char(10) NULL ,
[NotaI] char(50) NULL ,
[NotaII] char(50) NULL ,
[NotaIII] char(50) NULL ,
[Posicion] int NULL ,
[ClienteII] char(6) NULL ,
[Linea] char(4) NULL ,
[Tipo] char(4) NULL ,
[Fragancia] char(4) NULL ,
[Calidad] char(4) NULL ,
[Tamano] char(4) NULL ,
[Moneda] int NULL ,
[PrecioII] float(53) NULL ,
[Dto] float(53) NULL ,
[ImporteII] float(53) NULL ,
[Observa] char(100) NULL ,
[Fabrica] float(53) NULL ,
[Entregado] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Precios
-- ----------------------------
DROP TABLE [dbo].[Precios]
GO
CREATE TABLE [dbo].[Precios] (
[Clave] char(30) NULL ,
[Codigo] char(25) NULL ,
[Linea] char(4) NULL ,
[Tipo] char(4) NULL ,
[Fragancia] char(4) NULL ,
[Calidad] char(4) NULL ,
[Tamano] char(4) NULL ,
[Lista] char(4) NULL ,
[Desde] char(10) NULL ,
[Hasta] char(10) NULL ,
[OrdDesde] char(8) NULL ,
[OrdHasta] char(8) NULL ,
[Tope1] float(53) NULL ,
[Valor1] float(53) NULL ,
[Tope2] float(53) NULL ,
[Valor2] float(53) NULL ,
[Tope3] float(53) NULL ,
[Valor3] float(53) NULL ,
[Tope4] float(53) NULL ,
[Valor4] float(53) NULL ,
[Moneda] int NULL 
)


GO

-- ----------------------------
-- Table structure for PreciosHistorial
-- ----------------------------
DROP TABLE [dbo].[PreciosHistorial]
GO
CREATE TABLE [dbo].[PreciosHistorial] (
[Clave] char(30) NULL ,
[Codigo] char(25) NULL ,
[Linea] char(4) NULL ,
[Tipo] char(4) NULL ,
[Fragancia] char(4) NULL ,
[Calidad] char(4) NULL ,
[Tamano] char(4) NULL ,
[Lista] char(4) NULL ,
[Desde] char(10) NULL ,
[Hasta] char(10) NULL ,
[OrdDesde] char(8) NULL ,
[OrdHasta] char(8) NULL ,
[Tope1] float(53) NULL ,
[Valor1] float(53) NULL ,
[Tope2] float(53) NULL ,
[Valor2] float(53) NULL ,
[Tope3] float(53) NULL ,
[Valor3] float(53) NULL ,
[Tope4] float(53) NULL ,
[Valor4] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Proveedor
-- ----------------------------
DROP TABLE [dbo].[Proveedor]
GO
CREATE TABLE [dbo].[Proveedor] (
[Proveedor] char(8) NULL ,
[Nombre] char(50) NULL ,
[Direccion] char(50) NULL ,
[Localidad] char(50) NULL ,
[Provincia] int NULL ,
[Postal] char(20) NULL ,
[Cuit] char(15) NULL ,
[Telefono] char(30) NULL ,
[Email] char(50) NULL ,
[Observaciones] char(50) NULL ,
[Ganancia] int NULL ,
[Iva] int NULL ,
[Dias] int NULL ,
[Empresa] int NULL ,
[Tipo] float(53) NULL ,
[Importe1] float(53) NULL ,
[Importe2] float(53) NULL ,
[Importe3] float(53) NULL ,
[Importe4] float(53) NULL ,
[Importe5] float(53) NULL ,
[Importe6] float(53) NULL ,
[NombreCheque] char(50) NULL ,
[ReteIva] int NULL ,
[PorceReteIva] float(53) NULL ,
[CodigoEmpresa] int NULL ,
[Ordena] float(53) NULL 
)


GO

-- ----------------------------
-- Table structure for Provincia
-- ----------------------------
DROP TABLE [dbo].[Provincia]
GO
CREATE TABLE [dbo].[Provincia] (
[Codigo] char(2) NULL ,
[Descripcion] char(50) NULL 
)


GO

-- ----------------------------
-- Table structure for Recibos
-- ----------------------------
DROP TABLE [dbo].[Recibos]
GO
CREATE TABLE [dbo].[Recibos] (
[Clave] char(8) NULL ,
[Recibo] char(6) NULL ,
[Renglon] char(2) NULL ,
[Cliente] char(6) NULL ,
[Fecha] char(10) NULL ,
[Fechaord] char(8) NULL ,
[TipoRec] char(1) NULL ,
[RetGanancias] float(53) NULL ,
[RetIva] float(53) NULL ,
[RetOtra] float(53) NULL ,
[Retencion] float(53) NULL ,
[TipoReg] char(1) NULL ,
[Tipo1] char(2) NULL ,
[Letra1] char(1) NULL ,
[Punto1] char(4) NULL ,
[Numero1] char(8) NULL ,
[Importe1] float(53) NULL ,
[Tipo2] char(2) NULL ,
[Numero2] char(8) NULL ,
[Fecha2] char(10) NULL ,
[banco2] char(20) NULL ,
[Importe2] float(53) NULL ,
[Estado2] char(1) NULL ,
[Empresa] int NULL ,
[FechaOrd2] char(8) NULL ,
[Importe] float(53) NULL ,
[Observaciones] char(50) NULL ,
[Impolist] float(53) NULL ,
[Impo1list] float(53) NULL ,
[Destino] char(50) NULL ,
[Cuenta] char(50) NULL ,
[Orden] int NULL ,
[Deposito] int NULL ,
[NroRetGanancias] int NULL ,
[NroRetIva] int NULL ,
[NroRetOtra] int NULL ,
[RetSuss] float(53) NULL ,
[NroRetSuss] int NULL ,
[CodigoEmpresa] int NULL ,
[FechaRetIva] char(10) NULL ,
[FechaRetSuss] char(10) NULL ,
[FechaRetOtra] char(10) NULL ,
[FechaRetGanancias] char(10) NULL ,
[OrdFechaRetIva] char(8) NULL ,
[OrdFechaRetSuss] char(8) NULL ,
[OrdFechaRetOtra] char(8) NULL ,
[OrdFechaRetGanancias] char(8) NULL ,
[Juridiccion] int NULL ,
[Partida] char(1) NULL ,
[ImporteI] float(53) NULL ,
[ImporteII] float(53) NULL ,
[ImporteIII] float(53) NULL ,
[Descuento] float(53) NULL ,
[CodigoBanco] int NULL ,
[SucursalCheque] int NULL ,
[TipoCheque] char(1) NULL ,
[ClaseCheque] char(1) NULL ,
[ClaveLectora] char(31) NULL ,
[ProveedorSalida] int NULL ,
[BancoSalida] int NULL ,
[Vendedor] int NULL ,
[Lista] char(1) NULL ,
[Neto] float(53) NULL ,
[Porce] float(53) NULL ,
[Comision] float(53) NULL ,
[Titulo] char(50) NULL ,
[TituloII] char(50) NULL ,
[Periodo] char(6) NULL ,
[Cuit] char(15) NULL ,
[PorceDto] float(53) NULL ,
[TipoComi] int NULL ,
[RetOtraII] float(53) NULL ,
[NroRetOtraII] int NULL ,
[RetOtraIII] float(53) NULL ,
[NroRetOtraIII] int NULL ,
[RetOtraIV] float(53) NULL ,
[NroRetOtraIV] int NULL 
)


GO

-- ----------------------------
-- Table structure for Sector
-- ----------------------------
DROP TABLE [dbo].[Sector]
GO
CREATE TABLE [dbo].[Sector] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for Tamano
-- ----------------------------
DROP TABLE [dbo].[Tamano]
GO
CREATE TABLE [dbo].[Tamano] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for TipoArticulo
-- ----------------------------
DROP TABLE [dbo].[TipoArticulo]
GO
CREATE TABLE [dbo].[TipoArticulo] (
[Codigo] int NOT NULL ,
[Descripcion] char(50) NULL DEFAULT '' ,
[CodigoEmpresa] int NULL DEFAULT ((1)) 
)


GO

-- ----------------------------
-- Table structure for TipoClie
-- ----------------------------
DROP TABLE [dbo].[TipoClie]
GO
CREATE TABLE [dbo].[TipoClie] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for TipoPro
-- ----------------------------
DROP TABLE [dbo].[TipoPro]
GO
CREATE TABLE [dbo].[TipoPro] (
[Codigo] char(4) NULL ,
[Descripcion] char(50) NULL ,
[CodigoEmpresa] int NULL 
)


GO

-- ----------------------------
-- Table structure for Vendedor
-- ----------------------------
DROP TABLE [dbo].[Vendedor]
GO
CREATE TABLE [dbo].[Vendedor] (
[Codigo] int NULL ,
[Nombre] char(50) NULL ,
[Comision] float(53) NULL ,
[ComisionII] float(53) NULL ,
[CodigoEmpresa] int NULL ,
[Telefono] char(30) NULL ,
[Cuit] char(15) NULL ,
[Ordena] float(53) NULL 
)


GO

-- ----------------------------
-- Indexes structure for table dtproperties
-- ----------------------------

-- ----------------------------
-- Primary Key structure for table dtproperties
-- ----------------------------
ALTER TABLE [dbo].[dtproperties] ADD PRIMARY KEY ([id], [property])
GO
