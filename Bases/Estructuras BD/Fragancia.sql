if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Articulo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Articulo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Auxiliar]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Auxiliar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Banco]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Banco]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Bcra]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Bcra]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Calidad]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Calidad]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Clave]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Clave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cliente]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClienteBonifica]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClienteBonifica]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ClienteLista]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ClienteLista]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Combo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Combo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ConceptoStock]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ConceptoStock]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Conceptos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Conceptos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CondPago]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CondPago]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Configuracion]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Configuracion]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CtaCte]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CtaCte]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[CtaCtePrv]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[CtaCtePrv]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Cuenta]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Cuenta]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Depositos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Depositos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Desccomp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Desccomp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Dolar]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Dolar]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Empresa]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Empresa]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Esencia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Esencia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Estadistica]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Estadistica]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Expreso]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Expreso]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Factura]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Factura]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Familia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Familia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Formula]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Formula]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Fragancia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Fragancia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[HistorialCliente]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[HistorialCliente]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Hoja]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Hoja]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ImpCyb]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ImpCyb]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ImpProy]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ImpProy]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Impctacte]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Impctacte]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ImpreOrd]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ImpreOrd]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[ImpreRecibo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[ImpreRecibo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Imputac]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Imputac]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Insumo]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Insumo]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[InsumoHistorial]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[InsumoHistorial]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[IvaComp]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[IvaComp]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Lineas]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Lineas]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Lista]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Lista]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Movstk]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Movstk]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[NroRet]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[NroRet]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Operador]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Operador]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Orden]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Orden]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pagos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Pagos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Parametro]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Parametro]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Pedido]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Pedido]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Precios]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Precios]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[PreciosHistorial]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[PreciosHistorial]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Proveedor]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Proveedor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Provincia]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Provincia]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Recibos]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Recibos]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Sector]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Sector]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Tamano]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Tamano]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TipoClie]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TipoClie]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[TipoPro]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[TipoPro]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Vendedor]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Vendedor]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[iva]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[iva]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[movban]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[movban]
GO

CREATE TABLE [dbo].[Articulo] (
	[Codigo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Linea] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fragancia] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Calidad] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tamano] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[DescripcionII] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Activo] [int] NULL ,
	[FechaInactivo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Facturable] [int] NULL ,
	[Sector] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Etiqueta] [int] NULL ,
	[Insumo] [char] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[Stock] [float] NULL ,
	[StockI] [float] NULL ,
	[StockII] [float] NULL ,
	[StockIII] [float] NULL ,
	[StockIV] [float] NULL ,
	[StockV] [float] NULL ,
	[StockVI] [float] NULL ,
	[InsumoII] [char] (25) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Auxiliar] (
	[Empresa] [smallint] NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Direccion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[Actividad] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaRetgan] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaRetIva] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaRetotro] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ctadeudores] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaEfectivo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaCheque] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaDocumentos] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaProveedores] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIva21] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIva5] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIva27] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIb] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaTerceros] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Auxi1] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Auxi2] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Auxi3] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Auxi4] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Varios] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Auxi5] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Auxi6] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Dolar] [float] NULL ,
	[Porce] [float] NULL ,
	[TipoCosto] [int] NULL ,
	[Impo1] [float] NULL ,
	[Impo2] [float] NULL ,
	[Impo3] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Banco] (
	[Banco] [int] NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Bcra] (
	[Codigo] [int] NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Calidad] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Clave] (
	[Clave] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nivel] [int] NULL ,
	[Vto] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cliente] (
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Razon] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Direccion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Localidad] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Provincia] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Postal] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[Email] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fax] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Telefono] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[Iva] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe1] [float] NULL ,
	[Importe2] [float] NULL ,
	[Importe3] [float] NULL ,
	[Importe4] [float] NULL ,
	[Importe5] [float] NULL ,
	[Importe6] [float] NULL ,
	[Dias] [int] NULL ,
	[Empresa] [int] NULL ,
	[Vendedor] [int] NULL ,
	[Descuento] [float] NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Expreso] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Partida] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descuento1] [float] NULL ,
	[Descuento2] [float] NULL ,
	[Descuento3] [float] NULL ,
	[EntregaII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[EntregaIII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[EntregaIV] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[EntregaV] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[UltimaCompra] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdUltimaCompra] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Zona] [int] NULL ,
	[NroLista] [int] NULL ,
	[Condicion] [int] NULL ,
	[UltimaLista] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdUltimaLista] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Marca] [int] NULL ,
	[Ordena] [float] NULL ,
	[ClienteII] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreRazon] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreDireccion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreLocalidad] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreProvincia] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreBultos] [int] NULL ,
	[ImpreDespacho] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoClie] [int] NULL ,
	[Fantasia] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[DireccionII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaAlta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[PorceIva] [float] NULL ,
	[NombreI] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[TelefonoI] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[EmailI] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[NombreII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[TelefonoII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[EmailII] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[NombreIII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[TelefonoIII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[EmailIII] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[LocalidadII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ProvinciaII] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[PostalII] [char] (15) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClienteBonifica] (
	[Clave] [char] (35) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Codigo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Linea] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fragancia] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Calidad] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tamano] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Desde] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Hasta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdDesde] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdHasta] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tope1] [float] NULL ,
	[Valor1] [float] NULL ,
	[Tope2] [float] NULL ,
	[Valor2] [float] NULL ,
	[Tope3] [float] NULL ,
	[Valor3] [float] NULL ,
	[Tope4] [float] NULL ,
	[Valor4] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ClienteLista] (
	[Clave] [char] (18) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Linea] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Lista] [char] (3) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Combo] (
	[Clave] [char] (12) COLLATE Modern_Spanish_CI_AS NULL ,
	[Codigo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Renglon] [int] NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoProceso] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Insumo] [char] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [float] NULL ,
	[Costo] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ConceptoStock] (
	[Concepto] [int] NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Conceptos] (
	[Concepto] [int] NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Rubro] [int] NULL ,
	[Agrupa] [int] NULL ,
	[DesAgrupa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[TituloI] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[TituloII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impo1] [float] NULL ,
	[Impo2] [float] NULL ,
	[Impo3] [float] NULL ,
	[Impo4] [float] NULL ,
	[Impo5] [float] NULL ,
	[Impo6] [float] NULL ,
	[Impo7] [float] NULL ,
	[Impo8] [float] NULL ,
	[Impo9] [float] NULL ,
	[Impo10] [float] NULL ,
	[Impo11] [float] NULL ,
	[Impo12] [float] NULL ,
	[Marca] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [int] NULL ,
	[ImporteII] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CondPago] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Dias] [int] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Configuracion] (
	[Clave] [int] NULL ,
	[Iva1] [float] NULL ,
	[Iva2] [float] NULL ,
	[Percepcion] [float] NULL ,
	[Punto] [int] NULL ,
	[CantiFac] [int] NULL ,
	[CantiRem] [int] NULL ,
	[CantiArti] [int] NULL ,
	[IvaServicio] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CtaCte] (
	[Clave] [char] (17) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[Punto] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [char] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[Renglon] [char] (4) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Estado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Vencimiento] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Total] [float] NULL ,
	[Saldo] [float] NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdVencimiento] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Neto] [float] NULL ,
	[Iva1] [float] NULL ,
	[Iva2] [float] NULL ,
	[Pedido] [int] NULL ,
	[Remito] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Orden] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Provincia] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Vendedor] [int] NULL ,
	[Costo] [float] NULL ,
	[Importe1] [float] NULL ,
	[Importe2] [float] NULL ,
	[Importe3] [float] NULL ,
	[Importe4] [float] NULL ,
	[Importe5] [float] NULL ,
	[Importe6] [float] NULL ,
	[Importe7] [float] NULL ,
	[Tipoventa] [int] NULL ,
	[Proyecto] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Paridad] [float] NULL ,
	[TotalUs] [float] NULL ,
	[SaldoUs] [float] NULL ,
	[Remito1] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Remito2] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Busqueda] [char] (13) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descuento] [float] NULL ,
	[Partida] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Pago] [int] NULL ,
	[Lista] [int] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Linea] [int] NULL ,
	[Exento] [float] NULL ,
	[NetoTotal] [float] NULL ,
	[Imprime] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Comision] [int] NULL ,
	[Expreso] [int] NULL ,
	[TipoIva] [int] NULL ,
	[NroRemito] [int] NULL ,
	[ClienteII] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cae] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[VtoCae] [char] (10) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[CtaCtePrv] (
	[Clave] [char] (21) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [int] NOT NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[Punto] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [char] (8) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Estado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Vencimiento] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Vencimiento1] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Total] [float] NULL ,
	[Saldo] [float] NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdVencimiento] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[SaldoList] [float] NULL ,
	[NroInterno] [int] NULL ,
	[Lista] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Acumulado] [float] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Empresa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreObservaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreOrdFecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Orden] [int] NULL ,
	[Item] [int] NULL ,
	[ClaveOrden] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[SubItem] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Cuenta] (
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Depositos] (
	[Clave] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Deposito] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Renglon] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Banco] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaOrd] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[Acredita] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[AcreditaOrd] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo2] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero2] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha2] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe2] [real] NULL ,
	[Observaciones2] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Empresa] [int] NULL ,
	[Impolista] [float] NULL ,
	[ClaveCheque] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[ClaveLectora] [char] (31) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Desccomp] (
	[Clave] [char] (17) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Punto] [int] NULL ,
	[Numero] [int] NULL ,
	[Renglon] [int] NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[Empresa] [int] NULL ,
	[WDate] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Dolar] (
	[Codigo] [int] NULL ,
	[Paridad] [float] NULL ,
	[ParidadII] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Empresa] (
	[Empresa] [smallint] NULL ,
	[Nombre] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Direccion] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Localidad] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Actividad] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaRetgan] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaRetIva] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaRetotro] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ctadeudores] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaEfectivo] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaCheque] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaDocumentos] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaProveedores] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIva21] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIva5] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIva27] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIb] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaGanancia] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaChequeRecha] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIvaVen] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ctaventas] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Telefono] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Condiva] [nvarchar] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[IngBrutos] [nvarchar] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[InicioAct] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaIva105] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaFondoFijo] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaImpInterno] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaImpCombustible] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[CtaRetSuss] [nvarchar] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[NombreBase] [char] (20) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Esencia] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Estadistica] (
	[Clave] [char] (17) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [int] NULL ,
	[Punto] [int] NULL ,
	[Numero] [int] NULL ,
	[Renglon] [int] NULL ,
	[Articulo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [int] NULL ,
	[Precio] [float] NULL ,
	[PrecioUs] [float] NULL ,
	[Importe] [float] NULL ,
	[ImporteUs] [float] NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Paridad] [float] NULL ,
	[Vendedor] [int] NULL ,
	[Rubro] [int] NULL ,
	[Linea] [int] NULL ,
	[Costo1] [float] NULL ,
	[Costo2] [float] NULL ,
	[Coeficiente] [float] NULL ,
	[Pedido] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe1] [float] NULL ,
	[Importe2] [float] NULL ,
	[Importe3] [float] NULL ,
	[Importe4] [float] NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[WArticulo] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Remito] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[WDate] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Marca] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[ClaveCtacte] [char] (12) COLLATE Modern_Spanish_CI_AS NULL ,
	[Imprefactura] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[NroFactura] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Partida] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descuento] [float] NULL ,
	[CantidadII] [int] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[PrecioII] [float] NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Comision] [float] NULL ,
	[TipoComision] [float] NULL ,
	[Lista] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoII] [int] NULL ,
	[PrecioSalva] [float] NULL ,
	[TipoIII] [int] NULL ,
	[CantidadImpre] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Expreso] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Direccion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Localidad] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Provincia] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Postal] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[Email] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fax] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Telefono] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[Iva] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Estado] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Factura] (
	[Clave] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [int] NULL ,
	[Punto] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Factura] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Renglon] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Direccion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Localidad] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Partida] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Remito] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[Neto] [float] NULL ,
	[Dto] [float] NULL ,
	[Neto1] [float] NULL ,
	[Iva1] [float] NULL ,
	[Iva2] [float] NULL ,
	[Total] [float] NULL ,
	[Percepcion] [float] NULL ,
	[Imprepago] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre1] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre2] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[CondIva] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Item] [int] NULL ,
	[Articulo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [float] NULL ,
	[Precio] [float] NULL ,
	[Dias] [int] NULL ,
	[TipoIva1] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoIva2] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Pago1] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Pago2] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Dia] [int] NULL ,
	[Mes] [int] NULL ,
	[Ano] [int] NULL ,
	[Impre3] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre4] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[PorceIva] [float] NULL ,
	[PordeDto] [float] NULL ,
	[Postal] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cae] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[VtoCae] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreBarra] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreBarraII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[DescriII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CantiII] [int] NULL ,
	[PrecioII] [float] NULL ,
	[ImpreIva] [char] (20) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Familia] (
	[Codigo] [int] NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Margen] [float] NULL ,
	[PorceIva] [float] NULL ,
	[Ubicacion] [int] NULL ,
	[Estado] [int] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Ordena] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Formula] (
	[Clave] [char] (27) COLLATE Modern_Spanish_CI_AS NULL ,
	[Articulo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Renglon] [int] NULL ,
	[Combo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoProceso] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Insumo] [char] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[Terminado] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [float] NULL ,
	[Costo] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Fragancia] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[HistorialCliente] (
	[Clave] [char] (9) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cliente] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Renglon] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ordfecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Hoja] (
	[Clave] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [int] NULL ,
	[Renglon] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Pedidio] [int] NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tiporeg] [int] NULL ,
	[Articulo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Insumo] [char] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[SemiTerminado] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [float] NULL ,
	[Envase] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[DesEnvase] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ImpCyb] (
	[Clave] [char] (23) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [int] NULL ,
	[Tipo] [int] NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Punto] [int] NULL ,
	[Numero] [int] NULL ,
	[Renglon] [int] NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Debito] [float] NULL ,
	[Credito] [float] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ImpProy] (
	[Clave] [char] (23) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [int] NULL ,
	[Tipo] [int] NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Punto] [int] NULL ,
	[Numero] [int] NULL ,
	[Renglon] [int] NULL ,
	[Proyecto] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpoList] [float] NULL ,
	[Concepto] [int] NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Impctacte] (
	[Clave] [char] (17) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[Punto] [int] NULL ,
	[Numero] [int] NOT NULL ,
	[Renglon] [int] NOT NULL ,
	[Cliente] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Estado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Vencimiento] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Total] [float] NULL ,
	[Saldo] [float] NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdVencimiento] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe1] [float] NULL ,
	[Importe2] [float] NULL ,
	[Importe3] [float] NULL ,
	[Importe4] [float] NULL ,
	[Importe5] [float] NULL ,
	[Importe6] [float] NULL ,
	[Importe7] [float] NULL ,
	[Periodo] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[DesEmpresa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Agrupa] [char] (17) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ImpreOrd] (
	[Orden] [int] NULL ,
	[Renglon] [int] NULL ,
	[Proveedor] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoReg] [int] NULL ,
	[Tipo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha1] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Total] [float] NULL ,
	[Retencion] [float] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Dia] [int] NULL ,
	[Mes] [int] NULL ,
	[Ano] [int] NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[RetIb] [float] NULL ,
	[NroRet] [int] NULL ,
	[NroRet1] [int] NULL ,
	[Tasa] [float] NULL ,
	[Impo1] [float] NULL ,
	[Impo2] [float] NULL ,
	[Exepcion] [float] NULL ,
	[Impo3] [float] NULL ,
	[Impo4] [float] NULL ,
	[NombreEmpresa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[DireccionEmpresa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[LocalidadEmpresa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CuitEmpresa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Empresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[ImpreRecibo] (
	[Copia] [int] NULL ,
	[Recibo] [int] NULL ,
	[Renglon] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Razon] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[PesosI] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[PesosII] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[Total] [float] NULL ,
	[FechaI] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[NumeroI] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImporteI] [float] NULL ,
	[Banco] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Sucursal] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[NumeroII] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaII] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImporteII] [float] NULL ,
	[Estructura] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImporteIII] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Imputac] (
	[Clave] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoMovi] [int] NULL ,
	[Comprobante] [int] NULL ,
	[TipoComp] [int] NULL ,
	[LetraComp] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[PuntoComp] [int] NULL ,
	[NroComp] [int] NULL ,
	[Renglon] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Debito] [float] NULL ,
	[Credito] [float] NULL ,
	[FechaOrd] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Titulo] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Titulolist] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre] [char] (8) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Insumo] (
	[Codigo] [char] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Color] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Linea] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[UnidadCaja] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Costo] [float] NULL ,
	[FechaCosto] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaCierre] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaUltimaEntrada] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaUltimaSalida] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Minimo] [float] NULL ,
	[Entradas] [float] NULL ,
	[Salidas] [float] NULL ,
	[Stock] [float] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[OrdFechaCosto] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoProveedor] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ubicacion] [int] NULL ,
	[Faltante] [float] NULL ,
	[CostoAnterior] [float] NULL ,
	[FechaCostoAnterior] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Articulo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Notas] [ntext] COLLATE Modern_Spanish_CI_AS NULL ,
	[Moneda] [int] NULL ,
	[StockI] [float] NULL ,
	[StockII] [float] NULL ,
	[StockIII] [float] NULL ,
	[StockIV] [float] NULL ,
	[StockV] [float] NULL ,
	[StockVI] [float] NULL ,
	[Asociado] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[InsumoAsociado] [char] (16) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [dbo].[InsumoHistorial] (
	[Codigo] [char] (16) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Costo] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[IvaComp] (
	[Clave] [char] (21) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [int] NULL ,
	[Tipo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Punto] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Vencimiento] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Periodo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Neto] [float] NULL ,
	[Iva21] [float] NULL ,
	[Iva5] [float] NULL ,
	[Iva27] [float] NULL ,
	[Ib] [float] NULL ,
	[Exento] [float] NULL ,
	[Contado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impre] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ordfecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Netolist] [float] NULL ,
	[ExentoList] [float] NULL ,
	[Concepto] [int] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Iva105] [float] NULL ,
	[ProveedorIva] [int] NULL ,
	[MovStk] [int] NULL ,
	[Banco] [int] NULL ,
	[ImpInterno] [float] NULL ,
	[ImpCombustible] [float] NULL ,
	[Centro] [int] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[OrdPeriodo] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpreNumero] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Agrupa] [int] NULL ,
	[DesAgrupa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Lineas] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Lista] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Movstk] (
	[Clave] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [int] NOT NULL ,
	[Renglon] [char] (2) COLLATE Modern_Spanish_CI_AS NOT NULL ,
	[fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Articulo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [float] NULL ,
	[Auxiliar] [float] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CantidadII] [float] NULL ,
	[Stock] [float] NULL ,
	[StkAnt] [float] NULL ,
	[StkAct] [float] NULL ,
	[Carpeta] [int] NULL ,
	[Deposito] [int] NULL ,
	[StkAntII] [float] NULL ,
	[StkAntIII] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[NroRet] (
	[Clave] [int] NULL ,
	[Numero] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Operador] (
	[Operador] [int] NULL ,
	[Clave] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[ClaveII] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nivel] [int] NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Orden] (
	[Clave] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [int] NULL ,
	[Renglon] [int] NULL ,
	[Proveedor] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Articulo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [float] NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Pagos] (
	[Clave] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Orden] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Renglon] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [int] NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaOrd] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoOrd] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[RetGanancias] [real] NULL ,
	[RetIva] [real] NULL ,
	[RetOtra] [real] NULL ,
	[Retencion] [real] NULL ,
	[TipoReg] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo1] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra1] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Punto1] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero1] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe1] [real] NULL ,
	[Tipo2] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero2] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha2] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[banco2] [int] NULL ,
	[Importe2] [real] NULL ,
	[Observaciones2] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Concepto] [int] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[FechaOrd2] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImpoList] [float] NULL ,
	[Cuenta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Solicitud] [int] NULL ,
	[ClaveCheque] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nroret] [int] NULL ,
	[NroRet1] [int] NULL ,
	[PorceIva] [float] NULL ,
	[PorceRIva] [float] NULL ,
	[Exepcion] [float] NULL ,
	[Impo1] [float] NULL ,
	[Impo2] [float] NULL ,
	[Impo3] [float] NULL ,
	[Impo4] [float] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[XBrutoIva] [float] NULL ,
	[XNetoIva] [float] NULL ,
	[XIvaIva] [float] NULL ,
	[ClaveLectora] [char] (31) COLLATE Modern_Spanish_CI_AS NULL ,
	[Agrupa] [int] NULL ,
	[DesAgrupa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Parametro] (
	[Clave] [int] NULL ,
	[Minimo1] [float] NULL ,
	[Minimo2] [float] NULL ,
	[Minimo3] [float] NULL ,
	[Escala1] [float] NULL ,
	[Escala2] [float] NULL ,
	[Escala3] [float] NULL ,
	[Escala4] [float] NULL ,
	[Escala5] [float] NULL ,
	[RetMinima] [float] NULL ,
	[PorceBienes] [float] NULL ,
	[PorceServicios] [float] NULL ,
	[PorceTranspo] [float] NULL ,
	[MinimoIva] [float] NULL ,
	[IvaInscripto] [float] NULL ,
	[IvaNoInscripto] [float] NULL ,
	[TasaGen] [float] NULL ,
	[TasaBienes] [float] NULL ,
	[Tasa1] [float] NULL ,
	[Tasa2] [float] NULL ,
	[Tasa3] [float] NULL ,
	[Tasa4] [float] NULL ,
	[Tasa5] [float] NULL ,
	[TasaNoInscripto] [float] NULL ,
	[Minimo4] [float] NULL ,
	[Mes1] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Mes2] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Mes3] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Mes4] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Mes5] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Mes6] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[IvaServicio] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Pedido] (
	[Clave] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [int] NULL ,
	[Renglon] [int] NULL ,
	[Articulo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cantidad] [float] NULL ,
	[Precio] [float] NULL ,
	[Importe] [float] NULL ,
	[Cliente] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe1] [float] NULL ,
	[Importe2] [float] NULL ,
	[Importe3] [float] NULL ,
	[Importe4] [float] NULL ,
	[OrdFecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descuento] [float] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[FecEntrega] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFecEntrega] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Facturado] [float] NULL ,
	[Cotiza] [int] NULL ,
	[Pago] [int] NULL ,
	[Partida] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle1] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle2] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle3] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle4] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle5] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle6] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle7] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle8] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle9] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Talle10] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ajuste] [float] NULL ,
	[Descuento1] [float] NULL ,
	[Descuento2] [float] NULL ,
	[Descuento3] [float] NULL ,
	[Lista] [int] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Saldo] [float] NULL ,
	[Marca] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[MarcaII] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Grupo] [int] NULL ,
	[Alternativa] [int] NULL ,
	[OCompra] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[NotaI] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[NotaII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[NotaIII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Posicion] [int] NULL ,
	[ClienteII] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Linea] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fragancia] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Calidad] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tamano] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Moneda] [int] NULL ,
	[PrecioII] [float] NULL ,
	[Dto] [float] NULL ,
	[ImporteII] [float] NULL ,
	[Observa] [char] (100) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fabrica] [float] NULL ,
	[Entregado] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Precios] (
	[Clave] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Codigo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Linea] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fragancia] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Calidad] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tamano] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Lista] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Desde] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Hasta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdDesde] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdHasta] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tope1] [float] NULL ,
	[Valor1] [float] NULL ,
	[Tope2] [float] NULL ,
	[Valor2] [float] NULL ,
	[Tope3] [float] NULL ,
	[Valor3] [float] NULL ,
	[Tope4] [float] NULL ,
	[Valor4] [float] NULL ,
	[Moneda] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[PreciosHistorial] (
	[Clave] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Codigo] [char] (25) COLLATE Modern_Spanish_CI_AS NULL ,
	[Linea] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fragancia] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Calidad] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tamano] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Lista] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Desde] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Hasta] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdDesde] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdHasta] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tope1] [float] NULL ,
	[Valor1] [float] NULL ,
	[Tope2] [float] NULL ,
	[Valor2] [float] NULL ,
	[Tope3] [float] NULL ,
	[Valor3] [float] NULL ,
	[Tope4] [float] NULL ,
	[Valor4] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Proveedor] (
	[Proveedor] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Direccion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Localidad] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Provincia] [int] NULL ,
	[Postal] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[Telefono] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Email] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ganancia] [int] NULL ,
	[Iva] [int] NULL ,
	[Dias] [int] NULL ,
	[Empresa] [int] NULL ,
	[Tipo] [float] NULL ,
	[Importe1] [float] NULL ,
	[Importe2] [float] NULL ,
	[Importe3] [float] NULL ,
	[Importe4] [float] NULL ,
	[Importe5] [float] NULL ,
	[Importe6] [float] NULL ,
	[NombreCheque] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[ReteIva] [int] NULL ,
	[PorceReteIva] [float] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Ordena] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Provincia] (
	[Codigo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Recibos] (
	[Clave] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Recibo] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Renglon] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cliente] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fechaord] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[TipoRec] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[RetGanancias] [float] NULL ,
	[RetIva] [float] NULL ,
	[RetOtra] [float] NULL ,
	[Retencion] [float] NULL ,
	[TipoReg] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Tipo1] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra1] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Punto1] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero1] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe1] [float] NULL ,
	[Tipo2] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero2] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha2] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[banco2] [char] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe2] [float] NULL ,
	[Estado2] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Empresa] [int] NULL ,
	[FechaOrd2] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Importe] [float] NULL ,
	[Observaciones] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Impolist] [float] NULL ,
	[Impo1list] [float] NULL ,
	[Destino] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuenta] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Orden] [int] NULL ,
	[Deposito] [int] NULL ,
	[NroRetGanancias] [int] NULL ,
	[NroRetIva] [int] NULL ,
	[NroRetOtra] [int] NULL ,
	[RetSuss] [float] NULL ,
	[NroRetSuss] [int] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[FechaRetIva] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaRetSuss] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaRetOtra] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[FechaRetGanancias] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFechaRetIva] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFechaRetSuss] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFechaRetOtra] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[OrdFechaRetGanancias] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Juridiccion] [int] NULL ,
	[Partida] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[ImporteI] [float] NULL ,
	[ImporteII] [float] NULL ,
	[ImporteIII] [float] NULL ,
	[Descuento] [float] NULL ,
	[CodigoBanco] [int] NULL ,
	[SucursalCheque] [int] NULL ,
	[TipoCheque] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[ClaseCheque] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[ClaveLectora] [char] (31) COLLATE Modern_Spanish_CI_AS NULL ,
	[ProveedorSalida] [int] NULL ,
	[BancoSalida] [int] NULL ,
	[Vendedor] [int] NULL ,
	[Lista] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Neto] [float] NULL ,
	[Porce] [float] NULL ,
	[Comision] [float] NULL ,
	[Titulo] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[TituloII] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Periodo] [char] (6) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[PorceDto] [float] NULL ,
	[TipoComi] [int] NULL ,
	[RetOtraII] [float] NULL ,
	[NroRetOtraII] [int] NULL ,
	[RetOtraIII] [float] NULL ,
	[NroRetOtraIII] [int] NULL ,
	[RetOtraIV] [float] NULL ,
	[NroRetOtraIV] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Sector] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Tamano] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TipoClie] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[TipoPro] (
	[Codigo] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Descripcion] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[CodigoEmpresa] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[Vendedor] (
	[Codigo] [int] NULL ,
	[Nombre] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Comision] [float] NULL ,
	[ComisionII] [float] NULL ,
	[CodigoEmpresa] [int] NULL ,
	[Telefono] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[Cuit] [char] (15) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ordena] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[iva] (
	[Clave] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [int] NULL ,
	[Tipo] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Letra] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Punto] [char] (4) COLLATE Modern_Spanish_CI_AS NULL ,
	[Numero] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Vencimiento] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Periodo] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Neto] [float] NULL ,
	[Iva21] [float] NULL ,
	[Iva5] [float] NULL ,
	[Iva27] [float] NULL ,
	[Ib] [float] NULL ,
	[Exento] [float] NULL ,
	[Contado] [char] (1) COLLATE Modern_Spanish_CI_AS NULL ,
	[Concepto] [int] NULL ,
	[Impre] [char] (2) COLLATE Modern_Spanish_CI_AS NULL ,
	[Ordfecha] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Empresa] [int] NULL ,
	[Iva105] [float] NULL ,
	[ImpInterno] [float] NULL ,
	[ImpCombustible] [float] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[movban] (
	[banco] [int] NULL ,
	[fecha] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[fechaord] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[Acredita] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[AcreditaOrd] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[observaciones] [char] (30) COLLATE Modern_Spanish_CI_AS NULL ,
	[numero] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[debito] [float] NULL ,
	[credito] [float] NULL ,
	[comprobante] [char] (8) COLLATE Modern_Spanish_CI_AS NULL ,
	[empresa] [int] NULL ,
	[Tipocomp] [char] (10) COLLATE Modern_Spanish_CI_AS NULL ,
	[Saldo] [float] NULL ,
	[DesEmpresa] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Periodo] [char] (50) COLLATE Modern_Spanish_CI_AS NULL ,
	[Proveedor] [char] (50) COLLATE Modern_Spanish_CI_AS NULL 
) ON [PRIMARY]
GO

