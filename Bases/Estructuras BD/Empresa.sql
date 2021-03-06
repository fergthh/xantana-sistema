if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Clave]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Clave]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[Empresa]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[Empresa]
GO

CREATE TABLE [dbo].[Clave] (
	[Clave] [nvarchar] (20) COLLATE Modern_Spanish_CI_AS NULL ,
	[Nivel] [int] NULL ,
	[Vto] [nvarchar] (8) COLLATE Modern_Spanish_CI_AS NULL 
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

