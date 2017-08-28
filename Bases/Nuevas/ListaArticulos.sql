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

Date: 2017-08-28 19:14:31
*/


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
