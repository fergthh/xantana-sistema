VERSION 5.00
Begin VB.Form PrgEspejovarios 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "PrgEspejovarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WImpoDto1 As Double
Private WImpoDto2 As Double
Private WImpoDto3 As Double
Private WDescuento As Double
Private WCodIva As String
Private WDias As Integer
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private WTipo As String
Private WTipoIva As String
Private WProvincia As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private Mes(0 To 30) As String
Private XIndice As Single
Private WTipopro As Integer
Private XTalle As String
Private XColor As String
Private XArticulo As String
Private XTexto1 As String
Private XTexto2 As String
Private WPlazo1 As Integer
Private WVencimiento As String
Dim WPedido(1000) As String
Dim WSaldo As Double
Dim CantiFac As Integer
Dim CantiRem As Integer
Dim CantiArti As Integer
Dim ZMes As String
Dim ZAno As String
Dim ZZCambia As String
Dim ZZPasaImpre As Integer
Dim ZZPasaDatos(100, 10) As String



Dim ZZClave As String
Dim ZZLetra As String
Dim ZZTipo As String
Dim ZZPunto As String
Dim ZZNumero As String
Dim ZZRenglon As String
Dim ZZCliente As String
Dim ZZfecha As String
Dim ZZEstado As String
Dim ZZVencimiento As String
Dim ZZTotal As String
Dim ZZSaldo As String
Dim ZZOrdFecha As String
Dim ZZOrdVencimiento As String
Dim ZZImpre As String
Dim ZZNeto As String
Dim ZZIva1 As String
Dim ZZIva2 As String
Dim ZZPedido As String
Dim ZZRemito As String
Dim ZZOrden As String
Dim ZZProvincia As String
Dim ZZVendedor As String
Dim ZZCosto As String
Dim ZZImporte1 As String
Dim ZZImporte2 As String
Dim ZZImporte3 As String
Dim ZZImporte4 As String
Dim ZZImporte5 As String
Dim ZZImporte6 As String
Dim ZZImporte7 As String
Dim ZZTipoventa As String
Dim ZZProyecto As String
Dim ZZParidad As String
Dim ZZTotalUs As String
Dim ZZSaldoUs As String
Dim ZZRemito1 As String
Dim ZZRemito2 As String
Dim ZZBusqueda As String
Dim ZZDescuento As String
Dim ZZPartida As String
Dim ZZPago As String
Dim ZBaja As String
Dim ZZDireccionExpreso As String

Dim ZZCantidad As String
Dim ZZCantidadII As String

Dim WVector(100, 10) As String

Dim WWArticulo As String
Dim WWDescripcion As String
Dim WWCantidad As String
Dim WWPrecio As String
Dim WWImpre As Double
Dim WWDto(10) As Double
Dim WWIva(10) As Double

Dim ZFactuImpre(1000, 5) As String
Dim ZZImpreBarra As String
Dim ZZImpreBarraII As String
Dim ZZImpreExibidor(100, 10) As String

Dim XXPrecio As Double
Dim XXDto As Double
Dim XXComision As Double
Dim XXImpoComision As Double
Dim ZZZLetra As String
Dim ZZZPunto As String
Dim ZZZNumeroFac As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String



Private Sub Calcula_FechaVto()

    Rem With rstPago
    Rem    .Index = "Pago"
    Rem    .Seek "=", WPago1
    Rem    If .NoMatch = False Then
    Rem        WPlazo1 = !Plazo
    Rem        WTasa = !Tasa
    Rem        WDescuento = !Descuento
    Rem        WPago = !Nombre
    Rem    End If
    Rem End With
    
    Rem WFecha = Fecha.Text
    Rem Call Calcula_vencimiento(WFecha, WPlazo1, Wvencimiento)
    
    Rem With rstPago
    Rem     .Index = "Pago"
    Rem     .Seek "=", WPago2
    Rem     If .NoMatch = False Then
    Rem         WPlazo2 = !Plazo
    Rem     End If
    Rem End With
    
    Rem Call Calcula_vencimiento(WFecha, WPlazo2, WVencimiento1)

End Sub

Private Sub Anula_Click()

    T$ = "Anulacion de Comprobantes"
    M$ = "Desea Anular el Comprobante "
    Respuesta% = MsgBox(M$, 32 + 4, T$)
    If Respuesta% = 6 Then

        T$ = "Anulacion de Comprobantes"
        M$ = "Esta Seguro que Desea Anular el Comprobante "
        Respuesta% = MsgBox(M$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            ZBaja = "N"
            If Val(Pedido.Text) <> 0 Then
                T$ = "Baja de Comprobantes"
                M$ = "Desea Restaurar el saldo del pedido"
                Respuesta% = MsgBox(M$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    ZBaja = "S"
                End If
            End If
        
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "01"
                
            ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
               
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                ZSaldo = rstCtaCte!Saldo
                ZTotal = rstCtaCte!Total
                rstCtaCte.Close
                If ZSaldo <> ZTotal Then
                    M$ = "El comprobante se encuentra total o parcialmente cancelado"
                    a% = MsgBox(M$, 0, "Eliminacion de Comprobantes")
                    Exit Sub
                End If
            End If
            
            Erase WVector
        
            For WRenglon = 1 To 50
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "01" + Auxi + Auxi1
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                
                    Articulo = rstEstadistica!Articulo
                    Cantidad = rstEstadistica!Cantidad
                    CantidadII = rstEstadistica!CantidadII
                        
                    rstEstadistica.Close
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Salidas = Salidas - " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
                    If ZBaja = "S" Then
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Pedido SET "
                        ZSql = ZSql + " Facturado = Facturado - " + "'" + Str$(Cantidad) + "'"
                        ZSql = ZSql + " Where Numero = " + "'" + Pedido.Text + "'"
                        ZSql = ZSql + " and Articulo = " + "'" + Articulo + "'"
                        spPedido = ZSql
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                End If
                
            Next WRenglon
            
            ZSql = ""
            ZSql = ZSql + "DELETE Estadistica"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CtaCte SET"
            ZSql = ZSql + " Total = 0 ,"
            ZSql = ZSql + " Saldo = 0 ,"
            ZSql = ZSql + " TotalUs = 0 ,"
            ZSql = ZSql + " SaldoUs = 0 ,"
            ZSql = ZSql + " Neto = 0 ,"
            ZSql = ZSql + " NetoTotal = 0 ,"
            ZSql = ZSql + " Iva1 = 0 ,"
            ZSql = ZSql + " Iva2 = 0"
            ZSql = ZSql + " Where Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            Call Limpia_Click
            
            CLIENTE.SetFocus
            
        End If
        
    End If


End Sub



Private Sub Impresion_FacturaFe()


    ZSql = ""
    ZSql = ZSql + "DELETE Factura"
    spFactura = ZSql
    Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        ZZProvincia = rstCliente!Provincia
        ZZCodIva = rstCliente!Iva
        ZZRazon = rstCliente!Razon
        ZZDireccion = rstCliente!Direccion
        ZZLocalidad = rstCliente!Localidad
        ZZPostal = rstCliente!Postal
        ZZCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZCuitII = rstExpreso!Cuit
        rstExpreso.Close
    End If
    
    
    
    
    ZZLetra = "X"
    ZZTipo = "01"
    ZZPunto = "0001"
    Auxi1 = Numero.Text
    Call Ceros(Auxi1, 8)
    ZZFactura = Auxi1
    ZZfecha = Fecha.Text
    ZZCliente = CLIENTE.Text
    ZZNombre = Trim(ZZRazon)
    ZZDireccion = Trim(ZZDireccion)
    ZZLocalidad = Trim(ZZLocalidad)
    ZZPartida = Partida.Text
    ZZNeto = Neto.Caption
    ZZDto = Dto.Caption
    ZZNeto1 = SubTotal.Caption
    ZZIva1 = Iva1.Caption
    ZZIva2 = Iva2.Caption
    ZZTotal = Total.Caption
    ZZImprepago = Left$(DesPago.Caption, 35)
    ZZImpreIva = Iva(Val(ZZCodIva))
    If TipoIva.ListIndex = 0 Then
        ZZPorceIva = "21"
            Else
        ZZPorceIva = "10.5"
    End If
    ZZPorceDto = Descuento.Text
    Select Case Partida.Text
        Case "/"
            ZZPostal = "CP:B" + WPostal + "BIE"
        Case "?"
            ZZPostal = "CP%B" + WPostal + "BIE"
        Case Else
            ZZPostal = "CP B" + WPostal + "BIE"
    End Select
    
    Call Calcula_Barra
    ZZLugarFactura = 0
    
    For a = 1 To 38
    
        If Trim(WVector1.TextMatrix(a, 1)) <> "" Then
            
            ZZLugarFactura = ZZLugarFactura + 1
    
            ZZRenglon = Str$(ZZLugarFactura)
            Auxi1 = ZZRenglon
            Call Ceros(Auxi1, 2)
            ZZRenglon = Auxi1
            
            ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
            
            ZZItem = Str$(a)
            
            ZZArticulo = WVector1.TextMatrix(a, 1)
            ZZDescripcion = WVector1.TextMatrix(a, 2)
            ZZZCantidad = Val(WVector1.TextMatrix(a, 3))
            If ZZPasaImpre = 0 Then
                If Val(ZZZCantidad) <> 0 Then
                    Select Case Partida.Text
                        Case "/"
                            ZZZCantidad = ZZZCantidad / 2
                        Case "?"
                            ZZZCantidad = ZZZCantidad / 12
                        Case Else
                    End Select
                End If
            End If
            ZZZPrecio = Val(WVector1.TextMatrix(a, 4))
            ZZZImporte = ZZZPrecio * ZZZCantidad
            
            ZZCantidad = Str$(ZZZCantidad)
            ZZPrecio = Str$(ZZZPrecio)
            ZZImporte = Str$(ZZZImporte)
            
            If Trim(ZZArticulo) = "" Then
                ZZItem = ""
                ZZArticulo = ""
                ZZDescripcion = ""
                ZZCantidad = ""
                ZZPrecio = ""
                ZZImporte = ""
            End If
            
            ZZDescriII = ""
            ZZCantiII = ""
            ZZPrecioII = ""
            
            ZZImpre1 = Remito.Text
            Auxi2 = Numero.Text
            Call Ceros(Auxi2, 8)
            ZZImpre3 = Auxi2
            ZZImpre4 = "FACTURA"
            
            ZZRemito = OCompra.Text
            
            ZZExibidor = "N"
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Exibidores"
            ZSql = ZSql + " Where Exibidores.Codigo = " + "'" + ZZArticulo + "'"
            spExibidores = ZSql
            Set rstExibidores = db.OpenRecordset(spExibidores, dbOpenSnapshot, dbSQLPassThrough)
            If rstExibidores.RecordCount > 0 Then
                rstExibidores.Close
                ZZExibidor = "S"
            End If
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Factura ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Factura ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "fecha ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Postal ,"
            ZSql = ZSql + "Partida ,"
            ZSql = ZSql + "Cuit  ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Dto ,"
            ZSql = ZSql + "Neto ,"
            ZSql = ZSql + "Neto1 ,"
            ZSql = ZSql + "Iva1 ,"
            ZSql = ZSql + "Iva2 ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Item ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "Imprepago ,"
            ZSql = ZSql + "CondIva ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "Impre1 ,"
            ZSql = ZSql + "Impre3 ,"
            ZSql = ZSql + "Impre4 ,"
            ZSql = ZSql + "Cae ,"
            ZSql = ZSql + "VtoCae ,"
            ZSql = ZSql + "ImpreBarra ,"
            ZSql = ZSql + "ImpreBarraII ,"
            ZSql = ZSql + "DescriII ,"
            ZSql = ZSql + "CantiII ,"
            ZSql = ZSql + "PrecioII ,"
            ZSql = ZSql + "PorceIva ,"
            ZSql = ZSql + "PordeDto )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZPunto + "',"
            ZSql = ZSql + "'" + ZZFactura + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZNombre + "',"
            ZSql = ZSql + "'" + ZZDireccion + "',"
            ZSql = ZSql + "'" + ZZLocalidad + "',"
            ZSql = ZSql + "'" + ZZPostal + "',"
            ZSql = ZSql + "'" + ZZPartida + "',"
            ZSql = ZSql + "'" + ZZCuit + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZDto + "',"
            ZSql = ZSql + "'" + ZZNeto + "',"
            ZSql = ZSql + "'" + ZZNeto1 + "',"
            ZSql = ZSql + "'" + ZZIva1 + "',"
            ZSql = ZSql + "'" + ZZIva2 + "',"
            ZSql = ZSql + "'" + ZZTotal + "',"
            ZSql = ZSql + "'" + ZZItem + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZImprepago + "',"
            ZSql = ZSql + "'" + ZZImpreIva + "',"
            ZSql = ZSql + "'" + ZZRemito + "',"
            ZSql = ZSql + "'" + ZZImpre1 + "',"
            ZSql = ZSql + "'" + ZZImpre3 + "',"
            ZSql = ZSql + "'" + ZZImpre4 + "',"
            ZSql = ZSql + "'" + Cae.Text + "',"
            ZSql = ZSql + "'" + VtoCae.Text + "',"
            ZSql = ZSql + "'" + ZZImpreBarra + "',"
            ZSql = ZSql + "'" + ZZImpreBarraII + "',"
            ZSql = ZSql + "'" + ZZDescriII + "',"
            ZSql = ZSql + "'" + ZZCantiII + "',"
            ZSql = ZSql + "'" + ZZPrecioII + "',"
            ZSql = ZSql + "'" + ZZPorceIva + "',"
            ZSql = ZSql + "'" + ZZPorceDto + "')"
                                    
            spFactura = ZSql
            Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
            If ZZExibidor = "S" Then
                    
                Erase ZZImpreExibidor
                ZZLugarExibidor = 0
                    
                For WRenglon = 1 To 100
                
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    
                    WClave = Trim(ZZArticulo) + Auxi1
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Exibidores"
                    ZSql = ZSql + " Where Exibidores.Clave = " + "'" + WClave + "'"
                    spExibidores = ZSql
                    Set rstExibidores = db.OpenRecordset(spExibidores, dbOpenSnapshot, dbSQLPassThrough)
                    If rstExibidores.RecordCount > 0 Then
                        
                        ZZLugarExibidor = ZZLugarExibidor + 1
                            
                        Auxi = Trim(rstExibidores!Articulo)
                        ZZImpreExibidor(ZZLugarExibidor, 1) = Trim(rstExibidores!Articulo)
                        ZZImpreExibidor(ZZLugarExibidor, 3) = Str$(rstExibidores!Cantidad)
                        
                        rstExibidores.Close
                            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Articulo"
                        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            ZZImpreExibidor(ZZLugarExibidor, 2) = rstArticulo!Descripcion
                            ZZImpreExibidor(ZZLugarExibidor, 4) = Str$(rstArticulo!Precio)
                            rstArticulo.Close
                        End If
                                
                    End If
                
                Next WRenglon
                
                For WRenglon = 1 To ZZLugarExibidor
                
                    ZZLugarFactura = ZZLugarFactura + 1
            
                    ZZRenglon = Str$(ZZLugarFactura)
                    Auxi1 = ZZRenglon
                    Call Ceros(Auxi1, 2)
                    ZZRenglon = Auxi1
                    
                    ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
                    
                    ZZItem = ""
                    ZZArticulo = ""
                    ZZDescripcion = ""
                    ZZCantidad = ""
                    ZZPrecio = ""
                    ZZImporte = ""
                    
                    If WRenglon = 1 Then
                        ZZArticulo = "Contenido"
                    End If
                    
                    Rem ZZArticulo = ZZImpreExibidor(WRenglon, 1)
                    ZZDescriII = ZZImpreExibidor(WRenglon, 2)
                    ZZCantiII = ZZImpreExibidor(WRenglon, 3)
                    ZZPrecioII = ZZImpreExibidor(WRenglon, 4)
                    
                    ZZImpre1 = Remito.Text
                    Auxi2 = Numero.Text
                    Call Ceros(Auxi2, 8)
                    ZZImpre3 = Auxi2
                    ZZImpre4 = "FACTURA"
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Factura ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "Letra ,"
                    ZSql = ZSql + "Tipo ,"
                    ZSql = ZSql + "Punto ,"
                    ZSql = ZSql + "Factura ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "fecha ,"
                    ZSql = ZSql + "Cliente ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "Direccion ,"
                    ZSql = ZSql + "Localidad ,"
                    ZSql = ZSql + "Postal ,"
                    ZSql = ZSql + "Partida ,"
                    ZSql = ZSql + "Cuit  ,"
                    ZSql = ZSql + "Descripcion ,"
                    ZSql = ZSql + "Importe ,"
                    ZSql = ZSql + "Dto ,"
                    ZSql = ZSql + "Neto ,"
                    ZSql = ZSql + "Neto1 ,"
                    ZSql = ZSql + "Iva1 ,"
                    ZSql = ZSql + "Iva2 ,"
                    ZSql = ZSql + "Total ,"
                    ZSql = ZSql + "Item ,"
                    ZSql = ZSql + "Articulo ,"
                    ZSql = ZSql + "Cantidad ,"
                    ZSql = ZSql + "Precio ,"
                    ZSql = ZSql + "Imprepago ,"
                    ZSql = ZSql + "CondIva ,"
                    ZSql = ZSql + "Remito ,"
                    ZSql = ZSql + "Impre1 ,"
                    ZSql = ZSql + "Impre3 ,"
                    ZSql = ZSql + "Impre4 ,"
                    ZSql = ZSql + "Cae ,"
                    ZSql = ZSql + "VtoCae ,"
                    ZSql = ZSql + "ImpreBarra ,"
                    ZSql = ZSql + "ImpreBarraII ,"
                    ZSql = ZSql + "DescriII ,"
                    ZSql = ZSql + "CantiII ,"
                    ZSql = ZSql + "PrecioII ,"
                    ZSql = ZSql + "PorceIva ,"
                    ZSql = ZSql + "PordeDto )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + ZZClave + "',"
                    ZSql = ZSql + "'" + ZZLetra + "',"
                    ZSql = ZSql + "'" + ZZTipo + "',"
                    ZSql = ZSql + "'" + ZZPunto + "',"
                    ZSql = ZSql + "'" + ZZFactura + "',"
                    ZSql = ZSql + "'" + ZZRenglon + "',"
                    ZSql = ZSql + "'" + ZZfecha + "',"
                    ZSql = ZSql + "'" + ZZCliente + "',"
                    ZSql = ZSql + "'" + ZZNombre + "',"
                    ZSql = ZSql + "'" + ZZDireccion + "',"
                    ZSql = ZSql + "'" + ZZLocalidad + "',"
                    ZSql = ZSql + "'" + ZZPostal + "',"
                    ZSql = ZSql + "'" + ZZPartida + "',"
                    ZSql = ZSql + "'" + ZZCuit + "',"
                    ZSql = ZSql + "'" + ZZDescripcion + "',"
                    ZSql = ZSql + "'" + ZZImporte + "',"
                    ZSql = ZSql + "'" + ZZDto + "',"
                    ZSql = ZSql + "'" + ZZNeto + "',"
                    ZSql = ZSql + "'" + ZZNeto1 + "',"
                    ZSql = ZSql + "'" + ZZIva1 + "',"
                    ZSql = ZSql + "'" + ZZIva2 + "',"
                    ZSql = ZSql + "'" + ZZTotal + "',"
                    ZSql = ZSql + "'" + ZZItem + "',"
                    ZSql = ZSql + "'" + ZZArticulo + "',"
                    ZSql = ZSql + "'" + ZZCantidad + "',"
                    ZSql = ZSql + "'" + ZZPrecio + "',"
                    ZSql = ZSql + "'" + ZZImprepago + "',"
                    ZSql = ZSql + "'" + ZZImpreIva + "',"
                    ZSql = ZSql + "'" + ZZRemito + "',"
                    ZSql = ZSql + "'" + ZZImpre1 + "',"
                    ZSql = ZSql + "'" + ZZImpre3 + "',"
                    ZSql = ZSql + "'" + ZZImpre4 + "',"
                    ZSql = ZSql + "'" + Cae.Text + "',"
                    ZSql = ZSql + "'" + VtoCae.Text + "',"
                    ZSql = ZSql + "'" + ZZImpreBarra + "',"
                    ZSql = ZSql + "'" + ZZImpreBarraII + "',"
                    ZSql = ZSql + "'" + ZZDescriII + "',"
                    ZSql = ZSql + "'" + ZZCantiII + "',"
                    ZSql = ZSql + "'" + ZZPrecioII + "',"
                    ZSql = ZSql + "'" + ZZPorceIva + "',"
                    ZSql = ZSql + "'" + ZZPorceDto + "')"
                                            
                    spFactura = ZSql
                    Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)
                
                Next WRenglon
            
            End If
    
        End If
    
    Next a
    
    
    For a = ZZLugarFactura + 1 To 38

        ZZLugarFactura = ZZLugarFactura + 1

        ZZRenglon = Str$(ZZLugarFactura)
        Auxi1 = ZZRenglon
        Call Ceros(Auxi1, 2)
        ZZRenglon = Auxi1
        
        ZZClave = ZZLetra + ZZTipo + ZZPunto + ZZFactura + ZZRenglon
        
        ZZItem = ""
        ZZArticulo = ""
        ZZDescripcion = ""
        ZZCantidad = ""
        ZZPrecio = ""
        ZZImporte = ""
        
        ZZImpre1 = Remito.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi2, 8)
        ZZImpre3 = Auxi2
        ZZImpre4 = "FACTURA"
    
        ZZDescriII = ""
        ZZCantiII = ""
        ZZPrecioII = ""
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Factura ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Factura ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Cuit  ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Dto ,"
        ZSql = ZSql + "Neto ,"
        ZSql = ZSql + "Neto1 ,"
        ZSql = ZSql + "Iva1 ,"
        ZSql = ZSql + "Iva2 ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "Item ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Imprepago ,"
        ZSql = ZSql + "CondIva ,"
        ZSql = ZSql + "Remito ,"
        ZSql = ZSql + "Impre1 ,"
        ZSql = ZSql + "Impre3 ,"
        ZSql = ZSql + "Impre4 ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "VtoCae ,"
        ZSql = ZSql + "ImpreBarra ,"
        ZSql = ZSql + "ImpreBarraII ,"
        ZSql = ZSql + "DescriII ,"
        ZSql = ZSql + "CantiII ,"
        ZSql = ZSql + "PrecioII ,"
        ZSql = ZSql + "PorceIva ,"
        ZSql = ZSql + "PordeDto )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZFactura + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZNombre + "',"
        ZSql = ZSql + "'" + ZZDireccion + "',"
        ZSql = ZSql + "'" + ZZLocalidad + "',"
        ZSql = ZSql + "'" + ZZPostal + "',"
        ZSql = ZSql + "'" + ZZPartida + "',"
        ZSql = ZSql + "'" + ZZCuit + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZImporte + "',"
        ZSql = ZSql + "'" + ZZDto + "',"
        ZSql = ZSql + "'" + ZZNeto + "',"
        ZSql = ZSql + "'" + ZZNeto1 + "',"
        ZSql = ZSql + "'" + ZZIva1 + "',"
        ZSql = ZSql + "'" + ZZIva2 + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZItem + "',"
        ZSql = ZSql + "'" + ZZArticulo + "',"
        ZSql = ZSql + "'" + ZZCantidad + "',"
        ZSql = ZSql + "'" + ZZPrecio + "',"
        ZSql = ZSql + "'" + ZZImprepago + "',"
        ZSql = ZSql + "'" + ZZImpreIva + "',"
        ZSql = ZSql + "'" + ZZRemito + "',"
        ZSql = ZSql + "'" + ZZImpre1 + "',"
        ZSql = ZSql + "'" + ZZImpre3 + "',"
        ZSql = ZSql + "'" + ZZImpre4 + "',"
        ZSql = ZSql + "'" + Cae.Text + "',"
        ZSql = ZSql + "'" + VtoCae.Text + "',"
        ZSql = ZSql + "'" + ZZImpreBarra + "',"
        ZSql = ZSql + "'" + ZZImpreBarraII + "',"
        ZSql = ZSql + "'" + ZZDescriII + "',"
        ZSql = ZSql + "'" + ZZCantiII + "',"
        ZSql = ZSql + "'" + ZZPrecioII + "',"
        ZSql = ZSql + "'" + ZZPorceIva + "',"
        ZSql = ZSql + "'" + ZZPorceDto + "')"
                                
        spFactura = ZSql
        Set rstFactura = db.OpenRecordset(spFactura, dbOpenSnapshot, dbSQLPassThrough)

    Next a
    
    
    Listado.WindowTitle = "Impresion de Proforma"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Factura.Factura, Factura.Renglon, Factura.Fecha, Factura.Cliente, Factura.Nombre, Factura.Direccion, Factura.Localidad, Factura.Cuit, Factura.Descripcion, Factura.Neto, Factura.Dto, Factura.Neto1, Factura.Iva1, Factura.Iva2, Factura.Total, Factura.Imprepago, Factura.CondIva, Factura.Item, Factura.Articulo, Factura.Cantidad, Factura.Precio, Factura.PordeDto, Factura.Postal " _
            + "From " _
            + DSQ + ".dbo.Factura Factura " _
            + "Where " _
            + "Factura.Item >= 0 AND " _
            + "Factura.Item <= 99"
    
    Listado.Connect = Connect()
    
    Uno = "{Factura.Item} in 0 to 99"
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    If Val(Vendedor.Text) <> 1 And Val(Vendedor.Text) <> 7 Then
        Listado.CopiesToPrinter = 4
            Else
        Listado.CopiesToPrinter = 3
    End If
    Rem Listado.Destination = 0
    
    If Letra.Text = "A" Then
        Listado.ReportFileName = "ImpreFacturaElectronicaNuevoA.rpt"
            Else
        Listado.ReportFileName = "ImpreFacturaElectronicaNuevoB.rpt"
    End If
    
    Listado.Action = 1


End Sub

Private Sub Consulta_Click()

    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Pedidos"
    Opcion.AddItem "Condicion de Pago"
    Opcion.AddItem "Articulos"

    Opcion.Visible = True
     
 End Sub

Private Sub Impresion_Click()

    Rem Call Impresion_Factura_Reimpre
    
    T$ = "Emision de Facturas"
    M$ = "Desea reimprimir la factura electronica"
    Respuesta% = MsgBox(M$, 32 + 4, T$)
    If Respuesta% = 6 Then
        ZZPasaImpre = 1
        Call Impresion_FacturaFe
    End If
    
    T$ = "Emision de Facturas"
    M$ = "Desea reimprimir el remito"
    Respuesta% = MsgBox(M$, 32 + 4, T$)
    If Respuesta% = 6 Then
        ZZPasaImpre = 1
        Call Impresion_Remito
    End If

    WVector1.Col = 1
    WVector1.Row = 1
        
    Numero.SetFocus
    
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
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
                            IngresaItem = !CLIENTE + " " + !Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !CLIENTE
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 1
            Erase WPedido
            LugarPedido = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Cliente = " + "'" + CLIENTE + "'"
            ZSql = ZSql + " Order by Pedido.Numero"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                With rstPedido
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Saldo = rstPedido!Cantidad - rstPedido!facturado
                            If Saldo > 0 Then
                                Entra = "S"
                                For Ciclo = 1 To LugarPedido
                                    If Val(WPedido(Ciclo)) = rstPedido!Numero Then
                                        Entra = "N"
                                        Exit For
                                    End If
                                Next Ciclo
                                If Entra = "S" Then
                                    LugarPedido = LugarPedido + 1
                                    WPedido(LugarPedido) = Str$(rstPedido!Numero)
                                    WNumero = Str$(rstPedido!Numero)
                                    Call Ceros(WNumero, 8)
                                    IngresaItem = WNumero + " " + rstPedido!Fecha + " " + rstPedido!Observaciones
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstPedido!Numero
                                    WIndice.AddItem IngresaItem
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
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Order by CondPago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
            End If
            
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
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
                rstArticulo.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Calcula_Click()

    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    WTipo = "01"
    ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
        
    ZZExiste = ""
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ctacte"
    ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        rstCtaCte.Close
        ZZExiste = "S"
    End If

    WNeto = 0
    
    For a = 1 To 50
    
        WCantidad = Val(WVector1.TextMatrix(a, 3))
        WPrecio = Val(WVector1.TextMatrix(a, 4))
        
        If Letra.Text = "B" Then
            If TipoIva.ListIndex = 0 Then
                WWImpre = WPrecio * (1 + (ConfigIva1) / 100)
                    Else
                WWImpre = WPrecio * (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(WWImpre)
            WPrecio = WWImpre
        End If
        
        If Val(WCantidad) <> 0 And ZZExiste = "" Then
            Select Case Partida.Text
                Case "/"
                    WCantidad = WCantidad / 2
                Case "?"
                    WCantidad = WCantidad / 12
                Case Else
            End Select
        End If
        
        WNeto = WNeto + (WPrecio * WCantidad)
        
    Next a
    
    Call Calcula_Importe
    
End Sub

Private Sub CalculaReal_Click()

    WNeto = 0
    
    For a = 1 To 50
    
        WCantidad = Val(WVector1.TextMatrix(a, 3))
        WPrecio = Val(WVector1.TextMatrix(a, 4))
        
        If Letra.Text = "B" Then
            If TipoIva.ListIndex = 0 Then
                WWImpre = WPrecio * (1 + (ConfigIva1) / 100)
                    Else
                WWImpre = WPrecio * (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(WWImpre)
            WPrecio = WWImpre
        End If
        
        If Val(WCantidad) <> 0 Then
            Select Case Partida.Text
                Case "/"
                    WCantidad = WCantidad / 2
                Case "?"
                    WCantidad = WCantidad / 12
                Case Else
            End Select
        End If
        
        WNeto = WNeto + (WPrecio * WCantidad)
        
    Next a
    
    Call Calcula_Importe
    
End Sub


Private Sub Calcula_Importe()

    WImpoDto = 0
    WDescuento = Val(Descuento.Text)
    
    WDescuento = WDescuento
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    WIva1 = 0
    WIva2 = 0
    
    If Letra.Text = "A" Then
        Select Case Val(WCodIva)
            Case 2
                WIva1 = WNeto * ((ConfigIva1) / 100)
                WIva2 = WNeto * ((ConfigIva2) / 100)
                Call Redondeo(WIva1)
                Call Redondeo(WIva2)
            Case Else
                If TipoIva.ListIndex = 0 Then
                    WIva1 = WNeto * ((ConfigIva1) / 100)
                    Call Redondeo(WIva1)
                        Else
                    WIva1 = WNeto * ((ConfigIva2) / 100)
                    Call Redondeo(WIva1)
                End If
        End Select
    End If
    
    WWIva(1) = WIva1
    WWIva(2) = WIva2
    
    WTotal = WNeto + WIva1 + WIva2
    
    SubTotal.Caption = Str$(WNeto + WImpoDto)
    Dto.Caption = Str$(WImpoDto)
    Neto.Caption = Str$(WNeto)
    Iva1.Caption = Str$(WIva1)
    Iva2.Caption = Str$(WIva2)
    Total.Caption = Str$(WTotal)
    
    SubTotal.Caption = Pusing("###,###.##", SubTotal.Caption)
    Dto.Caption = Pusing("###,###.##", Dto.Caption)
    Neto.Caption = Pusing("###,###.##", Neto.Caption)
    Iva1.Caption = Pusing("###,###.##", Iva1.Caption)
    Iva2.Caption = Pusing("###,###.##", Iva2.Caption)
    Total.Caption = Pusing("###,###.##", Total.Caption)

End Sub

Private Sub cmdClose_Click()
    PrgFactura.Hide
    Unload Me
    Menu4.Show
End Sub

Private Sub Graba_Click()

    ZZLineas = 0
    Erase ZZPasaDatos
    
    For CicloII = 1 To 50
        
        ZZArticulo = WVector1.TextMatrix(CicloII, 1)
        ZZDesArticulo = WVector1.TextMatrix(CicloII, 2)
        ZZCantidad = WVector1.TextMatrix(CicloII, 3)
        ZZPrecio = WVector1.TextMatrix(CicloII, 4)
        ZZImporte = WVector1.TextMatrix(CicloII, 5)
        ZZStock = WVector1.TextMatrix(CicloII, 6)
        
        If Val(ZZCantidad) <> 0 Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Exibidores"
            ZSql = ZSql + " Where Exibidores.Codigo = " + "'" + ZZArticulo + "'"
            spExibidores = ZSql
            Set rstExibidores = db.OpenRecordset(spExibidores, dbOpenSnapshot, dbSQLPassThrough)
            If rstExibidores.RecordCount > 0 Then
            
                ZZLineas = ZZLineas + 1
                ZZPasaDatos(ZZLineas, 1) = ZZArticulo
                ZZPasaDatos(ZZLineas, 2) = ZZDesArticulo
                ZZPasaDatos(ZZLineas, 3) = ZZCantidad
                ZZPasaDatos(ZZLineas, 4) = ZZPrecio
                ZZPasaDatos(ZZLineas, 5) = ZZImporte
                ZZPasaDatos(ZZLineas, 6) = ZZStock
                ZZPasaDatos(ZZLineas, 7) = "1"
            
                rstExibidores.Close
                
                For WRenglon = 1 To 100
                
                    Auxi1 = WRenglon
                    Call Ceros(Auxi1, 2)
                    
                    WClave = Trim(ZZArticulo) + Auxi1
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Exibidores"
                    ZSql = ZSql + " Where Exibidores.Clave = " + "'" + WClave + "'"
                    spExibidores = ZSql
                    Set rstExibidores = db.OpenRecordset(spExibidores, dbOpenSnapshot, dbSQLPassThrough)
                    If rstExibidores.RecordCount > 0 Then
                        ZZLineas = ZZLineas + 1
                        ZZPasaDatos(ZZLineas, 1) = rstExibidores!Articulo
                        ZZPasaDatos(ZZLineas, 3) = Str$(Val(ZZCantidad) * rstExibidores!Cantidad)
                        ZZPasaDatos(ZZLineas, 7) = "2"
                        rstExibidores.Close
                    End If
                
                Next WRenglon
                    
                    Else
                
                ZZLineas = ZZLineas + 1
                ZZPasaDatos(ZZLineas, 1) = ZZArticulo
                ZZPasaDatos(ZZLineas, 2) = ZZDesArticulo
                ZZPasaDatos(ZZLineas, 3) = ZZCantidad
                ZZPasaDatos(ZZLineas, 4) = ZZPrecio
                ZZPasaDatos(ZZLineas, 5) = ZZImporte
                ZZPasaDatos(ZZLineas, 6) = ZZStock
                ZZPasaDatos(ZZLineas, 7) = "0"
                
            End If
            
        End If
    
    Next CicloII

    If ZZLineas > 38 Then
        M$ = "La factura a emitor supera los 38 renglones"
        a% = MsgBox(M$, 0, "Ingreso de Facturas")
        Exit Sub
    End If
    
    For CicloII = 1 To 50
    
        If Val(ZZPasaDatos(CicloII, 7)) = 2 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZZPasaDatos(CicloII, 1) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZPasaDatos(CicloII, 2) = rstArticulo!Descripcion
                ZZPasaDatos(CicloII, 4) = Str$(rstArticulo!Precio)
                ZZPasaDatos(CicloII, 5) = Str$(Val(ZZPasaDatos(CicloII, 3)) * Val(ZZPasaDatos(CicloII, 4)))
                ZZPasaDatos(CicloII, 6) = Str$(rstArticulo!Stock)
                rstArticulo.Close
            End If
            
        End If
        
    Next CicloII
    
    Call Calcula_Click
    
    WNeto = Val(Neto.Caption)
    WIva1 = Val(Iva1.Caption)
    WIva2 = Val(Iva2.Caption)
    WTotal = Val(Total.Caption)
    
    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
            
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
            
    WTipo = "01"
            
    ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ctacte"
    ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        rstCtaCte.Close
        M$ = "Factura ya emitida"
        a% = MsgBox(M$, 0, "Ingreso de Facturas")
        Exit Sub
    End If

    Call Calcula_Click
    
    WNeto = Val(Neto.Caption)
    WIva1 = Val(Iva1.Caption)
    WIva2 = Val(Iva2.Caption)
    WTotal = Val(Total.Caption)
    
    Pasa = "S"
    
    For a = 1 To 50
    
        Articulo = WVector1.TextMatrix(a, 1)
        ZDescripcion = WVector1.TextMatrix(a, 2)
        Cantidad = Val(WVector1.TextMatrix(a, 3))
        
        If Val(Cantidad) <> 0 Then
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZIva = rstArticulo!Iva
                rstArticulo.Close
            End If
            
            If ZZIva <> TipoIva.ListIndex Then
                M$ = "La condicion de Iva del Articulo " + ZDescripcion + " no coincide con el informado en la factura"
                a% = MsgBox(M$, 0, "Emision de Facturas")
                Exit Sub
            End If
            
            Select Case Partida.Text
                Case "/"
                    MiResultado = Val(Cantidad) Mod 2
                    If MiResultado <> 0 Then
                        M$ = "Las cantidades no son concordantes con el tipo de facturacion en el articulo " + ZDescripcion
                        a% = MsgBox(M$, 0, "Emision de Facturas")
                        Exit Sub
                    End If
                Case "?"
                    MiResultado = Val(Cantidad) Mod 12
                    If MiResultado <> 0 Then
                        M$ = "Las cantidades no son concordantes con el tipo de facturacion"
                        a% = MsgBox(M$, 0, "Emision de Facturas")
                        Exit Sub
                    End If
                Case Else
            End Select
            
        End If
        
                                        
    Next a
    
    ZAlta = "S"
    CLIENTE.Text = Trim(UCase(CLIENTE.Text))
    If CLIENTE.Text = "C-183" Then
        T$ = "Emision de Facturas"
        M$ = "Desea Actualizar el stock"
        Respuesta% = MsgBox(M$, 32 + 4, T$)
        If Respuesta% = 6 Then
            ZAlta = "S"
                Else
            ZAlta = "N"
        End If
    End If
        
    If Letra.Text = "B" Then
        If TipoIva.ListIndex = 0 Then
            WNeto = Val(Total.Caption) / (1 + ((ConfigIva1) / 100))
                Else
            WNeto = Val(Total.Caption) / (1 + ((ConfigIva2) / 100))
        End If
        Call Redondeo(WNeto)
        WIva1 = WTotal - WNeto
        Neto.Caption = Str$(WNeto)
        Iva1.Caption = Str$(WIva1)
    End If
        
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    If Trim(Cae.Text) = "" Then
        Call Calcula_Cae
        If Trim(Cae.Text) = "" Then
            Exit Sub
        End If
    End If
        
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
            
    WPunto = Punto.Text
    Call Ceros(WPunto, 4)
    
    ZZTipo = "01"
    ZZImpre = "FC"
            
    ZZPunto = WPunto
    ZZLetra = Letra.Text
    ZZNumero = Auxi
    ZZRenglon = "01"
    ZZCliente = CLIENTE.Text
    ZZfecha = Fecha.Text
    ZZEstado = "0"
    ZZVencimiento = Fecha.Text
    ZZTotal = Str$(WTotal)
    ZZSaldo = Str$(WTotal)
    If Letra.Text = "B" Then
        If TipoIva.ListIndex = 0 Then
            WNeto = WTotal / (1 + ((ConfigIva1) / 100))
                Else
            WNeto = WTotal / (1 + ((ConfigIva2) / 100))
        End If
        Call Redondeo(WNeto)
        WIva1 = WTotal - WNeto
        WIva2 = 0
        ZZNeto = Str$(WNeto)
        ZZIva1 = Str$(WIva1)
        ZZIva2 = Str$(WIva2)
            Else
        ZZNeto = Str$(WNeto)
        ZZIva1 = Str$(WIva1)
        ZZIva2 = Str$(WIva2)
    End If
    
    Select Case Partida.Text
        Case "/"
            ZZTotalUs = Str$(WTotal + WNeto)
            ZZSaldoUs = Str$(WTotal + WNeto)
        Case "?"
            ZZTotalUs = Str$(WTotal + (WNeto * 11))
            ZZSaldoUs = Str$(WTotal + (WNeto * 11))
        Case Else
            ZZTotalUs = Str$(WTotal)
            ZZSaldoUs = Str$(WTotal)
    End Select
    
    ZZExento = "0"
    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZZOrdVencimiento = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZZPedido = Pedido.Text
    ZZRemito = ""
    ZZOrden = OCompra.Text
    ZZProvincia = WProvincia
    ZZVendedor = Vendedor.Text
    ZZCosto = "0"
    ZZImporte1 = "0"
    ZZImporte2 = "0"
    ZZImporte3 = "0"
    ZZImporte4 = "0"
    ZZImporte5 = "0"
    ZZImporte6 = "0"
    ZZImporte7 = "0"
    ZZTipoventa = "0"
    ZZProyecto = ""
    ZZParidad = "0"
    ZZRemito1 = ""
    ZZRemito2 = ""
    ZZBusqueda = ZZLetra + WPunto + Auxi
    
    ZZDescuento = Descuento.Text
    ZZPago = Pago.Text
    ZZPartida = Partida.Text
    ZZExpreso = Expreso.Text
    ZZTipoIva = Str$(TipoIva.ListIndex)
    ZZComision = Str$(Comision.Value)
    ZZRemito = Remito.Text
    
    ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
    
    ZZLinea = ""
    
    ZZNetoTotal = ZZNeto
    If ZZPartida = "/" Then
        ZZNetoTotal = Str$(WNeto * 2)
    End If
    If ZZPartida = "?" Then
        ZZNetoTotal = Str$(WNeto * 12)
    End If
    
    ZZCae = Cae.Text
    ZZVtoCae = VtoCae.Text
    
    ZSql = ""
    ZSql = ZSql + "INSERT INTO CtaCte ("
    ZSql = ZSql + "Clave ,"
    ZSql = ZSql + "Letra ,"
    ZSql = ZSql + "Tipo ,"
    ZSql = ZSql + "Punto ,"
    ZSql = ZSql + "Numero ,"
    ZSql = ZSql + "Renglon ,"
    ZSql = ZSql + "Cliente ,"
    ZSql = ZSql + "fecha ,"
    ZSql = ZSql + "Estado ,"
    ZSql = ZSql + "Vencimiento ,"
    ZSql = ZSql + "Total ,"
    ZSql = ZSql + "Saldo ,"
    ZSql = ZSql + "OrdFecha  ,"
    ZSql = ZSql + "OrdVencimiento ,"
    ZSql = ZSql + "Impre ,"
    ZSql = ZSql + "Neto ,"
    ZSql = ZSql + "NetoTotal ,"
    ZSql = ZSql + "Iva1 ,"
    ZSql = ZSql + "Iva2 ,"
    ZSql = ZSql + "Exento ,"
    ZSql = ZSql + "Pedido ,"
    ZSql = ZSql + "Remito ,"
    ZSql = ZSql + "Orden ,"
    ZSql = ZSql + "Provincia ,"
    ZSql = ZSql + "Vendedor ,"
    ZSql = ZSql + "Costo ,"
    ZSql = ZSql + "Importe1 ,"
    ZSql = ZSql + "Importe2 ,"
    ZSql = ZSql + "Importe3 ,"
    ZSql = ZSql + "Importe4 ,"
    ZSql = ZSql + "Importe5 ,"
    ZSql = ZSql + "Importe6 ,"
    ZSql = ZSql + "Importe7 ,"
    ZSql = ZSql + "Tipoventa ,"
    ZSql = ZSql + "Proyecto ,"
    ZSql = ZSql + "Paridad ,"
    ZSql = ZSql + "TotalUs ,"
    ZSql = ZSql + "SaldoUs ,"
    ZSql = ZSql + "Remito1 ,"
    ZSql = ZSql + "Remito2 ,"
    ZSql = ZSql + "Descuento ,"
    ZSql = ZSql + "Partida ,"
    ZSql = ZSql + "Pago ,"
    ZSql = ZSql + "Linea ,"
    ZSql = ZSql + "Expreso ,"
    ZSql = ZSql + "TipoIva ,"
    ZSql = ZSql + "Comision ,"
    ZSql = ZSql + "NroRemito ,"
    ZSql = ZSql + "Cae ,"
    ZSql = ZSql + "VtoCae ,"
    ZSql = ZSql + "Busqueda )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + ZZClave + "',"
    ZSql = ZSql + "'" + ZZLetra + "',"
    ZSql = ZSql + "'" + ZZTipo + "',"
    ZSql = ZSql + "'" + ZZPunto + "',"
    ZSql = ZSql + "'" + ZZNumero + "',"
    ZSql = ZSql + "'" + ZZRenglon + "',"
    ZSql = ZSql + "'" + ZZCliente + "',"
    ZSql = ZSql + "'" + ZZfecha + "',"
    ZSql = ZSql + "'" + ZZEstado + "',"
    ZSql = ZSql + "'" + ZZVencimiento + "',"
    ZSql = ZSql + "'" + ZZTotal + "',"
    ZSql = ZSql + "'" + ZZSaldo + "',"
    ZSql = ZSql + "'" + ZZOrdFecha + "',"
    ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
    ZSql = ZSql + "'" + ZZImpre + "',"
    ZSql = ZSql + "'" + ZZNeto + "',"
    ZSql = ZSql + "'" + ZZNetoTotal + "',"
    ZSql = ZSql + "'" + ZZIva1 + "',"
    ZSql = ZSql + "'" + ZZIva2 + "',"
    ZSql = ZSql + "'" + ZZExento + "',"
    ZSql = ZSql + "'" + ZZPedido + "',"
    ZSql = ZSql + "'" + ZZRemito + "',"
    ZSql = ZSql + "'" + ZZOrden + "',"
    ZSql = ZSql + "'" + ZZProvincia + "',"
    ZSql = ZSql + "'" + ZZVendedor + "',"
    ZSql = ZSql + "'" + ZZCosto + "',"
    ZSql = ZSql + "'" + ZZImporte1 + "',"
    ZSql = ZSql + "'" + ZZImporte2 + "',"
    ZSql = ZSql + "'" + ZZImporte3 + "',"
    ZSql = ZSql + "'" + ZZImporte4 + "',"
    ZSql = ZSql + "'" + ZZImporte5 + "',"
    ZSql = ZSql + "'" + ZZImporte6 + "',"
    ZSql = ZSql + "'" + ZZImporte7 + "',"
    ZSql = ZSql + "'" + ZZTipoventa + "',"
    ZSql = ZSql + "'" + ZZProyecto + "',"
    ZSql = ZSql + "'" + ZZParidad + "',"
    ZSql = ZSql + "'" + ZZTotalUs + "',"
    ZSql = ZSql + "'" + ZZSaldoUs + "',"
    ZSql = ZSql + "'" + ZZRemito1 + "',"
    ZSql = ZSql + "'" + ZZRemito2 + "',"
    ZSql = ZSql + "'" + ZZDescuento + "',"
    ZSql = ZSql + "'" + ZZPartida + "',"
    ZSql = ZSql + "'" + ZZPago + "',"
    ZSql = ZSql + "'" + ZZLinea + "',"
    ZSql = ZSql + "'" + ZZExpreso + "',"
    ZSql = ZSql + "'" + ZZTipoIva + "',"
    ZSql = ZSql + "'" + ZZComision + "',"
    ZSql = ZSql + "'" + ZZRemito + "',"
    ZSql = ZSql + "'" + ZZCae + "',"
    ZSql = ZSql + "'" + ZZVtoCae + "',"
    ZSql = ZSql + "'" + ZZBusqueda + "')"
                            
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    
    
    
    
    
    Renglon = 0
    WRenglon = 0
    ZZSumaComision = 0
        
    For a = 1 To 50
        
        WRenglon = WRenglon + 1
        
        Articulo = UCase(ZZPasaDatos(WRenglon, 1))
        DesArticulo = ZZPasaDatos(WRenglon, 2)
        Cantidad = Val(ZZPasaDatos(WRenglon, 3))
        Precio = Val(ZZPasaDatos(WRenglon, 4))
        TipoII = Val(ZZPasaDatos(WRenglon, 7))
        PrecioSalva = Val(ZZPasaDatos(WRenglon, 4))
            
        If Cantidad <> 0 Then
                    
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Numero.Text)
            Call Ceros(Auxi1, 8)
            
            If ZAlta = "N" Then
                ZNoCantidad = Cantidad
                Cantidad = 0
            End If
            
            If TipoII = 1 Then
                Precio = 0
            End If
            
            ZZCosto1 = "0"
            ZZComision = "0"
            ZZTipoComision = Str$(Comision.Value)
            
            If TipoII <> 1 Then
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZCosto1 = Str$(rstArticulo!Costo)
                    ZZComision = Str$(rstArticulo!Comision)
                    rstArticulo.Close
                End If
            End If
            
            XXPrecio = Precio
            If Val(Descuento.Text) <> 0 Then
                XXDto = XXPrecio * (Val(Descuento.Text) / 100)
                Call Redondeo(XXDto)
                XXPrecio = XXPrecio - XXDto
                Call Redondeo(XXPrecio)
            End If
            
            If Comision.Value = 0 Then
                XXComision = ZZComision
                    Else
                XXComision = ZZComision * 0.5
            End If
            
            XXImpoComision = XXPrecio * (XXComision / 100)
            Call Redondeo(XXImpoComision)
            
            XXPrecio = XXPrecio - XXImpoComision
            Call Redondeo(XXPrecio)

            ZZTipo = "01"
            ZZNumero = Numero.Text
            ZZRenglon = Renglon
            ZZArticulo = Articulo
            ZZDescripcion = DesArticulo
            ZZCantidad = Str$(Cantidad)
            ZZCantidadII = Str$(Cantidad)
            ZZPrecio = Str$(Precio)
            ZZPrecioSalva = Str$(PrecioSalva)
            ZZPrecioUs = Str$(XXPrecio)
            ZZImporte = Str$(Precio * Cantidad)
            ZZImporteUs = Str$(XXPrecio * Cantidad)
            ZZCliente = CLIENTE.Text
            ZZParidad = "0"
            ZZVendedor = "0"
            ZZRubro = "0"
            ZZLinea = "0"
            ZZCosto2 = "0"
            ZZCoeficiente = "0"
            ZZPedido = "0"
            ZZfecha = Fecha.Text
            ZZImporte1 = "0"
            ZZImporte2 = "0"
            ZZImporte3 = "0"
            ZZImporte4 = "0"
            ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZZWArticulo = ""
            ZZRemito = ""
            ZZTipoII = Str$(TipoII)
            
            ZZZLetra = Letra.Text
            
            ZZZPunto = Punto.Text
            Call Ceros(ZZZPunto, 1)
            
            ZZZNumeroFac = Numero.Text
            Call Ceros(ZZZNumeroFac, 6)
            
            ZZClave = "01" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi
            
            ZZWDate = Date$
            ZZClaveCtacte = "01" + Auxi1 + "01"
            
            ZZImprefactura = "FACTURA"
            ZZNroFactura = Auxi1
            ZZTalle = Talle
            ZZColor = XXColor
            ZZCuenta = WCuenta
            ZZDescuento = Descuento.Text
            ZZPartida = Partida.Text
            
            ZZCantidadII = ZZCantidad
            If ZZPartida = "/" Then
                ZZCantidadII = Str$(Val(ZZCantidad) / 2)
            End If
            If ZZPartida = "?" Then
                ZZCantidadII = Str$(Val(ZZCantidad) / 12)
            End If
            
            If ZAlta = "N" Then
                ZZCantidadII = Str$(ZNoCantidad)
            End If
            
            ZZPrecioII = Str$(XXPrecio)
            
            ZZSumaComision = ZZSumaComision + (XXImpoComision * Val(ZZCantidadII))
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Estadistica ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "CantidadII ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "PrecioII ,"
            ZSql = ZSql + "PrecioSalva ,"
            ZSql = ZSql + "PrecioUs ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "ImporteUs ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Paridad ,"
            ZSql = ZSql + "Vendedor ,"
            ZSql = ZSql + "Rubro ,"
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Costo1 ,"
            ZSql = ZSql + "Costo2 ,"
            ZSql = ZSql + "Comision ,"
            ZSql = ZSql + "TipoComision ,"
            ZSql = ZSql + "Coeficiente ,"
            ZSql = ZSql + "Pedido ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Importe1 ,"
            ZSql = ZSql + "Importe2 ,"
            ZSql = ZSql + "Importe3 ,"
            ZSql = ZSql + "Importe4 ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "WArticulo ,"
            ZSql = ZSql + "Remito ,"
            ZSql = ZSql + "WDate ,"
            ZSql = ZSql + "Marca ,"
            ZSql = ZSql + "TipoII ,"
            ZSql = ZSql + "ClaveCtacte ,"
            ZSql = ZSql + "Imprefactura ,"
            ZSql = ZSql + "NroFactura ,"
            ZSql = ZSql + "Descuento ,"
            ZSql = ZSql + "Partida )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZPunto + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZDescripcion + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZCantidadII + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZPrecioII + "',"
            ZSql = ZSql + "'" + ZZPrecioSalva + "',"
            ZSql = ZSql + "'" + ZZPrecioUs + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZImporteUs + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZParidad + "',"
            ZSql = ZSql + "'" + ZZVendedor + "',"
            ZSql = ZSql + "'" + ZZRubro + "',"
            ZSql = ZSql + "'" + ZZLinea + "',"
            ZSql = ZSql + "'" + ZZCosto1 + "',"
            ZSql = ZSql + "'" + ZZCosto2 + "',"
            ZSql = ZSql + "'" + ZZComision + "',"
            ZSql = ZSql + "'" + ZZTipoComision + "',"
            ZSql = ZSql + "'" + ZZCoeficiente + "',"
            ZSql = ZSql + "'" + ZZPedido + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZImporte1 + "',"
            ZSql = ZSql + "'" + ZZImporte2 + "',"
            ZSql = ZSql + "'" + ZZImporte3 + "',"
            ZSql = ZSql + "'" + ZZImporte4 + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + ZZWArticulo + "',"
            ZSql = ZSql + "'" + ZZRemito + "',"
            ZSql = ZSql + "'" + ZZWDate + "',"
            ZSql = ZSql + "'" + ZZMarca + "',"
            ZSql = ZSql + "'" + ZZTipoII + "',"
            ZSql = ZSql + "'" + ZZClaveCtacte + "',"
            ZSql = ZSql + "'" + ZZImprefactura + "',"
            ZSql = ZSql + "'" + ZZNroFactura + "',"
            ZSql = ZSql + "'" + ZZDescuento + "',"
            ZSql = ZSql + "'" + ZZPartida + "')"
                            
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                        
        End If
                                        
    Next a
    
    Auxi = Numero.Text
    Call Ceros(Auxi, 8)
    
    ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Costo = " + "'" + Str$(ZZSumaComision) + "'"
    ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    For WRenglon = 1 To 50
            
        Articulo = ZZPasaDatos(WRenglon, 1)
        Cantidad = Val(ZZPasaDatos(WRenglon, 3))
            
        Rem If Partida.Text = "/" Then
        Rem     Cantidad = Cantidad * 2
        Rem End If
        Rem If Partida.Text = "?" Then
        Rem     Cantidad = Cantidad * 12
        Rem End If
            
        If Cantidad <> 0 Then
        
            If ZAlta = "S" Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " FechaUltimaSalida = " + "'" + Fecha.Text + "',"
                ZSql = ZSql + " Salidas = Salidas + " + "'" + Str$(Cantidad) + "',"
                ZSql = ZSql + " Stock = Stock - " + "'" + Str$(Cantidad) + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            If Val(Pedido.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Pedido SET "
                ZSql = ZSql + " Facturado = Facturado + " + "'" + Str$(Cantidad) + "',"
                ZSql = ZSql + " Marca = " + "'" + "" + "'"
                ZSql = ZSql + " Where Numero = " + "'" + Pedido.Text + "'"
                ZSql = ZSql + " and Articulo = " + "'" + Articulo + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
        End If
            
    Next WRenglon
    
    If Val(Pedido.Text) = 0 Then
        
        Pedido.Text = "1"
        ZSql = ""
        ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
        ZSql = ZSql + " FROM Pedido"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            rstPedido.MoveLast
            ZUltimo = IIf(IsNull(rstPedido!NumeroMayor), "0", rstPedido!NumeroMayor)
            Pedido.Text = ZUltimo + 1
            rstPedido.Close
        End If
        
        Renglon = 0
    
        For WRenglon = 1 To 50
            
            WVector1.Row = WRenglon
            
            WVector1.Col = 1
            Articulo = WVector1.Text
                    
            WVector1.Col = 3
            Cantidad = Val(WVector1.Text)
            
            Rem If Partida.Text = "/" Then
            Rem     Cantidad = Cantidad * 2
            Rem End If
            Rem If Partida.Text = "?" Then
            Rem     Cantidad = Cantidad * 12
            Rem End If
            
            If Cantidad <> 0 Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Pedido.Text)
                Call Ceros(Auxi1, 8)
                    
                ZZNumero = Pedido.Text
                ZZRenglon = Str$(Renglon)
                ZZArticulo = Articulo
                ZZCantidad = Str$(Cantidad)
                ZZPrecio = ""
                ZZCliente = CLIENTE.Text
                ZZImporte = ""
                ZZfecha = Fecha.Text
                ZZImporte1 = "0"
                ZZImporte2 = "0"
                ZZImporte3 = "0"
                ZZImporte4 = "0"
                ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                ZZObservaciones = "Pedido Automatico"
                ZZFecEntrega = Fecha.Text
                ZZOrdFecEntrega = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                ZZFacturado = Str$(Cantidad)
                ZZCotiza = "0"
                ZZMarca = ""
            
                ZZPago = Pago.Text
                ZZPartida = Partida.Text
                ZZDescuento = Descuento.Text
                ZZAjuste = ""
                ZZTalle1 = ""
                ZZTalle2 = ""
                ZZTalle3 = ""
                ZZTalle4 = ""
                ZZTalle5 = ""
                ZZTalle6 = ""
                ZZTalle7 = ""
                ZZTalle8 = ""
                ZZTalle9 = ""
                ZZTalle10 = ""
            
                ZZClave = Auxi1 + Auxi
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Pedido ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Articulo ,"
                ZSql = ZSql + "Cantidad ,"
                ZSql = ZSql + "Precio ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Importe1 ,"
                ZSql = ZSql + "Importe2 ,"
                ZSql = ZSql + "Importe3 ,"
                ZSql = ZSql + "Importe4 ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Descuento,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "FecEntrega  ,"
                ZSql = ZSql + "OrdFecEntrega ,"
                ZSql = ZSql + "Facturado ,"
                ZSql = ZSql + "Ajuste ,"
                ZSql = ZSql + "Pago ,"
                ZSql = ZSql + "Partida ,"
                ZSql = ZSql + "Marca ,"
                ZSql = ZSql + "Talle1 ,"
                ZSql = ZSql + "Talle2 ,"
                ZSql = ZSql + "Talle3 ,"
                ZSql = ZSql + "Talle4 ,"
                ZSql = ZSql + "Talle5 ,"
                ZSql = ZSql + "Talle6 ,"
                ZSql = ZSql + "Talle7 ,"
                ZSql = ZSql + "Talle8 ,"
                ZSql = ZSql + "Talle9 ,"
                ZSql = ZSql + "Talle10 ,"
                ZSql = ZSql + "Cotiza )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZZClave + "',"
                ZSql = ZSql + "'" + ZZNumero + "',"
                ZSql = ZSql + "'" + ZZRenglon + "',"
                ZSql = ZSql + "'" + ZZArticulo + "',"
                ZSql = ZSql + "'" + ZZCantidad + "',"
                ZSql = ZSql + "'" + ZZPrecio + "',"
                ZSql = ZSql + "'" + ZZImporte + "',"
                ZSql = ZSql + "'" + ZZCliente + "',"
                ZSql = ZSql + "'" + ZZfecha + "',"
                ZSql = ZSql + "'" + ZZImporte1 + "',"
                ZSql = ZSql + "'" + ZZImporte2 + "',"
                ZSql = ZSql + "'" + ZZImporte3 + "',"
                ZSql = ZSql + "'" + ZZImporte4 + "',"
                ZSql = ZSql + "'" + ZZOrdFecha + "',"
                ZSql = ZSql + "'" + ZZDescuento + "',"
                ZSql = ZSql + "'" + ZZObservaciones + "',"
                ZSql = ZSql + "'" + ZZFecEntrega + "',"
                ZSql = ZSql + "'" + ZZOrdFecEntrega + "',"
                ZSql = ZSql + "'" + ZZFacturado + "',"
                ZSql = ZSql + "'" + ZZAjuste + "',"
                ZSql = ZSql + "'" + ZZPago + "',"
                ZSql = ZSql + "'" + ZZPartida + "',"
                ZSql = ZSql + "'" + ZZMarca + "',"
                ZSql = ZSql + "'" + ZZTalle1 + "',"
                ZSql = ZSql + "'" + ZZTalle2 + "',"
                ZSql = ZSql + "'" + ZZTalle3 + "',"
                ZSql = ZSql + "'" + ZZTalle4 + "',"
                ZSql = ZSql + "'" + ZZTalle5 + "',"
                ZSql = ZSql + "'" + ZZTalle6 + "',"
                ZSql = ZSql + "'" + ZZTalle7 + "',"
                ZSql = ZSql + "'" + ZZTalle8 + "',"
                ZSql = ZSql + "'" + ZZTalle9 + "',"
                ZSql = ZSql + "'" + ZZTalle10 + "',"
                ZSql = ZSql + "'" + ZZCotiza + "')"
            
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
        Next WRenglon
        
    End If
    
    
    WOrdUtimaCompra = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    ZSql = ""
    ZSql = ZSql + "UPDATE Cliente SET "
    ZSql = ZSql + " UltimaCompra = " + "'" + Fecha.Text + "',"
    ZSql = ZSql + " OrdUltimaCompra = " + "'" + WOrdUltimaCompra + "'"
    ZSql = ZSql + " Where Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
    ZZPasaImpre = 0
    
    Call Impresion_FacturaFe
    Call Impresion_Remito

    Rem T$ = "Emision de Facturas"
    Rem m$ = "Desea Imprimir la Factura"
    Rem Respuesta% = MsgBox(m$, 32 + 4, T$)
    Rem If Respuesta% = 6 Then
    Rem     Call WImpresion
    Rem End If
        
    Call Limpia_Click
        
    CLIENTE.SetFocus
        
End Sub

Private Sub CmdDelete_Click()

    T$ = "Baja de Comprobantes"
    M$ = "Desea Borrar el Comprobante "
    Respuesta% = MsgBox(M$, 32 + 4, T$)
    If Respuesta% = 6 Then

        T$ = "Baja de Comprobantes"
        M$ = "Esta seguro que Desea Borrar el Comprobante "
        Respuesta% = MsgBox(M$, 32 + 4, T$)
        If Respuesta% = 6 Then
        
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
                
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            WTipo = "01"
                
            ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
               
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                ZSaldo = rstCtaCte!Saldo
                ZTotal = rstCtaCte!Total
                rstCtaCte.Close
                If ZSaldo <> ZTotal Then
                    M$ = "El comprobante se encuentra total o parcialmente cancelado"
                    a% = MsgBox(M$, 0, "Eliminacion de Comprobantes")
                    Exit Sub
                End If
            End If
        
            ZBaja = "N"
            If Val(Pedido.Text) <> 0 Then
                T$ = "Baja de Comprobantes"
                M$ = "Desea Restaurar el saldo del pedido"
                Respuesta% = MsgBox(M$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    ZBaja = "S"
                End If
            End If
        
            Erase WVector
        
            For WRenglon = 1 To 50
            
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
                WClave = "01" + Auxi + Auxi1
                    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                
                    Articulo = rstEstadistica!Articulo
                    Cantidad = rstEstadistica!Cantidad
                    CantidadII = rstEstadistica!CantidadII
                        
                    rstEstadistica.Close
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " Salidas = Salidas - " + "'" + Str$(Cantidad) + "',"
                    ZSql = ZSql + " Stock = Stock + " + "'" + Str$(Cantidad) + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + Articulo + "'"
                    spArticulo = ZSql
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
                    If ZBaja = "S" Then
                    
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Pedido SET "
                        ZSql = ZSql + " Facturado = Facturado - " + "'" + Str$(Cantidad) + "'"
                        ZSql = ZSql + " Where Numero = " + "'" + Pedido.Text + "'"
                        ZSql = ZSql + " and Articulo = " + "'" + Articulo + "'"
                        spPedido = ZSql
                        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                End If
                
            Next WRenglon
            
            ZSql = ""
            ZSql = ZSql + "DELETE Estadistica"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Numero.Text + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            
            ZSql = ""
            ZSql = ZSql + "DELETE CtaCte"
            ZSql = ZSql + " Where Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Auxi + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            Call Limpia_Click
            
            CLIENTE.SetFocus
            
        End If
    
    End If

End Sub

Private Sub Limpia_Click()

    Call Limpia_Vector

    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    CLIENTE.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Descuento.Text = ""
    Pedido.Text = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Vendedor.Text = ""
    Expreso.Text = ""
    DesVendedor.Caption = ""
    DesExpreso.Caption = ""
    Comision.Value = 0
    Partida.Text = ""
    Remito.Text = ""
    OCompra.Text = ""
    Cae.Text = ""
    VtoCae.Text = ""
    
    Renglon = 0
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
    TipoIva.ListIndex = 0

    Graba.Enabled = True
    CLIENTE.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            CLIENTE.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Pedido.Text = WIndice.List(Indice)
            Call Pedido_KeyPress(13)
            
        Case 2
            Indice = Pantalla.ListIndex
            Pago.Text = WIndice.List(Indice)
            Call Pago_Keypress(13)
            
        Case 3
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Indice = Pantalla.ListIndex
            ClaveVen$ = WIndice.List(Indice)
            
            ZPasa = "S"
            If Val(Pedido.Text) <> 0 Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Pedido"
                ZSql = ZSql + " Where Numero = " + "'" + Pedido.Text + "'"
                ZSql = ZSql + " and Articulo = " + "'" + ClaveVen$ + "'"
                spPedido = ZSql
                Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                If rstPedido.RecordCount > 0 Then
                    rstPedido.Close
                        Else
                    M$ = "El articulo no esta en el pedido"
                    a% = MsgBox(M$, 0, "Carga de Articulos")
                    ZPasa = "N"
                End If
                
            End If
            
            If ZPasa = "S" Then
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ClaveVen$ + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 1
                    WVector1.Text = rstArticulo!Codigo
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    WVector1.Col = 4
                    WVector1.Text = Str$(rstArticulo!Precio)
                    WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    WVector1.Col = 6
                    WVector1.Text = Str$(rstArticulo!Stock)
                    WVector1.Col = 3
                    rstArticulo.Close
                    Call StartEdit
                End If
                
                    Else
                    
                WVector1.Col = 1
                WVector1.Text = ""
                WVector1.Col = 2
                WVector1.Text = ""
                WVector1.Col = 4
                WVector1.Text = ""
                    
                WVector1.Col = 1
                Call StartEdit
            
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "1Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "C.Final"
    Iva(4) = "Exento"
    Iva(5) = "MOnotributo"
    Iva(6) = "Exterior"
    
    Mes(1) = "Enero"
    Mes(2) = "Febrero"
    Mes(3) = "Marzo"
    Mes(4) = "Abril"
    Mes(5) = "Mayo"
    Mes(6) = "Junio"
    Mes(7) = "Julio"
    Mes(8) = "Agosto"
    Mes(9) = "Septiembre"
    Mes(10) = "Octubre"
    Mes(11) = "Noviembre"
    Mes(12) = "Diciembre"
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"
    
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    CLIENTE.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Descuento.Text = ""
    Pedido.Text = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Vendedor.Text = ""
    Expreso.Text = ""
    DesVendedor.Caption = ""
    DesExpreso.Caption = ""
    Comision.Value = 0
    Partida.Text = ""
    Remito.Text = ""
    OCompra.Text = ""
    Cae.Text = ""
    VtoCae.Text = ""
    
    SubTotal.Caption = ""
    Neto.Caption = ""
    Iva1.Caption = ""
    Iva2.Caption = ""
    Dto.Caption = ""
    Total.Caption = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Configuracion"
    ZSql = ZSql + " Where Configuracion.Clave = 1"
    spConfiguracion = ZSql
    Set rstConfiguracion = db.OpenRecordset(spConfiguracion, dbOpenSnapshot, dbSQLPassThrough)
    If rstConfiguracion.RecordCount > 0 Then
        ConfigIva1 = rstConfiguracion!Iva1
        ConfigIva2 = rstConfiguracion!Iva2
        ConfigPercepcion = rstConfiguracion!Percepcion
        Rem ConfigPunto = rstConfiguracion!Punto
        ConfigPunto = 3
        CantiFac = rstConfiguracion!CantiFac
        CantiRem = rstConfiguracion!CantiRem
        CantiArti = rstConfiguracion!CantiArti
        rstConfiguracion.Close
    End If
    
    TipoIva.Clear
    
    TipoIva.AddItem "21 %"
    TipoIva.AddItem "10.5 %"
    
    TipoIva.ListIndex = 0
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 50
        
        If Val(Punto.Text) = 1 Then
    
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
                
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "01" + Auxi + Auxi1
            
                Else
            
            ZZZLetra = Letra.Text
            
            ZZZPunto = Punto.Text
            Call Ceros(ZZZPunto, 1)
            
            ZZZNumeroFac = Numero.Text
            Call Ceros(ZZZNumeroFac, 6)
                
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            
            WClave = "01" + ZZZLetra + ZZZPunto + ZZZNumeroFac + Auxi1
                
        End If
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Estadistica"
        ZSql = ZSql + " Where Estadistica.Clave = " + "'" + WClave + "'"
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
        
            ZZTipoII = IIf(IsNull(rstEstadistica!TipoII), "o", rstEstadistica!TipoII)
            If ZZTipoII = 0 Or ZZTipoII = 1 Then
            
                Canti = rstEstadistica!Cantidad
                
                Renglon = Renglon + 1
                        
                WVector1.Row = Renglon
                        
                WVector1.Col = 1
                WVector1.Text = rstEstadistica!Articulo
                Auxi1 = rstEstadistica!Articulo
                    
                WVector1.Col = 2
                WVector1.Text = IIf(IsNull(rstEstadistica!Descripcion), "", rstEstadistica!Descripcion)
                    
                WVector1.Col = 3
                WVector1.Text = Pusing("###,###", Str$(rstEstadistica!CantidadII))
                    
                If ZZTipoII = 0 Then
                
                    WVector1.Col = 4
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Precio))
                    
                    WVector1.Col = 5
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!Precio * rstEstadistica!CantidadII))
                    
                        Else
                
                    WVector1.Col = 4
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!PrecioSalva))
                    
                    WVector1.Col = 5
                    WVector1.Text = Pusing("###,###.##", Str$(rstEstadistica!PrecioSalva * rstEstadistica!CantidadII))
                
                End If
                
                rstEstadistica.Close
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                        
            End If
                
        End If
    
    Next WRenglon

    Call Calcula_Click
    
    Graba.Enabled = True

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Trim(CLIENTE.Text) <> "" Then
            Auxi = UCase(Left$(CLIENTE.Text, 1))
            Auxi1 = Mid$(CLIENTE.Text, 2, 5)
            Call Ceros(Auxi1, 3)
            CLIENTE.Text = Auxi + "-" + Auxi1
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = Trim(rstCliente!Razon)
            Descuento.Text = Str$(rstCliente!Descuento)
            Descuento.Text = Pusing("###,###.##", Descuento.Text)
            Vendedor.Text = rstCliente!Vendedor
            Pago.Text = rstCliente!Condicion
            Expreso.Text = rstCliente!Expreso
            WProvincia = rstCliente!Provincia
            WCodIva = rstCliente!Iva
            WRazon = Trim(rstCliente!Razon)
            WDireccion = Trim(rstCliente!Direccion)
            WLocalidad = Trim(rstCliente!Localidad)
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            Select Case Val(WCodIva)
                Case 1, 2
                    Letra.Text = "A"
                Case Else
                    Letra.Text = "B"
            End Select
            ZMarca = IIf(IsNull(rstCliente!Marca), "0", rstCliente!Marca)
            
            Rem If Letra.Text = "B" Then
            Rem     m$ = "COLOQUE EL FORMULARIO B"
            Rem     a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Rem End If
            
            rstCliente.Close
                
            
            Rem ZSql = ""
            Rem ZSql = ZSql + "Select *"
            Rem ZSql = ZSql + " FROM ClienteAdicional"
            Rem ZSql = ZSql + " Where ClienteAdicional.Cliente = " + "'" + Cliente.Text + "'"
            Rem ZSql = ZSql + " and ClienteAdicional.Linea = " + "'" + Linea.Text + "'"
            Rem spClienteAdicional = ZSql
            Rem Set rstClienteAdicional = db.OpenRecordset(spClienteAdicional, dbOpenSnapshot, dbSQLPassThrough)
            Rem If rstClienteAdicional.RecordCount > 0 Then
            Rem     Descuento.Text = Str$(rstClienteAdicional!Descuento)
            Rem     Descuento.Text = Pusing("###,###.##", Descuento.Text)
            Rem     Vendedor.Text = rstClienteAdicional!Vendedor
            Rem     rstClienteAdicional.Close
            Rem End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Codigo = " + "'" + Vendedor.Text + "'"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                DesVendedor.Caption = rstVendedor!Nombre
                rstVendedor.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Expreso"
            ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
            spExpreso = ZSql
            Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
            If rstExpreso.RecordCount > 0 Then
                DesExpreso.Caption = rstExpreso!Nombre
                ZZDireccionExpreso = rstExpreso!Direccion
                rstExpreso.Close
                    Else
                DesExpreso.Caption = ""
                ZZDireccionExpreso = ""
            End If

            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                DesPago.Caption = rstCondPago!Nombre
                rstCondPago.Close
                    Else
                DesPago.Caption = ""
            End If
            
            WPunto = Str(ConfigPunto)
            Call Ceros(WPunto, 4)
            Punto.Text = WPunto
                
            Numero.Text = "1"
            WTipo = "01"
            
            ZSql = ""
            ZSql = ZSql + "Select CtaCte.Letra, CtaCte.Punto, CtaCte.Numero, CtaCte.Tipo, CtaCTe.Remito"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Letra = " + "'" + Letra.Text + "'"
            ZSql = ZSql + " and Ctacte.Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
            ZSql = ZSql + " Order by Ctacte.Numero"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                With rstCtaCte
                    .MoveLast
                    Do
                        If .BOF = False Then
                    
                            If Letra.Text = rstCtaCte!Letra And Punto.Text = rstCtaCte!Punto Then
                                Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 3 Then
                                    Numero.Text = Str$(Val(rstCtaCte!Numero) + 1)
                                    Exit Do
                                End If
                            End If
                                
                            .MovePrevious
                            
                            If .BOF = True Then
                                Exit Do
                            End If
                                
                                Else
                            
                            Exit Do
                    
                        End If
                    Loop
                End With
                rstCtaCte.Close
            End If
            
            
            ZRemito1 = "0"
            ZRemito2 = "0"
            
            ZSql = ""
            ZSql = ZSql + "Select CtaCte.Letra, CtaCte.Punto, CtaCte.Numero, CtaCte.Tipo, CtaCTe.NroRemito"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Ctacte.Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Ctacte.Letra = " + "'" + "A" + "'"
            ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
            ZSql = ZSql + " Order by Ctacte.Numero"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                With rstCtaCte
                    .MoveLast
                    Do
                        If .BOF = False Then
                    
                            If Punto.Text = rstCtaCte!Punto Then
                                Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                If Val(rstCtaCte!Tipo) = 1 Then
                                    ZRemito1 = Str$(rstCtaCte!NroRemito + 1)
                                    Exit Do
                                End If
                            End If
                                
                            .MovePrevious
                            
                            If .BOF = True Then
                                Exit Do
                            End If
                                
                                Else
                            
                            Exit Do
                    
                        End If
                    Loop
                End With
                rstCtaCte.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select CtaCte.Letra, CtaCte.Punto, CtaCte.Numero, CtaCte.Tipo, CtaCTe.NroRemito"
            ZSql = ZSql + " FROM Ctacte"
            ZSql = ZSql + " Where Ctacte.Punto = " + "'" + Punto.Text + "'"
            ZSql = ZSql + " and Ctacte.Tipo = " + "'" + "01" + "'"
            ZSql = ZSql + " and Ctacte.Letra = " + "'" + "B" + "'"
            ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
            ZSql = ZSql + " Order by Ctacte.Numero"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                With rstCtaCte
                    .MoveLast
                    Do
                        If .BOF = False Then
                    
                            If Punto.Text = rstCtaCte!Punto Then
                                Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                                If Val(rstCtaCte!Tipo) = 1 Then
                                    ZRemito2 = Str$(rstCtaCte!NroRemito + 1)
                                    Exit Do
                                End If
                            End If
                                
                            .MovePrevious
                            
                            If .BOF = True Then
                                Exit Do
                            End If
                                
                                Else
                            
                            Exit Do
                    
                        End If
                    Loop
                End With
                rstCtaCte.Close
            End If
            
            If Val(ZRemito1) > Val(ZRemito2) Then
                Remito.Text = ZRemito1
                    Else
                Remito.Text = ZRemito2
            End If
            
            Pedido.SetFocus
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM HistorialCliente"
            ZSql = ZSql + " Where HistorialCliente.Cliente = " + "'" + CLIENTE.Text + "'"
            spHistorialCliente = ZSql
            Set rstHistorialCliente = db.OpenRecordset(spHistorialCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstHistorialCliente.RecordCount > 0 Then
                rstHistorialCliente.Close
                ZZPasaCliente = CLIENTE.Text
                ZZPasaProceso = 0
                PrgHistorialClienteConsulta.Show
            End If
            
                Else
                
            CLIENTE.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        CLIENTE.Text = ""
        DesCliente.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Numero.Text
        Call Ceros(WPunto, 4)
        
        Numero.Text = "1"
        WTipo = "01"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Letra = " + "'" + Letra.Text + "'"
        ZSql = ZSql + " and Ctacte.Punto = " + "'" + Punto.Text + "'"
        ZSql = ZSql + " and Ctacte.Numero <= " + "'" + "99999999" + "'"
        ZSql = ZSql + " Order by Ctacte.Numero"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            With rstCtaCte
                .MoveLast
                Do
                    If .BOF = False Then
                    
                        If Letra.Text = rstCtaCte!Letra And Punto.Text = rstCtaCte!Punto Then
                            Rem If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 2 Or Val(rstCtaCte!Tipo) = 3 Or Val(rstCtaCte!Tipo) = 4 Or Val(rstCtaCte!Tipo) = 5 Then
                            If Val(rstCtaCte!Tipo) = 1 Or Val(rstCtaCte!Tipo) = 3 Then
                                Numero.Text = Str$(Val(rstCtaCte!Numero) + 1)
                                Exit Do
                            End If
                        End If
                                
                        .MovePrevious
                            
                        If .BOF = True Then
                            Exit Do
                        End If
                                
                            Else
                            
                        Exit Do
                    
                    End If
                Loop
            End With
            rstCtaCte.Close
        End If
        
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
            
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
            
        WTipo = "01"
            
        ClaveVen$ = Letra.Text + WTipo + WPunto + Auxi + "01"
           
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Clave = " + "'" + ClaveVen$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
            
            Fecha.Text = rstCtaCte!Fecha
            CLIENTE.Text = rstCtaCte!CLIENTE
            Pedido.Text = Str$(Val(rstCtaCte!Pedido))
            Descuento.Text = Str$(rstCtaCte!Descuento)
            Descuento.Text = Pusing("###,###.##", Descuento.Text)
            Pago.Text = rstCtaCte!Pago
            Partida.Text = rstCtaCte!Partida
            Comision.Value = rstCtaCte!Comision
            Expreso.Text = rstCtaCte!Expreso
            TipoIva.ListIndex = rstCtaCte!TipoIva
            Comision.Value = rstCtaCte!Comision
            Remito.Text = rstCtaCte!NroRemito
            OCompra.Text = IIf(IsNull(rstCtaCte!Orden), "", rstCtaCte!Orden)
            Cae.Text = IIf(IsNull(rstCtaCte!Cae), "", rstCtaCte!Cae)
            VtoCae.Text = IIf(IsNull(rstCtaCte!VtoCae), "", rstCtaCte!VtoCae)

            rstCtaCte.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!Razon
                Rem Descuento.Text = Str$(rstCliente!Descuento)
                Rem Descuento.Text = Pusing("###,###.##", Descuento.Text)
                WProvincia = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                Vendedor.Text = rstCliente!Vendedor
                rstCliente.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                DesPago.Caption = Trim(rstCondPago!Nombre)
                rstCondPago.Close
            End If
            
            Call Proceso_Click
                
                Else
                    
            Graba.Enabled = True
            WNumero = Numero.Text
            Numero.Text = WNumero
            Fecha.SetFocus
                
        End If
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Pedido.SetFocus
                Else
            M$ = "Formato de fecha invalido"
            a% = MsgBox(M$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub FechaRecon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaRecon.Text, Auxi)
        If Auxi = "S" Then
        
            PantaRecon.Visible = False
        
            Erase ZFactuImpre
            ZLugar = 0
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Fecha = " + "'" + FechaRecon.Text + "'"
            ZSql = ZSql + " Order by CtaCte.Clave"
        
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
            
                With rstCtaCte
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            Select Case Val(rstCtaCte!Tipo)
                                Case 1
                                    ZLugar = ZLugar + 1
                                    ZFactuImpre(ZLugar, 1) = rstCtaCte!Tipo
                                    ZFactuImpre(ZLugar, 2) = rstCtaCte!Punto
                                    ZFactuImpre(ZLugar, 3) = rstCtaCte!Letra
                                    ZFactuImpre(ZLugar, 4) = rstCtaCte!Numero
                                    ZFactuImpre(ZLugar, 5) = rstCtaCte!CLIENTE
                                Case Else
                            End Select
                                
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                
                rstCtaCte.Close
            End If
            
            For Ciclo = 1 To ZLugar
            
                Punto.Text = ZFactuImpre(Ciclo, 2)
                Letra.Text = ZFactuImpre(Ciclo, 3)
                Numero.Text = ZFactuImpre(Ciclo, 4)
                CLIENTE.Text = ZFactuImpre(Ciclo, 5)
                
                Call Numero_Keypress(13)
                Call Impresion_Facturaii
                If Val(Vendedor.Text) <> 1 And Val(Vendedor.Text) <> 7 Then
                    Call Impresion_Facturaii
                End If
                Call Limpia_Click
                
            Next Ciclo
            
                Else
                
            M$ = "Formato de fecha invalido"
            a% = MsgBox(M$, 0, "Emision de Comprobante varios")
            FechaRecon.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        FechaRecon.Text = "  /  /    "
    End If
End Sub


Private Sub Pedido_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Pedido.Text) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Pedido.Numero = " + "'" + Pedido.Text + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZZCliente = rstPedido!CLIENTE
                ZZDescuento = rstPedido!Descuento
                ZZPago = rstPedido!Pago
                ZZPartida = rstPedido!Partida
                ZZOCompra = IIf(IsNull(rstPedido!OCompra), "", rstPedido!OCompra)
                
                rstPedido.Close
                If Trim(UCase(ZZCliente)) <> Trim(UCase(CLIENTE.Text)) Then
                    M$ = "El cliente informado no concuerda con el del pedido"
                    a% = MsgBox(M$, 0, "Ingreso de Facturas")
                    Exit Sub
                End If
                Descuento.Text = ZZDescuento
                Descuento.Text = Pusing("###,###.##", Descuento.Text)
                Pago.Text = ZZPago
                Partida.Text = ZZPartida
                OCompra.Text = ZZOCompra
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CondPago"
                ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
                spCondPago = ZSql
                Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
                If rstCondPago.RecordCount > 0 Then
                    DesPago.Caption = Trim(rstCondPago!Nombre)
                    rstCondPago.Close
                End If
                
                    Else
                M$ = "Pedido Inexistente"
                a% = MsgBox(M$, 0, "Ingreso de Facturas")
                Exit Sub
            End If
        
            Descuento.SetFocus
        
                Else
                
            Pago.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Pedido.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descuento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descuento.Text = Pusing("###,###.##", Descuento.Text)
        Pago.SetFocus
    End If
    If KeyAscii = 27 Then
        Descuento.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CondPago"
        ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Pago.Text + "'"
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstCondPago.RecordCount > 0 Then
            DesPago.Caption = rstCondPago!Nombre
            rstCondPago.Close
            Expreso.SetFocus
                Else
            Pago.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Pago.Text = ""
        DesPago.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub EXPRESO_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Expreso"
        ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
        spExpreso = ZSql
        Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
        If rstExpreso.RecordCount > 0 Then
            DesExpreso.Caption = rstExpreso!Nombre
            ZZDireccionExpreso = rstExpreso!Direccion
            rstExpreso.Close
            CONFIRMA.Text = Partida.Text
            PantallaConfirma.Visible = True
            CONFIRMA.SetFocus
                Else
            Expreso.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Expreso.Text = ""
        DesExpreso.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Confirma_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CONFIRMA.Text = Trim(UCase(CONFIRMA.Text))
        If CONFIRMA.Text = "S" Or CONFIRMA.Text = "N" Or CONFIRMA.Text = "/" Or CONFIRMA.Text = "?" Then
            PantallaConfirma.Visible = False
            If CONFIRMA.Text <> "N" Then
                Partida.Text = CONFIRMA.Text
                Call Lee_Pedido
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
            End If
        End If
    End If
    If KeyAscii = 27 Then
        CONFIRMA.Text = ""
    End If
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
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !CLIENTE + " " + !Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !CLIENTE
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by CondPago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
            End If
    
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
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
                rstArticulo.Close
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


Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub Reconstruccion_Click()
    FechaRecon.Text = "  /  /    "
    PantaRecon.Visible = True
    FechaRecon.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,f5,f9,F10
        Case 112, 113, 114, 115, 116, 120, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,f5,f9,F10
        Case 112, 113, 114, 115, 116, 120, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,f5,f9,F10
        Case 112, 113, 114, 115, 116, 120, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 4
            Rem If WVector1.Row < WVector1.Rows - 1 Then
            If WVector1.Row < 38 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Trim(WVector1.Text) <> "" Then
                ZZVeri = UCase(Left$(WVector1.Text, 1))
                If ZZVeri < "A" Or ZZVeri > "Z" Then
                    ZZVeri = Left$(WVector1.TextMatrix(WVector1.Row - 1, 1), 1)
                    WVector1.Text = ZZVeri + WVector1.Text
                End If
                Auxi = UCase(Left$(WVector1.Text, 1))
                Auxi1 = Mid$(WVector1.Text, 2, 5)
                Call Ceros(Auxi1, 5)
                WVector1.Text = Auxi + Auxi1
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pedido"
            ZSql = ZSql + " Where Numero = " + "'" + Pedido.Text + "'"
            ZSql = ZSql + " and Articulo = " + "'" + WVector1.Text + "'"
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            If rstPedido.RecordCount > 0 Then
                ZPasa = "S"
                rstPedido.Close
                    Else
                Rem m$ = "El articulo no esta en el pedido"
                Rem a% = MsgBox(m$, 0, "Carga de Articulos")
                Rem ZPasa = "N"
                Rem WControl = "N"
                ZPasa = "S"
            End If
            
            If ZPasa = "S" Then
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZIva = rstArticulo!Iva
                    If ZZIva <> TipoIva.ListIndex Then
                        M$ = "La condicion de Iva del Articulo " + ZDescripcion + " no coincide con el informado en la factura"
                        a% = MsgBox(M$, 0, "Emision de Facturas")
                        WControl = "N"
                            Else
                        WVector1.Col = 2
                        WVector1.Text = rstArticulo!Descripcion
                        WVector1.Col = 4
                        WVector1.Text = Str$(rstArticulo!Precio)
                        WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        WVector1.Col = 6
                        WVector1.Text = Str$(rstArticulo!Stock)
                        WVector1.Col = 2
                    End If
                    rstArticulo.Close
                        Else
                    WControl = "N"
                End If
        
            End If
            
        Case 3
            Select Case Partida.Text
                Case "/"
                    MiResultado = Val(WVector1.TextMatrix(WVector1.Row, 3)) Mod 2
                    If MiResultado <> 0 Then
                        M$ = "Las cantidades no son concordantes con el tipo de facturacion en el articulo " + ZDescripcion
                        a% = MsgBox(M$, 0, "Emision de Facturas")
                        WVector1.TextMatrix(WVector1.Row, 3) = ""
                        WControl = "N"
                    End If
                Case "?"
                    MiResultado = Val(WVector1.TextMatrix(WVector1.Row, 3)) Mod 12
                    If MiResultado <> 0 Then
                        M$ = "Las cantidades no son concordantes con el tipo de facturacion"
                        a% = MsgBox(M$, 0, "Emision de Facturas")
                        WVector1.TextMatrix(WVector1.Row, 3) = ""
                        WControl = "N"
                    End If
                Case Else
            End Select
            
            WCantidad = Val(WVector1.Text)
            If Val(WCantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        WCantidad = WCantidad / 2
                    Case "?"
                        WCantidad = WCantidad / 12
                    Case Else
                End Select
            End If
        
            WVector1.TextMatrix(WVector1.Row, 5) = Str$(WCantidad * Val(WVector1.TextMatrix(WVector1.Row, 4)))
            WVector1.TextMatrix(WVector1.Row, 5) = Pusing("###,###.##", WVector1.TextMatrix(WVector1.Row, 5))
            
        Case 4
            WCantidad = Val(WVector1.TextMatrix(WVector1.Row, 3))
            If Val(WCantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        WCantidad = WCantidad / 2
                    Case "?"
                        WCantidad = WCantidad / 12
                    Case Else
                End Select
            End If
        
            WVector1.TextMatrix(WVector1.Row, 5) = Str$(WCantidad * Val(WVector1.Text))
            WVector1.TextMatrix(WVector1.Row, 5) = Pusing("###,###.##", WVector1.TextMatrix(WVector1.Row, 5))
            
        Case Else
            WVector1.Col = XColumna
    End Select
    Call Calcula_Click
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        WVector1.Col = 3
        WAuxi3 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    Call Calcula_Click
    
    End If
    
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 3
    
    Call Opcion_Click
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 51
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Articulo"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 3100
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Precio"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 6
                WVector1.Text = "Stock"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho
    
    For Ciclo = 1 To 50
        WVector1.TextMatrix(Ciclo, 0) = Trim(Str$(Ciclo))
    Next Ciclo

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Cliente_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulo"
    Opcion.AddItem "Pedidos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Pedido_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Pedidos"
    Opcion.AddItem "Condicion"
    Opcion.AddItem "Articulo"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Pago_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Pedidos"
    Opcion.AddItem "Condicion"
    Opcion.AddItem "Articulo"

    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub PedidoAyuda_Click()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulo"
    Opcion.AddItem "Pedidos"

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = ""
    sCentimos = "centavos"
    
    Numero = CStr(Val(Total.Caption))
    
    XTexto1 = Numero2Letra(Numero, , sMoneda & " ", sCentimos & " ")
    XTexto1 = XTexto1 + Space$(100)
    
    Pasa = 0
    
    For da = 60 To 1 Step -1
        If Mid$(XTexto1, da, 1) = Space$(1) Then
            Pasa = 1
        End If
        If Pasa = 1 Then
            If Mid$(XTexto1, da, 1) <> Space$(1) Then
                Exit For
            End If
        End If
    Next da
    
    XTexto2 = Mid$(XTexto1, da + 2, 100)
    XTexto1 = Left$(XTexto1, da)
    
End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Letra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Punto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Pedido_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descuento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Remito_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Partida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Pago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Graba_Click
        Case 113
            Call CmdDelete_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call PedidoAyuda_Click
        Case 120
            Call Impresion_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub


Sub Impresion_Factura()


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WProvincia = rstCliente!Provincia
        WCodIva = rstCliente!Iva
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZZClienteII = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Numero = " + "'" + Pedido.Text + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        ZZClienteII = IIf(IsNull(rstPedido!ClienteII), "", rstPedido!ClienteII)
        rstPedido.Close
    End If

    WProvinciaII = WProvincia
    WDireccionII = WDireccion
    WLocalidadII = WLocalidad
    WPostalII = WPostal

    If Trim(ZZClienteII) <> "" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZZClienteII + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WProvinciaII = rstCliente!Provincia
            WDireccionII = rstCliente!Direccion
            WLocalidadII = rstCliente!Localidad
            WPostalII = rstCliente!Postal
            rstCliente.Close
        End If
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZCuitII = rstExpreso!Cuit
        rstExpreso.Close
    End If
    

    If Letra.Text = "A" Then
    

        Open "lpt3" For Output As #1
        Rem Open "dada3.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(66); Fecha.Text;
        Print #1, Tab(126); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        If Comision.Value = 1 Then
            ImpreComi = "*"
                Else
            ImpreComi = ""
        End If
        
        Print #1, Tab(10); Trim(WRazon); " "; CLIENTE.Text; " "; ZZClienteII; " "; Numero.Text; " "; ImpreComi;
        Print #1, Tab(85); Trim(Left$(WRazon, 30)); " "; CLIENTE.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad);
        Print #1, Tab(85); Trim(WDireccionII); " "; Trim(WLocalidadII)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE";
                Print #1, Tab(85); "CP:B" + WPostalII + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE";
                Print #1, Tab(85); "CP%B" + WPostalII + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE";
                Print #1, Tab(85); "CP B" + WPostalII + "BIE"
        End Select
        
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit;
        Print #1, Tab(85); Iva(Val(WCodIva));
        Print #1, Tab(121); Trim(WCuit)
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35);
        Print #1, Tab(55); Remito.Text;
        Print #1, Tab(91); Left$(DesPago.Caption, 35)
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL";
        Print #1, Tab(82); "Item";
        Print #1, Tab(88); "Uni.";
        Print #1, Tab(93); "Descripcion"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            If Val(Cantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        Cantidad = Cantidad / 2
                    Case "?"
                        Cantidad = Cantidad / 12
                    Case Else
                End Select
            End If
            Precio = Val(WVector1.TextMatrix(a, 4))
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
            
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(22); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial));
                Print #1, Tab(82); a;
                Print #1, Tab(86); Alinea("#####", Str$(Cantidad));
                Rem Print #1, Tab(93); Left$(Articulo, 8);
                Print #1, Tab(92); Left$(ZDescripcion, 17);
                
                ZZDespacho = ""
                ZZNroDespacho = ""
                ZZOrigen = ""
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(Articulo) + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZDespacho = rstArticulo!Despacho
                    rstArticulo.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Despacho"
                ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Str$(ZZDespacho) + "'"
                spDespacho = ZSql
                Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
                If rstDespacho.RecordCount > 0 Then
                    ZZNroDespacho = rstDespacho!Numero
                    ZZOrigen = rstDespacho!Origen
                    rstDespacho.Close
                End If
                
                Print #1, Tab(112); Left$(ZZNroDespacho, 12);
                Print #1, Tab(125); Left$(ZZOrigen, 7)
                
                    Else
                    
                Print #1, ""
                
            End If
            
        Next a
        
        If Trim(OCompra.Text) <> "" Then
            Print #1, Tab(35); "Orden de Compra : " + OCompra.Text;
            Print #1, Tab(85); "Orden de Compra : " + OCompra.Text
                Else
            Print #1, ""
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(17); "Dto."; Alinea("###.##", Descuento.Text);
        End If
        If TipoIva.ListIndex = 0 Then
            Print #1, Tab(47); "21";
                Else
            Print #1, Tab(47); "10.5";
        End If
        Print #1, Tab(85); DesExpreso.Caption
        
        Print #1, Tab(2); Alinea("###,###.##", SubTotal.Caption);
        Print #1, Tab(15); Alinea("###,###.##", Dto.Caption);
        Print #1, Tab(27); Alinea("###,###.##", Neto.Caption);
        Print #1, Tab(38); Alinea("###,###.##", Iva1.Caption);
        Print #1, Tab(51); Alinea("###,###.##", Iva2.Caption);
        Print #1, Tab(63); Alinea("###,###.##", Total.Caption);
        Print #1, Tab(85); Left$(ZZDireccionExpreso, 30)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
            Else
    
        Open "lpt1" For Output As #1
        Rem Open "dada11.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(66); Fecha.Text;
        Print #1, Tab(126); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        If Comision.Value = 1 Then
            ImpreComi = "*"
                Else
            ImpreComi = ""
        End If
        
        Print #1, Tab(10); Trim(WRazon); " "; CLIENTE.Text; " "; Numero.Text; " "; ImpreComi;
        Print #1, Tab(85); Trim(WRazon); " "; CLIENTE.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad);
        Print #1, Tab(85); Trim(WDireccion); " "; Trim(WLocalidad)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE";
                Print #1, Tab(85); "CP:B" + WPostal + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE";
                Print #1, Tab(85); "CP%B" + WPostal + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE";
                Print #1, Tab(85); "CP B" + WPostal + "BIE"
        End Select
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit;
        Print #1, Tab(85); Iva(Val(WCodIva));
        Print #1, Tab(121); Trim(WCuit)
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35);
        Print #1, Tab(55); Remito.Text;
        Print #1, Tab(91); Left$(DesPago.Caption, 35)
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL";
        Print #1, Tab(82); "Item";
        Print #1, Tab(88); "Uni.";
        Print #1, Tab(93); "Descripcion"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            If Val(Cantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        Cantidad = Cantidad / 2
                    Case "?"
                        Cantidad = Cantidad / 12
                    Case Else
                End Select
            End If
            
            Precio = Val(WVector1.TextMatrix(a, 4))
            If TipoIva.ListIndex = 0 Then
                WWImpre = Precio * (1 + (ConfigIva1) / 100)
                    Else
                WWImpre = Precio * (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(WWImpre)
            Precio = WWImpre
            
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
            
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(21); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial));
                Print #1, Tab(82); a;
                Print #1, Tab(86); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(92); Left$(ZDescripcion, 17);
                
                ZZDespacho = ""
                ZZNroDespacho = ""
                ZZOrigen = ""
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(Articulo) + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZDespacho = rstArticulo!Despacho
                    rstArticulo.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Despacho"
                ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Str$(ZZDespacho) + "'"
                spDespacho = ZSql
                Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
                If rstDespacho.RecordCount > 0 Then
                    ZZNroDespacho = rstDespacho!Numero
                    ZZOrigen = rstDespacho!Origen
                    rstDespacho.Close
                End If
                
                Print #1, Tab(112); Left$(ZZNroDespacho, 12);
                Print #1, Tab(125); Left$(ZZOrigen, 7)
                
                    Else
                    
                Print #1, ""
                
            End If
            
        Next a
        
        If Trim(OCompra.Text) <> "" Then
            Print #1, Tab(35); "Orden de Compra : " + OCompra.Text;
            Print #1, Tab(85); "Orden de Compra : " + OCompra.Text
                Else
            Print #1, ""
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(85); DesExpreso.Caption
        Print #1, ""
        
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(7); "SubTotal:"; Alinea("###,###.##", SubTotal.Caption);
            Print #1, Tab(31); "Descuento:"; Alinea("###.##", Descuento.Text); "% "; Alinea("###,###.##", Dto.Caption);
        End If
        
        Print #1, Tab(65); Alinea("###,###.##", Total.Caption);
        Print #1, Tab(85); Left$(ZZDireccionExpreso, 30)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
    End If
            
            

End Sub



Sub Impresion_Remito()


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WProvincia = rstCliente!Provincia
        WCodIva = rstCliente!Iva
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZZClienteII = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Numero = " + "'" + Pedido.Text + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        ZZClienteII = IIf(IsNull(rstPedido!ClienteII), "", rstPedido!ClienteII)
        rstPedido.Close
    End If

    WProvinciaII = WProvincia
    WDireccionII = WDireccion
    WLocalidadII = WLocalidad
    WPostalII = WPostal

    If Trim(ZZClienteII) <> "" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZZClienteII + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WProvinciaII = rstCliente!Provincia
            WDireccionII = rstCliente!Direccion
            WLocalidadII = rstCliente!Localidad
            WPostalII = rstCliente!Postal
            rstCliente.Close
        End If
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZCuitII = rstExpreso!Cuit
        rstExpreso.Close
    End If
    

    Open "lpt3" For Output As #1
    Rem Open "dada3.txt" For Output As #1
    
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Tab(51); Fecha.Text;
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    If Comision.Value = 1 Then
        ImpreComi = "*"
            Else
        ImpreComi = ""
    End If
    
    Print #1, Tab(10); Trim(Left$(WRazon, 30)); " "; CLIENTE.Text; " "; Letra.Text + "-"; Numero.Text
    Print #1, Tab(10); Trim(WDireccionII); " "; Trim(WLocalidadII)
    Select Case Partida.Text
        Case "/"
            Print #1, Tab(10); "CP:B" + WPostal + "BIE"
        Case "?"
            Print #1, Tab(10); "CP%B" + WPostal + "BIE"
        Case Else
            Print #1, Tab(10); "CP B" + WPostal + "BIE"
    End Select
    
    Print #1, Tab(10); Iva(Val(WCodIva));
    Print #1, Tab(44); Trim(WCuit)
    Print #1, ""
    
    Print #1, Tab(15); Left$(DesPago.Caption, 35);
    Print #1, ""
    Print #1, Tab(5); "Item";
    Print #1, Tab(11); "Uni.";
    Print #1, Tab(17); "Descripcion"
    Print #1, ""
    
    Impre = 0
    
    For a = 1 To 40
        
        Articulo = WVector1.TextMatrix(a, 1)
        ZDescripcion = WVector1.TextMatrix(a, 2)
        Cantidad = Val(WVector1.TextMatrix(a, 3))
        If ZZPasaImpre = 0 Then
            If Val(Cantidad) <> 0 Then
                Select Case Partida.Text
                    Case "/"
                        Cantidad = Cantidad / 2
                    Case "?"
                        Cantidad = Cantidad / 12
                    Case Else
                End Select
            End If
        End If
        Precio = Val(WVector1.TextMatrix(a, 4))
        parcial = Precio * Cantidad
    
        If Articulo <> "" Then
        
            Print #1, Tab(5); a;
            Print #1, Tab(10); Alinea("#####", Str$(Cantidad));
            Print #1, Tab(17); Left$(ZDescripcion, 20);
            
            ZZDespacho = ""
            ZZNroDespacho = ""
            ZZOrigen = ""
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(Articulo) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZDespacho = rstArticulo!Despacho
                rstArticulo.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Despacho"
            ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Str$(ZZDespacho) + "'"
            spDespacho = ZSql
            Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
            If rstDespacho.RecordCount > 0 Then
                ZZNroDespacho = rstDespacho!Numero
                ZZOrigen = rstDespacho!Origen
                rstDespacho.Close
            End If
            
            Print #1, Tab(40); Left$(ZZNroDespacho, 12);
            Print #1, Tab(55); Left$(ZZOrigen, 7)
            
                Else
                
            Print #1, ""
            
        End If
        
    Next a
    
    If Trim(OCompra.Text) <> "" Then
        Print #1, Tab(25); "Orden de Compra : " + OCompra.Text
            Else
        Print #1, ""
    End If
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, Tab(10); DesExpreso.Caption
    Print #1, Tab(10); Left$(ZZDireccionExpreso, 30)
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    Print #1, ""
    
    Rem Print #1, Chr$(12)
    
    Close #1

End Sub


Sub Impresion_Factura_Reimpre()


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WProvincia = rstCliente!Provincia
        WCodIva = rstCliente!Iva
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZZClienteII = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Numero = " + "'" + Pedido.Text + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        ZZClienteII = IIf(IsNull(rstPedido!ClienteII), "", rstPedido!ClienteII)
        rstPedido.Close
    End If

    WProvinciaII = WProvincia
    WDireccionII = WDireccion
    WLocalidadII = WLocalidad
    WPostalII = WPostal

    If Trim(ZZClienteII) <> "" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZZClienteII + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WProvinciaII = rstCliente!Provincia
            WDireccionII = rstCliente!Direccion
            WLocalidadII = rstCliente!Localidad
            WPostalII = rstCliente!Postal
            rstCliente.Close
        End If
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZCuitII = rstExpreso!Cuit
        rstExpreso.Close
    End If
    

    If Letra.Text = "A" Then
    

        Open "lpt3" For Output As #1
        Rem Open "dada3.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(66); Fecha.Text;
        Print #1, Tab(126); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        If Comision.Value = 1 Then
            ImpreComi = "*"
                Else
            ImpreComi = ""
        End If
        
        Print #1, Tab(10); Trim(WRazon); " "; CLIENTE.Text; " "; ZZClienteII; " "; Numero.Text; " "; ImpreComi;
        Print #1, Tab(85); Trim(Left$(WRazon, 30)); " "; CLIENTE.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad);
        Print #1, Tab(85); Trim(WDireccionII); " "; Trim(WLocalidadII)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE";
                Print #1, Tab(85); "CP:B" + WPostalII + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE";
                Print #1, Tab(85); "CP%B" + WPostalII + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE";
                Print #1, Tab(85); "CP B" + WPostalII + "BIE"
        End Select
        
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit;
        Print #1, Tab(85); Iva(Val(WCodIva));
        Print #1, Tab(121); Trim(WCuit)
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35);
        Print #1, Tab(55); Remito.Text;
        Print #1, Tab(91); Left$(DesPago.Caption, 35)
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL";
        Print #1, Tab(82); "Item";
        Print #1, Tab(88); "Uni.";
        Print #1, Tab(93); "Descripcion"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            Rem If Val(Cantidad) <> 0 Then
            Rem     Select Case Partida.Text
            Rem         Case "/"
            Rem             Cantidad = Cantidad / 2
            Rem         Case "?"
            Rem             Cantidad = Cantidad / 12
            Rem         Case Else
            Rem     End Select
            Rem End If
            Precio = Val(WVector1.TextMatrix(a, 4))
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
            
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(22); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial));
                Print #1, Tab(82); a;
                Print #1, Tab(86); Alinea("#####", Str$(Cantidad));
                Rem Print #1, Tab(93); Left$(Articulo, 8);
                Print #1, Tab(92); Left$(ZDescripcion, 17);
                
                ZZDespacho = ""
                ZZNroDespacho = ""
                ZZOrigen = ""
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(Articulo) + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZDespacho = rstArticulo!Despacho
                    rstArticulo.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Despacho"
                ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Str$(ZZDespacho) + "'"
                spDespacho = ZSql
                Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
                If rstDespacho.RecordCount > 0 Then
                    ZZNroDespacho = rstDespacho!Numero
                    ZZOrigen = rstDespacho!Origen
                    rstDespacho.Close
                End If
                
                Print #1, Tab(112); Left$(ZZNroDespacho, 12);
                Print #1, Tab(125); Left$(ZZOrigen, 7)
                
                    Else
                    
                Print #1, ""
                
            End If
            
        Next a
        
        If Trim(OCompra.Text) <> "" Then
            Print #1, Tab(35); "Orden de Compra : " + OCompra.Text;
            Print #1, Tab(85); "Orden de Compra : " + OCompra.Text
                Else
            Print #1, ""
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(17); "Dto."; Alinea("###.##", Descuento.Text);
        End If
        If TipoIva.ListIndex = 0 Then
            Print #1, Tab(47); "21";
                Else
            Print #1, Tab(47); "10.5";
        End If
        Print #1, Tab(85); DesExpreso.Caption
        
        Print #1, Tab(2); Alinea("###,###.##", SubTotal.Caption);
        Print #1, Tab(15); Alinea("###,###.##", Dto.Caption);
        Print #1, Tab(27); Alinea("###,###.##", Neto.Caption);
        Print #1, Tab(38); Alinea("###,###.##", Iva1.Caption);
        Print #1, Tab(51); Alinea("###,###.##", Iva2.Caption);
        Print #1, Tab(63); Alinea("###,###.##", Total.Caption);
        Print #1, Tab(85); Left$(ZZDireccionExpreso, 30)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
            Else
    
        Open "lpt1" For Output As #1
        Rem Open "dada11.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(66); Fecha.Text;
        Print #1, Tab(126); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        If Comision.Value = 1 Then
            ImpreComi = "*"
                Else
            ImpreComi = ""
        End If
        
        Print #1, Tab(10); Trim(WRazon); " "; CLIENTE.Text; " "; Numero.Text; " "; ImpreComi;
        Print #1, Tab(85); Trim(WRazon); " "; CLIENTE.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad);
        Print #1, Tab(85); Trim(WDireccion); " "; Trim(WLocalidad)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE";
                Print #1, Tab(85); "CP:B" + WPostal + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE";
                Print #1, Tab(85); "CP%B" + WPostal + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE";
                Print #1, Tab(85); "CP B" + WPostal + "BIE"
        End Select
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit;
        Print #1, Tab(85); Iva(Val(WCodIva));
        Print #1, Tab(121); Trim(WCuit)
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35);
        Print #1, Tab(55); Remito.Text;
        Print #1, Tab(91); Left$(DesPago.Caption, 35)
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL";
        Print #1, Tab(82); "Item";
        Print #1, Tab(88); "Uni.";
        Print #1, Tab(93); "Descripcion"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            Rem If Val(Cantidad) <> 0 Then
            Rem     Select Case Partida.Text
            Rem         Case "/"
            Rem             Cantidad = Cantidad / 2
            Rem         Case "?"
            Rem             Cantidad = Cantidad / 12
            Rem         Case Else
            Rem     End Select
            Rem End If
            
            Precio = Val(WVector1.TextMatrix(a, 4))
            If TipoIva.ListIndex = 0 Then
                WWImpre = Precio * (1 + (ConfigIva1) / 100)
                    Else
                WWImpre = Precio * (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(WWImpre)
            Precio = WWImpre
            
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
            
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(21); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial));
                Print #1, Tab(82); a;
                Print #1, Tab(86); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(92); Left$(ZDescripcion, 17);
                
                ZZDespacho = ""
                ZZNroDespacho = ""
                ZZOrigen = ""
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(Articulo) + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZZDespacho = rstArticulo!Despacho
                    rstArticulo.Close
                End If
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Despacho"
                ZSql = ZSql + " Where Despacho.Codigo = " + "'" + Str$(ZZDespacho) + "'"
                spDespacho = ZSql
                Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
                If rstDespacho.RecordCount > 0 Then
                    ZZNroDespacho = rstDespacho!Numero
                    ZZOrigen = rstDespacho!Origen
                    rstDespacho.Close
                End If
                
                Print #1, Tab(112); Left$(ZZNroDespacho, 12);
                Print #1, Tab(125); Left$(ZZOrigen, 7)
                
                    Else
                    
                Print #1, ""
                
            End If
            
        Next a
        
        If Trim(OCompra.Text) <> "" Then
            Print #1, Tab(35); "Orden de Compra : " + OCompra.Text;
            Print #1, Tab(85); "Orden de Compra : " + OCompra.Text
                Else
            Print #1, ""
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(85); DesExpreso.Caption
        Print #1, ""
        
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(7); "SubTotal:"; Alinea("###,###.##", SubTotal.Caption);
            Print #1, Tab(31); "Descuento:"; Alinea("###.##", Descuento.Text); "% "; Alinea("###,###.##", Dto.Caption);
        End If
        
        Print #1, Tab(65); Alinea("###,###.##", Total.Caption);
        Print #1, Tab(85); Left$(ZZDireccionExpreso, 30)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
    End If
            
            

End Sub








Sub Impresion_Facturaii()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WProvincia = rstCliente!Provincia
        WCodIva = rstCliente!Iva
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WCuit = rstCliente!Cuit
        rstCliente.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZCuitII = rstExpreso!Cuit
        rstExpreso.Close
    End If
    

    If Letra.Text = "A" Then
    
        Open "lpt1" For Output As #1
        Rem Open "dada34.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(66); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(10); Trim(WRazon); " "; CLIENTE.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE"
        End Select
        
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35)
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            Rem If Val(Cantidad) <> 0 Then
            Rem     Select Case Partida.Text
            Rem         Case "/"
            Rem             Cantidad = Cantidad / 2
            Rem         Case "?"
            Rem             Cantidad = Cantidad / 12
            Rem         Case Else
            Rem     End Select
            Rem End If
            Precio = Val(WVector1.TextMatrix(a, 4))
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(22); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial))
                    Else
                Print #1, ""
            End If
            
        Next a
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(17); "Dto."; Alinea("###.##", Descuento.Text);
        End If
        If TipoIva.ListIndex = 0 Then
            Print #1, Tab(47); "21";
                Else
            Print #1, Tab(47); "10.5";
        End If
        Print #1, ""
        
        Print #1, Tab(2); Alinea("###,###.##", SubTotal.Caption);
        Print #1, Tab(15); Alinea("###,###.##", Dto.Caption);
        Print #1, Tab(27); Alinea("###,###.##", Neto.Caption);
        Print #1, Tab(38); Alinea("###,###.##", Iva1.Caption);
        Print #1, Tab(51); Alinea("###,###.##", Iva2.Caption);
        Print #1, Tab(63); Alinea("###,###.##", Total.Caption)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
            Else
    

        Open "lpt1" For Output As #1
        Rem Open "dada1.txt" For Output As #1
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(66); Fecha.Text;
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(10); Trim(WRazon); " "; CLIENTE.Text; " "; Numero.Text
        Print #1, Tab(10); Trim(WDireccion); " "; Trim(WLocalidad)
        Select Case Partida.Text
            Case "/"
                Print #1, Tab(10); "CP:B" + WPostal + "BIE"
            Case "?"
                Print #1, Tab(10); "CP%B" + WPostal + "BIE"
            Case Else
                Print #1, Tab(10); "CP B" + WPostal + "BIE"
        End Select
        Print #1, Tab(10); Iva(Val(WCodIva));
        Print #1, Tab(61); WCuit
        Print #1, ""
        
        Print #1, Tab(15); Left$(DesPago.Caption, 35);
        Print #1, Tab(55); Remito.Text
        Print #1, ""
        Print #1, Tab(3); "Item";
        Print #1, Tab(9); "Uni.";
        Print #1, Tab(14); "Codigo";
        Print #1, Tab(22); "Descripcion";
        Print #1, Tab(54); "Pr.Unitario";
        Print #1, Tab(68); "TOTAL"
        Print #1, ""
        
        Impre = 0
        
        For a = 1 To 40
            
            Articulo = WVector1.TextMatrix(a, 1)
            ZDescripcion = WVector1.TextMatrix(a, 2)
            Cantidad = Val(WVector1.TextMatrix(a, 3))
            Rem If Val(Cantidad) <> 0 Then
            Rem     Select Case Partida.Text
            Rem         Case "/"
            Rem             Cantidad = Cantidad / 2
            Rem         Case "?"
            Rem             Cantidad = Cantidad / 12
            Rem         Case Else
            Rem     End Select
            Rem End If
            
            Precio = Val(WVector1.TextMatrix(a, 4))
            If TipoIva.ListIndex = 0 Then
                WWImpre = Precio * (1 + (ConfigIva1) / 100)
                    Else
                WWImpre = Precio * (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(WWImpre)
            Precio = WWImpre
            
            parcial = Precio * Cantidad
        
            If Articulo <> "" Then
                Print #1, Tab(3); a;
                Print #1, Tab(8); Alinea("#####", Str$(Cantidad));
                Print #1, Tab(14); Left$(Articulo, 6);
                Print #1, Tab(21); Left$(ZDescripcion, 30);
                Print #1, Tab(52); Alinea("###,###.##", Str$(Precio));
                Print #1, Tab(62); Alinea("###,###.##", Str$(parcial))
                    Else
                Print #1, ""
            End If
            
        Next a
        
        If Trim(OCompra.Text) <> "" Then
            Print #1, Tab(35); "Orden de Compra : " + OCompra.Text;
            Print #1, Tab(85); "Orden de Compra : " + OCompra.Text
                Else
            Print #1, ""
        End If
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        If Val(Descuento.Text) <> 0 Then
            Print #1, Tab(7); "SubTotal:"; Alinea("###,###.##", SubTotal.Caption);
            Print #1, Tab(31); "Descuento:"; Alinea("###.##", Descuento.Text); "% "; Alinea("###,###.##", Dto.Caption);
        End If
        
        Print #1, Tab(65); Alinea("###,###.##", Total.Caption)
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
        Rem Print #1, Chr$(12)
        
        Close #1
        
        
    End If
            
            

End Sub











Private Sub Lee_Pedido()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 99
    
        Auxi = Pedido.Text
        Call Ceros(Auxi, 8)
            
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Clave = " + "'" + WClave + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Trim(rstPedido!Articulo) + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZIva = rstArticulo!Iva
                ZZPrecio = rstArticulo!Precio
                rstArticulo.Close
            End If
            
            If ZZIva = TipoIva.ListIndex And rstPedido!facturado = 0 Then
            Rem If ZZIva = TipoIva.ListIndex Then
        
                Canti = rstPedido!Cantidad
                
                Renglon = Renglon + 1
                        
                WVector1.Row = Renglon
                        
                WVector1.Col = 1
                WVector1.Text = Trim(rstPedido!Articulo)
                Auxi1 = rstPedido!Articulo
                    
                WVector1.Col = 3
                WVector1.Text = Pusing("###,###", Str$(rstPedido!Cantidad))
                    
                WVector1.Col = 4
                WVector1.Text = Pusing("###,###.##", Str$(ZZPrecio))
                
                WCantidad = rstPedido!Cantidad
                If Val(WCantidad) <> 0 Then
                    Select Case Partida.Text
                        Case "/"
                            WCantidad = WCantidad / 2
                        Case "?"
                            WCantidad = WCantidad / 12
                        Case Else
                    End Select
                End If
                
                
                WVector1.Col = 5
                WVector1.Text = Pusing("###,###.##", Str$(ZZPrecio * WCantidad))
                    
                rstPedido.Close
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WVector1.Col = 2
                    WVector1.Text = rstArticulo!Descripcion
                    WVector1.Col = 6
                    WVector1.Text = Str$(rstArticulo!Stock)
                    rstArticulo.Close
                End If
                
                    Else
                    
                rstPedido.Close
                
            End If
            
        End If
    
    Next WRenglon
    Call Calcula_Click
    
End Sub

Private Sub Calcula_Cae()
    
    Dim WSAA As Object, WSFEv1 As Object
    
    On Error GoTo ManejoError
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    Rem Cae.Text = "12345678901234"
    Rem VtoCae.Text = "14/04/2011"
    Rem Exit Sub
    
    ' Crear objeto interface Web Service Autenticaci?n y Autorizaci?n
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    'Debug.Print WSAA.InstallDir
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
    tra = WSAA.CreateTRA("wsfe")
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
        
    ZPath = "c:\salva\"
    ZNombre = "celugama"
    ZCuit = "30637671622"
    punto_vta = 3
    
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Rem Certificado = "..\..\reingart.crt" ' certificado de prueba
    Rem ClavePrivada = "..\..\reingart.key" ' clave privada de prueba
    
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    ' Llamar al web service para autenticar:
    proxy = "" '"usuario:clave@localhost:8000"
    Rem ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms", proxy) ' Homologaci?n
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms", proxy) ' Homologaci?n

    ' Imprimir el ticket de acceso, ToKen y Sign de autorizaci?n
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 24 horas
    ' (este per?odo se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electr?nica de Mercado Interno
    Set WSFEv1 = CreateObject("WSFEv1")
    Debug.Print WSFEv1.Version
    'Debug.Print WSFEv1.InstallDir
    
    ' Setear tocken y sing de autorizaci?n (pasos previos)
    WSFEv1.Token = WSAA.Token
    WSFEv1.Sign = WSAA.Sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFEv1.Cuit = ZCuit
    
    ' Conectar al Servicio Web de Facturaci?n
    proxy = "" ' "usuario:clave@localhost:8000"
    wsdl = "https://servicios1.afip.gov.ar/wsfev1/service.asmx?WSDL"
    cache = ""    'Rem Path
        
    ok = WSFEv1.Conectar(cache, wsdl, proxy, "") ' homologaci?n
    Debug.Print WSFEv1.Version
    
    ' mostrar bit?cora de depuraci?n:
    Debug.Print WSFEv1.DebugLog
    
    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFEv1.Dummy
    Debug.Print "appserver status", WSFEv1.AppServerStatus
    Debug.Print "dbserver status", WSFEv1.DbServerStatus
    Debug.Print "authserver status", WSFEv1.AuthServerStatus
       
    ' Establezco los valores de la factura a autorizar:
    If Letra.Text = "A" Then
        tipo_cbte = 1
            Else
        tipo_cbte = 6
    End If
    
    cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
    If cbte_nro = "" Then
        cbte_nro = 0                ' no hay comprobantes emitidos
            Else
        cbte_nro = CLng(cbte_nro)   ' convertir a entero largo
    End If
    
    If cbte_nro + 1 <> Val(Numero.Text) Then
        M$ = "Numero de comprobante no coincide con el de la afip"
        a% = MsgBox(M$, 0, "Ingreso de Facturas")
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WCuit = rstCliente!Cuit
        rstCliente.Close
        Call Eval
    End If
    
    Rem Fecha = Format(Date, "yyyymmdd")
    
    Rem CONCEPTO   1-PRODUCTO    2-SERVICIOS     3-PRODUCTOS Y SERVICIOS
    concepto = 1
    
    Rem TIPO DE DOCUMENTO
    If Len(WCuit) = 11 Then
        tipo_doc = 80
            Else
        tipo_doc = 96
    End If
    
    Rem NUMERO DE DOCUMENTO
    nro_doc = Left$(WCuit + Space$(11), 11)
    
    Rem NUMERO DE DOCUMENTO
    cbte_nro = cbte_nro + 1
    cbt_desde = cbte_nro
    cbt_hasta = cbte_nro
    
    Rem IMPORTE TOTAL
    imp_total = Val(Total.Caption)
    
    Rem IMPORTE DE CONCEPTOS NO GRAVADOS POR EL IVA
    imp_tot_conc = 0
    
    Rem IMPORTE NETO
    imp_neto = Val(Neto.Caption)
    
    Rem IMPORTE IVA
    imp_iva = Val(Iva1.Caption)
    
    Rem suma de importes de otros impuestos
    imp_trib = 0
    
    Rem IMPORTE EXENTO DE IVA
    imp_op_ex = 0
    
    Rem FECHA
    ZZfecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    fecha_cbte = ZZfecha
    
    Rem VENCIMIENTO
    fecha_venc_pago = ""
    
    Rem FECHAS DE SERVICIOS PARA SERVICIOS
    ' Fechas del per?odo del servicio facturado (solo si concepto = 1?)
    fecha_serv_desde = ""
    fecha_serv_hasta = ""
    
    Rem MONEDA
    moneda_id = "PES"
    
    Rem COTIZACION
    moneda_ctz = 1

    ok = WSFEv1.CrearFactura(concepto, tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        imp_iva, imp_trib, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta, _
        moneda_id, moneda_ctz)
    
    ' Agrego los comprobantes asociados:
    Rem If False Then ' solo nc/nd
    Rem     tipo = 19
    Rem     pto_vta = 2
    Rem     nro = 1234
    Rem     ok = WSFEv1.AgregarCmpAsoc(tipo, pto_vta, nro)
    Rem End If
        
    ' Agrego impuestos varios
    Rem id = 99
    Rem Desc = "Impuesto Municipal Matanza'"
    Rem base_imp = "100.00"
    Rem alic = "1.00"
    Rem importe = "1.00"
    Rem ok = WSFEv1.AgregarTributo(id, Desc, base_imp, alic, importe)

    ' Agrego tasas de IVA
    If TipoIva.ListIndex = 0 Then
        Id = 5 ' 21%
            Else
        Id = 4 ' 10.50%
    End If
    
    If Val(Iva1.Caption) = 0 Then
        Id = 3
    End If
    
    base_imp = Val(Neto.Caption)
    IMPORTE = Val(Iva1.Caption)
    ok = WSFEv1.AgregarIva(Id, base_imp, IMPORTE)
    
    ' Habilito reprocesamiento autom?tico (predeterminado):
    WSFEv1.Reprocesar = True

    ' Solicito CAE:
    Cae = WSFEv1.CAESolicitar()
    
    Debug.Print "Resultado", WSFEv1.resultado
    Debug.Print "CAE", WSFEv1.Cae

    Debug.Print "Numero de comprobante:", WSFEv1.CbteNro
    
    ' Imprimo pedido y respuesta XML para depuraci?n (errores de formato)
    Debug.Print WSFEv1.XmlRequest
    Debug.Print WSFEv1.XmlResponse
    
    Debug.Print "Reprocesar:", WSFEv1.Reprocesar
    Debug.Print "Reproceso:", WSFEv1.Reproceso
    Debug.Print "CAE:", WSFEv1.Cae
    Debug.Print "EmisionTipo:", WSFEv1.EmisionTipo

    MsgBox "Resultado:" & WSFEv1.resultado & " CAE: " & Cae & " Venc: " & WSFEv1.vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
    ' Muestro los errores
    If WSFEv1.ErrMsg <> "" Then
        MsgBox WSFEv1.ErrMsg, vbExclamation, "Error"
    End If
    
    ' Muestro los eventos (mantenimiento programados y otros mensajes de la AFIP)
    For Each evento In WSFEv1.eventos:
        MsgBox evento, vbInformation, "Evento"
    Next
    
    ' Buscar la factura
    cae2 = WSFEv1.CompConsultar(tipo_cbte, punto_vta, cbte_nro)
    
    Debug.Print "Fecha Comprobante:", WSFEv1.FechaCbte
    Debug.Print "Fecha Vencimiento CAE", WSFEv1.vencimiento
    Debug.Print "Importe Total:", WSFEv1.ImpTotal
    Debug.Print "Resultado:", WSFEv1.resultado
    
    If Cae <> cae2 Then
        MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!: " & Cae & " vs " & cae2
    Else
        MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
    End If
        
    If WSFEv1.resultado = "A" Then
        Cae.Text = Trim(Cae)
        If Len(Trim(WSFEv1.vencimiento)) = 8 Then
            VtoCae.Text = Right$(WSFEv1.vencimiento, 2) + "/" + Mid$(WSFEv1.vencimiento, 5, 2) + "/" + Left$(WSFEv1.vencimiento, 4)
                Else
            VtoCae.Text = WSFEv1.vencimiento
        End If
    End If

    Exit Sub
ManejoError:
    ' Si hubo error:
    Debug.Print WSFEv1.Excepcion
    Debug.Print Err.Description            ' descripci?n error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Print WSFEv1.XmlRequest
            Debug.Print WSFEv1.XmlResponse
            Debug.Print WSFEv1.traceback
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select
    Debug.Print WSFEv1.XmlRequest
    Debug.Assert False
    Debug.Print WSFEv1.traceback
End Sub




Private Sub Eval()

    Es = WCuit

    X = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               X = X + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               X = X + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               X = X + Y: MinusOk = 0

        End If

    Next

    WCuit = X

End Sub



Private Sub Calcula_Barra()
    
    Dim ZZCara(1000) As String
    
    ZZNumero = "30637671622"
    
    If Letra.Text = "A" Then
        ZZNumero = ZZNumero + "01"
            Else
        ZZNumero = ZZNumero + "06"
    End If
            
    Auxi1 = Punto.Text
    Call Ceros(Auxi1, 4)
    ZZNumero = ZZNumero + Auxi1
    
    ZZNumero = ZZNumero + Trim(Cae.Text)
    
    ZZFechaCae = VtoCae.Text
    ZZOrdFechaCae = Right$(ZZFechaCae, 4) + Mid$(ZZFechaCae, 4, 2) + Left$(ZZFechaCae, 2)
    ZZNumero = ZZNumero + ZZOrdFechaCae
    
    ZZCara(0) = "!"
    ZZCara(1) = Chr$(34)
    ZZCara(2) = "#"
    ZZCara(3) = "$"
    ZZCara(4) = "%"
    ZZCara(5) = "&"
    ZZCara(6) = "?"
    ZZCara(7) = "("
    ZZCara(8) = ")"
    ZZCara(9) = "*"
    ZZCara(10) = "+"
    ZZCara(11) = ","
    ZZCara(12) = "-"
    ZZCara(13) = "."
    ZZCara(14) = "/"
    ZZCara(15) = "0"
    ZZCara(16) = "1"
    ZZCara(17) = "2"
    ZZCara(18) = "3"
    ZZCara(19) = "4"
    ZZCara(20) = "5"
    ZZCara(21) = "6"
    ZZCara(22) = "7"
    ZZCara(23) = "8"
    ZZCara(24) = "9"
    ZZCara(25) = ":"
    ZZCara(26) = ";"
    ZZCara(27) = "<"
    ZZCara(28) = "="
    ZZCara(29) = ">"
    ZZCara(30) = "?"
    ZZCara(31) = "@"
    ZZCara(32) = "A"
    ZZCara(33) = "B"
    ZZCara(34) = "C"
    ZZCara(35) = "D"
    ZZCara(36) = "E"
    ZZCara(37) = "F"
    ZZCara(38) = "G"
    ZZCara(39) = "H"
    ZZCara(40) = "I"
    ZZCara(41) = "J"
    ZZCara(42) = "K"
    ZZCara(43) = "L"
    ZZCara(44) = "M"
    ZZCara(45) = "N"
    ZZCara(46) = "O"
    ZZCara(47) = "P"
    ZZCara(48) = "Q"
    ZZCara(49) = "R"
    ZZCara(50) = "S"
    ZZCara(51) = "T"
    ZZCara(52) = "U"
    ZZCara(53) = "V"
    ZZCara(54) = "W"
    ZZCara(55) = "X"
    ZZCara(56) = "Y"
    ZZCara(57) = "Z"
    ZZCara(58) = "["
    ZZCara(59) = "\"
    ZZCara(60) = "]"
    ZZCara(61) = "^"
    ZZCara(62) = "_"
    ZZCara(63) = "`"
    ZZCara(64) = "a"
    ZZCara(65) = "b"
    ZZCara(66) = "c"
    ZZCara(67) = "d"
    ZZCara(68) = "e"
    ZZCara(69) = "f"
    ZZCara(70) = "g"
    ZZCara(71) = "h"
    ZZCara(72) = "i"
    ZZCara(73) = "j"
    ZZCara(74) = "k"
    ZZCara(75) = "l"
    ZZCara(76) = "m"
    ZZCara(77) = "n"
    ZZCara(78) = "o"
    ZZCara(79) = "p"
    ZZCara(80) = "q"
    ZZCara(81) = "r"
    ZZCara(82) = "s"
    ZZCara(83) = "t"
    ZZCara(84) = "u"
    ZZCara(85) = "v"
    ZZCara(86) = "w"
    ZZCara(87) = "x"
    ZZCara(88) = "y"
    ZZCara(89) = "z"
    ZZCara(90) = "¡"
    ZZCara(91) = "¢"
    ZZCara(92) = "£"
    ZZCara(93) = "¤"
    ZZCara(94) = "¥"
    ZZCara(95) = "¦"
    ZZCara(96) = "§"
    ZZCara(97) = "¨"
    ZZCara(98) = "©"
    ZZCara(99) = "ª"
    
    Rem ZZNumero = "3070306062119000260321213344273201008198"
    Rem ZZNumero = "000102030405060708091011121314151617181920"
    Rem ZZNumero = "2122232425262728293031323334353637383940"
    Rem ZZNumero = "4142434445464748495051525354555657585960"
    Rem ZZNumero = "6162636465666768697071727374757677787980"
    Rem ZZNumero = "81828384858687888990919293949596979899"
    Rem ZZNumero = "307030606211900026032121334427320100819"
    
    ZZSumaI = 0
    ZZSumaII = 0
    
    For Ciclo = 1 To 39 Step 2
        ZZSumaI = ZZSumaI + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    ZZSumaI = ZZSumaI * 3
    
    For Ciclo = 2 To 39 Step 2
        ZZSumaII = ZZSumaII + Val(Mid$(ZZNumero, Ciclo, 1))
    Next Ciclo
    
    ZZSuma = ZZSumaI + ZZSumaII
    ZZVerifica = ZZSuma
    ZZDigi = 0
    
    Do
    
        ZZVerifi = Int(ZZVerifica / 10) * 10
        
        If ZZVerifi = ZZVerifica Then
            Exit Do
        End If
        
        ZZDigi = ZZDigi + 1
        
        ZZVerifica = ZZSuma + ZZDigi
        
    Loop
    
    ZZNumero = ZZNumero + Trim(Str$(ZZDigi))
    
    lccar = ""
    barralargo = ZZNumero
    
    For lni = 1 To Len(barralargo) Step 2
        ZZLugar = Val(Mid(barralargo, lni, 2))
        lccar = lccar + ZZCara(ZZLugar)
    Next
    
    Rem barralargo = "{" + lccar + "}"
    barralargo = "(" + lccar + ")"
    
    ZZImpreBarra = barralargo
    ZZImpreBarraII = ZZNumero

End Sub



