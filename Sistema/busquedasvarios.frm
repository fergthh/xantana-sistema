VERSION 5.00
Begin VB.Form busquedasvarias 
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
Attribute VB_Name = "busquedasvarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                DesCuenta.Caption = IIf(IsNull(rstCuenta!Descripcion), "", rstCuenta!Descripcion)
                rstCuenta.Close
                Descripcion.SetFocus
                    Else
                DesCuenta.Caption = ""
                Cuenta.SetFocus
            End If
        
        
        
        
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoPro"
    ZSql = ZSql + " Where TipoPro.Codigo = " + "'" + Tipo.Text + "'"
    spTipoPro = ZSql
    Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoPro.RecordCount > 0 Then
        DesTipo.Caption = rstTipoPro!Descripcion
        rstTipoPro.Close
            Else
        DesTipo.Caption = ""
    End If
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        Nombre.Text = rstConceptos!Nombre
        Cuenta.Text = rstConceptos!Cuenta
        rstConceptos.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
    
    

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        Nombre.Text = rstProveedor!Nombre
        Direccion.Text = rstProveedor!Direccion
        Localidad.Text = rstProveedor!Localidad
        Postal.Text = rstProveedor!Postal
        Cuit.Text = rstProveedor!Cuit
        Telefono.Text = rstProveedor!Telefono
        email.Text = rstProveedor!email
        Observaciones.Text = rstProveedor!Observaciones
        Dias.Text = rstProveedor!Dias
        Tipo.Text = rstProveedor!Tipo
        Iva.ListIndex = rstProveedor!Iva
        Ganancia.ListIndex = rstProveedor!Ganancia
        Provincia.ListIndex = rstProveedor!Provincia
        NombreCheque.Text = rstProveedor!NombreCheque
        PorceReteIva.Text = rstProveedor!PorceReteIva
        Reteiva.ListIndex = rstProveedor!Reteiva
        rstProveedor.Close
        Call Format_datos
    End If
    
    

        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proyecto"
    ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + Codigo.Text + "'"
    spProyecto = ZSql
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        Descripcion.Text = rstProyecto!Descripcion
        rstProyecto.Close
        Call Format_datos
        Call Imprime_Nombre
    End If




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
        ConfigPunto = rstConfiguracion!Punto
        rstConfiguracion.Close
    End If


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Banco"
    ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
    spBanco = ZSql
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        DesBanco.Caption = rstBanco!Nombre
        rstBanco.Close
            Else
        DesBanco.Caption = ""
    End If


        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pagos"
        ZSql = ZSql + " Where Pago.Orden = " + "'" + Orden.Text + "'"
        ZSql = ZSql + " and Pago.Renglon = " + "'" + Auxi1 + "'"
        spPago = ZSql
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            rstPago.Close
        End If

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        DesProveedor.Caption = rstProveedor!Nombre
        WPrvDireccion = rstProveedor!Direccion
        WPrvCuit = rstProveedor!Cuit
        WTipoprv = rstProveedor!Ganancia
        WTipoiva = rstProveedor!Iva
        WTipoReteiva = rstProveedor!Reteiva
        WExepcion = rstProveedor!PorceReteIva
        rstProveedor.Close
            Else
        DesProveedor.Caption = ""
        WPrvDireccion = ""
        WPrvCuit = ""
        WTipoprv = 0
        WTipoiva = 0
        WTipoReteiva = 0
        WExepcion = 0
    End If


    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cierre"
    ZSql = ZSql + " Where Cierre.Clave = " + "'" + ZClave + "'"
    spCierre = ZSql
    Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
    If rstCierre.RecordCount > 0 Then
        ZEstado = rstCierre!Estado
        rstCierre.Close
    End If
    


    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Chequera"
                ZSql = ZSql + " Where Chequera.Banco = " + "'" + WBancoCheque + "'"
                ZSql = ZSql + " Order by Chequera.Desde"
                spChequera = ZSql
                Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
                If rstChequera.RecordCount > 0 Then
                    With rstChequera
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                If WNumeroCheque >= rstChequera!Desde And WNumeroCheque <= rstChequera!Hasta Then
                                    Entra = "S"
                                    Exit Do
                                End If
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstChequera.Close
                End If
    



    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Chequera"
    ZSql = ZSql + " Where Chequera.Codigo = " + "'" + Codigo.Text + "'"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        Codigo.Text = !Codigo
        Banco.Text = !Banco
        Fecha.Text = !Fecha
        Desde.Text = !Desde
        Hasta.Text = !Hasta
        rstChequera.Close
        Call Format_datos
        Call Imprime_Nombre
    End If




        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM NroRet"
        ZSql = ZSql + " Where NroRet.Clave = 1"
        spNroRet = ZSql
        Set rstNroRet = db.OpenRecordset(spNroRet, dbOpenSnapshot, dbSQLPassThrough)
        If rstNroRet.RecordCount > 0 Then
            rstNroRet.Close
        End If




                    Indice = Pantalla.ListIndex
                        
            ZSql = ""
            ZSq = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCtePrv"
            ZSql = ZSql + " Where CtaCtePrv.Clave = " + "'" + ZZClave + "'"
            spCtaCtePrv = ZSql
            Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCtePrv.RecordCount > 0 Then
                WSaldo = rstCtaCtePrv!Saldo
                WTotal = rstCtaCtePrv!Total
                rstCtaCtePrv.Close
                If WSaldo <> WTotal Then
                    m$ = "El anticipo no se puede eliminar debido a que ya a sido aplicado"
                    A% = MsgBox(m$, 0, "Baja de Ordenes de Pago")
                    Exit Sub
                End If
            End If
                    
                    
                    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.Clave = " + "'" + WIndice.List(Indice) + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
                    




                    txtOdbc = "Empresa"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Empresa"
                    ZSql = ZSql + " Where Empresa.Empresa = " + "'" + WEmpresa + "'"
                    spEmpresa = ZSql
                    Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEmpresa.RecordCount > 0 Then
                        WNombreBase = Trim(rstEmpresa!NombreBase)
                        WNombreEmpresa = Trim(rstEmpresa!Nombre)
                        WCtaProveedor = rstEmpresa!CtaProveedores
                        WCtaEfectivo = rstEmpresa!CtaEfectivo
                        WCtaCheques = rstEmpresa!CtaCheque
                        rstEmpresa.Close
                    End If
    
                    txtOdbc = WNombreBase
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)




    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    ZSql = ZSql + " Where IvaComp.Clave = " + "'" + WClave + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then





    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Parametro"
    ZSql = ZSql + " Where Parametro.Clave = 1"
    spParametro = ZSql
    Set rstParametro = db.OpenRecordset(spParametro, dbOpenSnapshot, dbSQLPassThrough)
    If rstParametro.RecordCount > 0 Then
        WMinimo1 = rstParametro!Minimo1
        WMinimo2 = rstParametro!Minimo2
        WMinimo3 = rstParametro!Minimo3
        WMinimo4 = rstParametro!Minimo4
        WEscala1 = rstParametro!Escala1
        WEscala2 = rstParametro!Escala2
        WEscala3 = rstParametro!Escala3
        WEscala4 = rstParametro!Escala4
        WEscala5 = rstParametro!Escala5
        XTasa1 = rstParametro!Tasa1
        XTasa2 = rstParametro!Tasa2
        XTasa3 = rstParametro!Tasa3
        XTasa4 = rstParametro!Tasa4
        XTasa5 = rstParametro!Tasa5
        WRetMinima = rstParametro!RetMinima
        WPorceBienes = rstParametro!PorceBienes / 100
        WPorceServicios = rstParametro!PorceServicios / 100
        WPorceTranspo = rstParametro!PorceTranspo / 100
        WMinimoIva = rstParametro!MinimoIva
        WIvaInscripto = rstParametro!IvaInscripto
        WIvaNoInscripto = rstParametro!IvaNoInscripto
        WTasaGen = rstParametro!TasaGen / 100
        WTasaBienes = rstParametro!TasaBienes / 100
        WTasaNoInscripto = rstParametro!TasaNoInscripto / 100
        rstParametro.Close
    End If




        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Depositos"
        ZSql = ZSql + " Where Depositos.Clave = " + "'" + WClave + "'"
        spDepositos = ZSql
        Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
        If rstDepositos.RecordCount > 0 Then









    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.razon = " + "'" + Clientes.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Clientes.Text = rstCliente!Cliente
        DesClientes.Caption = rstCliente!Razon
        WRazon = rstCliente!Razon
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WProv = rstCliente!Provincia
        rstCliente.Close
        Call Format_datos
    End If
    



    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Lineas"
    ZSql = ZSql + " Where Lineas.Linea = " + "'" + Linea.Text + "'"
    spLinea = ZSql
    Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstLinea.RecordCount > 0 Then
        Nombre.Text = Trim(rstLinea!Nombre)
        rstLinea.Close
        Call Format_datos
        Call Imprime_Nombre
    End If




    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Vendedor"
    ZSql = ZSql + " Where Vendedor.Codigo = " + "'" + Vendedor.Text + "'"
    spVendedor = ZSql
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstVendedor.RecordCount > 0 Then
        DesVendedor.Caption = !Nombre
        rstVendedor.Close
            Else
        DesVendedor.Caption = ""
    End If



    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Codigo.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
    End If

