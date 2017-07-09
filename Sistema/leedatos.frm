VERSION 5.00
Begin VB.Form PrgLeeDatos 
   Caption         =   "Traspaso de Datos"
   ClientHeight    =   3510
   ClientLeft      =   2805
   ClientTop       =   915
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3510
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Cancela 
         Caption         =   "Menu "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         MouseIcon       =   "leedatos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "leedatos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Salida"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Confirma "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   960
         MouseIcon       =   "leedatos.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "leedatos.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Confirma Proceso de Grabacion"
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "PrgLeeDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WClave As String
Dim WCodigo As String
Dim WRenglon As String
Dim WProveedor As String
Dim WConcepto As String
Dim WFecha As String
Dim WOrdFecha As String
Dim WNumero As String
Dim WNeto As String
Dim WIva As String
Dim WTotal As String
Dim WCamion As String
Dim WTanque As String
Dim WLitros As String
Dim WPrecio As String
Dim WFechaCarga As String
Dim WOrdFechaCarga As String
Dim WChofer As String
Dim WPunto As String

Dim ZZClaveChofer As String
Dim ZZChofer As String
Dim ZZLetra As String
Dim ZZTipo As String
Dim ZZPunto As String
Dim ZZNumero As String
Dim ZZfecha As String
Dim ZZVencimiento As String
Dim ZZTotal As String
Dim ZZSaldo As String
Dim ZZObservaciones As String
Dim ZZOrdFecha As String
Dim ZZOrdVencimiento As String
Dim ZZProveedor As String
Dim ZZCai As String
Dim ZZVtoCai As String
Dim ZZImpre As String

Dim WWNumero As String
Dim WWChofer As String
Dim WWNumeroII As String
Dim WWChoferII As String
Dim WWNumeroIII As String
Dim WWChoferIII As String

Dim ZZRecibo As String
Dim ZZRenglon As String
Dim ZZCliente As String
Dim ZZFechaOrd As String
Dim ZZTipoRec As String
Dim ZZRetGanancias As String
Dim ZZRetIva As String
Dim ZZRetOtra As String
Dim ZZRetSuss As String
Dim ZZNroRetganancias As String
Dim ZZNroRetIva As String
Dim ZZNroRetOtra As String
Dim ZZNroRetSuss As String
Dim ZZRetencion As String
Dim ZZTipoReg As String
Dim ZZTipo1 As String
Dim ZZLetra1 As String
Dim ZZPunto1 As String
Dim ZZNumero1 As String
Dim ZZImporte1 As String
Dim ZZTipo2 As String
Dim ZZNumero2 As String
Dim ZZFecha2 As String
Dim ZZFechaOrd2 As String
Dim ZZBanco2 As String
Dim ZZImporte2 As String
Dim ZZEstado2 As String
Dim ZZEmpresa As String
Dim ZZClave As String
Dim ZZImporte As String
Dim ZZCuenta As String
Dim ZZDestino As String
Dim ZZOrden As String
Dim ZZDeposito As String

Dim ZZRazon As String
Dim ZZDireccion As String
Dim ZZLocalidad As String
Dim ZZPostal As String
Dim ZZTelefono As String
Dim ZZCuit As String
Dim ZZEmail As String
Dim ZZFax As String
Dim ZZProvincia As String
Dim ZZIva As String
Dim ZZVendedor As String
Dim ZZDescuento As String
Dim ZZComision1 As String
Dim ZZComision2 As String


Dim ZZEstado As String
Dim ZZTipoFac As String


Dim ZSuma As Double
Dim ZPago As Double
Dim ZTotalAnte As Double
Dim ZSaldoAnte As Double

Dim ZZPrecio As Double
Dim ZZMargen As Double

Dim ZZDto(100) As String
Dim ZZAplica As String


Private Sub Acepta_Click()
    Call Proceso
    Call Cancela_Click
End Sub

Private Sub Cancela_Click()
    PrgLeeDatos.Hide
    Unload Me
    MenuAdminis.SetFocus
End Sub

Private Sub Proceso()




aa = WEmpresa





    GoTo da
    
    

    ZCodigo = 900001
    ZRenglon = 0


    Open "c:\datos\banmov.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WBanco = Mid$(WDato, 2, 3)
        WNumeroII = Mid$(WDato, 5, 6)
        WNumero = Mid$(WDato, 11, 8)
        WSucursal = Mid$(WDato, 19, 3)
        WStatus = Mid$(WDato, 22, 1)
        WFecha = Mid$(WDato, 23, 8)
        WVendedor = Mid$(WDato, 31, 2)
        WCliente = Mid$(WDato, 33, 1) + "-" + Mid$(WDato, 34, 3)
        WImporte = Mid$(WDato, 37, 12)
        WFactura = Mid$(WDato, 49, 10)
        WFechaII = Mid$(WDato, 59, 8)
        WFechaIII = Mid$(WDato, 67, 8)
        WDestino = Mid$(WDato, 75, 1)
        WReceptor = Mid$(WDato, 76, 4)
        WTipoCheque = Mid$(WDato, 80, 1)
        WClaseCheque = Mid$(WDato, 81, 1)
        
        If Val(WReceptor) = 0 Then
        
        If Val(WDestino) = 0 Then
            ZZEstado2 = "P"
                Else
            ZZEstado2 = "X"
        End If
        
        ZRenglon = ZRenglon + 1
        If ZRenglon > 99 Then
            ZCodigo = ZCodigo + 1
            ZRenglon = 1
        End If
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        
        ZZRecibo = Str(ZCodigo)
        ZZRenglon = Auxi1
        
        Call Ceros(ZZRecibo, 6)
        Call Ceros(ZZRenglon, 2)
        
        ZZCliente = WCliente
        ZZfecha = Right$(WFecha, 2) + "/" + Mid$(WFecha, 5, 2) + "/" + Left$(WFecha, 4)
        ZZFechaOrd = Right$(ZZfecha, 4) + Mid$(ZZfecha, 4, 2) + Left$(ZZfecha, 2)
        ZZTipoRec = "3"
        
        ZZRetGanancias = "0"
        ZZRetIva = "0"
        ZZRetOtra = "0"
        ZZRetSuss = "0"
        ZZNroRetganancias = "0"
        ZZNroRetIva = "0"
        ZZNroRetOtra = "0"
        ZZNroRetSuss = "0"
        ZZRetencion = "0"
        ZZTipoReg = "2"
        ZZTipo1 = ""
        ZZLetra1 = ""
        ZZPunto1 = ""
        ZZNumero1 = ""
        ZZImporte1 = "0"
        
        ZZTipo2 = "02"
        ZZNumero2 = WNumero
        ZZFecha2 = Right$(WFechaII, 2) + "/" + Mid$(WFechaII, 5, 2) + "/" + Left$(WFechaII, 4)
        ZZFechaOrd2 = Right$(ZZFecha2, 4) + Mid$(ZZFecha2, 4, 2) + Left$(ZZFecha2, 2)
        ZZPeriodo = Right$(ZZFecha2, 4) + Mid$(ZZFecha2, 4, 2)
        ZZBanco2 = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Bcra"
        ZSql = ZSql + " Where Bcra.Codigo = " + "'" + WBanco + "'"
        spBcra = ZSql
        Set rstBcra = db.OpenRecordset(spBcra, dbOpenSnapshot, dbSQLPassThrough)
        If rstBcra.RecordCount > 0 Then
            ZZBanco2 = Left$(rstBcra!Descripcion, 20)
            rstBcra.Close
        End If
        ZZImporte2 = WImporte
        ZZObservaciones = ""
        ZZEmpresa = "1"
        ZZClave = ZZRecibo + ZZRenglon
        ZZImporte = WImporte
        ZCuenta = "1"
        ZZDestino = ""
        ZZOrden = "0"
        ZZDeposito = "0"
        If ZZEstado2 = "X" Then
            ZZOrden = "1"
        End If
        
        
        ZZCodigoBanco = WBanco
        ZZSucursalCheque = WSucursal
        ZZTipoCheque = ""
        ZZClaseCheque = ""
        ZZProveedorSalida = "0"
        ZZBancoSalidaSalida = "0"
        ZZClaveLectora = ""
        
        If Val(WDestino) = 0 Then
        
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Recibos ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Recibo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "FechaOrd ,"
            ZSql = ZSql + "TipoRec ,"
            ZSql = ZSql + "RetGanancias ,"
            ZSql = ZSql + "RetIva ,"
            ZSql = ZSql + "RetOtra ,"
            ZSql = ZSql + "Retencion ,"
            ZSql = ZSql + "TipoReg ,"
            ZSql = ZSql + "Tipo1  ,"
            ZSql = ZSql + "Letra1 ,"
            ZSql = ZSql + "Punto1 ,"
            ZSql = ZSql + "Numero1 ,"
            ZSql = ZSql + "Importe1 ,"
            ZSql = ZSql + "Tipo2 ,"
            ZSql = ZSql + "Numero2 ,"
            ZSql = ZSql + "Fecha2 ,"
            ZSql = ZSql + "banco2 ,"
            ZSql = ZSql + "Importe2 ,"
            ZSql = ZSql + "Estado2 ,"
            ZSql = ZSql + "Empresa ,"
            ZSql = ZSql + "FechaOrd2 ,"
            ZSql = ZSql + "Periodo ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Impolist ,"
            ZSql = ZSql + "Impo1list ,"
            ZSql = ZSql + "Destino ,"
            ZSql = ZSql + "Cuenta ,"
            ZSql = ZSql + "Orden ,"
            ZSql = ZSql + "Deposito ,"
            ZSql = ZSql + "CodigoBanco ,"
            ZSql = ZSql + "SucursalCheque ,"
            ZSql = ZSql + "TipoCheque ,"
            ZSql = ZSql + "ProveedorSalida ,"
            ZSql = ZSql + "BancoSalida ,"
            ZSql = ZSql + "ClaseCheque ,"
            ZSql = ZSql + "ClaveLectora ,"
            ZSql = ZSql + "NroRetGanancias ,"
            ZSql = ZSql + "NroRetIva ,"
            ZSql = ZSql + "NroRetOtra ,"
            ZSql = ZSql + "RetSuss ,"
            ZSql = ZSql + "NroRetSuss )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZRecibo + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZFechaOrd + "',"
            ZSql = ZSql + "'" + ZZTipoRec + "',"
            ZSql = ZSql + "'" + ZZRetGanancias + "',"
            ZSql = ZSql + "'" + ZZRetIva + "',"
            ZSql = ZSql + "'" + ZZRetOtra + "',"
            ZSql = ZSql + "'" + ZZRetencion + "',"
            ZSql = ZSql + "'" + ZZTipoReg + "',"
            ZSql = ZSql + "'" + ZZTipo1 + "',"
            ZSql = ZSql + "'" + ZZLetra1 + "',"
            ZSql = ZSql + "'" + ZZPunto1 + "',"
            ZSql = ZSql + "'" + ZZNumero1 + "',"
            ZSql = ZSql + "'" + ZZImporte1 + "',"
            ZSql = ZSql + "'" + ZZTipo2 + "',"
            ZSql = ZSql + "'" + ZZNumero2 + "',"
            ZSql = ZSql + "'" + ZZFecha2 + "',"
            ZSql = ZSql + "'" + ZZBanco2 + "',"
            ZSql = ZSql + "'" + ZZImporte2 + "',"
            ZSql = ZSql + "'" + ZZEstado2 + "',"
            ZSql = ZSql + "'" + ZZEmpresa + "',"
            ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
            ZSql = ZSql + "'" + ZZPeriodo + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZImpoList + "',"
            ZSql = ZSql + "'" + ZZImpo1list + "',"
            ZSql = ZSql + "'" + ZZDestino + "',"
            ZSql = ZSql + "'" + ZZCuenta + "',"
            ZSql = ZSql + "'" + ZZOrden + "',"
            ZSql = ZSql + "'" + ZZDeposito + "',"
            ZSql = ZSql + "'" + ZZCodigoBanco + "',"
            ZSql = ZSql + "'" + ZZSucursalCheque + "',"
            ZSql = ZSql + "'" + ZZTipoCheque + "',"
            ZSql = ZSql + "'" + ZZProveedorSalida + "',"
            ZSql = ZSql + "'" + ZZBancoSalida + "',"
            ZSql = ZSql + "'" + ZZClaseCheque + "',"
            ZSql = ZSql + "'" + ZZClaveLectora + "',"
            ZSql = ZSql + "'" + ZZNroRetganancias + "',"
            ZSql = ZSql + "'" + ZZNroRetIva + "',"
            ZSql = ZSql + "'" + ZZNroRetOtra + "',"
            ZSql = ZSql + "'" + ZZRetSuss + "',"
            ZSql = ZSql + "'" + ZZNroRetSuss + "')"
                
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        End If
        
    Loop
    
    Close #1







Stop














    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
            WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
            
            WNombre = Mid$(WDato, 9, 30)
            WColor = Mid$(WDato, 39, 5)
            WFob = Mid$(WDato, 44, 7)
            WMinimoVenta = Mid$(WDato, 51, 3)
            WGrupo = Mid$(WDato, 54, 3)
            WProveedor = Mid$(WDato, 57, 4)
            WUnidad = Mid$(WDato, 60, 3)
            WCosto = Mid$(WDato, 104, 10)
            WStockMinimo = Mid$(WDato, 164, 5)
            WMargen = Mid$(WDato, 227, 5)
            WMargen = Str$(Val(WMargen) / 100)
            
            WCodigoBara = Mid$(WDato, 240, 13)
            
            
            WMargenFuturo = "0"
            WCif = "0"
            WCostoAnterior = "0"
            WFechaCostoAnterior = "  /  /    "
            WFechaCosto = "  /  /    "
            WOrdFechaCosto = ""
            WCostoFuturo = "0"
            
            WFechaCierre = "  /  /    "
            WFechaUltimaEntrada = "  /  /    "
            WFechaUltimaSalida = "  /  /    "
            WMinimo = "0"
            
            If WCodigo = "P00200" Then Stop
            
            
            WStockAnterior = Mid$(WDato, 169, 6)
            WEntradas = Mid$(WDato, 175, 6)
            WSalidas = Mid$(WDato, 181, 6)
            WStock = Str$(Val(WStockAnterior) + Val(WEntradas) - Val(WSalidas))
            WIva = "0"
            WVenta1 = "0"
            WVenta2 = "0"
            WVenta3 = "0"
            WVenta4 = "0"
            WVenta5 = "0"
            WVenta6 = "0"
            WPosicion = "0"
            WPosicionII = "0"
            WComision = "0"
            WDespacho = "0"
            WPrecio = "0"
            
            If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
                ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
                Call Redondeo(ZZMargen)
                ZZPrecio = Val(WCosto) + ZZMargen
            End If
            WPrecio = Str$(ZZPrecio)
            
            
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Articulo SET "
            Rem ZSql = ZSql + " StockAnterior = " + "'" + WStockAnterior + "',"
            Rem ZSql = ZSql + " Entradas = " + "'" + WEntradas + "',"
            Rem ZSql = ZSql + " Salidas = " + "'" + WSalidas + "',"
            Rem ZSql = ZSql + " Stock = " + "'" + WStock + "'"
            Rem ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            Rem spArticulo = ZSql
            Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Articulo SET "
            Rem ZSql = ZSql + " StockAnterior = " + "'" + WStock + "',"
            Rem ZSql = ZSql + " Entradas = " + "'" + "0" + "',"
            Rem ZSql = ZSql + " Salidas = " + "'" + "0" + "',"
            Rem ZSql = ZSql + " Stock = " + "'" + "0" + "'"
            Rem ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            Rem spArticulo = ZSql
            Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    Loop
    
    Close #1













Stop































    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " StockAnterior = " + "'" + "0" + "',"
    ZSql = ZSql + " Entradas = " + "'" + "0" + "',"
    ZSql = ZSql + " Salidas = " + "'" + "0" + "',"
    ZSql = ZSql + " Stock = " + "'" + "0" + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)













    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
            WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
            
            WNombre = Mid$(WDato, 9, 30)
            WColor = Mid$(WDato, 39, 5)
            WFob = Mid$(WDato, 44, 7)
            WMinimoVenta = Mid$(WDato, 51, 3)
            WGrupo = Mid$(WDato, 54, 3)
            WProveedor = Mid$(WDato, 57, 4)
            WUnidad = Mid$(WDato, 60, 3)
            WCosto = Mid$(WDato, 104, 10)
            WStockMinimo = Mid$(WDato, 164, 5)
            WMargen = Mid$(WDato, 227, 5)
            WMargen = Str$(Val(WMargen) / 100)
            
            WCodigoBara = Mid$(WDato, 240, 13)
            
            
            WMargenFuturo = "0"
            WCif = "0"
            WCostoAnterior = "0"
            WFechaCostoAnterior = "  /  /    "
            WFechaCosto = "  /  /    "
            WOrdFechaCosto = ""
            WCostoFuturo = "0"
            
            WFechaCierre = "  /  /    "
            WFechaUltimaEntrada = "  /  /    "
            WFechaUltimaSalida = "  /  /    "
            WMinimo = "0"
            
            Rem If WCodigo = "A00263" Then Stop
            
            
            WStockAnterior = Mid$(WDato, 169, 6)
            WEntradas = Mid$(WDato, 175, 6)
            WSalidas = Mid$(WDato, 181, 6)
            WStock = Str$(Val(WStockAnterior) + Val(WEntradas) - Val(WSalidas))
            WIva = "0"
            WVenta1 = "0"
            WVenta2 = "0"
            WVenta3 = "0"
            WVenta4 = "0"
            WVenta5 = "0"
            WVenta6 = "0"
            WPosicion = "0"
            WPosicionII = "0"
            WComision = "0"
            WDespacho = "0"
            WPrecio = "0"
            
            If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
                ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
                Call Redondeo(ZZMargen)
                ZZPrecio = Val(WCosto) + ZZMargen
            End If
            WPrecio = Str$(ZZPrecio)
            
            
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Articulo SET "
            Rem ZSql = ZSql + " StockAnterior = " + "'" + WStockAnterior + "',"
            Rem ZSql = ZSql + " Entradas = " + "'" + WEntradas + "',"
            Rem ZSql = ZSql + " Salidas = " + "'" + WSalidas + "',"
            Rem ZSql = ZSql + " Stock = " + "'" + WStock + "'"
            Rem ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            Rem spArticulo = ZSql
            Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " StockAnterior = " + "'" + WStock + "',"
            ZSql = ZSql + " Entradas = " + "'" + "0" + "',"
            ZSql = ZSql + " Salidas = " + "'" + "0" + "',"
            ZSql = ZSql + " Stock = " + "'" + "0" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    Loop
    
    Close #1






























Stop












    ZZDto(1) = "0"
    ZZDto(2) = "5"
    ZZDto(3) = "10"
    ZZDto(4) = "15"
    ZZDto(5) = "20"
    ZZDto(6) = "24"
    ZZDto(7) = "28"
    ZZDto(8) = "36"
    ZZDto(9) = "40"
    ZZDto(10) = "15"
    ZZDto(11) = "46"
    ZZDto(12) = "30"
    ZZDto(13) = "50"
    ZZDto(14) = "32"
    ZZDto(15) = "60"
    ZZDto(16) = "44"
    ZZDto(17) = "100"
    ZZDto(18) = "17"
    ZZDto(19) = "22"
    ZZDto(20) = "95"
    ZZDto(21) = "1"
    ZZDto(22) = "42"
    ZZDto(23) = "22"
    ZZDto(24) = "64"
    ZZDto(25) = "14"



    Open "c:\datos\VP200808.dat" For Input As #1
    
    Pasa = 0
    Renglon = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        
        WTipo = Mid$(WDato, 2, 1)
        WDocu = Mid$(WDato, 3, 1)
        WCliente = Mid$(WDato, 4, 4)
        WNumero = Mid$(WDato, 8, 5)
        WFecha = Mid$(WDato, 13, 6)
        WArticulo = Mid$(WDato, 19, 7)
        WCantidad = Mid$(WDato, 26, 6)
        WPrecio = Mid$(WDato, 32, 10)
        WVendedor = Mid$(WDato, 42, 2)
        WCondicion = Mid$(WDato, 44, 2)
        WPrograma = Mid$(WDato, 46, 1)
        WBonificacion = Mid$(WDato, 47, 2)
        WLista = Mid$(WDato, 49, 1)
        WProveedor = Mid$(WDato, 50, 3)
        WCosto = Mid$(WDato, 53, 6)
        WComision = Mid$(WDato, 59, 1)
        WIva = Mid$(WDato, 60, 1)
        
        
        Renglon = Renglon + 1
        If Renglon > 99 Then
            Renglon = 1
        End If
        
        Auxi = Str$(Renglon)
        Call Ceros(Auxi, 2)
                    
        Auxi1 = Str$(Numero.Text)
        Call Ceros(Auxi1, 8)
        
        ZZTipo = "01"
        ZZNumero = Numero.Text
        ZZRenglon = Renglon
        ZZArticulo = Articulo
        ZZDescripcion = DesArticulo
        ZZCantidad = Str$(Cantidad)
        ZZCantidadII = Str$(Cantidad)
        ZZPrecio = Str$(Precio)
        ZZPrecioUs = Str$(Precio)
        ZZImporte = Str$(Precio * Cantidad)
        ZZImporteUs = Str$(Precio * Cantidad)
        ZZCliente = Cliente.Text
        ZZParidad = "0"
        ZZVendedor = "0"
        ZZRubro = "0"
        ZZLinea = "0"
        ZZCosto1 = "0"
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
        ZZClave = "01" + Auxi1 + Auxi
        ZZWDate = Date$
        ZZClaveCtacte = Left$(ZZClave, 10) + "01"
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
        
        ZZPrecioII = ZZPrecio
        
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
        ZSql = ZSql + "'" + ZZClaveCtacte + "',"
        ZSql = ZSql + "'" + ZZImprefactura + "',"
        ZSql = ZSql + "'" + ZZNroFactura + "',"
        ZSql = ZSql + "'" + ZZDescuento + "',"
        ZSql = ZSql + "'" + ZZPartida + "')"
                        
        spEstadistica = ZSql
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        
        
    Loop
    
    Close #1

























Stop














    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
            WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
            
            WNombre = Mid$(WDato, 9, 30)
            WColor = Mid$(WDato, 39, 5)
            WFob = Mid$(WDato, 44, 7)
            WMinimoVenta = Mid$(WDato, 51, 3)
            WGrupo = Mid$(WDato, 54, 3)
            WProveedor = Mid$(WDato, 57, 4)
            WUnidad = Mid$(WDato, 60, 3)
            WCosto = Mid$(WDato, 104, 10)
            WStockMinimo = Mid$(WDato, 164, 5)
            WMargen = Mid$(WDato, 227, 5)
            WMargen = Str$(Val(WMargen) / 100)
            WCodigoBara = Mid$(WDato, 240, 13)
            WStockAnterior = Mid$(WDato, 169, 6)
            WEntradas = Mid$(WDato, 175, 6)
            WSalidas = Mid$(WDato, 181, 6)
            WStock = Str$(Val(WStockAnterior) + Val(WEntradas) - Val(WSalidas))
            WIva = "0"
            WVenta1 = "0"
            WVenta2 = "0"
            WVenta3 = "0"
            WVenta4 = "0"
            WVenta5 = "0"
            WVenta6 = "0"
            WPosicion = "0"
            WPosicionII = "0"
            WComision = "0"
            WDespacho = "0"
            WPrecio = "0"
            
            If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
                ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
                Call Redondeo(ZZMargen)
                ZZPrecio = Val(WCosto) + ZZMargen
            End If
            WPrecio = Str$(ZZPrecio)
            
            WPosicion = Mid$(WDato, 225, 2)
            WComision = Mid$(WDato, 235, 5)
            If Val(WComision) = 0 Then
                WComision = "10"
                    Else
                WComision = "5"
            End If
            
            
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Articulo SET "
            Rem ZSql = ZSql + " StockAnterior = " + "'" + WStockAnterior + "',"
            Rem ZSql = ZSql + " Entradas = " + "'" + WEntradas + "',"
            Rem ZSql = ZSql + " Salidas = " + "'" + WSalidas + "',"
            Rem ZSql = ZSql + " Stock = " + "'" + WStock + "'"
            Rem ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            Rem spArticulo = ZSql
            Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Comision = " + "'" + WComision + "',"
            ZSql = ZSql + " Posicion = " + "'" + WPosicion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    Loop
    
    Close #1

























Stop
















    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " StockAnterior = " + "'" + "0" + "',"
    ZSql = ZSql + " Entradas = " + "'" + "0" + "',"
    ZSql = ZSql + " Salidas = " + "'" + "0" + "',"
    ZSql = ZSql + " Stock = " + "'" + "0" + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)













    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
            WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
            
            WNombre = Mid$(WDato, 9, 30)
            WColor = Mid$(WDato, 39, 5)
            WFob = Mid$(WDato, 44, 7)
            WMinimoVenta = Mid$(WDato, 51, 3)
            WGrupo = Mid$(WDato, 54, 3)
            WProveedor = Mid$(WDato, 57, 4)
            WUnidad = Mid$(WDato, 60, 3)
            WCosto = Mid$(WDato, 104, 10)
            WStockMinimo = Mid$(WDato, 164, 5)
            WMargen = Mid$(WDato, 227, 5)
            WMargen = Str$(Val(WMargen) / 100)
            
            WCodigoBara = Mid$(WDato, 240, 13)
            
            
            WMargenFuturo = "0"
            WCif = "0"
            WCostoAnterior = "0"
            WFechaCostoAnterior = "  /  /    "
            WFechaCosto = "  /  /    "
            WOrdFechaCosto = ""
            WCostoFuturo = "0"
            
            WFechaCierre = "  /  /    "
            WFechaUltimaEntrada = "  /  /    "
            WFechaUltimaSalida = "  /  /    "
            WMinimo = "0"
            
            Rem If WCodigo = "A00263" Then Stop
            
            
            WStockAnterior = Mid$(WDato, 169, 6)
            WEntradas = Mid$(WDato, 175, 6)
            WSalidas = Mid$(WDato, 181, 6)
            WStock = Str$(Val(WStockAnterior) + Val(WEntradas) - Val(WSalidas))
            WIva = "0"
            WVenta1 = "0"
            WVenta2 = "0"
            WVenta3 = "0"
            WVenta4 = "0"
            WVenta5 = "0"
            WVenta6 = "0"
            WPosicion = "0"
            WPosicionII = "0"
            WComision = "0"
            WDespacho = "0"
            WPrecio = "0"
            
            If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
                ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
                Call Redondeo(ZZMargen)
                ZZPrecio = Val(WCosto) + ZZMargen
            End If
            WPrecio = Str$(ZZPrecio)
            
            
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Articulo SET "
            Rem ZSql = ZSql + " StockAnterior = " + "'" + WStockAnterior + "',"
            Rem ZSql = ZSql + " Entradas = " + "'" + WEntradas + "',"
            Rem ZSql = ZSql + " Salidas = " + "'" + WSalidas + "',"
            Rem ZSql = ZSql + " Stock = " + "'" + WStock + "'"
            Rem ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            Rem spArticulo = ZSql
            Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " StockAnterior = " + "'" + WStock + "',"
            ZSql = ZSql + " Entradas = " + "'" + "0" + "',"
            ZSql = ZSql + " Salidas = " + "'" + "0" + "',"
            ZSql = ZSql + " Stock = " + "'" + "0" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    Loop
    
    Close #1






























Stop








    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
            WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
            WDespacho = Mid$(WDato, 231, 3)
            
            WDespacho = Mid$(WDato, 232, 3)
            Waa = Mid$(WDato, 229, 10)
            Rem WDespacho = ""
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Despacho = " + "'" + WDespacho + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    Loop
    
    Close #1






























Stop








    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
            WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
            WDespacho = Mid$(WDato, 231, 3)
            
            WDespacho = Mid$(WDato, 232, 3)
            Waa = Mid$(WDato, 229, 10)
            Rem WDespacho = ""
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Despacho = " + "'" + WDespacho + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    Loop
    
    Close #1






Stop















Stop





    Open "c:\datos\vendespa.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 3)
        WDescripcion = Mid$(WDato, 6, 35)
        WNumero = Mid$(WDato, 41, 15)
        WOrigen = Mid$(WDato, 56, 10)
        WAduana = Mid$(WDato, 66, 10)
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Despacho ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Posicion ,"
        ZSql = ZSql + "Aduana ,"
        ZSql = ZSql + "Puerto ,"
        ZSql = ZSql + "Importador )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WNumero + "',"
        ZSql = ZSql + "'" + WDescripcion + "',"
        ZSql = ZSql + "'" + WOrigen + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + WAduana + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "')"
        spDespacho = ZSql
        Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1

















Stop








    Open "c:\datos\venhisto.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        Rem dada
        Rem dada
        Rem dada
        
        WCliente = Mid$(WDato, 2, 1) + "-" + Mid$(WDato, 3, 3)
        WFecha = Mid$(WDato, 12, 2) + "/" + Mid$(WDato, 10, 2) + "/" + Mid$(WDato, 6, 4)
        WDescripcion = Mid$(WDato, 14, 50)
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM HistorialCliente"
        ZSql = ZSql + " Where HistorialCliente.Cliente = " + "'" + WCliente + "'"
        ZSql = ZSql + " Order by Renglon desc"
        spHistorialCliente = ZSql
        Set rstHistorialCliente = db.OpenRecordset(spHistorialCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstHistorialCliente.RecordCount > 0 Then
            WRenglon = rstHistorialCliente!Renglon + 1
            rstHistorialCliente.MoveLast
                Else
            WRenglon = 1
        End If
        
        Auxi = Str$(WRenglon)
        Call Ceros(Auxi, 3)
        
        ZZClave = WCliente + Auxi
        ZZRenglon = Str$(WRenglon)
        ZZfecha = WFecha
        ZZObservaciones = WDescripcion
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO HistorialCliente ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Observaciones )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + WCliente + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "')"
        spHistorialCliente = ZSql
        Set rstHistorialCliente = db.OpenRecordset(spHistorialCliente, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1






Stop




















    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
            WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
            
            If WCodigo = "A00061" Then Stop
            
            WNombre = Mid$(WDato, 9, 30)
            WColor = Mid$(WDato, 39, 5)
            WFob = Mid$(WDato, 44, 7)
            WMinimoVenta = Mid$(WDato, 51, 3)
            WGrupo = Mid$(WDato, 54, 3)
            WProveedor = Mid$(WDato, 57, 4)
            WUnidad = Mid$(WDato, 60, 3)
            WCosto = Mid$(WDato, 104, 10)
            WStockMinimo = Mid$(WDato, 164, 5)
            WMargen = Mid$(WDato, 227, 5)
            WMargen = Str$(Val(WMargen) / 100)
            
            WCodigoBara = Mid$(WDato, 240, 13)
            
            
            WMargenFuturo = "0"
            WCif = "0"
            WCostoAnterior = "0"
            WFechaCostoAnterior = "  /  /    "
            WFechaCosto = "  /  /    "
            WOrdFechaCosto = ""
            WCostoFuturo = "0"
            
            WFechaCierre = "  /  /    "
            WFechaUltimaEntrada = "  /  /    "
            WFechaUltimaSalida = "  /  /    "
            WMinimo = "0"
            WStockAnterior = Mid$(WDato, 169, 6)
            WEntradas = Mid$(WDato, 175, 6)
            WSalidas = Mid$(WDato, 181, 6)
            WStock = Str$(Val(WStockAnterior) + Val(WEntradas) - Val(WSalidas))
            WIva = "0"
            WVenta1 = "0"
            WVenta2 = "0"
            WVenta3 = "0"
            WVenta4 = "0"
            WVenta5 = "0"
            WVenta6 = "0"
            WPosicion = "0"
            WPosicionII = "0"
            WComision = "0"
            WDespacho = "0"
            WPrecio = "0"
            
            If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
                ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
                Call Redondeo(ZZMargen)
                ZZPrecio = Val(WCosto) + ZZMargen
            End If
            WPrecio = Str$(ZZPrecio)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Stock = " + "'" + WStock + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
            
        
        End If
        
        
    Loop
    
    Close #1






Stop

















    
    

    ZCodigo = 900001
    ZRenglon = 0


    Open "c:\datos\banmov.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WBanco = Mid$(WDato, 2, 3)
        WNumeroII = Mid$(WDato, 5, 6)
        WNumero = Mid$(WDato, 11, 8)
        WSucursal = Mid$(WDato, 19, 3)
        WStatus = Mid$(WDato, 22, 1)
        WFecha = Mid$(WDato, 23, 8)
        WVendedor = Mid$(WDato, 31, 2)
        WCliente = Mid$(WDato, 33, 1) + "-" + Mid$(WDato, 34, 3)
        WImporte = Mid$(WDato, 37, 12)
        WFactura = Mid$(WDato, 49, 10)
        WFechaII = Mid$(WDato, 59, 8)
        WFechaIII = Mid$(WDato, 67, 8)
        WDestino = Mid$(WDato, 75, 1)
        WReceptor = Mid$(WDato, 76, 4)
        WTipoCheque = Mid$(WDato, 80, 1)
        WClaseCheque = Mid$(WDato, 81, 1)
        
        If Val(WDestino) = 0 Then
            ZZEstado2 = "P"
                Else
            ZZEstado2 = "X"
        End If
        
        ZRenglon = ZRenglon + 1
        If ZRenglon > 99 Then
            ZCodigo = ZCodigo + 1
            ZRenglon = 1
        End If
    
        Renglon = Renglon + 1
        Auxi1 = Str$(Renglon)
        Call Ceros(Auxi1, 2)
        
        ZZRecibo = Str(ZCodigo)
        ZZRenglon = Auxi1
        
        Call Ceros(ZZRecibo, 6)
        Call Ceros(ZZRenglon, 2)
        
        ZZCliente = WCliente
        ZZfecha = Right$(WFecha, 2) + "/" + Mid$(WFecha, 5, 2) + "/" + Left$(WFecha, 4)
        ZZFechaOrd = Right$(ZZfecha, 4) + Mid$(ZZfecha, 4, 2) + Left$(ZZfecha, 2)
        ZZTipoRec = "3"
        
        ZZRetGanancias = "0"
        ZZRetIva = "0"
        ZZRetOtra = "0"
        ZZRetSuss = "0"
        ZZNroRetganancias = "0"
        ZZNroRetIva = "0"
        ZZNroRetOtra = "0"
        ZZNroRetSuss = "0"
        ZZRetencion = "0"
        ZZTipoReg = "2"
        ZZTipo1 = ""
        ZZLetra1 = ""
        ZZPunto1 = ""
        ZZNumero1 = ""
        ZZImporte1 = "0"
        
        ZZTipo2 = "02"
        ZZNumero2 = WNumero
        ZZFecha2 = Right$(WFechaII, 2) + "/" + Mid$(WFechaII, 5, 2) + "/" + Left$(WFechaII, 4)
        ZZFechaOrd2 = Right$(ZZFecha2, 4) + Mid$(ZZFecha2, 4, 2) + Left$(ZZFecha2, 2)
        ZZBanco2 = ""
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Bcra"
        ZSql = ZSql + " Where Bcra.Codigo = " + "'" + WBanco + "'"
        spBcra = ZSql
        Set rstBcra = db.OpenRecordset(spBcra, dbOpenSnapshot, dbSQLPassThrough)
        If rstBcra.RecordCount > 0 Then
            ZZBanco2 = Left$(rstBcra!Descripcion, 20)
            rstBcra.Close
        End If
        ZZImporte2 = WImporte
        ZZObservaciones = ""
        ZZEmpresa = "1"
        ZZClave = ZZRecibo + ZZRenglon
        ZZImporte = WImporte
        ZCuenta = "1"
        ZZDestino = ""
        ZZOrden = "0"
        ZZDeposito = "0"
        If ZZEstado2 = "X" Then
            ZZOrden = "1"
        End If
        
        
        ZZCodigoBanco = WBanco
        ZZSucursalCheque = WSucursal
        ZZTipoCheque = ""
        ZZClaseCheque = ""
        ZZProveedorSalida = "0"
        ZZBancoSalidaSalida = "0"
        ZZClaveLectora = ""
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Recibos ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Recibo ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "FechaOrd ,"
        ZSql = ZSql + "TipoRec ,"
        ZSql = ZSql + "RetGanancias ,"
        ZSql = ZSql + "RetIva ,"
        ZSql = ZSql + "RetOtra ,"
        ZSql = ZSql + "Retencion ,"
        ZSql = ZSql + "TipoReg ,"
        ZSql = ZSql + "Tipo1  ,"
        ZSql = ZSql + "Letra1 ,"
        ZSql = ZSql + "Punto1 ,"
        ZSql = ZSql + "Numero1 ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Tipo2 ,"
        ZSql = ZSql + "Numero2 ,"
        ZSql = ZSql + "Fecha2 ,"
        ZSql = ZSql + "banco2 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Estado2 ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "FechaOrd2 ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Impolist ,"
        ZSql = ZSql + "Impo1list ,"
        ZSql = ZSql + "Destino ,"
        ZSql = ZSql + "Cuenta ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Deposito ,"
        ZSql = ZSql + "CodigoBanco ,"
        ZSql = ZSql + "SucursalCheque ,"
        ZSql = ZSql + "TipoCheque ,"
        ZSql = ZSql + "ProveedorSalida ,"
        ZSql = ZSql + "BancoSalida ,"
        ZSql = ZSql + "ClaseCheque ,"
        ZSql = ZSql + "ClaveLectora ,"
        ZSql = ZSql + "NroRetGanancias ,"
        ZSql = ZSql + "NroRetIva ,"
        ZSql = ZSql + "NroRetOtra ,"
        ZSql = ZSql + "RetSuss ,"
        ZSql = ZSql + "NroRetSuss )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + ZZRecibo + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZTipoRec + "',"
        ZSql = ZSql + "'" + ZZRetGanancias + "',"
        ZSql = ZSql + "'" + ZZRetIva + "',"
        ZSql = ZSql + "'" + ZZRetOtra + "',"
        ZSql = ZSql + "'" + ZZRetencion + "',"
        ZSql = ZSql + "'" + ZZTipoReg + "',"
        ZSql = ZSql + "'" + ZZTipo1 + "',"
        ZSql = ZSql + "'" + ZZLetra1 + "',"
        ZSql = ZSql + "'" + ZZPunto1 + "',"
        ZSql = ZSql + "'" + ZZNumero1 + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZTipo2 + "',"
        ZSql = ZSql + "'" + ZZNumero2 + "',"
        ZSql = ZSql + "'" + ZZFecha2 + "',"
        ZSql = ZSql + "'" + ZZBanco2 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZEstado2 + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZFechaOrd2 + "',"
        ZSql = ZSql + "'" + ZZImporte + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "',"
        ZSql = ZSql + "'" + ZZImpoList + "',"
        ZSql = ZSql + "'" + ZZImpo1list + "',"
        ZSql = ZSql + "'" + ZZDestino + "',"
        ZSql = ZSql + "'" + ZZCuenta + "',"
        ZSql = ZSql + "'" + ZZOrden + "',"
        ZSql = ZSql + "'" + ZZDeposito + "',"
        ZSql = ZSql + "'" + ZZCodigoBanco + "',"
        ZSql = ZSql + "'" + ZZSucursalCheque + "',"
        ZSql = ZSql + "'" + ZZTipoCheque + "',"
        ZSql = ZSql + "'" + ZZProveedorSalida + "',"
        ZSql = ZSql + "'" + ZZBancoSalida + "',"
        ZSql = ZSql + "'" + ZZClaseCheque + "',"
        ZSql = ZSql + "'" + ZZClaveLectora + "',"
        ZSql = ZSql + "'" + ZZNroRetganancias + "',"
        ZSql = ZSql + "'" + ZZNroRetIva + "',"
        ZSql = ZSql + "'" + ZZNroRetOtra + "',"
        ZSql = ZSql + "'" + ZZRetSuss + "',"
        ZSql = ZSql + "'" + ZZNroRetSuss + "')"
            
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1







    ZCodigo = 0

    Open "c:\datos\banche.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        ZCodigo = ZCodigo + 1
        WCodigo = Str$(ZCodigo)
            
        WBanco = Mid$(WDato, 2, 2)
        WNumero = Mid$(WDato, 4, 8)
        ZFecha = Mid$(WDato, 12, 8)
        WFecha = Right$(ZFecha, 2) + "/" + Mid$(ZFecha, 5, 2) + "/" + Left$(ZFecha, 4)
        WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WImporte = Mid$(WDato, 20, 12)
        WObservaciones = Mid$(WDato, 32, 30)
        WTipo = Mid$(WDato, 62, 1)
        
        If Val(WTipo) = 1 Or Val(WTipo) = 3 Then
                WTipoMOvimiento = "1"
                Else
            WTipoMOvimiento = "0"
        End If
        
        Rem WObservaciones = "MARVAL"
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO GastosBancarios ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Banco ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Comprobante ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "TipoMovimiento )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WFecha + "',"
        ZSql = ZSql + "'" + WOrdFecha + "',"
        ZSql = ZSql + "'" + WBanco + "',"
        ZSql = ZSql + "'" + WImporte + "',"
        ZSql = ZSql + "'" + WNumero + "',"
        ZSql = ZSql + "'" + WObservaciones + "',"
        ZSql = ZSql + "'" + WTipoMOvimiento + "')"
                                
        spGastosBancarios = ZSql
        Set rstGastosBancarios = db.OpenRecordset(spGastosBancarios, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1




























Stop



aa = WEmpresa


    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Mid$(WDato, 7, 2) = "00" Then
        
        WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
        WNombre = Mid$(WDato, 9, 30)
        WColor = Mid$(WDato, 39, 5)
        WFob = Mid$(WDato, 44, 7)
        WMinimoVenta = Mid$(WDato, 51, 3)
        WGrupo = Mid$(WDato, 54, 3)
        WProveedor = Mid$(WDato, 57, 4)
        WUnidad = Mid$(WDato, 60, 3)
        WCosto = Mid$(WDato, 104, 10)
        WStockMinimo = Mid$(WDato, 164, 5)
        WMargen = Mid$(WDato, 227, 5)
        WMargen = Str$(Val(WMargen) / 100)
        
        WMargenFuturo = "0"
        WCif = "0"
        WCostoAnterior = "0"
        WFechaCostoAnterior = "  /  /    "
        WFechaCosto = "  /  /    "
        WOrdFechaCosto = ""
        WCostoFuturo = "0"
        
        WFechaCierre = "  /  /    "
        WFechaUltimaEntrada = "  /  /    "
        WFechaUltimaSalida = "  /  /    "
        WMinimo = "0"
        WStockAnterior = Mid$(WDato, 169, 6)
        WEntradas = Mid$(WDato, 175, 6)
        WSalidas = Mid$(WDato, 181, 6)
        WStock = Str$(Val(WStockAnterior) + Val(WEntradas) - Val(WSalidas))
        WIva = "0"
        WVenta1 = "0"
        WVenta2 = "0"
        WVenta3 = "0"
        WVenta4 = "0"
        WVenta5 = "0"
        WVenta6 = "0"
        WPosicion = "0"
        WPosicionII = "0"
        WComision = "0"
        WDespacho = "0"
        WCodigoBarra = "0"
        WPrecio = "0"
        
        If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
            ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
            Call Redondeo(ZZMargen)
            ZZPrecio = Val(WCosto) + ZZMargen
        End If
        WPrecio = Str$(ZZPrecio)
        
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Articulo ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Color ,"
        ZSql = ZSql + "Grupo ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "MinimoVenta ,"
        ZSql = ZSql + "UnidadCaja ,"
        ZSql = ZSql + "Margen ,"
        ZSql = ZSql + "MargenFuturo ,"
        ZSql = ZSql + "Fob ,"
        ZSql = ZSql + "Cif ,"
        ZSql = ZSql + "CostoAnterior ,"
        ZSql = ZSql + "FechaCostoAnterior ,"
        ZSql = ZSql + "Costo ,"
        ZSql = ZSql + "FechaCosto ,"
        ZSql = ZSql + "OrdFechaCosto ,"
        ZSql = ZSql + "CostoFuturo ,"
        ZSql = ZSql + "FechaCierre ,"
        ZSql = ZSql + "FechaUltimaEntrada ,"
        ZSql = ZSql + "FechaUltimaSalida ,"
        ZSql = ZSql + "Minimo ,"
        ZSql = ZSql + "Entradas ,"
        ZSql = ZSql + "Salidas ,"
        ZSql = ZSql + "Stock ,"
        ZSql = ZSql + "StockAnterior ,"
        ZSql = ZSql + "Iva ,"
        ZSql = ZSql + "Venta1 ,"
        ZSql = ZSql + "Venta2 ,"
        ZSql = ZSql + "Venta3 ,"
        ZSql = ZSql + "Venta4 ,"
        ZSql = ZSql + "Venta5 ,"
        ZSql = ZSql + "Venta6 ,"
        ZSql = ZSql + "Posicion ,"
        ZSql = ZSql + "PosicionII ,"
        ZSql = ZSql + "Comision ,"
        ZSql = ZSql + "Despacho ,"
        ZSql = ZSql + "CodigoBarra ,"
        ZSql = ZSql + "Precio )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WNombre + "',"
        ZSql = ZSql + "'" + WColor + "',"
        ZSql = ZSql + "'" + WGrupo + "',"
        ZSql = ZSql + "'" + WProveedor + "',"
        ZSql = ZSql + "'" + WMinimoVenta + "',"
        ZSql = ZSql + "'" + WUnidad + "',"
        ZSql = ZSql + "'" + WMargen + "',"
        ZSql = ZSql + "'" + WMargenFuturo + "',"
        ZSql = ZSql + "'" + WFob + "',"
        ZSql = ZSql + "'" + WCif + "',"
        ZSql = ZSql + "'" + WCostoAnterior + "',"
        ZSql = ZSql + "'" + WFechaCostoAnterior + "',"
        ZSql = ZSql + "'" + WCosto + "',"
        ZSql = ZSql + "'" + WFechaCosto + "',"
        ZSql = ZSql + "'" + WOrdFechaCosto + "',"
        ZSql = ZSql + "'" + WCostoFuturo + "',"
        ZSql = ZSql + "'" + WFechaCierre + "',"
        ZSql = ZSql + "'" + WFechaUltimaEntrada + "',"
        ZSql = ZSql + "'" + WFechaUltimaSalida + "',"
        ZSql = ZSql + "'" + WMinimo + "',"
        ZSql = ZSql + "'" + WEntradas + "',"
        ZSql = ZSql + "'" + WSalidas + "',"
        ZSql = ZSql + "'" + WStock + "',"
        ZSql = ZSql + "'" + WStockAnterior + "',"
        ZSql = ZSql + "'" + WIva + "',"
        ZSql = ZSql + "'" + WVenta1 + "',"
        ZSql = ZSql + "'" + WVenta2 + "',"
        ZSql = ZSql + "'" + WVenta3 + "',"
        ZSql = ZSql + "'" + WVenta4 + "',"
        ZSql = ZSql + "'" + WVenta5 + "',"
        ZSql = ZSql + "'" + WVenta6 + "',"
        ZSql = ZSql + "'" + WPosicion + "',"
        ZSql = ZSql + "'" + WPosicionII + "',"
        ZSql = ZSql + "'" + WComision + "',"
        ZSql = ZSql + "'" + WDespacho + "',"
        ZSql = ZSql + "'" + WCodigoBarra + "',"
        ZSql = ZSql + "'" + WPrecio + "')"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        
    Loop
    
    Close #1






Stop


























    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
        WNombre = Mid$(WDato, 9, 30)
        WColor = Mid$(WDato, 39, 5)
        WFob = Mid$(WDato, 44, 7)
        WMinimoVenta = Mid$(WDato, 51, 3)
        WGrupo = Mid$(WDato, 54, 3)
        WProveedor = Mid$(WDato, 57, 4)
        WUnidad = Mid$(WDato, 60, 3)
        WCosto = Mid$(WDato, 104, 10)
        WStockMinimo = Mid$(WDato, 164, 5)
        WMargen = Mid$(WDato, 227, 5)
        WMargen = Str$(Val(WMargen) / 100)
        
        WMargenFuturo = "0"
        WCif = "0"
        WCostoAnterior = "0"
        WFechaCostoAnterior = "  /  /    "
        WFechaCosto = "  /  /    "
        WOrdFechaCosto = ""
        WCostoFuturo = "0"
        
        WFechaCierre = "  /  /    "
        WFechaUltimaEntrada = "  /  /    "
        WFechaUltimaSalida = "  /  /    "
        WMinimo = "0"
        WStockAnterior = Mid$(WDato, 169, 6)
        WEntradas = Mid$(WDato, 175, 6)
        WSalidas = Mid$(WDato, 181, 6)
        WStock = Str$(Val(WStockAnterior) + Val(WEntradas) + Val(WSalidas))
        WIva = "0"
        WVenta1 = "0"
        WVenta2 = "0"
        WVenta3 = "0"
        WVenta4 = "0"
        WVenta5 = "0"
        WVenta6 = "0"
        WPosicion = "0"
        WPosicionII = "0"
        WComision = "0"
        WDespacho = "0"
        WCodigoBarra = "0"
        WPrecio = "0"
        
        If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
            ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
            Call Redondeo(ZZMargen)
            ZZPrecio = Val(WCosto) + ZZMargen
        End If
        WPrecio = Str$(ZZPrecio)
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Articulo ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Color ,"
        ZSql = ZSql + "Grupo ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "MinimoVenta ,"
        ZSql = ZSql + "UnidadCaja ,"
        ZSql = ZSql + "Margen ,"
        ZSql = ZSql + "MargenFuturo ,"
        ZSql = ZSql + "Fob ,"
        ZSql = ZSql + "Cif ,"
        ZSql = ZSql + "CostoAnterior ,"
        ZSql = ZSql + "FechaCostoAnterior ,"
        ZSql = ZSql + "Costo ,"
        ZSql = ZSql + "FechaCosto ,"
        ZSql = ZSql + "OrdFechaCosto ,"
        ZSql = ZSql + "CostoFuturo ,"
        ZSql = ZSql + "FechaCierre ,"
        ZSql = ZSql + "FechaUltimaEntrada ,"
        ZSql = ZSql + "FechaUltimaSalida ,"
        ZSql = ZSql + "Minimo ,"
        ZSql = ZSql + "Entradas ,"
        ZSql = ZSql + "Salidas ,"
        ZSql = ZSql + "Stock ,"
        ZSql = ZSql + "StockAnterior ,"
        ZSql = ZSql + "Iva ,"
        ZSql = ZSql + "Venta1 ,"
        ZSql = ZSql + "Venta2 ,"
        ZSql = ZSql + "Venta3 ,"
        ZSql = ZSql + "Venta4 ,"
        ZSql = ZSql + "Venta5 ,"
        ZSql = ZSql + "Venta6 ,"
        ZSql = ZSql + "Posicion ,"
        ZSql = ZSql + "PosicionII ,"
        ZSql = ZSql + "Comision ,"
        ZSql = ZSql + "Despacho ,"
        ZSql = ZSql + "CodigoBarra ,"
        ZSql = ZSql + "Precio )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WNombre + "',"
        ZSql = ZSql + "'" + WColor + "',"
        ZSql = ZSql + "'" + WGrupo + "',"
        ZSql = ZSql + "'" + WProveedor + "',"
        ZSql = ZSql + "'" + WMinimoVenta + "',"
        ZSql = ZSql + "'" + WUnidad + "',"
        ZSql = ZSql + "'" + WMargen + "',"
        ZSql = ZSql + "'" + WMargenFuturo + "',"
        ZSql = ZSql + "'" + WFob + "',"
        ZSql = ZSql + "'" + WCif + "',"
        ZSql = ZSql + "'" + WCostoAnterior + "',"
        ZSql = ZSql + "'" + WFechaCostoAnterior + "',"
        ZSql = ZSql + "'" + WCosto + "',"
        ZSql = ZSql + "'" + WFechaCosto + "',"
        ZSql = ZSql + "'" + WOrdFechaCosto + "',"
        ZSql = ZSql + "'" + WCostoFuturo + "',"
        ZSql = ZSql + "'" + WFechaCierre + "',"
        ZSql = ZSql + "'" + WFechaUltimaEntrada + "',"
        ZSql = ZSql + "'" + WFechaUltimaSalida + "',"
        ZSql = ZSql + "'" + WMinimo + "',"
        ZSql = ZSql + "'" + WEntradas + "',"
        ZSql = ZSql + "'" + WSalidas + "',"
        ZSql = ZSql + "'" + WStock + "',"
        ZSql = ZSql + "'" + WStockAnterior + "',"
        ZSql = ZSql + "'" + WIva + "',"
        ZSql = ZSql + "'" + WVenta1 + "',"
        ZSql = ZSql + "'" + WVenta2 + "',"
        ZSql = ZSql + "'" + WVenta3 + "',"
        ZSql = ZSql + "'" + WVenta4 + "',"
        ZSql = ZSql + "'" + WVenta5 + "',"
        ZSql = ZSql + "'" + WVenta6 + "',"
        ZSql = ZSql + "'" + WPosicion + "',"
        ZSql = ZSql + "'" + WPosicionII + "',"
        ZSql = ZSql + "'" + WComision + "',"
        ZSql = ZSql + "'" + WDespacho + "',"
        ZSql = ZSql + "'" + WCodigoBarra + "',"
        ZSql = ZSql + "'" + WPrecio + "')"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1









    Open "c:\datos\venconpa.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 2)
        WNombre = Mid$(WDato, 5, 40)
        WDias = Mid$(WDato, 45, 2)
        WObservaciones = ""
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO CondPago ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Dias )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WNombre + "',"
        ZSql = ZSql + "'" + WObservaciones + "',"
        ZSql = ZSql + "'" + WDias + "')"
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1






Stop



    Open "c:\datos\vendespa.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 4)
        WDescripcion = Mid$(WDato, 6, 35)
        WNumero = Mid$(WDato, 41, 15)
        WOrigen = Mid$(WDato, 56, 10)
        WAduana = Mid$(WDato, 66, 10)
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Despacho ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Origen ,"
        ZSql = ZSql + "Posicion ,"
        ZSql = ZSql + "Aduana ,"
        ZSql = ZSql + "Puerto ,"
        ZSql = ZSql + "Importador )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WNumero + "',"
        ZSql = ZSql + "'" + WDescripcion + "',"
        ZSql = ZSql + "'" + WOrigen + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + WAduana + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "')"
        spDespacho = ZSql
        Set rstDespacho = db.OpenRecordset(spDespacho, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1










    Open "c:\datos\venpro.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 1) + "0" + Mid$(WDato, 3, 4)
        WNombre = Mid$(WDato, 9, 30)
        WColor = Mid$(WDato, 39, 5)
        WFob = Mid$(WDato, 44, 7)
        WMinimoVenta = Mid$(WDato, 51, 3)
        WGrupo = Mid$(WDato, 54, 3)
        WProveedor = Mid$(WDato, 57, 4)
        WUnidad = Mid$(WDato, 60, 3)
        WCosto = Mid$(WDato, 104, 10)
        WStockMinimo = Mid$(WDato, 164, 5)
        WMargen = Mid$(WDato, 227, 5)
        WMargen = Str$(Val(WMargen) / 100)
        
        WMargenFuturo = "0"
        WCif = "0"
        WCostoAnterior = "0"
        WFechaCostoAnterior = "  /  /    "
        WFechaCosto = "  /  /    "
        WOrdFechaCosto = ""
        WCostoFuturo = "0"
        
        WFechaCierre = "  /  /    "
        WFechaUltimaEntrada = "  /  /    "
        WFechaUltimaSalida = "  /  /    "
        WMinimo = "0"
        WEntradas = "0"
        WSalidas = ""
        WStock = "0"
        WStockAnterior = "0"
        WIva = "0"
        WVenta1 = "0"
        WVenta2 = "0"
        WVenta3 = "0"
        WVenta4 = "0"
        WVenta5 = "0"
        WVenta6 = "0"
        WPosicion = "0"
        WPosicionII = "0"
        WComision = "0"
        WDespacho = "0"
        WCodigoBarra = "0"
        WPrecio = "0"
        
        If Val(WCosto) <> 0 And Val(WMargen) <> 0 Then
            ZZMargen = Val(WCosto) * (Val(WMargen) / 100)
            Call Redondeo(ZZMargen)
            ZZPrecio = Val(WCosto) + ZZMargen
        End If
        WPrecio = Str$(ZZPrecio)
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Articulo ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Color ,"
        ZSql = ZSql + "Grupo ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "MinimoVenta ,"
        ZSql = ZSql + "UnidadCaja ,"
        ZSql = ZSql + "Margen ,"
        ZSql = ZSql + "MargenFuturo ,"
        ZSql = ZSql + "Fob ,"
        ZSql = ZSql + "Cif ,"
        ZSql = ZSql + "CostoAnterior ,"
        ZSql = ZSql + "FechaCostoAnterior ,"
        ZSql = ZSql + "Costo ,"
        ZSql = ZSql + "FechaCosto ,"
        ZSql = ZSql + "OrdFechaCosto ,"
        ZSql = ZSql + "CostoFuturo ,"
        ZSql = ZSql + "FechaCierre ,"
        ZSql = ZSql + "FechaUltimaEntrada ,"
        ZSql = ZSql + "FechaUltimaSalida ,"
        ZSql = ZSql + "Minimo ,"
        ZSql = ZSql + "Entradas ,"
        ZSql = ZSql + "Salidas ,"
        ZSql = ZSql + "Stock ,"
        ZSql = ZSql + "StockAnterior ,"
        ZSql = ZSql + "Iva ,"
        ZSql = ZSql + "Venta1 ,"
        ZSql = ZSql + "Venta2 ,"
        ZSql = ZSql + "Venta3 ,"
        ZSql = ZSql + "Venta4 ,"
        ZSql = ZSql + "Venta5 ,"
        ZSql = ZSql + "Venta6 ,"
        ZSql = ZSql + "Posicion ,"
        ZSql = ZSql + "PosicionII ,"
        ZSql = ZSql + "Comision ,"
        ZSql = ZSql + "Despacho ,"
        ZSql = ZSql + "CodigoBarra ,"
        ZSql = ZSql + "Precio )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WNombre + "',"
        ZSql = ZSql + "'" + WColor + "',"
        ZSql = ZSql + "'" + WGrupo + "',"
        ZSql = ZSql + "'" + WProveedor + "',"
        ZSql = ZSql + "'" + WMinimoVenta + "',"
        ZSql = ZSql + "'" + WUnidad + "',"
        ZSql = ZSql + "'" + WMargen + "',"
        ZSql = ZSql + "'" + WMargenFuturo + "',"
        ZSql = ZSql + "'" + WFob + "',"
        ZSql = ZSql + "'" + WCif + "',"
        ZSql = ZSql + "'" + WCostoAnterior + "',"
        ZSql = ZSql + "'" + WFechaCostoAnterior + "',"
        ZSql = ZSql + "'" + WCosto + "',"
        ZSql = ZSql + "'" + WFechaCosto + "',"
        ZSql = ZSql + "'" + WOrdFechaCosto + "',"
        ZSql = ZSql + "'" + WCostoFuturo + "',"
        ZSql = ZSql + "'" + WFechaCierre + "',"
        ZSql = ZSql + "'" + WFechaUltimaEntrada + "',"
        ZSql = ZSql + "'" + WFechaUltimaSalida + "',"
        ZSql = ZSql + "'" + WMinimo + "',"
        ZSql = ZSql + "'" + WEntradas + "',"
        ZSql = ZSql + "'" + WSalidas + "',"
        ZSql = ZSql + "'" + WStock + "',"
        ZSql = ZSql + "'" + WStockAnterior + "',"
        ZSql = ZSql + "'" + WIva + "',"
        ZSql = ZSql + "'" + WVenta1 + "',"
        ZSql = ZSql + "'" + WVenta2 + "',"
        ZSql = ZSql + "'" + WVenta3 + "',"
        ZSql = ZSql + "'" + WVenta4 + "',"
        ZSql = ZSql + "'" + WVenta5 + "',"
        ZSql = ZSql + "'" + WVenta6 + "',"
        ZSql = ZSql + "'" + WPosicion + "',"
        ZSql = ZSql + "'" + WPosicionII + "',"
        ZSql = ZSql + "'" + WComision + "',"
        ZSql = ZSql + "'" + WDespacho + "',"
        ZSql = ZSql + "'" + WCodigoBarra + "',"
        ZSql = ZSql + "'" + WPrecio + "')"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1




    
    Open "c:\datos\vencli.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        If Pasa = 0 Then
            Pasa = 1
            Line Input #1, WDato
        End If
        
        WCliente = Mid$(WDato, 2, 1) + "-" + Mid$(WDato, 3, 3)
        WRazon = Mid$(WDato, 6, 30)
        WDireccion = Mid$(WDato, 36, 25)
        WLocalidad = Mid$(WDato, 61, 15)
        WPostal = Mid$(WDato, 76, 4)
        WTelefono = Trim(Mid$(WDato, 82, 55))
        WCuit = Mid$(WDato, 138, 15)
        WIb = Mid$(WDato, 153, 12)
        WExpreso = Mid$(WDato, 281, 3)
        WVendedor = Mid$(WDato, 168, 2)
        WCondicion = Mid$(WDato, 165, 2)
        WNroLista = Mid$(WDato, 280, 1)
        WZona = Mid$(WDato, 280, 1)
        WObservaciones = ""
        WDescuento = "20"
        WUltimaCompra = Mid$(WDato, 191, 2) + "/" + Mid$(WDato, 190, 2) + "/" + Mid$(WDato, 186, 4)
        WUltimaLista = Mid$(WDato, 274, 2) + "/" + Mid$(WDato, 272, 2) + "/" + Mid$(WDato, 268, 4)
        WOrdUltimaCompra = ""
        WOrdUltimaLista = ""
        WEmail = ""
        WFax = ""
        WPartida = ""
        WMarca = "0"
        
        WTipIva = Mid$(WDato, 137, 1)
        If Val(WTipoiva) = 4 Then
            WIva = "5"
                Else
            WIva = "1"
        End If
        
        WProvincia = Mid$(WDato, 80, 2)
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Cliente ("
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Telefono ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "Email ,"
        ZSql = ZSql + "Fax ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "Marca ,"
        ZSql = ZSql + "Iva ,"
        ZSql = ZSql + "Expreso ,"
        ZSql = ZSql + "Vendedor ,"
        ZSql = ZSql + "Descuento ,"
        ZSql = ZSql + "UltimaCompra ,"
        ZSql = ZSql + "OrdUltimaCompra ,"
        ZSql = ZSql + "UltimaLista ,"
        ZSql = ZSql + "OrdUltimaLista ,"
        ZSql = ZSql + "Zona ,"
        ZSql = ZSql + "NroLista ,"
        ZSql = ZSql + "Condicion ,"
        ZSql = ZSql + "Partida )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCliente + "',"
        ZSql = ZSql + "'" + WRazon + "',"
        ZSql = ZSql + "'" + WDireccion + "',"
        ZSql = ZSql + "'" + WLocalidad + "',"
        ZSql = ZSql + "'" + WPostal + "',"
        ZSql = ZSql + "'" + WTelefono + "',"
        ZSql = ZSql + "'" + WObservaciones + "',"
        ZSql = ZSql + "'" + WCuit + "',"
        ZSql = ZSql + "'" + WEmail + "',"
        ZSql = ZSql + "'" + WFax + "',"
        ZSql = ZSql + "'" + WProvincia + "',"
        ZSql = ZSql + "'" + WMarca + "',"
        ZSql = ZSql + "'" + WIva + "',"
        ZSql = ZSql + "'" + WExpreso + "',"
        ZSql = ZSql + "'" + WVendedor + "',"
        ZSql = ZSql + "'" + WDescuento + "',"
        ZSql = ZSql + "'" + WUltimaCompra + "',"
        ZSql = ZSql + "'" + WOrdUltimaCompra + "',"
        ZSql = ZSql + "'" + WUltimaLista + "',"
        ZSql = ZSql + "'" + WOrdUltimaLista + "',"
        ZSql = ZSql + "'" + WZona + "',"
        ZSql = ZSql + "'" + WNroLista + "',"
        ZSql = ZSql + "'" + WCondicion + "',"
        ZSql = ZSql + "'" + WPartida + "')"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1
    




    Linea% = 0

    ZZDto(1) = "0"
    ZZDto(2) = "5"
    ZZDto(3) = "10"
    ZZDto(4) = "15"
    ZZDto(5) = "20"
    ZZDto(6) = "24"
    ZZDto(7) = "28"
    ZZDto(8) = "36"
    ZZDto(9) = "40"
    ZZDto(10) = "15"
    ZZDto(11) = "46"
    ZZDto(12) = "30"
    ZZDto(13) = "50"
    ZZDto(14) = "32"
    ZZDto(15) = "60"
    ZZDto(16) = "44"
    ZZDto(17) = "100"
    ZZDto(18) = "17"
    ZZDto(19) = "22"
    ZZDto(20) = "95"
    ZZDto(21) = "1"
    ZZDto(22) = "42"
    ZZDto(23) = "22"
    ZZDto(24) = "64"
    ZZDto(25) = "14"

    Open "c:\datos\venMOV.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        Linea% = Linea% + 1
        
        WTipo = Mid$(WDato, 2, 1)
        WDocu = Mid$(WDato, 3, 1)
        WNumero = Mid$(WDato, 4, 8)
        WOrdFecha = Mid$(WDato, 12, 8)
        WFecha = Right$(WOrdFecha, 2) + "/" + Mid$(WOrdFecha, 5, 2) + "/" + Left$(WOrdFecha, 4)
        WOrdvencimiento = Mid$(WDato, 20, 8)
        WVencimiento = Right$(WOrdvencimiento, 2) + "/" + Mid$(WOrdvencimiento, 5, 2) + "/" + Left$(WOrdvencimiento, 4)
        WFechaII = Mid$(WDato, 28, 8)
        WImporte = Mid$(WDato, 36, 12)
        WIva = Mid$(WDato, 48, 12)
        WIvaII = Mid$(WDato, 60, 12)
        WImpInt = Mid$(WDato, 72, 12)
        WCliente = Mid$(WDato, 84, 1) + "-" + Mid$(WDato, 85, 3)
        WBonificacion = Mid$(WDato, 88, 2)
        WCondicion = Mid$(WDato, 90, 2)
        WVendedor = Mid$(WDato, 92, 2)
        WLista = Mid$(WDato, 94, 1)
        WUsaRemito = Mid$(WDato, 95, 2)
        WRutinas = Mid$(WDato, 97, 1)
        WAplica = Mid$(WDato, 98, 5)
        WExento = Mid$(WDato, 103, 12)
        WPuntero = Mid$(WDato, 115, 5)
        WComision = Mid$(WDato, 120, 12)
        
        If Val(WDocu) = 2 Or Val(WDocu) = 4 Or Val(WDocu) = 6 Then
            WImporte = Str$(Val(WImporte) * -1)
            WIva = Str$(Val(WIva) * -1)
            WIvaII = Str$(Val(WIvaII) * -1)
            WExento = Str$(Val(WExento) * -1)
        End If
        
        WTotal = Str$(Val(WImporte) + Val(WIva) + Val(WIvaII) + Val(WExento))
        
        
        If Val(WTotal) <> 0 Then
        
    
        
            Auxi = WNumero
            Call Ceros(Auxi, 8)
                    
            WPunto = "1"
            Call Ceros(WPunto, 4)
            
            Select Case WDocu
                Case "1"
                    ZZTipo = "01"
                    ZZImpre = "FC"
                Case "2"
                    ZZTipo = "02"
                    ZZImpre = "DV"
                Case "3", "5"
                    ZZTipo = "04"
                    ZZImpre = "ND"
                Case "4"
                    ZZTipo = "05"
                    ZZImpre = "NC"
                Case "6"
                    ZZTipo = "06"
                    ZZImpre = "RC"
                Case Else
                    Stop
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WCodIva = rstCliente!Iva
                WExpreso = Str$(rstCliente!Expreso)
                WProvincia = Str$(rstCliente!Provincia)
                Select Case Val(WCodIva)
                    Case 1, 2
                        WLetra = "A"
                    Case Else
                        WLetra = "B"
                End Select
                rstCliente.Close
            End If
    
            ZZPunto = WPunto
            ZZLetra = WLetra
            ZZNumero = Auxi
            ZZRenglon = "01"
            ZZCliente = WCliente
            ZZfecha = WFecha
            ZZEstado = "0"
            ZZVencimiento = WVencimiento
            
            ZZTotal = WTotal
            ZZSaldo = WTotal
            
            ZZNeto = WImporte
            ZZIva1 = WIva
            ZZIva2 = WIvaII
            ZZExento = WExento
            
            Select Case WTipo
                Case "1"
                    ZZPartida = "/"
                Case "3"
                    ZZPartida = "?"
                Case Else
                    ZZPartida = "S"
            End Select
            
            Select Case ZZPartida
                Case "/"
                    ZZTotalUs = Str$(Val(WTotal) + Val(WImporte))
                    ZZSaldoUs = Str$(Val(WTotal) + Val(WImporte))
                Case "?"
                    ZZTotalUs = Str$(Val(WTotal) + (Val(WImporte) * 11))
                    ZZSaldoUs = Str$(Val(WTotal) + (Val(WImporte) * 11))
                Case Else
                    ZZTotalUs = WTotal
                    ZZSaldoUs = WTotal
            End Select
            
            ZZOrdFecha = WOrdFecha
            ZZOrdVencimiento = WOrdvencimiento
            ZZPedido = ""
            ZZRemito = ""
            ZZOrden = ""
            ZZProvincia = Trim(WProvincia)
            ZZVendedor = WVendedor
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
            ZZBusqueda = ZZLetra + ZZPunto + Auxi
            
            ZZDescuento = ZZDto(Val(WBonificacion))
            ZZPago = WCondicion
            ZZExpreso = WExpreso
            ZZTipoIva = "0"
            If Val(ZZIva2) <> 0 Then
                ZZTipoIva = "1"
            End If
            ZZComision = "0"
            ZZRemito = ""
            
            ZZClave = ZZLetra + ZZTipo + ZZPunto + Auxi + "01"
            
            ZZLinea = ""
            
            ZZNetoTotal = ZZNeto
            If ZZPartida = "/" Then
                ZZNetoTotal = Str$(Val(ZZNeto) * 2)
            End If
            If ZZPartida = "?" Then
                ZZNetoTotal = Str$(Val(ZZNeto) * 12)
            End If
            
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
            ZSql = ZSql + "'" + ZZBusqueda + "')"
                                    
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
    Loop
    
    Close #1















    Linea% = 0

    Open "c:\datos\venMOV.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        Linea% = Linea% + 1
        
        WTipo = Mid$(WDato, 2, 1)
        WDocu = Mid$(WDato, 3, 1)
        WNumero = Mid$(WDato, 4, 8)
        WOrdFecha = Mid$(WDato, 12, 8)
        WFecha = Right$(WOrdFecha, 2) + "/" + Mid$(WOrdFecha, 5, 2) + "/" + Left$(WOrdFecha, 4)
        WOrdvencimiento = Mid$(WDato, 20, 8)
        WVencimiento = Right$(WOrdvencimiento, 2) + "/" + Mid$(WOrdvencimiento, 5, 2) + "/" + Left$(WOrdvencimiento, 4)
        WFechaII = Mid$(WDato, 28, 8)
        WImporte = Mid$(WDato, 36, 12)
        WIva = Mid$(WDato, 48, 12)
        WIvaII = Mid$(WDato, 60, 12)
        WImpInt = Mid$(WDato, 72, 12)
        WCliente = Mid$(WDato, 84, 1) + "-" + Mid$(WDato, 85, 3)
        WBonificacion = Mid$(WDato, 88, 2)
        WCondicion = Mid$(WDato, 90, 2)
        WVendedor = Mid$(WDato, 92, 2)
        WLista = Mid$(WDato, 94, 1)
        WUsaRemito = Mid$(WDato, 95, 2)
        WRutinas = Mid$(WDato, 97, 1)
        WAplica = Mid$(WDato, 98, 5)
        WExento = Mid$(WDato, 103, 12)
        WPuntero = Mid$(WDato, 115, 5)
        WComision = Mid$(WDato, 120, 12)
        
        If Val(WDocu) = 2 Or Val(WDocu) = 4 Or Val(WDocu) = 6 Then
            WImporte = Str$(Val(WImporte) * -1)
            WIva = Str$(Val(WIva) * -1)
            WIvaII = Str$(Val(WIvaII) * -1)
            WExento = Str$(Val(WExento) * -1)
        End If
        
        WTotal = Str$(Val(WImporte) + Val(WIva) + Val(WIvaII) + Val(WExento))
        
        
        If Val(WTotal) <> 0 And Val(WAplica) <> 0 And Val(WDocu) > 1 Then
        
        If Val(WAplica) <> Val(WNumero) Then
        
            Auxi = WNumero
            Call Ceros(Auxi, 8)
                    
            WPunto = "1"
            Call Ceros(WPunto, 4)
            
            Select Case WDocu
                Case "1"
                    ZZTipo = "01"
                    ZZImpre = "FC"
                Case "2"
                    ZZTipo = "02"
                    ZZImpre = "DV"
                Case "3", "5"
                    ZZTipo = "04"
                    ZZImpre = "ND"
                Case "4"
                    ZZTipo = "05"
                    ZZImpre = "NC"
                Case "6"
                    ZZTipo = "06"
                    ZZImpre = "RC"
                Case Else
                    Stop
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                WCodIva = rstCliente!Iva
                WExpreso = Str$(rstCliente!Expreso)
                WProvincia = Str$(rstCliente!Provincia)
                Select Case Val(WCodIva)
                    Case 1, 2
                        WLetra = "A"
                    Case Else
                        WLetra = "B"
                End Select
                rstCliente.Close
            End If
    
            ZZPunto = WPunto
            ZZLetra = WLetra
            ZZNumero = Auxi
            ZZRenglon = "01"
            ZZCliente = WCliente
            ZZfecha = WFecha
            ZZEstado = "0"
            ZZVencimiento = WVencimiento
            
            ZZTotal = WTotal
            ZZSaldo = WTotal
            
            ZZNeto = WImporte
            ZZIva1 = WIva
            ZZIva2 = WIvaII
            ZZExento = WExento
            
            Select Case WTipo
                Case "1"
                    ZZPartida = "/"
                Case "3"
                    ZZPartida = "?"
                Case Else
                    ZZPartida = "S"
            End Select
            
            Select Case ZZPartida
                Case "/"
                    ZZTotalUs = Str$(Val(WTotal) + Val(WImporte))
                    ZZSaldoUs = Str$(Val(WTotal) + Val(WImporte))
                Case "?"
                    ZZTotalUs = Str$(Val(WTotal) + (Val(WImporte) * 11))
                    ZZSaldoUs = Str$(Val(WTotal) + (Val(WImporte) * 11))
                Case Else
                    ZZTotalUs = WTotal
                    ZZSaldoUs = WTotal
            End Select
            
            ZZOrdFecha = WOrdFecha
            ZZOrdVencimiento = WOrdvencimiento
            ZZPedido = ""
            ZZRemito = ""
            ZZOrden = ""
            ZZProvincia = Trim(WProvincia)
            ZZVendedor = WVendedor
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
            ZZBusqueda = ZZLetra + ZZPunto + Auxi
            
            ZZDescuento = ZZDto(Val(WBonificacion))
            ZZPago = WCondicion
            ZZExpreso = WExpreso
            ZZTipoIva = "0"
            If Val(ZZIva2) <> 0 Then
                ZZTipoIva = "1"
            End If
            ZZComision = "0"
            ZZRemito = ""
            
            ZZClave = ZZLetra + ZZTipo + ZZPunto + Auxi + "01"
            
            ZZLinea = ""
            
            ZZNetoTotal = ZZNeto
            If ZZPartida = "/" Then
                ZZNetoTotal = Str$(Val(ZZNeto) * 2)
            End If
            If ZZPartida = "?" Then
                ZZNetoTotal = Str$(Val(ZZNeto) * 12)
            End If
            
            ZZAplica = WAplica
            Call Ceros(ZZAplica, 8)
            
            If Val(ZZAplica) = 70767 Then Stop
            
            If Val(WDocu) = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE CtaCte SET "
                ZSql = ZSql + " Saldo = 0 " + ","
                ZSql = ZSql + " SaldoUs = 0 "
                ZSql = ZSql + " Where Numero = " + "'" + Auxi + "'"
                ZSql = ZSql + " and Cliente = " + "'" + ZZCliente + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE CtaCte SET "
                ZSql = ZSql + " Saldo = 0 " + ","
                ZSql = ZSql + " SaldoUs = 0 "
                ZSql = ZSql + " Where Numero = " + "'" + ZZAplica + "'"
                ZSql = ZSql + " and Cliente = " + "'" + ZZCliente + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
                    
                ZSql = ""
                ZSql = ZSql + "UPDATE CtaCte SET "
                ZSql = ZSql + " Saldo = 0" + ","
                ZSql = ZSql + " SaldoUs = 0 "
                ZSql = ZSql + " Where Numero = " + "'" + Auxi + "'"
                ZSql = ZSql + " and Cliente = " + "'" + ZZCliente + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE CtaCte SET "
                ZSql = ZSql + " Saldo = Saldo + " + "'" + ZZTotal + "',"
                ZSql = ZSql + " SaldoUs = SaldoUs + " + "'" + ZZTotalUs + "'"
                ZSql = ZSql + " Where Numero = " + "'" + ZZAplica + "'"
                ZSql = ZSql + " and Cliente = " + "'" + ZZCliente + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
        
        End If
        
        End If
        
    Loop
    
    Close #1
    


    Open "c:\datos\venprove.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WProveedor = Mid$(WDato, 2, 4)
        WNombre = Mid$(WDato, 6, 25)
        WDireccion = Mid$(WDato, 31, 20)
        WLocalidad = Mid$(WDato, 51, 15)
        WPostal = Mid$(WDato, 66, 4)
        WTelefono = Mid$(WDato, 70, 20)
        WObservaciones = Mid$(WDato, 90, 20) + " " + Mid$(WDato, 109, 8)
        WCuit = Mid$(WDato, 118, 15)
        WEmail = ""
        WDias = "0"
        WGanancia = "0"
        WIva = "0"
        WProvincia = "0"
        WNombreCheque = ""
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Proveedor ("
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "Telefono ,"
        ZSql = ZSql + "Email ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Dias ,"
        ZSql = ZSql + "Ganancia ,"
        ZSql = ZSql + "Iva ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "NombreCheque )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WProveedor + "',"
        ZSql = ZSql + "'" + WNombre + "',"
        ZSql = ZSql + "'" + WDireccion + "',"
        ZSql = ZSql + "'" + WLocalidad + "',"
        ZSql = ZSql + "'" + WPostal + "',"
        ZSql = ZSql + "'" + WCuit + "',"
        ZSql = ZSql + "'" + WTelefono + "',"
        ZSql = ZSql + "'" + WEmail + "',"
        ZSql = ZSql + "'" + WObservaciones + "',"
        ZSql = ZSql + "'" + WDias + "',"
        ZSql = ZSql + "'" + WGanancia + "',"
        ZSql = ZSql + "'" + WIva + "',"
        ZSql = ZSql + "'" + WProvincia + "',"
        ZSql = ZSql + "'" + WNombreCheque + "')"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1


















    
    
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem dada
    Rem dada
    Rem dada
    Rem dada
    Rem dada








































































Stop







Stop












    Open "c:\datos\venfami.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 3)
        WDescripcion = Mid$(WDato, 6, 30)
        WEstado = "0"
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Familia ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Estado )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WDescripcion + "',"
        ZSql = ZSql + "'" + WEstado + "')"
        spFamilia = ZSql
        Set rstFamilia = db.OpenRecordset(spFamilia, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1



    Open "c:\datos\venzonas.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 3)
        WDescripcion = Mid$(WDato, 5, 25)
        WEstado = "0"
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Zona ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Estado )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WDescripcion + "',"
        ZSql = ZSql + "'" + WEstado + "')"
        spZona = ZSql
        Set rstZona = db.OpenRecordset(spZona, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1




    Open "c:\datos\venexpre.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 3)
        WNombre = Mid$(WDato, 6, 25)
        WDireccion = Mid$(WDato, 31, 20)
        WLocalidad = Mid$(WDato, 51, 15)
        WTelefono = Mid$(WDato, 66, 15)
        WCuit = Mid$(WDato, 81, 15)
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Expreso ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Nombre ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Postal ,"
        ZSql = ZSql + "Telefono ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "Email ,"
        ZSql = ZSql + "Fax ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "Iva )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WNombre + "',"
        ZSql = ZSql + "'" + WDireccion + "',"
        ZSql = ZSql + "'" + WLocalidad + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + WTelefono + "',"
        ZSql = ZSql + "'" + WCuit + "',"
        ZSql = ZSql + "'" + "0" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "0" + "',"
        ZSql = ZSql + "'" + "1" + "')"
        spExpreso = ZSql
        Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1



da:















    Open "c:\datos\venprnte.dat" For Input As #1
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 3, 4)
        WDescripcion = Mid$(WDato, 9, 30)
        WColor = Mid$(WDato, 39, 5)
        WLinea = ""
        WProveedor = Mid$(WDato, 57, 4)
        WUbicacion = Mid$(WDato, 222, 3)
        WUnidadCaja = Mid$(WDato, 61, 3)
        WCosto = Mid$(WDato, 64, 10)
        WFechaCosto = "  /  /    "
        WOrdFechaCosto = "00000000"
        WFechaCierre = "  /  /    "
        WFechaUltimaEntrada = "  /  /    "
        WFechaUltimaSalida = "  /  /    "
        WMinimo = Mid$(WDato, 164, 5)
        WStockAnterior = Mid$(WDato, 169, 6)
        WEntradas = Mid$(WDato, 175, 6)
        WSalidas = Mid$(WDato, 181, 6)
        WStock = Str$(Val(WStockAnterior) + Val(WEntradas) - Val(WSalidas))
        WCodigoProveedor = ""
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Insumo ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Color ,"
        ZSql = ZSql + "Linea ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "Ubicacion ,"
        ZSql = ZSql + "UnidadCaja ,"
        ZSql = ZSql + "Costo ,"
        ZSql = ZSql + "FechaCosto ,"
        ZSql = ZSql + "OrdFechaCosto ,"
        ZSql = ZSql + "FechaCierre ,"
        ZSql = ZSql + "FechaUltimaEntrada ,"
        ZSql = ZSql + "FechaUltimaSalida ,"
        ZSql = ZSql + "Minimo ,"
        ZSql = ZSql + "Entradas ,"
        ZSql = ZSql + "Salidas ,"
        ZSql = ZSql + "Stock ,"
        ZSql = ZSql + "CodigoProveedor )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WCodigo + "',"
        ZSql = ZSql + "'" + WDescripcion + "',"
        ZSql = ZSql + "'" + WColor + "',"
        ZSql = ZSql + "'" + WLinea + "',"
        ZSql = ZSql + "'" + WProveedor + "',"
        ZSql = ZSql + "'" + WUbicacion + "',"
        ZSql = ZSql + "'" + WUnidadCaja + "',"
        ZSql = ZSql + "'" + WCosto + "',"
        ZSql = ZSql + "'" + WFechaCosto + "',"
        ZSql = ZSql + "'" + WOrdFechaCosto + "',"
        ZSql = ZSql + "'" + WFechaCierre + "',"
        ZSql = ZSql + "'" + WFechaUltimaEntrada + "',"
        ZSql = ZSql + "'" + WFechaUltimaSalida + "',"
        ZSql = ZSql + "'" + WMinimo + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + "" + "',"
        ZSql = ZSql + "'" + WStock + "',"
        ZSql = ZSql + "'" + WCodigoProveedor + "')"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1




























Stop


















    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub

