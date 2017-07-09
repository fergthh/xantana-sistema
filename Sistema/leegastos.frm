VERSION 5.00
Begin VB.Form PrgLeeGastos 
   Caption         =   "Traspaso de Gastos"
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
         MouseIcon       =   "leegastos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "leegastos.frx":030A
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
         MouseIcon       =   "leegastos.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "leegastos.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Confirma Proceso de Grabacion"
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "PrgLeeGastos"
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

Private Sub Acepta_Click()
    Call Proceso
    Call Cancela_Click
End Sub

Private Sub Cancela_Click()
    PrgLeeGastos.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Proceso()
    
    
    Open "c:\Gastos\Gastos.txt" For Input As #1
    If Val(WEmpresa) = 1 Then
         Open "c:\Gastos\Cliedon.txt" For Input As #2
         Open "c:\Gastos\Ccdon.txt" For Input As #3
             Else
         Open "c:\Gastos\ClieSu.txt" For Input As #2
         Open "c:\Gastos\CcSu.txt" For Input As #3
    End If
    
    Pasa = 0
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 2, 8)
        WCodigo = Str$(Val(WCodigo))
        
        Call Ceros(WCodigo, 6)
        WRenglon = Mid$(WDato, 10, 9)
        WRenglon = Str$(Val(WRenglon))
        Call Ceros(WRenglon, 2)
        WClave = WCodigo + WRenglon
        WProveedor = Mid$(WDato, 20, 9)
        WConcepto = Mid$(WDato, 30, 9)
        WFecha = Mid$(WDato, 40, 9)
        WFecha = Mid$(WFecha, 1, 2) + "/" + Mid$(WFecha, 3, 2) + "/20" + Mid$(WFecha, 5, 2)
        WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        WNumero = Mid$(WDato, 50, 9)
        WNeto = Mid$(WDato, 60, 19)
        WIva = Mid$(WDato, 80, 19)
        WTotal = Str$(Val(WNeto) + Val(WIva))
        WCamion = Mid$(WDato, 100, 9)
        WTanque = Mid$(WDato, 110, 9)
        WLetra = Mid$(WDato, 120, 1)
        WLitros = Mid$(WDato, 130, 19)
        WPrecio = Mid$(WDato, 150, 19)
        WFechaCarga = Mid$(WDato, 170, 9)
        WFechaCarga = Mid$(WFechaCarga, 1, 2) + "/" + Mid$(WFechaCarga, 3, 2) + "/20" + Mid$(WFechaCarga, 5, 2)
        WOrdFechaCarga = Right$(WFechaCarga, 4) + Mid$(WFechaCarga, 4, 2) + Left$(WFechaCarga, 2)
        WChofer = Mid$(WDato, 180, 9)
        
        If (Val(WEmpresa) = 1 And Val(WChofer) >= 100) Or (Val(WEmpresa) = 2 And Val(WChofer) < 100) Then
        
            If Pasa = 0 Then
            
                Pasa = 1
            
                Corte = WCodigo
                CorteChofer = WChofer
                CorteFecha = WFecha
            
                ZSuma = 0
            
                ZSql = ""
                ZSql = ZSql + "DELETE Gastos"
                ZSql = ZSql + " Where Codigo = " + "'" + Corte + "'"
                spGastos = ZSql
                Set rstGastos = db.OpenRecordset(spGastos, dbOpenSnapshot, dbSQLPassThrough)
            
                WWLetra = "X"
                WWTipo = "01"
                WWPunto = "9999"
                WWNumero = Corte
                WWChofer = CorteChofer
        
                Call Ceros(WWNumero, 8)
                Call Ceros(WWChofer, 6)
        
                WWClave = WWChofer + WWLetra + WWTipo + WWPunto + WWNumero
                
                ZPago = 0
                ZTotalAnte = 0
                ZSaldoAnte = 0
                ZChoferAnte = 0
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCteChofer"
                ZSql = ZSql + " Where CtaCteChofer.Letra = " + "'" + WWLetra + "'"
                ZSql = ZSql + " and CtaCteChofer.Tipo = " + "'" + WWTipo + "'"
                ZSql = ZSql + " and CtaCteChofer.Punto = " + "'" + WWPunto + "'"
                ZSql = ZSql + " and CtaCteChofer.Numero = " + "'" + WWNumero + "'"
                spCtaCteChofer = ZSql
                Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCteChofer.RecordCount > 0 Then
                    ZPago = rstCtaCteChofer!Total - rstCtaCteChofer!Saldo
                    ZTotalAnte = rstCtaCteChofer!Total
                    ZSaldoAnte = rstCtaCteChofer!Saldo
                    ZChoferAnte = rstCtaCteChofer!Chofer
                    Call Redondeo(ZPago)
                    Call Redondeo(ZTotalAnte)
                    Call Redondeo(ZSaldoAnte)
                    rstCtaCteChofer.Close
                End If
                
                If ZChoferAnte <> Val(WWChofer) And ZPago <> 0 Then
                
                    WGrabaTotal = "0"
                    WGrabaSaldo = Str$(ZPago * -1)
                    
                    WWLetraII = "X"
                    WWTipoII = "01"
                    WWPuntoII = "8888"
                    WWNumeroII = WWNumero
                    WWChoferII = Str$(ZChoferAnte)
        
                    Call Ceros(WWNumeroII, 8)
                    Call Ceros(WWChoferII, 6)
        
                    WWClaveII = WWChoferII + WWLetraII + WWTipoII + WWPuntoII + WWNumeroII
                
                    WWLetraIII = "X"
                    WWTipoIII = "01"
                    WWPuntoIII = "9999"
                    WWNumeroIII = WWNumero
                    WWChoferIII = Str$(ZChoferAnte)
        
                    Call Ceros(WWNumeroIII, 8)
                    Call Ceros(WWChoferIII, 6)
        
                    WWClaveIII = WWChoferIII + WWLetraIII + WWTipoIII + WWPuntoIII + WWNumeroIII
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CtaCteChofer SET "
                    ZSql = ZSql + " CtaCteChofer.Total = " + "'" + WGrabaTotal + "',"
                    ZSql = ZSql + " CtaCteChofer.Saldo = " + "'" + WGrabaSaldo + "',"
                    ZSql = ZSql + " CtaCteChofer.Punto = " + "'" + WWPuntoII + "',"
                    ZSql = ZSql + " CtaCteChofer.Clave = " + "'" + WWClaveII + "'"
                    ZSql = ZSql + " Where CtaCteChofer.Clave = " + "'" + WWClaveIII + "'"
                    spCtaCteChofer = ZSql
                    Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZPago = 0
            
                        Else
                
                    ZSql = ""
                    ZSql = ZSql + "DELETE CtaCteChofer"
                    ZSql = ZSql + " Where CtaCteChofer.Letra = " + "'" + WWLetra + "'"
                    ZSql = ZSql + " and CtaCteChofer.Tipo = " + "'" + WWTipo + "'"
                    ZSql = ZSql + " and CtaCteChofer.Punto = " + "'" + WWPunto + "'"
                    ZSql = ZSql + " and CtaCteChofer.Numero = " + "'" + WWNumero + "'"
                    spCtaCteChofer = ZSql
                    Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
            End If
        
            If Corte <> WCodigo Then
        
                ZZChofer = CorteChofer
                ZZLetra = "X"
                ZZTipo = "01"
                ZZPunto = "9999"
                ZZNumero = Corte
                Call Ceros(ZZChofer, 6)
                Call Ceros(ZZTipo, 2)
                Call Ceros(ZZPunto, 4)
                Call Ceros(ZZNumero, 8)
                ZZClaveChofer = ZZChofer + ZZLetra + ZZTipo + ZZPunto + ZZNumero
                ZZfecha = CorteFecha
                ZZVencimiento = CorteFecha
                ZZTotal = Str$(ZSuma)
                ZZSaldo = Str$(ZSuma - ZPago)
                ZZObservaciones = "Constancia"
                ZZOrdFecha = Right$(CorteFecha, 4) + Mid$(CorteFecha, 4, 2) + Left$(CorteFecha, 2)
                ZZOrdVencimiento = Right$(CorteFecha, 4) + Mid$(CorteFecha, 4, 2) + Left$(CorteFecha, 2)
                ZZProveedor = "1"
                ZZCai = ""
                ZZVtoCai = ""
                ZZImpre = "CO"
                
                If Val(ZZTotal) <> 0 Then
            
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO CtaCteChofer ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "Chofer ,"
                    ZSql = ZSql + "Letra ,"
                    ZSql = ZSql + "Tipo ,"
                    ZSql = ZSql + "Punto ,"
                    ZSql = ZSql + "Numero ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Estado ,"
                    ZSql = ZSql + "Vencimiento ,"
                    ZSql = ZSql + "Total ,"
                    ZSql = ZSql + "Saldo ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "OrdFecha ,"
                    ZSql = ZSql + "OrdVencimiento ,"
                    ZSql = ZSql + "Proveedor ,"
                    ZSql = ZSql + "Cai ,"
                    ZSql = ZSql + "VtoCai ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + ZZClaveChofer + "',"
                    ZSql = ZSql + "'" + ZZChofer + "',"
                    ZSql = ZSql + "'" + ZZLetra + "',"
                    ZSql = ZSql + "'" + ZZTipo + "',"
                    ZSql = ZSql + "'" + ZZPunto + "',"
                    ZSql = ZSql + "'" + ZZNumero + "',"
                    ZSql = ZSql + "'" + ZZfecha + "',"
                    ZSql = ZSql + "'" + "1" + "',"
                    ZSql = ZSql + "'" + ZZVencimiento + "',"
                    ZSql = ZSql + "'" + ZZTotal + "',"
                    ZSql = ZSql + "'" + ZZSaldo + "',"
                    ZSql = ZSql + "'" + ZZObservaciones + "',"
                    ZSql = ZSql + "'" + ZZOrdFecha + "',"
                    ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
                    ZSql = ZSql + "'" + ZZProveedor + "',"
                    ZSql = ZSql + "'" + ZZCai + "',"
                    ZSql = ZSql + "'" + ZZVtoCai + "',"
                    ZSql = ZSql + "'" + ZZImpre + "')"
                    spCtaCteChofer = ZSql
                    Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
            
                Corte = WCodigo
                CorteChofer = WChofer
                CorteFecha = WFecha
            
                ZSuma = 0
            
                ZSql = ""
                ZSql = ZSql + "DELETE Gastos"
                ZSql = ZSql + " Where Codigo = " + "'" + Corte + "'"
                spGastos = ZSql
                Set rstGastos = db.OpenRecordset(spGastos, dbOpenSnapshot, dbSQLPassThrough)
            
                WWLetra = "X"
                WWTipo = "01"
                WWPunto = "9999"
                WWNumero = Corte
                WWChofer = CorteChofer
        
                Call Ceros(WWNumero, 8)
                Call Ceros(WWChofer, 6)
        
                WWClave = WWChofer + WWLetra + WWTipo + WWPunto + WWNumero
                
                ZPago = 0
                ZTotalAnte = 0
                ZSaldoAnte = 0
                ZChoferAnte = 0
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCteChofer"
                ZSql = ZSql + " Where CtaCteChofer.Letra = " + "'" + WWLetra + "'"
                ZSql = ZSql + " and CtaCteChofer.Tipo = " + "'" + WWTipo + "'"
                ZSql = ZSql + " and CtaCteChofer.Punto = " + "'" + WWPunto + "'"
                ZSql = ZSql + " and CtaCteChofer.Numero = " + "'" + WWNumero + "'"
                spCtaCteChofer = ZSql
                Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCteChofer.RecordCount > 0 Then
                    ZPago = rstCtaCteChofer!Total - rstCtaCteChofer!Saldo
                    ZTotalAnte = rstCtaCteChofer!Total
                    ZSaldoAnte = rstCtaCteChofer!Saldo
                    ZChoferAnte = rstCtaCteChofer!Chofer
                    Call Redondeo(ZPago)
                    Call Redondeo(ZTotalAnte)
                    Call Redondeo(ZSaldoAnte)
                    rstCtaCteChofer.Close
                End If
                
                If ZChoferAnte <> Val(WWChofer) And ZPago <> 0 Then
                
                    WGrabaTotal = "0"
                    WGrabaSaldo = Str$(ZPago * -1)
                    
                    WWLetraII = "X"
                    WWTipoII = "01"
                    WWPuntoII = "8888"
                    WWNumeroII = WWNumero
                    WWChoferII = Str$(ZChoferAnte)
        
                    Call Ceros(WWNumeroII, 8)
                    Call Ceros(WWChoferII, 6)
        
                    WWClaveII = WWChoferII + WWLetraII + WWTipoII + WWPuntoII + WWNumeroII
                
                    WWLetraIII = "X"
                    WWTipoIII = "01"
                    WWPuntoIII = "9999"
                    WWNumeroIII = WWNumero
                    WWChoferIII = Str$(ZChoferAnte)
        
                    Call Ceros(WWNumeroIII, 8)
                    Call Ceros(WWChoferIII, 6)
        
                    WWClaveIII = WWChoferIII + WWLetraIII + WWTipoIII + WWPuntoIII + WWNumeroIII
                
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CtaCteChofer SET "
                    ZSql = ZSql + " CtaCteChofer.Total = " + "'" + WGrabaTotal + "',"
                    ZSql = ZSql + " CtaCteChofer.Saldo = " + "'" + WGrabaSaldo + "',"
                    ZSql = ZSql + " CtaCteChofer.Punto = " + "'" + WWPuntoII + "',"
                    ZSql = ZSql + " CtaCteChofer.Clave = " + "'" + WWClaveII + "'"
                    ZSql = ZSql + " Where CtaCteChofer.Clave = " + "'" + WWClaveIII + "'"
                    spCtaCteChofer = ZSql
                    Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZPago = 0
            
                        Else
                
                    ZSql = ""
                    ZSql = ZSql + "DELETE CtaCteChofer"
                    ZSql = ZSql + " Where CtaCteChofer.Letra = " + "'" + WWLetra + "'"
                    ZSql = ZSql + " and CtaCteChofer.Tipo = " + "'" + WWTipo + "'"
                    ZSql = ZSql + " and CtaCteChofer.Punto = " + "'" + WWPunto + "'"
                    ZSql = ZSql + " and CtaCteChofer.Numero = " + "'" + WWNumero + "'"
                    spCtaCteChofer = ZSql
                    Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
            
            End If
        
            ZSuma = ZSuma + Val(WTotal)
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Gastos ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Concepto ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Neto ,"
            ZSql = ZSql + "Iva ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Camion ,"
            ZSql = ZSql + "Tanque ,"
            ZSql = ZSql + "Litros ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "FechaCarga ,"
            ZSql = ZSql + "OrdFechaCarga ,"
            ZSql = ZSql + "Chofer )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WRenglon + "',"
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WConcepto + "',"
            ZSql = ZSql + "'" + WFecha + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + WNumero + "',"
            ZSql = ZSql + "'" + WNeto + "',"
            ZSql = ZSql + "'" + WIva + "',"
            ZSql = ZSql + "'" + WTotal + "',"
            ZSql = ZSql + "'" + WCamion + "',"
            ZSql = ZSql + "'" + WTanque + "',"
            ZSql = ZSql + "'" + WLitros + "',"
            ZSql = ZSql + "'" + WPrecio + "',"
            ZSql = ZSql + "'" + WFechaCarga + "',"
            ZSql = ZSql + "'" + WOrdFechaCarga + "',"
            ZSql = ZSql + "'" + WChofer + "')"
        
            spGastos = ZSql
            Set rstGastos = db.OpenRecordset(spGastos, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Loop
    
    If Pasa <> 0 Then
        
        ZZChofer = CorteChofer
        ZZLetra = "X"
        ZZTipo = "01"
        ZZPunto = "9999"
        ZZNumero = Corte
        Call Ceros(ZZChofer, 6)
        Call Ceros(ZZTipo, 2)
        Call Ceros(ZZPunto, 4)
        Call Ceros(ZZNumero, 8)
        ZZClaveChofer = ZZChofer + ZZLetra + ZZTipo + ZZPunto + ZZNumero
        ZZfecha = CorteFecha
        ZZVencimiento = CorteFecha
        ZZTotal = Str$(ZSuma)
        ZZSaldo = Str$(ZSuma - ZPago)
        ZZObservaciones = "Constancia"
        ZZOrdFecha = Right$(CorteFecha, 4) + Mid$(CorteFecha, 4, 2) + Left$(CorteFecha, 2)
        ZZOrdVencimiento = Right$(CorteFecha, 4) + Mid$(CorteFecha, 4, 2) + Left$(CorteFecha, 2)
        ZZProveedor = "1"
        ZZCai = ""
        ZZVtoCai = ""
        ZZImpre = "CO"
        
        If Val(ZZTotal) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CtaCteChofer ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Chofer ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "Vencimiento ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Saldo ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "OrdVencimiento ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Cai ,"
            ZSql = ZSql + "VtoCai ,"
            ZSql = ZSql + "Impre )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClaveChofer + "',"
            ZSql = ZSql + "'" + ZZChofer + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZPunto + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + "1" + "',"
            ZSql = ZSql + "'" + ZZVencimiento + "',"
            ZSql = ZSql + "'" + ZZTotal + "',"
            ZSql = ZSql + "'" + ZZSaldo + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
            ZSql = ZSql + "'" + ZZProveedor + "',"
            ZSql = ZSql + "'" + ZZCai + "',"
            ZSql = ZSql + "'" + ZZVtoCai + "',"
            ZSql = ZSql + "'" + ZZImpre + "')"
            spCtaCteChofer = ZSql
            Set rstCtaCteChofer = db.OpenRecordset(spCtaCteChofer, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    End If
    
    
    Do
    
        Line Input #2, WDato
        If EOF(2) Then Exit Do
        
        WCodigo = Trim(Mid$(WDato, 1, 9))
        
        If Val(WCodigo) <> 0 Then
        
        WNombre = Trim(Mid$(WDato, 10, 50))
        WDireccion = Trim(Mid$(WDato, 60, 50))
        WLocalidad = Trim(Mid$(WDato, 110, 50))
        Select Case Val(Trim(Mid$(WDato, 170, 9)))
            Case 2
                WProvincia = "1"
            Case Else
                WProvincia = "0"
        End Select
        WTelefono = Trim(Mid$(WDato, 180, 20))
        WIva = Trim(Mid$(WDato, 200, 4))
        WCuit = Trim(Mid$(WDato, 205, 20))
        WObservaciones = Trim(Mid$(WDato, 230, 50))
        WDias = Trim(Mid$(WDato, 300, 9))
        WCuenta = Trim(Mid$(WDato, 310, 9))
        WComision1 = Trim(Mid$(WDato, 320, 19))
        WComision2 = Trim(Mid$(WDato, 240, 19))
        
        ZZCliente = WCodigo
        ZZRazon = WNombre
        ZZDireccion = WDireccion
        ZZLocalidad = WLocalidad
        ZZPostal = ""
        ZZTelefono = WTelefono
        ZZObservaciones = WObservaciones
        ZZCuit = WCuit
        ZZEmail = ""
        ZZFax = ""
        ZZProvincia = WProvincia
        ZZIva = WIva
        ZZDias = WDias
        ZZVendedor = ""
        ZZDescuento = ""
        ZZCuenta = WCuenta
        ZZComision1 = WComision1
        ZZComision2 = WComision2
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZZCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            rstCliente.Close
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + " Razon = " + "'" + ZZRazon + "',"
            ZSql = ZSql + " Direccion = " + "'" + ZZDireccion + "',"
            ZSql = ZSql + " Localidad = " + "'" + ZZLocalidad + "',"
            ZSql = ZSql + " Postal = " + "'" + ZZPostal + "',"
            ZSql = ZSql + " Telefono = " + "'" + ZZTelefono + "',"
            ZSql = ZSql + " Observaciones = " + "'" + ZZObservaciones + "',"
            ZSql = ZSql + " Cuit = " + "'" + ZZCuit + "',"
            ZSql = ZSql + " Email = " + "'" + ZZEmail + "',"
            ZSql = ZSql + " Fax = " + "'" + ZZFax + "',"
            ZSql = ZSql + " Provincia = " + "'" + ZZProvincia + "',"
            ZSql = ZSql + " Iva = " + "'" + ZZIva + "',"
            ZSql = ZSql + " Dias = " + "'" + ZZDias + "',"
            ZSql = ZSql + " Vendedor = " + "'" + ZZVendedor + "',"
            ZSql = ZSql + " Descuento = " + "'" + ZZDescuento + "',"
            ZSql = ZSql + " Cuenta = " + "'" + ZZCuenta + "'"
            ZSql = ZSql + " Where Cliente = " + "'" + ZZCliente + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
                Else
            
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
            ZSql = ZSql + "Iva ,"
            ZSql = ZSql + "Dias ,"
            ZSql = ZSql + "Vendedor ,"
            ZSql = ZSql + "Descuento ,"
            ZSql = ZSql + "Cuenta )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZRazon + "',"
            ZSql = ZSql + "'" + ZZDireccion + "',"
            ZSql = ZSql + "'" + ZZLocalidad + "',"
            ZSql = ZSql + "'" + ZZPostal + "',"
            ZSql = ZSql + "'" + ZZTelefono + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZCuit + "',"
            ZSql = ZSql + "'" + ZZEmail + "',"
            ZSql = ZSql + "'" + ZZFax + "',"
            ZSql = ZSql + "'" + ZZProvincia + "',"
            ZSql = ZSql + "'" + ZZIva + "',"
            ZSql = ZSql + "'" + ZZDias + "',"
            ZSql = ZSql + "'" + ZZVendedor + "',"
            ZSql = ZSql + "'" + ZZDescuento + "',"
            ZSql = ZSql + "'" + ZZCuenta + "')"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
        End If
        
    Loop
    
    
    
    Do
    
        Line Input #3, WDato
        If EOF(3) Then Exit Do
        
        WCliente = Trim(Mid$(WDato, 1, 9))
        WTipo = Trim(Mid$(WDato, 10, 9))
        WNumero = Trim(Mid$(WDato, 20, 9))
        WRenglon = Trim(Mid$(WDato, 30, 9))
        WFecha = Trim(Mid$(WDato, 40, 9))
        WFecha = Mid$(WFecha, 5, 2) + "/" + Mid$(WFecha, 3, 2) + "/20" + Mid$(WFecha, 1, 2)
        WEstado = Trim(Mid$(WDato, 50, 9))
        WTotal = Val(Mid$(WDato, 60, 19))
        WSaldo = Val(Mid$(WDato, 80, 19))
        WVencimiento = Trim(Mid$(WDato, 100, 19))
        WVencimiento = Mid$(WVencimiento, 5, 2) + "/" + Mid$(WVencimiento, 3, 2) + "/20" + Mid$(WVencimiento, 1, 2)
        WTipoFac = Trim(Mid$(WDato, 120, 9))
        
        Auxi = WNumero
        Call Ceros(Auxi, 8)
            
        ZZPunto = "0001"
        Select Case Val(WTipo)
            Case 1
                ZZImpre = "AN"
                ZZTipo = "07"
            Case 78
                ZZImpre = "ND"
                ZZTipo = "02"
            Case 79
                ZZImpre = "NC"
                ZZTipo = "03"
            Case Else
                ZZImpre = "FC"
                ZZTipo = "01"
        End Select
        ZZLetra = "A"
        Call Ceros(ZZTipo, 2)
        
        Claveven$ = ZZLetra + ZZTipo + ZZPunto + Auxi + "01"
           
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Ctacte"
        ZSql = ZSql + " Where Ctacte.Clave = " + "'" + Claveven$ + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCte.RecordCount > 0 Then
        
            rstCtaCte.Close
            
                Else
                
            ZZPunto = "0001"
            ZZLetra = "A"
            ZZNumero = Auxi
            ZZRenglon = "01"
            ZZCliente = WCliente
            ZZfecha = WFecha
            ZZEstado = WEstado
            ZZVencimiento = WVencimiento
            Select Case Val(ZZTipo)
                Case 7, 3
                    ZZTotal = Str$(WTotal * -1)
                    ZZSaldo = Str$(WSaldo * -1)
                    ZZTotalUs = Str$(WTotal * -1)
                    ZZSaldoUs = Str$(WSaldo * -1)
                Case Else
                    ZZTotal = Str$(WTotal)
                    ZZSaldo = Str$(WSaldo)
                    ZZTotalUs = Str$(WTotal)
                    ZZSaldoUs = Str$(WSaldo)
            End Select
                
            ZZNeto = "0"
            ZZIva1 = "0"
            ZZIva2 = "0"
            ZZOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            ZZOrdVencimiento = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
            ZZPedido = ""
            ZZRemito = ""
            ZZOrden = ""
            ZZProvincia = ""
            ZZVendedor = ""
            ZZCosto = "0"
            ZZImporte1 = "0"
            ZZImporte2 = "0"
            ZZImporte3 = "0"
            ZZImporte4 = "0"
            ZZImporte5 = "0"
            ZZImporte6 = "0"
            ZZImporte7 = "0"
            ZZTipoventa = WTipoFac
            ZZProyecto = ""
            ZZParidad = "0"
            ZZRemito1 = ""
            ZZRemito2 = ""
            ZZBusqueda = ZZLetra + ZZPunto + Auxi
            ZZDescuento = ""
    
    
            ZZClave = ZZLetra + ZZTipo + ZZPunto + Auxi + "01"
    
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
            ZSql = ZSql + "Iva1 ,"
            ZSql = ZSql + "Iva2 ,"
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
            ZSql = ZSql + "'" + ZZIva1 + "',"
            ZSql = ZSql + "'" + ZZIva2 + "',"
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
            ZSql = ZSql + "'" + ZZBusqueda + "')"
                                
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
        
    Loop
    
    
    
    
Da:
    
    Close #1
    Close #2
    Close #3
    
    
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub


Private Sub Form_Load()
    PrgLeeGastos.Caption = "Traspaso de Gastos: " + WNombreEmpresa
End Sub
