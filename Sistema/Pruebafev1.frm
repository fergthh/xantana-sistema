VERSION 5.00
Begin VB.Form Pruebafev1 
   Caption         =   "Form1"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox Cae 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox Letra 
      Height          =   405
      Left            =   1080
      TabIndex        =   3
      Text            =   "A"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox Numero 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Text            =   "1"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Fecha 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "29/03/2011"
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Vencimiento 
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Text            =   "15/04/2011"
      Top             =   2880
      Width           =   2415
   End
End
Attribute VB_Name = "Pruebafev1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Ejemplo de Uso de Interface COM con Web Service Factura Electr?nica Mercado Interno AFIP
' Seg?n RG2904 Art?culo 4 Opci?n B (sin detalle, Version 1)
' 2010 (C) Mariano Reingart <reingart@gmail.com>
' Licencia: GPLv3

Dim WCuit As String

Private Sub Command1_Click()
    
    Dim WSAA As Object, WSFEv1 As Object
    
    On Error GoTo ManejoError
    
    
    Stop
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    
    
    ' Crear objeto interface Web Service Autenticaci?n y Autorizaci?n
    Set WSAA = CreateObject("WSAA")
    Debug.Print WSAA.Version
    'Debug.Print WSAA.InstallDir
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA) para WSFEv1
    tra = WSAA.CreateTRA("wsfe")
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
        
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "surfactan"
    Rem ZCuit = "30549165083"
    Rem punto_vta = 9

    ZPath = "c:\salva\"
    ZNombre = "Pellital"
    ZCuit = "30610524598"
    punto_vta = 6
    
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "Yenadi"
    Rem ZCuit = "30629950415"
    Rem punto_vta = 4
    
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "Haz"
    Rem ZCuit = "30650608638"
    Rem punto_vta = 10
    
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "celugama"
    Rem ZCuit = "30637671622"
    Rem punto_vta = 3
    
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "Jinarg"
    Rem ZCuit = "30707062882"
    Rem punto_vta = 5
    
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "Delia"
    Rem ZCuit = "27061818087"
    Rem punto_vta = 2
    
    Rem ZPath = "c:\salva\"
    Rem ZNombre = "Tiny"
    Rem ZCuit = "30708486481"
    Rem punto_vta = 2
    
    
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
    tipo_cbte = 2
    
    Stop
    
    cbte_nro = WSFEv1.CompUltimoAutorizado(tipo_cbte, punto_vta)
    If cbte_nro = "" Then
        cbte_nro = 0                ' no hay comprobantes emitidos
            Else
        cbte_nro = CLng(cbte_nro)   ' convertir a entero largo
    End If
    
    
    
    Stop
    
    WRazon = "Surfacatn S.A."
    WCuit = "30-54916508-3"
    Call Eval
    
    
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
    imp_total = 36.18
    
    Rem IMPORTE DE CONCEPTOS NO GRAVADOS POR EL IVA
    imp_tot_conc = 0
    
    Rem IMPORTE NETO
    imp_neto = 29.9
    
    Rem IMPORTE IVA
    imp_iva = 6.28
    
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
    id = 5 ' 21%
    base_imp = 29.9
    IMPORTE = 6.28
    ok = WSFEv1.AgregarIva(id, base_imp, IMPORTE)
    
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

    MsgBox "Resultado:" & WSFEv1.resultado & " CAE: " & Cae & " Venc: " & WSFEv1.Vencimiento & " Obs: " & WSFEv1.obs & " Reproceso: " & WSFEv1.Reproceso, vbInformation + vbOKOnly
    
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
    Debug.Print "Fecha Vencimiento CAE", WSFEv1.Vencimiento
    Debug.Print "Importe Total:", WSFEv1.ImpTotal
    Debug.Print "Resultado:", WSFEv1.resultado
    
    If Cae <> cae2 Then
        MsgBox "El CAE de la factura no concuerdan con el recuperado en la AFIP!: " & Cae & " vs " & cae2
    Else
        MsgBox "El CAE de la factura concuerdan con el recuperado de la AFIP"
    End If
        
        
    If WSFE.resultado = "A" Then
        Cae.Text = Cae
        Vencimiento.Text = WSFE.Vencimiento
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

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub



















Private Sub Calcula_CaeV0()

    Dim WSAA As Object, WSFE As Object
    
    On Error GoTo ManejoError
    
    Cae.Text = "12345678901234"
    VtoCae.Text = "12/12/2011"
    Exit Sub
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    ' Crear objeto interface Web Service Autenticación y Autorización
    Set WSAA = CreateObject("WSAA")
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA)
    tra = WSAA.CreateTRA()
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Rem Path = CurDir() + "\"
    
    ZPath = "c:\salva\"
    ZNombre = "Yenadi"
    ZCuit = "30629950415"
        
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    ' Llamar al web service para autenticar:
    Rem ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms") ' Hologación
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms") ' Producción

    ' Imprimir el ticket de acceso, ToKen y Sign de autorización
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 6 horas
    ' (este período se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electrónica
    Set WSFE = CreateObject("WSFE")
    ' Setear tocken y sing de autorización (pasos previos)
    WSFE.Token = WSAA.Token
    WSFE.Sign = WSAA.Sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFE.Cuit = ZCuit
    
    ' Conectar al Servicio Web de Facturación
    Rem ok = WSFE.Conectar("https://wswhomo.afip.gov.ar/wsfe/service.asmx") ' homologación
    Rem ok = WSFE.Conectar("https://wsw.afip.gov.ar/wsfe/service.asmx") ' producción
    ok = WSFE.Conectar("https://servicios1.afip.gov.ar/wsfe/service.asmx") ' producción

    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFE.Dummy
    Debug.Print "appserver status", WSFE.AppServerStatus
    Debug.Print "dbserver status", WSFE.DbServerStatus
    Debug.Print "authserver status", WSFE.AuthServerStatus
    
    ' Recupera cantidad máxima de registros (opcional)
    qty = WSFE.RecuperarQty()
    
    ' Recupera último número de secuencia ID
    LastId = WSFE.UltNro()
    
    ' Recupero último número de comprobante para un punto de venta y tipo (opcional)
    
    If Letra.Text = "A" Then
        tipo_cbte = 1
        WTipo = "01"
            Else
        tipo_cbte = 6
        WTipo = "06"
    End If
    
    punto_vta = Val(Punto.Text)
    LastCBTE = WSFE.RecuperaLastCMP(punto_vta, tipo_cbte)
    
    ' Establezco los valores de la factura o lote a autorizar:

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WCuit = rstCliente!Cuit
        rstCliente.Close
        Call Eval
    End If
        
    Rem Fecha = Format(Date, "yyyymmdd")
    
    id = LastId + 1
    presta_serv = 0
    
    If Len(WCuit) = 11 Then
        tipo_doc = 80
            Else
        tipo_doc = 96
    End If
    nro_doc = Left$(WCuit + Space$(11), 11)
    
    If Val(Numero.Text) <> LastCBTE + 1 Then
        M$ = "El numero de comprobante no es igua al correlativo indicado por la afip " + Str$(LastCBTE + 1)
        a% = MsgBox(M$, 0, "Eliminacion de Comprobantes")
        Exit Sub
    End If
    
    cbt_desde = LastCBTE + 1
    cbt_hasta = LastCBTE + 1
    
    imp_total = Val(Total.Caption)
    imp_tot_conc = 0
    imp_neto = Val(Neto.Caption)
    impto_liq = Val(Iva1.Caption)
    impto_liq_rni = 0
    imp_op_ex = 0
    
    ZZfecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    fecha_cbte = ZZfecha
    
    ZZfecha = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
    fecha_venc_pago = ZZfecha
    
    
    ZZDesdeFecha = "0000000"
    ZZHastaFecha = "0000000"
    
    ' Fechas del período del servicio facturado (solo si presta_serv = 1)
    ZZfecha = Right$(ZZDesdeFecha, 4) + Mid$(ZZDesdeFecha, 4, 2) + Left$(ZZDesdeFecha, 2)
    Rem XXFecha = Format(ZZFecha, "yyyymmdd")
    fecha_serv_desde = ZZfecha
    
    
    
    ZZfecha = Right$(ZZHastaFecha, 4) + Mid$(ZZHastaFecha, 4, 2) + Left$(ZZHastaFecha, 2)
    Rem Fecha = Format(ZZFecha, "yyyymmdd")
    fecha_serv_hasta = ZZfecha
    
    
    fecha_serv_desde = ""
    fecha_serv_hasta = ""
    
    
    ' Llamo al WebService de Autorización para obtener el CAE
    Cae = WSFE.Aut(id, presta_serv, _
        tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        impto_liq, impto_liq_rni, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta) ' si presta_serv = 0 no pasar estas fechas
    
    Debug.Print "Vencimiento ", WSFE.Vencimiento ' Fecha de vencimiento o vencimiento de la autorización
    Debug.Print "Resultado: ", WSFE.resultado ' A=Aceptado, R=Rechazado
    Debug.Print "Motivo de rechazo o advertencia", WSFE.Motivo ' 00= No hay error
    Debug.Print "Reprocesado?", WSFE.Reproceso ' S=Si, N=No
    
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    If Cae = "" Then
        MsgBox "La página esta caida o la respuesta es inválida"
    ElseIf Cae = "NULL" Or WSFE.resultado <> "A" Then
        MsgBox "No se asignó CAE (Rechazado). Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    ElseIf WSFE.Motivo <> "NULL" And WSFE.Motivo <> "00" Then
        MsgBox "Se asignó CAE pero con advertencias. Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    End If
    
    ' Imprimo respuesta XML para depuración (errores de formato)
    Debug.Print WSFE.XmlResponse
    
    MsgBox "QTY: " & qty & vbCrLf & "LastId: " & LastId & vbCrLf & "LastCBTE:" & LastCBTE & vbCrLf & "CAE: " & Cae, vbInformation + vbOKOnly
    
    If WSFE.resultado = "A" Then
        Cae.Text = Cae
        VtoCae.Text = WSFE.Vencimiento
    End If
    
    Exit Sub

ManejoError:
    ' Si hubo error:
    Debug.Print Err.Description            ' descripción error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select

End Sub


