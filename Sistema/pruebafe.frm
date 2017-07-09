VERSION 5.00
Begin VB.Form Pruebafe 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Vencimiento 
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Text            =   "15/04/2011"
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Fecha 
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "15/03/2011"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Numero 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Text            =   "1"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox Letra 
      Height          =   405
      Left            =   840
      TabIndex        =   2
      Text            =   "A"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox Cae 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Pruebafe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Dim WSAA As Object, WSFE As Object
    
    On Error GoTo ManejoError
    
    
    Stop
    
    
    If Trim(Cae.Text) <> "" Then
        Exit Sub
    End If
    
    
    ' Crear objeto interface Web Service Autenticaci�n y Autorizaci�n
    Set WSAA = CreateObject("WSAA")
    
    ' Generar un Ticket de Requerimiento de Acceso (TRA)
    tra = WSAA.CreateTRA()
    Debug.Print tra
    
    ' Especificar la ubicacion de los archivos certificado y clave privada
    Rem Path = CurDir() + "\"
    
    ZPath = "c:\salva\"
    ZNombre = "surfa"
    ZCuit = "30549165083"
            
            
    Rem ZPath = "c:\delia\"
    Rem ZNombre = "delia"
    Rem ZCuit = "27061818087"
            
            
    ' Certificado: certificado es el firmado por la AFIP
    ' ClavePrivada: la clave privada usada para crear el certificado
    Certificado = ZPath + ZNombre + ".crt" ' certificado de prueba
    ClavePrivada = ZPath + ZNombre + ".key" ' clave privada de prueba
    
    ' Generar el mensaje firmado (CMS)
    cms = WSAA.SignTRA(tra, Path + Certificado, Path + ClavePrivada)
    Debug.Print cms
    
    ' Llamar al web service para autenticar:
    Rem ta = WSAA.CallWSAA(cms, "https://wsaahomo.afip.gov.ar/ws/services/LoginCms") ' Hologaci�n
    ta = WSAA.CallWSAA(cms, "https://wsaa.afip.gov.ar/ws/services/LoginCms") ' Producci�n

    ' Imprimir el ticket de acceso, ToKen y Sign de autorizaci�n
    Debug.Print ta
    Debug.Print "Token:", WSAA.Token
    Debug.Print "Sign:", WSAA.Sign
    
    ' Una vez obtenido, se puede usar el mismo token y sign por 6 horas
    ' (este per�odo se puede cambiar)
    
    ' Crear objeto interface Web Service de Factura Electr�nica
    Set WSFE = CreateObject("WSFE")
    ' Setear tocken y sing de autorizaci�n (pasos previos)
    WSFE.Token = WSAA.Token
    WSFE.Sign = WSAA.Sign
    
    ' CUIT del emisor (debe estar registrado en la AFIP)
    WSFE.Cuit = ZCuit
    
    ' Conectar al Servicio Web de Facturaci�n
    Rem ok = WSFE.Conectar("https://wswhomo.afip.gov.ar/wsfe/service.asmx") ' homologaci�n
    Rem ok = WSFE.Conectar("https://wsw.afip.gov.ar/wsfe/service.asmx") ' producci�n
    ok = WSFE.Conectar("https://servicios1.afip.gov.ar/wsfe/service.asmx") ' producci�n
    


    ' Llamo a un servicio nulo, para obtener el estado del servidor (opcional)
    WSFE.Dummy
    Debug.Print "appserver status", WSFE.AppServerStatus
    Debug.Print "dbserver status", WSFE.DbServerStatus
    Debug.Print "authserver status", WSFE.AuthServerStatus
    
    ' Recupera cantidad m�xima de registros (opcional)
    qty = WSFE.RecuperarQty()
    
    ' Recupera �ltimo n�mero de secuencia ID
    LastId = WSFE.UltNro()
    
    ' Recupero �ltimo n�mero de comprobante para un punto de venta y tipo (opcional)
    
    ZZTipo = "03"
    
    If Letra.Text = "A" Then
        Select Case Val(ZZTipo)
            Case 3
                WTipo = "01"
            Case 4
                WTipo = "02"
            Case 5
                WTipo = "03"
            Case Else
        End Select
            Else
        Select Case Val(ZZTipo)
            Case 3
                WTipo = "06"
            Case 4
                WTipo = "07"
            Case 5
                WTipo = "08"
            Case Else
        End Select
    End If
    
    
    tipo_cbte = Val(WTipo)
    punto_vta = 2
    LastCBTE = WSFE.RecuperaLastCMP(punto_vta, tipo_cbte)
    
    Stop
    
    WRazon = "Pellital S.A."
    WCuit = "30-61052459-8"
    Call Eval
    
    
    ' Establezco los valores de la factura o lote a autorizar:
    id = LastId + 1
    presta_serv = 0
    If Len(WCuit) = 11 Then
        tipo_doc = 80
            Else
        tipo_doc = 96
    End If
    nro_doc = Left$(WCuit + Space$(11), 11)
    
    If Val(Numero.Text) <> LastCBTE + 1 Then
        m$ = "El numero de comprobante no es igua al correlativo indicado por la afip " + Str$(LastCBTE + 1)
        A% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
        Exit Sub
    End If

    Rem dada

    cbt_desde = LastCBTE + 1
    cbt_hasta = LastCBTE + 1
    imp_total = "11.21"
    imp_tot_conc = "10"
    imp_neto = "10"
    impto_liq = "1.21"
    impto_liq_rni = "0.00"
    imp_op_ex = "0.00"
    
    ZZFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    Fecha = Format(ZZFecha, "yyyymmdd")
    fecha_cbte = Fecha
    
    ZZFecha = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
    Fecha = Format(ZZFecha, "yyyymmdd")
    fecha_venc_pago = Fecha
    
    ' Fechas del per�odo del servicio facturado (solo si presta_serv = 1)
    ZZFecha = "00000000"
    Fecha = Format(ZZFecha, "yyyymmdd")
    fecha_serv_desde = ZZFecha
    
    ZZFecha = "00000000"
    Fecha = Format(ZZFecha, "yyyymmdd")
    fecha_serv_hasta = ZZFecha
    
    ' Llamo al WebService de Autorizaci�n para obtener el CAE
    Cae = WSFE.Aut(id, presta_serv, _
        tipo_doc, nro_doc, tipo_cbte, punto_vta, _
        cbt_desde, cbt_hasta, imp_total, imp_tot_conc, imp_neto, _
        impto_liq, impto_liq_rni, imp_op_ex, fecha_cbte, fecha_venc_pago, _
        fecha_serv_desde, fecha_serv_hasta) ' si presta_serv = 0 no pasar estas fechas
    
    Debug.Print "Vencimiento ", WSFE.Vencimiento ' Fecha de vencimiento o vencimiento de la autorizaci�n
    Debug.Print "Resultado: ", WSFE.resultado ' A=Aceptado, R=Rechazado
    Debug.Print "Motivo de rechazo o advertencia", WSFE.Motivo ' 00= No hay error
    Debug.Print "Reprocesado?", WSFE.Reproceso ' S=Si, N=No
    
    If WSFE.resultado = "A" Then
        Cae.Text = Cae
        VtoCae.Text = WSFE.Vencimiento
    End If
    
    ' Verifico que no haya rechazo o advertencia al generar el CAE
    If Cae = "" Then
        MsgBox "La p�gina esta caida o la respuesta es inv�lida"
    ElseIf Cae = "NULL" Or WSFE.resultado <> "A" Then
        MsgBox "No se asign� CAE (Rechazado). Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    ElseIf WSFE.Motivo <> "NULL" And WSFE.Motivo <> "00" Then
        MsgBox "Se asign� CAE pero con advertencias. Motivos: " & WSFE.Motivo, vbInformation + vbOKOnly
    End If
    
    ' Imprimo respuesta XML para depuraci�n (errores de formato)
    Debug.Print WSFE.XmlResponse
    
    MsgBox "QTY: " & qty & vbCrLf & "LastId: " & LastId & vbCrLf & "LastCBTE:" & LastCBTE & vbCrLf & "CAE: " & Cae, vbInformation + vbOKOnly
    Exit Sub
ManejoError:
    ' Si hubo error:
    Debug.Print Err.Description            ' descripci�n error afip
    Debug.Print Err.Number - vbObjectError ' codigo error afip
    Select Case MsgBox(Err.Description, vbCritical + vbRetryCancel, "Error:" & Err.Number - vbObjectError & " en " & Err.Source)
        Case vbRetry
            Debug.Assert False
            Resume
        Case vbCancel
            Debug.Print Err.Description
    End Select




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



End Sub
