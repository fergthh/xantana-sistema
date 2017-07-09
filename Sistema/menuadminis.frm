VERSION 5.00
Begin VB.Form MenuAdminis 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Control de Gestion - Administracion  : "
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   660
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "menuadminis.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "menuadminis.frx":0442
   ScaleHeight     =   8310
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.Frame PantaClave 
      Height          =   3015
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox ClaveIII 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox ClaveII 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox ClaveI 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Nueva Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nueva Clave"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Clave Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Image Acepta 
         Height          =   480
         Left            =   1320
         MouseIcon       =   "menuadminis.frx":0884
         MousePointer    =   99  'Custom
         Picture         =   "menuadminis.frx":0B8E
         ToolTipText     =   "Confirma el Proceso"
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   2400
         MouseIcon       =   "menuadminis.frx":0FD0
         MousePointer    =   99  'Custom
         Picture         =   "menuadminis.frx":12DA
         ToolTipText     =   "Menu Principal"
         Top             =   2160
         Width           =   480
      End
   End
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Logo 
      Height          =   1005
      Left            =   8520
      Picture         =   "menuadminis.frx":1B1C
      Top             =   6360
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label VentasII 
      BackColor       =   &H8000000C&
      Caption         =   "Ventas y Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "menuadminis.frx":249A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Ventas y Stock"
      Top             =   7560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   960
      Picture         =   "menuadminis.frx":27A4
      Top             =   120
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.Menu Maestros 
      Caption         =   "Maestros"
      Begin VB.Menu Cuentas 
         Caption         =   "Ingreso de Cuentas Contables"
      End
      Begin VB.Menu Prove 
         Caption         =   "Ingreso de Proveedores"
      End
      Begin VB.Menu Concepto 
         Caption         =   "Ingreso de Conceptos de Compras"
      End
      Begin VB.Menu Banco 
         Caption         =   "Ingeso de Bancos"
      End
   End
   Begin VB.Menu fgnhghj 
      Caption         =   "Novedades"
      Begin VB.Menu Compras 
         Caption         =   "Ingreso de Comprobantes de Proveedores"
      End
      Begin VB.Menu Pago 
         Caption         =   "Ingreso de Pagos"
      End
      Begin VB.Menu recibo 
         Caption         =   "Ingreso de Cobranzas"
      End
      Begin VB.Menu Deposito 
         Caption         =   "Ingreso de Depositos"
      End
      Begin VB.Menu Gastatosbancarios 
         Caption         =   "Ingresos de Debitos y Creditos Bancarios"
      End
      Begin VB.Menu GastosCaja 
         Caption         =   "Ingresos de Gastos"
      End
      Begin VB.Menu Transferencia 
         Caption         =   "Ingreso de Transferencias"
      End
      Begin VB.Menu citicompras 
         Caption         =   "Generacion del CITI COMPRAS"
      End
   End
   Begin VB.Menu sdfsdfsdf 
      Caption         =   "Listados"
      Begin VB.Menu ccprv 
         Caption         =   "Listado de Cuenta Corrientes de Proveedoes"
      End
      Begin VB.Menu ccprv1 
         Caption         =   "Consulta de Cuenta Corriente de Proveedores"
      End
      Begin VB.Menu proyprv 
         Caption         =   "Proyeccion de Pagos"
      End
      Begin VB.Menu ccprvfecha 
         Caption         =   "Listado de Cuenta Corriente de Proveedores a Fecha"
      End
      Begin VB.Menu ListaPagosFceha 
         Caption         =   "Listado de Pagos por Fecha"
      End
      Begin VB.Menu ListaPagosProve 
         Caption         =   "Listado de Pagos por Proveedor"
      End
      Begin VB.Menu Subduariopagos 
         Caption         =   "Subdiario de Pagos"
      End
      Begin VB.Menu Listreci 
         Caption         =   "Listado de Cobranzas"
      End
      Begin VB.Menu MovCaja 
         Caption         =   "Subdiario de Caja"
      End
      Begin VB.Menu Movvan 
         Caption         =   "Listado de Movimientos Bancarios"
      End
      Begin VB.Menu cartera 
         Caption         =   "Listado de Cheques en Cartera"
      End
      Begin VB.Menu carteracliente 
         Caption         =   "Listado de Cheques en Cartera por Cliente"
      End
      Begin VB.Menu cateranumero 
         Caption         =   "Listado de Cheques en Cartera por Nro de Cheque"
      End
      Begin VB.Menu caretarprove 
         Caption         =   "Listado de Cheques en Cartera por Proveedor"
      End
      Begin VB.Menu caretaringreso 
         Caption         =   "Listado de Cheques en Cartera por Fecha de Ingreso"
      End
      Begin VB.Menu Ivacomp 
         Caption         =   "Iva Compras"
      End
      Begin VB.Menu compcon 
         Caption         =   "Listado de Compras por concepto"
      End
      Begin VB.Menu Impcyb 
         Caption         =   "Listado de Imputaciones Contables"
      End
      Begin VB.Menu Listacaja 
         Caption         =   "Listado de Caja"
      End
      Begin VB.Menu compconconsol 
         Caption         =   "Listado de Compras por concepto Consolidado"
      End
   End
   Begin VB.Menu fgh 
      Caption         =   "Procesos"
      Begin VB.Menu Escalas 
         Caption         =   "Ingreso de Configuracion del Sistema"
      End
      Begin VB.Menu parametro 
         Caption         =   "Ingreso de Datos de la Empresa"
      End
      Begin VB.Menu lledatos 
         Caption         =   "Traspaso de Datos"
         Visible         =   0   'False
      End
      Begin VB.Menu Calculadora 
         Caption         =   "Calculadora"
         Visible         =   0   'False
      End
      Begin VB.Menu Empre 
         Caption         =   "Cambio de Empresa"
         Visible         =   0   'False
      End
      Begin VB.Menu CambiaClave 
         Caption         =   "Cambio de Clave"
      End
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "MenuAdminis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cash_Click()
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgCash.Show
End Sub


Private Sub Banco_Click()
    PrgBanco.Show
End Sub

Private Sub Cancela_click()
    PantaClave.Visible = False
End Sub

Private Sub CambiaClave_Click()
    PantaClave.Visible = True
    ClaveI.Text = ""
    ClaveII.Text = ""
    ClaveIII.Text = ""
End Sub


Private Sub citicompras_Click()
    PrgCitinuevoCompras.Show
End Sub

Private Sub ClaveI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ClaveII.SetFocus
    End If
    If KeyAscii = 27 Then
        ClaveI.Text = ""
    End If
End Sub

Private Sub ClaveII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ClaveIII.SetFocus
    End If
    If KeyAscii = 27 Then
        ClaveII.Text = ""
    End If
End Sub

Private Sub ClaveIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ClaveI.SetFocus
    End If
    If KeyAscii = 27 Then
        ClaveIII.Text = ""
    End If
End Sub

Private Sub Acepta_Click()
        
    If ClaveII.Text <> ClaveIII.Text Or ClaveII.Text = "" Then
        Exit Sub
    End If
        
    If ZZNivel = 0 Then
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + ClaveI.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZZZZoperador = rstOperador!operador
            rstOperador.Close
            If Val(ZZOperador) <> ZZZZoperador Then
                Exit Sub
            End If
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Operador SET "
        ZSql = ZSql + " Clave = " + "'" + ClaveII.Text + "'"
        ZSql = ZSql + " Where Operador = " + "'" + ZZOperador + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        WEmpresa = "1"
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            
        ZZZZoperador = 0
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.ClaveII = " + "'" + ClaveI.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZZZZoperador = rstOperador!operador
            rstOperador.Close
        End If
        
        
        If Val(ZZOperador) <> ZZZZoperador Then
            WEmpresa = "2"
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            Exit Sub
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Operador SET "
        ZSql = ZSql + " ClaveII = " + "'" + ClaveII.Text + "'"
        ZSql = ZSql + " Where Operador = " + "'" + ZZOperador + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            
    End If

    PantaClave.Visible = False

    WEmpresa = "2"
    Rem txtUserName = "SA"
    Rem txtPassword = "Sw58125812"
    
    txtOdbc = "FraganciasII"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

End Sub





Private Sub caretaringreso_Click()
    PrgValcarFecha.Show
End Sub

Private Sub caretarprove_Click()
    PrgValcarProveedor.Show
End Sub

Private Sub cartera_Click()
    PrgValcar.Show
End Sub

Private Sub carteracliente_Click()
    PrgValcarCliente.Show
End Sub

Private Sub cateranumero_Click()
    PrgConsultaCheque.Show
End Sub

Private Sub ccprv_Click()
    PrgCcprv.Show
End Sub

Private Sub ccprv1_Click()
    PrgCcprv1.Show
End Sub

Private Sub ccprvfecha_Click()
    PrgCcprvFecha.Show
End Sub

Private Sub compcon_Click()
    PrgCompcon.Show
End Sub

Private Sub compconconsol_Click()
    PrgCompconConsol.Show
End Sub

Private Sub Compras_Click()
    PrgCompras.Show
End Sub

Private Sub Concepto_Click()
    PrgConcepto.Show
End Sub

Private Sub Cuentas_Click()
    PrgCuenta.Show
End Sub

Private Sub Agenda_Click()
    OPEN_FILE_Agenda
    PrgAgenda.Show
End Sub

Private Sub Aplica_Click()
    PrgAplica.Show
End Sub

Private Sub ActalizaCodigo_Click()
    prgActualizacionIndividual.Show
End Sub

Private Sub ActualizaCostoFuturo_Click()
    prgActualizacionCostoFuturo.Show
End Sub

Private Sub ActualizaGeneral_Click()
    prgActualizacionGeneral.Show
End Sub

Private Sub Calculadora_Click()
    Calculator.Show
End Sub

Private Sub Cotiza_Click()
    OPEN_FILE_Cotiza
    OPEN_FILE_Configuracion
    PrgCotiza.Show
End Sub

Private Sub Calidad_Click()
    PrgCalidad.Show
End Sub

Private Sub Command1_Click()






Stop


    Open "Clientes.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCliente = Trim(Mid$(WDato, 1, 10))
        WRazon = Mid$(WDato, 11, 50)
        WFantasia = Mid$(WDato, 61, 30)
        WIb = Mid$(WDato, 91, 20)
        WCuit = Mid$(WDato, 111, 20)
        WGanancia = Mid$(WDato, 131, 20)
        WCalle = Mid$(WDato, 151, 30)
        WNumero = Mid$(WDato, 181, 8)
        WExtension = Mid$(WDato, 189, 15)
        WPostal = Mid$(WDato, 204, 8)
        WTelefono = Mid$(WDato, 212, 30)
        WFax = Mid$(WDato, 242, 20)
        WEmail = Mid$(WDato, 262, 50)
        WCateIva = Mid$(WDato, 312, 2)
        WCateGana = Mid$(WDato, 314, 2)
        WCateIb = Mid$(WDato, 316, 2)
        WFechaAlta = Mid$(WDato, 318, 8)
        WFechaBaja = Mid$(WDato, 326, 8)
        WTipoFactu = Mid$(WDato, 334, 1)
        WCodCta = Mid$(WDato, 335, 9)
        WCtaDeuda = Mid$(WDato, 344, 9)
        WCondVta = Mid$(WDato, 353, 2)
        WContactoI = Mid$(WDato, 355, 30)
        WTelefonoI = Mid$(WDato, 385, 30)
        WPuestoI = Mid$(WDato, 415, 15)
        WContactoII = Mid$(WDato, 430, 30)
        WTelefonoII = Mid$(WDato, 460, 15)
        WPuestoII = Mid$(WDato, 475, 30)
        WContactoIII = Mid$(WDato, 505, 30)
        WTelefonoIII = Mid$(WDato, 535, 15)
        WPuestoIII = Mid$(WDato, 550, 30)
        WM1 = Mid$(WDato, 580, 10)
        WM2 = Mid$(WDato, 590, 10)
        WM3 = Mid$(WDato, 600, 10)
        WM4 = Mid$(WDato, 610, 10)
        WM5 = Mid$(WDato, 620, 10)
        WM6 = Mid$(WDato, 630, 10)
        WM7 = Mid$(WDato, 640, 10)
        WM8 = Mid$(WDato, 650, 10)
        WM9 = Mid$(WDato, 660, 10)
        WM10 = Mid$(WDato, 670, 10)
        WM11 = Mid$(WDato, 680, 10)
        WM12 = Mid$(WDato, 690, 10)
        WExpreso = Mid$(WDato, 700, 40)
        WLocalidad = Mid$(WDato, 740, 30)
        WProvincia = Mid$(WDato, 770, 30)
        WCabeza = Mid$(WDato, 800, 1)
        WPatro = Mid$(WDato, 801, 10)
        WCodGrupo = Mid$(WDato, 811, 20)
        WFra = Mid$(WDato, 831, 2)
        WPlan = Mid$(WDato, 833, 5)
        WLista = Mid$(WDato, 838, 3)
        WPP = Mid$(WDato, 841, 1)
        Rem WAviso = Mid$(WDato, 842, 4)
        WBonifica = Mid$(WDato, 842, 5)
        Rem WExporta = Mid$(WDato, 851, 4)
        WCalleII = Mid$(WDato, 847, 30)
        WNumeroII = Mid$(WDato, 877, 8)
        WExtensionII = Mid$(WDato, 885, 15)
        Wa4 = Mid$(WDato, 900, 8)
        Wa5 = Mid$(WDato, 908, 30)
        Wa6 = Mid$(WDato, 938, 30)
        WTipoClie = Mid$(WDato, 968, 2)
        WPotencial = Mid$(WDato, 970, 1)
        WConocio = Mid$(WDato, 971, 50)
        
        ZZDireccion = Trim(WCalle) + " " + Trim(WNumero) + " " + Trim(WExtension)
        ZZDireccionII = Trim(WCalleII) + " " + Trim(WNumeroII) + " " + Trim(WExtensionII)
        
        Select Case WCateIva
            Case "MO"
                WIva = "5"
            Case "CF"
                WIva = "3"
            Case "Ex"
                WIva = "4"
            Case Else
                WIva = "1"
        End Select
        
        ZZFechaAlta = Right$(WFechaAlta, 2) + "/" + Mid$(WFechaAlta, 5, 2) + "/" + Left$(WFechaAlta, 4)
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            rstCliente.Close
            
            
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
            ZSql = ZSql + "NombreI ,"
            ZSql = ZSql + "TelefonoI ,"
            ZSql = ZSql + "EmailI ,"
            ZSql = ZSql + "NombreII ,"
            ZSql = ZSql + "TelefonoII ,"
            ZSql = ZSql + "EmailII ,"
            ZSql = ZSql + "NombreIII ,"
            ZSql = ZSql + "TelefonoIII ,"
            ZSql = ZSql + "EmailIII ,"
            ZSql = ZSql + "Fantasia ,"
            ZSql = ZSql + "DireccionII ,"
            ZSql = ZSql + "FechaAlta ,"
            ZSql = ZSql + "Cuit ,"
            ZSql = ZSql + "Email ,"
            ZSql = ZSql + "Fax ,"
            ZSql = ZSql + "PorceIva ,"
            ZSql = ZSql + "Provincia ,"
            ZSql = ZSql + "Iva ,"
            ZSql = ZSql + "Expreso ,"
            ZSql = ZSql + "TipoClie ,"
            ZSql = ZSql + "NroLista ,"
            ZSql = ZSql + "Condicion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCliente + "',"
            ZSql = ZSql + "'" + WRazon + "',"
            ZSql = ZSql + "'" + ZZDireccion + "',"
            ZSql = ZSql + "'" + WLocalidad + "',"
            ZSql = ZSql + "'" + WPostal + "',"
            ZSql = ZSql + "'" + WTelefono + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WContactoI + "',"
            ZSql = ZSql + "'" + WTelefonoI + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WContactoII + "',"
            ZSql = ZSql + "'" + WTelefonoII + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WContactoIII + "',"
            ZSql = ZSql + "'" + WTelefonoIII + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WFantasia + "',"
            ZSql = ZSql + "'" + ZZDireccionII + "',"
            ZSql = ZSql + "'" + ZZFechaAlta + "',"
            ZSql = ZSql + "'" + WCuit + "',"
            ZSql = ZSql + "'" + WEmail + "',"
            ZSql = ZSql + "'" + WFax + "',"
            ZSql = ZSql + "'" + WPlan + "',"
            ZSql = ZSql + "'" + "0" + "',"
            ZSql = ZSql + "'" + WIva + "',"
            ZSql = ZSql + "'" + WExpreso + "',"
            ZSql = ZSql + "'" + WTipoClie + "',"
            ZSql = ZSql + "'" + WLista + "',"
            ZSql = ZSql + "'" + "" + "')"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
        
    Loop
    
    Close #1








Stop


    Open "precios.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WLista = Trim(Mid$(WDato, 1, 3))
        WCodigo = Mid$(WDato, 4, 15)
        WOrdDesde = Mid$(WDato, 19, 8)
        WOrdHasta = Mid$(WDato, 27, 8)
        WTope1 = Str$(Val(Mid$(WDato, 35, 8)))
        WValor1 = Str$(Val(Mid$(WDato, 43, 7)))
        WTope2 = Str$(Val(Mid$(WDato, 50, 8)))
        WValor2 = Str$(Val(Mid$(WDato, 58, 7)))
        WTope3 = Str$(Val(Mid$(WDato, 65, 8)))
        WValor3 = Str$(Val(Mid$(WDato, 73, 7)))
        WTope4 = Str$(Val(Mid$(WDato, 80, 8)))
        WValor4 = Str$(Val(Mid$(WDato, 88, 7)))
        
        WDesde = Right$(WOrdDesde, 2) + "/" + Mid$(WOrdDesde, 5, 2) + "/" + Left$(WOrdDesde, 4)
        WHasta = Right$(WOrdHasta, 2) + "/" + Mid$(WOrdHasta, 5, 2) + "/" + Left$(WOrdHasta, 4)
            
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
        
        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
        WClave = Trim(WLista) + ZZCodigo
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + WLista + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + WLinea + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + WFragancia + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + WCalidad + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + WTamano + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
        
            rstPrecios.Close
                
            ZZLugar = ZZLugar + 1
            dada.Text = ZZLugar
            dadaII.Text = ZZCodigo
            
            DoEvents
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Precios SET "
            ZSql = ZSql + " Desde = " + "'" + WDesde + "',"
            ZSql = ZSql + " Hasta = " + "'" + WHasta + "',"
            ZSql = ZSql + " OrdDesde = " + "'" + WOrdDesde + "',"
            ZSql = ZSql + " OrdHasta = " + "'" + WOrdHasta + "',"
            ZSql = ZSql + " Tope1 = " + "'" + WTope1 + "',"
            ZSql = ZSql + " Valor1 = " + "'" + WValor1 + "',"
            ZSql = ZSql + " Tope2 = " + "'" + WTope2 + "',"
            ZSql = ZSql + " Valor2 = " + "'" + WValor2 + "',"
            ZSql = ZSql + " Tope3 = " + "'" + WTope3 + "',"
            ZSql = ZSql + " Valor3 = " + "'" + WValor3 + "',"
            ZSql = ZSql + " Tope4 = " + "'" + WTope4 + "',"
            ZSql = ZSql + " Valor4 = " + "'" + WValor4 + "'"
            ZSql = ZSql + " Where Lista = " + "'" + WLista + "'"
            ZSql = ZSql + " and LInea = " + "'" + WLinea + "'"
            ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Fragancia = " + "'" + WFragancia + "'"
            ZSql = ZSql + " and Calidad = " + "'" + WCalidad + "'"
            ZSql = ZSql + " and Tamano = " + "'" + WTamano + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZZLugar = ZZLugar + 1
            dada.Text = ZZLugar
            dadaII.Text = ZZCodigo
            
            DoEvents
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Precios ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Fragancia ,"
            ZSql = ZSql + "Calidad ,"
            ZSql = ZSql + "Tamano ,"
            ZSql = ZSql + "Lista ,"
            ZSql = ZSql + "Desde ,"
            ZSql = ZSql + "Hasta ,"
            ZSql = ZSql + "OrdDesde ,"
            ZSql = ZSql + "OrdHasta ,"
            ZSql = ZSql + "Tope1 ,"
            ZSql = ZSql + "Valor1 ,"
            ZSql = ZSql + "Tope2 ,"
            ZSql = ZSql + "Valor2 ,"
            ZSql = ZSql + "Tope3 ,"
            ZSql = ZSql + "Valor3 ,"
            ZSql = ZSql + "Tope4 ,"
            ZSql = ZSql + "Valor4 )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + ZZCodigo + "',"
            ZSql = ZSql + "'" + WLinea + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + WFragancia + "',"
            ZSql = ZSql + "'" + WCalidad + "',"
            ZSql = ZSql + "'" + WTamano + "',"
            ZSql = ZSql + "'" + WLista + "',"
            ZSql = ZSql + "'" + WDesde + "',"
            ZSql = ZSql + "'" + WHasta + "',"
            ZSql = ZSql + "'" + WOrdDesde + "',"
            ZSql = ZSql + "'" + WOrdHasta + "',"
            ZSql = ZSql + "'" + WTope1 + "',"
            ZSql = ZSql + "'" + WValor1 + "',"
            ZSql = ZSql + "'" + WTope2 + "',"
            ZSql = ZSql + "'" + WValor2 + "',"
            ZSql = ZSql + "'" + WTope3 + "',"
            ZSql = ZSql + "'" + WValor3 + "',"
            ZSql = ZSql + "'" + WTope4 + "',"
            ZSql = ZSql + "'" + WValor4 + "')"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Loop
    
    Close #1






Stop


    Open "producto.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 1, 15)
        WDescripcion = Mid$(WDato, 16, 50)
        WDescripcionII = Mid$(WDato, 66, 20)
        WWFacturable = Mid$(WDato, 86, 1)
        WTipo = Mid$(WDato, 87, 2)
        WImportado = Mid$(WDato, 89, 1)
        WEstado = Mid$(WDato, 90, 1)
        WWFechaInactivo = Mid$(WDato, 91, 8)
        WSucursal = Mid$(WDato, 99, 2)
        WAreaSol = Mid$(WDato, 101, 2)
        WAreaRea = Mid$(WDato, 103, 2)
        WStock = Mid$(WDato, 105, 1)
        WSector = Mid$(WDato, 106, 2)
        WComision = Mid$(WDato, 108, 1)
        WWEtiqueta = Mid$(WDato, 109, 1)
        WInsumo = Mid$(WDato, 110, 10)
        WCodCombo = Mid$(WDato, 120, 8)
        WCosto = Mid$(WDato, 128, 10)
        WFechaCosto = Mid$(WDato, 138, 8)
            
        WStock = "0"
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
        
        If WEstado = "I" Then
            WActivo = "1"
            WFechaInactivo = Right$(WFechaInactivo, 2) + "/" + Mid$(WFechaInactivo, 5, 2) + "/" + Left$(WFechaInactivo, 4)
                Else
            WActivo = "0"
            WFechaInactivo = "  /  /    "
        End If
        
        If WWFacturable = "N" Then
            WFacturable = "1"
                Else
            WFacturable = "0"
        End If
        
        If WWEtiqueta = "N" Then
            WEtiqueta = "2"
                Else
            If WWEtiqueta = "N" Then
                WEtiqueta = "1"
                    Else
                WEtiqueta = "0"
            End If
        End If
                
        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
            
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.LInea = " + "'" + WLinea + "'"
        ZSql = ZSql + " and Articulo.Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Articulo.fragancia = " + "'" + WFragancia + "'"
        ZSql = ZSql + " and Articulo.Calidad = " + "'" + WCalidad + "'"
        ZSql = ZSql + " and Articulo.Tamano = " + "'" + WTamano + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            rstArticulo.Close
                
            ZZLugar = ZZLugar + 1
            dada.Text = ZZLugar
            dadaII.Text = ZZCodigo
            
            DoEvents
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "',"
            ZSql = ZSql + " DescripcionII = " + "'" + WDescripcionII + "',"
            ZSql = ZSql + " Stock = " + "'" + WStock + "',"
            ZSql = ZSql + " Insumo = " + "'" + WInsumo + "',"
            ZSql = ZSql + " Sector = " + "'" + WSector + "',"
            ZSql = ZSql + " Activo = " + "'" + WActivo + "',"
            ZSql = ZSql + " FechaInactivo = " + "'" + WFechaInactivo + "',"
            ZSql = ZSql + " Facturable = " + "'" + WFacturable + "',"
            ZSql = ZSql + " Etiqueta = " + "'" + WEtiqueta + "'"
            ZSql = ZSql + " Where LInea = " + "'" + WLinea + "'"
            ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Fragancia = " + "'" + WFragancia + "'"
            ZSql = ZSql + " and Calidad = " + "'" + WCalidad + "'"
            ZSql = ZSql + " and Tamano = " + "'" + WTamano + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZZLugar = ZZLugar + 1
            dada.Text = ZZLugar
            dadaII.Text = ZZCodigo
            
            DoEvents
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Articulo ("
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Fragancia ,"
            ZSql = ZSql + "Calidad ,"
            ZSql = ZSql + "Tamano ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "DescripcionII ,"
            ZSql = ZSql + "Stock ,"
            ZSql = ZSql + "Insumo ,"
            ZSql = ZSql + "Sector ,"
            ZSql = ZSql + "Activo ,"
            ZSql = ZSql + "FechaInactivo ,"
            ZSql = ZSql + "Facturable ,"
            ZSql = ZSql + "Etiqueta )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WLinea + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + WFragancia + "',"
            ZSql = ZSql + "'" + WCalidad + "',"
            ZSql = ZSql + "'" + WTamano + "',"
            ZSql = ZSql + "'" + ZZCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "',"
            ZSql = ZSql + "'" + WDescripcionII + "',"
            ZSql = ZSql + "'" + WStock + "',"
            ZSql = ZSql + "'" + WInsumo + "',"
            ZSql = ZSql + "'" + WSector + "',"
            ZSql = ZSql + "'" + WActivo + "',"
            ZSql = ZSql + "'" + WFechaInactivo + "',"
            ZSql = ZSql + "'" + WFacturable + "',"
            ZSql = ZSql + "'" + WEtiqueta + "')"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Loop
    
    Close #1



Stop








    Open "lineas.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 3)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Lineas"
        ZSql = ZSql + " Where Lineas.Codigo = " + "'" + WCodigo + "'"
        spLinea = ZSql
        Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstLinea.RecordCount > 0 Then
            rstLinea.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Lineas SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Lineas ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Loop
    
    Close #1



Stop



    GoTo da:








    Open "tipo.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoPro"
        ZSql = ZSql + " Where TipoPro.Codigo = " + "'" + WCodigo + "'"
        spTipoPro = ZSql
        Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoPro.RecordCount > 0 Then
            rstTipoPro.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE TipoPro SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO TipoPro ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1







    Open "fragancia.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Fragancia"
        ZSql = ZSql + " Where Fragancia.Codigo = " + "'" + WCodigo + "'"
        spFragancia = ZSql
        Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
        If rstFragancia.RecordCount > 0 Then
            rstFragancia.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Fragancia SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spFragancia = ZSql
            Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Fragancia ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spFragancia = ZSql
            Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1





    Open "calidad.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Calidad"
        ZSql = ZSql + " Where Calidad.Codigo = " + "'" + WCodigo + "'"
        spCalidad = ZSql
        Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
        If rstCalidad.RecordCount > 0 Then
            rstCalidad.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Calidad SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spCalidad = ZSql
            Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Calidad ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spCalidad = ZSql
            Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1













    Open "tamaño.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Tamano"
        ZSql = ZSql + " Where Tamano.Codigo = " + "'" + WCodigo + "'"
        spTamano = ZSql
        Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
        If rstTamano.RecordCount > 0 Then
            rstTamano.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Tamano SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spTamano = ZSql
            Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Tamano ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spTamano = ZSql
            Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1

















    Open "Sector.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Sector"
        ZSql = ZSql + " Where Sector.Codigo = " + "'" + WCodigo + "'"
        spSector = ZSql
        Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        If rstSector.RecordCount > 0 Then
            rstSector.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Sector SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Sector ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1






    Open "Insumos.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WProveedor = Mid$(WDato, 1, 8)
        WCodigo = Mid$(WDato, 9, 16)
        WDescripcion = Mid$(WDato, 25, 30)
        WMOneda = Mid$(WDato, 55, 1)
        WCosto = Mid$(WDato, 56, 10)
        WOrdFecha = Mid$(WDato, 66, 8)
        WTipo = Mid$(WDato, 74, 1)
        WFechaCosto = "  /  /    "
        WOrdFechaCosto = "00000000"
        
        WFecha = Mid$(WOrdFecha, 7, 2) + "/" + Mid(WOrdFecha, 5, 2) + "/" + Mid(WOrdFecha, 1, 4)
        
        If WMOneda = "P" Then
            WMOneda = "1"
                Else
            WMOneda = "2"
        End If
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Insumo"
        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + WCodigo + "'"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        If rstInsumo.RecordCount > 0 Then
        
            rstInsumo.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Insumo SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "',"
            ZSql = ZSql + " Linea = " + "'" + WTipo + "',"
            ZSql = ZSql + " Proveedor = " + "'" + WProveedor + "',"
            ZSql = ZSql + " Costo = " + "'" + WCosto + "',"
            ZSql = ZSql + " MOneda = " + "'" + WMOneda + "',"
            ZSql = ZSql + " FechaCosto = " + "'" + WFechaCosto + "',"
            ZSql = ZSql + " OrdFechaCosto = " + "'" + WOrdFechaCosto + "',"
            ZSql = ZSql + " Minimo = " + "'" + "" + "',"
            ZSql = ZSql + " Stock = " + "'" + "" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Insumo ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Costo ,"
            ZSql = ZSql + "Moneda ,"
            ZSql = ZSql + "FechaCosto ,"
            ZSql = ZSql + "OrdFechaCosto ,"
            ZSql = ZSql + "Minimo ,"
            ZSql = ZSql + "Stock )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WCosto + "',"
            ZSql = ZSql + "'" + WMOneda + "',"
            ZSql = ZSql + "'" + WFechaCosto + "',"
            ZSql = ZSql + "'" + WOrdFechaCosto + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "')"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
    Loop
    
    Close #1


da:


    Open "proveedores.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WProveedor = Mid$(WDato, 1, 8)
        WRazon = Mid$(WDato, 9, 50)
        WComercial = Mid$(WDato, 59, 30)
        WNroIb = Mid$(WDato, 89, 20)
        WCuit = Mid$(WDato, 109, 20)
        WGananacia = Mid$(WDato, 129, 20)
        WCalle = Mid$(WDato, 149, 30)
        WNro = Mid$(WDato, 179, 8)
        WExtension = Mid$(WDato, 187, 15)
        WPostal = Mid$(WDato, 202, 8)
        WTelefono = Mid$(WDato, 210, 30)
        WFax = Mid$(WDato, 240, 20)
        WEmail = Mid$(WDato, 260, 50)
        WDireccion = Trim(WCalle) + " " + WNro
        
        WLocalidad = ""
        WDias = "0"
        WNombreCheque = ""
        WGamancia = "0"
        WIva = "0"
        WProvincia = "0"
        WObservaciones = ""
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            rstProveedor.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + " Nombre = " + "'" + WRazon + "',"
            ZSql = ZSql + " Direccion = " + "'" + WDireccion + "',"
            ZSql = ZSql + " Localidad = " + "'" + WLocalidad + "',"
            ZSql = ZSql + " Postal = " + "'" + WPostal + "',"
            ZSql = ZSql + " Cuit = " + "'" + WCuit + "',"
            ZSql = ZSql + " Telefono = " + "'" + WTelefono + "',"
            ZSql = ZSql + " EMail = " + "'" + WEmail + "',"
            ZSql = ZSql + " Observaciones = " + "'" + WObservaciones + "',"
            ZSql = ZSql + " Dias = " + "'" + WDias + "',"
            ZSql = ZSql + " Ganancia = " + "'" + WGanancia + "',"
            ZSql = ZSql + " Iva = " + "'" + WIva + "',"
            ZSql = ZSql + " Provincia = " + "'" + WProvincia + "',"
            ZSql = ZSql + " NombreCheque = " + "'" + WNombreCheque + "'"
            ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                Else
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
            ZSql = ZSql + "'" + WRazon + "',"
            ZSql = ZSql + "'" + WDireccion + "',"
            ZSql = ZSql + "'" + WLocalidad + "',"
            ZSql = ZSql + "'" + WPostal + "',"
            ZSql = ZSql + "'" + Left$(WCuit, 15) + "',"
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
        End If
        
        
    Loop
    
    Close #1

















End Sub

Private Sub CondPago_Click()
    PrgCondPago.Show
End Sub

Private Sub ctactefecha_Click()
    PrgCtaCtefecha.Show
End Sub

Private Sub CtaCteCampana_Click()
    PrgCtaCteCampana.Show
End Sub

Private Sub ConsultaPrecios_Click()
    prgConsultaArticulo.Show
End Sub

Private Sub Despacho_Click()
    PrgDespacho.Show
End Sub

Private Sub ControlPedidos_Click()
    PrgControlPedido.Show
End Sub

Private Sub Deposito_Click()
    PrgDeposito.Show
End Sub

Private Sub Empre_Click()
    Empresa.Show
End Sub

Private Sub Escalas_Click()
    PrgEscalas.Show
End Sub

Private Sub Cliente_Click()
    prgcliente.Show
End Sub

Private Sub Comiven_Click()
    PrgVentVend.Show
End Sub

Private Sub ComparaVen_Click()
    PrgComparaVen.Show
End Sub

Private Sub CompArt_Click()
    PrgCompArt.Show
End Sub

Private Sub CtaCte_Click()
    PrgCtaCte.Show
End Sub

Private Sub ctacte2_Click()
    PrgCtaCte1.Show
End Sub

Private Sub CtacteVen_Click()
    PrgCtaCteVen.Show
End Sub

Private Sub devol_Click()
    PrgDevol.Show
End Sub

Private Sub Esencia_Click()
    PrgEsencia.Show
End Sub

Private Sub Expreso_Click()
    PrgExpreso.Show
End Sub

Private Sub factu_Click()
    PrgFactura.Show
End Sub

Private Sub FactuExpo_Click()
    OPEN_FILE_Configuracion
    WVarios = 1
    PrgFactuExpo.Show
End Sub

Private Sub Fallados_Click()
    PrgFallados.Show
End Sub

Private Sub Factura_Click()

End Sub

Private Sub Fin_Click()
    Rem Menu.WindowState = 1
    Rem a = Shell("Sistema.exe", 1)
    Close
    End
End Sub

Private Sub Form_Load()
    If ZZNivel = "0" Then
        MenuAdminis.Caption = "Sistema de Administracion : Mc Fragancias S.A."
            Else
        MenuAdminis.Caption = "Sistema de Administracion : Mc Fragancias"
    End If
End Sub

Private Sub Fragancia_Click()
    PrgFragancia.Show
End Sub

Private Sub GuiaTRansporte_Click()
    PrgGuiaTransporte.Show
End Sub

Private Sub HistorialCliente_Click()
    PrgHistorialCliente.Show
End Sub

Private Sub Insumo_Click()
    prgInsumo.Show
End Sub

Private Sub Gastatosbancarios_Click()
    PrgGastosBancarios.Show
End Sub

Private Sub Linea_Click()
    PrgLinea.Show
End Sub

Private Sub ListaClienteProvincia_Click()
    PrgListaClieProvincia.Show
End Sub

Private Sub ListaClienteVendedor_Click()
    PrgListaClieVende.Show
End Sub

Private Sub ListaClienteZona_Click()
    PrgListaClieZona.Show
End Sub

Private Sub ListaCostoProveedor_Click()
    PrgListaCostoProveedor.Show
End Sub

Private Sub listageneral_Click()
    PrgListaNovedadesPrecio.Show
End Sub

Private Sub ListaStockMinimo_Click()
    PrgListaStockMinimo.Show
End Sub

Private Sub ListaStockProve_Click()
    PrgListaStockProveedor.Show
End Sub

Private Sub Lista_Click()
    PrgLista.Show
End Sub

Private Sub ListaStockValora_Click()
    PrgListaStockValora.Show
End Sub

Private Sub ListaStocvGrupo_Click()
    PrgListaStockGrupo.Show
End Sub

Private Sub ListaventasSemestre_Click()
    PrgListaEvolucionVentas.Show
End Sub

Private Sub NotasEnvio_Click()
    PrgNotaEnvio.Show
End Sub

Private Sub PasajecostoFuturo_Click()
    prgPasajeCostoFuturo.Show
End Sub

Private Sub GastosCaja_Click()
    PrgGastosCaja.Show
End Sub

Private Sub Impcyb_Click()
    PrgImpcyb.Show
End Sub

Private Sub IvaComp_Click()
    PrgIvacomp.Show
End Sub

Private Sub Listacaja_Click()
    PrgMovCajaOtro.Show
End Sub

Private Sub ListaPagosFceha_Click()
    PrgListaPagosFecha.Show
End Sub

Private Sub ListaPagosProve_Click()
    PrgListaPagosProve.Show
End Sub

Private Sub Listreci_Click()
    PrgListreci.Show
End Sub

Private Sub MovCaja_Click()
    PrgMovCaja.Show
End Sub

Private Sub Movvan_Click()
    PrgMovban.Show
End Sub

Private Sub Pago_Click()
    Prgpago.Show
End Sub

Private Sub Prove_Click()
    ZZSistema = 1
    PrgProve.Show
End Sub

Private Sub IngrePtoVend_Click()
    PrgIngrePtoVend.Show
End Sub

Private Sub IvaCompo_Click()
    PrgIvaCompo.Show
End Sub

Private Sub Ivaven_Click()
    PrgIvaven.Show
End Sub

Private Sub Lineas_Click()
    PrgZona.Show
End Sub

Private Sub ListaClieVende_Click()
    PrgListaClieVende.Show
End Sub

Private Sub Listapre_Click()
    PrgPrecioFam.Show
End Sub

Private Sub ListCoti_Click()
    PrgListCoti.Show
End Sub

Private Sub Listmov_Click()
    PrgListMov.Show
End Sub

Private Sub ListPedArt_Click()
    PrgListPedaRT.Show
End Sub

Private Sub ListPedCli_Click()
    PrgListPedCli.Show
End Sub

Private Sub Minimo_Click()
    PrgMinimo.Show
End Sub

Private Sub MovStk_Click()
    PrgMovStk.Show
End Sub

Private Sub lledatos_Click()
    PrgLeeDatos.Show
End Sub

Private Sub MOvvar_Click()
    PrgMovStk.Show
End Sub

Private Sub OrdenFactura_Click()
    PrgOrdenFactura.Show
End Sub

Private Sub parametro_Click()
    PrgParametro.Show
End Sub

Private Sub Pedido_Click()
    PrgPedido.Show
End Sub

Private Sub Plantilla_Click()
    PrgPlantilla.Show
End Sub

Private Sub Prod_Click()
    prgArticulo.Show
End Sub

Private Sub Proyec_Click()
    PrgProyCta.Show
End Sub

Private Sub Proyecto_Click()
    PrgProyecto.Show
End Sub

Private Sub Tipopro_Click()
    PrgTipoPro.Show
End Sub

Private Sub Produccion_Click()
    PrgProduccion.Show
End Sub

Private Sub proyprv_Click()
    PrgProyPrv.Show
End Sub

Private Sub Recibo_Click()
    ZZSistema = 1
    PrgRecibos.Show
End Sub

Private Sub recibos_Click()
    PrgRecibos.Show
End Sub

Private Sub SaldosCta_Click()
    PrgSaldoCta.Show
End Sub

Private Sub Valua_Click()
    PrgValua.Show
End Sub

Private Sub Sector_Click()
    PrgSector.Show
End Sub

Private Sub SubLinea_Click()
    PrgFamilia.Show
End Sub

Private Sub TotalArticulo_Click()
    PrgTotalArticulo.Show
End Sub

Private Sub Subduariopagos_Click()
    PrgListaPagosSubDiario.Show
End Sub

Private Sub Tamaño_Click()
    PrgTamaño.Show
End Sub

Private Sub Tipo_Click()
    PrgTipoPro.Show
End Sub

Private Sub TipoClie_Click()
    PrgTipoClie.Show
End Sub

Private Sub Varios1_Click()
    WVarios = 1
    PrgVarios.Show
End Sub

Private Sub Varios2_Click()
    WVarios = 2
    PrgVarios.Show
End Sub

Private Sub Varios3_Click()
    WVarios = 3
    PrgVarios.Show
End Sub

Private Sub Vende_Click()
    PrgVendedor.Show
End Sub

Private Sub VentaCampana_Click()
    PrgVentaCampana.Show
End Sub

Private Sub VentArt_Click()
    PrgEstaart.Show
End Sub

Private Sub VentasCli_Click()
    PrgVentClie.Show
End Sub

Private Sub Ventasproy_Click()
    PrgVentProy.Show
End Sub

Private Sub ventclie_Click()
    PrgEstacli.Show
End Sub

Private Sub Form_Activate()
    Menu.Caption = "Sistema de Ventas : " + WNombreEmpresa
    If WEmpresa = "" Then
        Menu.Hide
        Empresa.Show
    End If
    
End Sub

Private Sub VentPcia_Click()
    PrgVentPcia.Show
End Sub

Private Sub Transferencia_Click()
    PrgTransferencia.Show
End Sub
