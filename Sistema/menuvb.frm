VERSION 5.00
Begin VB.Form Menuvb 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Control de Gestion - Proveedores - Caja y Bancos : "
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   585
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
   Icon            =   "menuvb.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "menuvb.frx":0442
   ScaleHeight     =   8190
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
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
      Left            =   5640
      TabIndex        =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label AdmiII 
      BackColor       =   &H8000000C&
      Caption         =   "Proveedores - Caja y Bancos"
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
      Height          =   495
      Left            =   960
      MouseIcon       =   "menuvb.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Proveedores y Caja y Bancos"
      Top             =   7440
      Width           =   4695
   End
   Begin VB.Image Logo 
      Height          =   1005
      Left            =   8520
      Picture         =   "menuvb.frx":0B8E
      Top             =   6360
      Width           =   2025
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   960
      Picture         =   "menuvb.frx":150C
      Top             =   120
      Width           =   9600
   End
   Begin VB.Menu Maestros 
      Caption         =   "Maestros"
      Begin VB.Menu Cuentas 
         Caption         =   "Ingreso de Cuentas Contables"
      End
      Begin VB.Menu Tipopro 
         Caption         =   "Ingreso de Tipo de Proveedores"
      End
      Begin VB.Menu Banco 
         Caption         =   "Ingreso de Bancos"
      End
      Begin VB.Menu Concepto 
         Caption         =   "Ingreso de Conceptos de Compras"
      End
      Begin VB.Menu Centro 
         Caption         =   "Ingreso de Centro de Costo"
      End
      Begin VB.Menu Prove 
         Caption         =   "Ingreso de Proveedores"
      End
   End
   Begin VB.Menu Nov 
      Caption         =   "Novedades"
      Begin VB.Menu Compras 
         Caption         =   "Ingreso de Comprobantes de Proveedores"
      End
      Begin VB.Menu Solic 
         Caption         =   "Ingreso de Solicitud de Orden de Pago"
         Visible         =   0   'False
      End
      Begin VB.Menu Pago 
         Caption         =   "Ingreso de Ordenes de Pago"
      End
      Begin VB.Menu Recibo 
         Caption         =   "Ingreso de Recibos"
      End
      Begin VB.Menu Ingresosvarios 
         Caption         =   "Ingreso Varios de Direro"
      End
      Begin VB.Menu Deposito 
         Caption         =   "Ingreso de Depositos"
      End
      Begin VB.Menu Chequera 
         Caption         =   "Ingreso de Chequeras"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu CtaCtePrv 
         Caption         =   "Listado de Cuenta Corriente de Proveedores"
      End
      Begin VB.Menu CtaCtePrv1 
         Caption         =   "Consulta de Cuenta Corriente de Proveedores"
      End
      Begin VB.Menu SaldoPrv 
         Caption         =   "Listado de Saldos de Cuenta Corriente de Proveedores"
      End
      Begin VB.Menu ListProyPrv 
         Caption         =   "Listado de Proyeccion de Pagos"
      End
      Begin VB.Menu ccprvfecha 
         Caption         =   "Listado de Cuenta Corriente de Proveedores a Fecha"
      End
      Begin VB.Menu IvaComp 
         Caption         =   "Listado de Iva Compras"
      End
      Begin VB.Menu PosiIva 
         Caption         =   "Listado de Posicion de Iva"
      End
      Begin VB.Menu Movban 
         Caption         =   "Listado de Movimientos de Bancos"
      End
      Begin VB.Menu ListCaja 
         Caption         =   "Listado de Subdiario de Caja Diaria"
      End
      Begin VB.Menu ListImpvar 
         Caption         =   "Listado de Imputaciones Contables"
      End
      Begin VB.Menu ListComp1 
         Caption         =   "Listado de Compras por Centro de Costo"
      End
      Begin VB.Menu ListComp2 
         Caption         =   "Listado de Compras por Concepto"
      End
      Begin VB.Menu ListComp3 
         Caption         =   "Listado de Compras por Proveedor"
      End
      Begin VB.Menu Listrete1 
         Caption         =   "Listado de Retenciones de Ganancias"
      End
      Begin VB.Menu Reteiva 
         Caption         =   "Listado de Retenciones de Iva"
      End
      Begin VB.Menu ListaReteRecibos 
         Caption         =   "Listado de Retenciones Recibidas"
      End
      Begin VB.Menu Listcob 
         Caption         =   "Listado de Cobranza"
      End
      Begin VB.Menu ListCartera 
         Caption         =   "Listado de Cheques en Cartera"
      End
      Begin VB.Menu ListPosdat 
         Caption         =   "Listado de Cheques Posdatados"
      End
      Begin VB.Menu ListCheque 
         Caption         =   "Listado de Cheques Emitidos"
      End
      Begin VB.Menu RankProv 
         Caption         =   "Listado de Ranking por Proveedor"
      End
      Begin VB.Menu RankConc 
         Caption         =   "Listado de Ranking por Concepto"
      End
      Begin VB.Menu ListCash 
         Caption         =   "Listado de Cash Flow"
      End
      Begin VB.Menu listadocontrol 
         Caption         =   "Listado de Control de Chequeras"
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
      Begin VB.Menu Grabaasi 
         Caption         =   "Grabacion de Asientos Contables"
      End
      Begin VB.Menu Procesoreteganancia 
         Caption         =   "Proceso de Traspaso de Retencines de Ganancias a SIAP"
      End
      Begin VB.Menu pasadatos 
         Caption         =   "Traspaso de Datos"
         Visible         =   0   'False
      End
      Begin VB.Menu CambiaClave 
         Caption         =   "Cambio de Clave de Seguridad"
      End
      Begin VB.Menu Calculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu Cierre 
         Caption         =   "Cierre de Mes"
      End
      Begin VB.Menu empresaii 
         Caption         =   "Cambio de Empresa"
         Visible         =   0   'False
      End
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menuvb"
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

Private Sub Ayuda_Click()
    Rem A = Shell("C:\Archivos de programa\Accesorios\WORDPAD.EXE ayuda.doc", 1)
End Sub

Private Sub Agenda_Click()
    OPEN_FILE_Agenda
    PrgAgenda.Show
End Sub

Private Sub Banco_Click()
    PrgBanco.Show
End Sub

Private Sub Calculadora_Click()
    Calculator.Show
End Sub

Private Sub CambiaClave_Click()
    PrgCambiaClave.Show
End Sub

Private Sub CambiaCuenta_Click()
    PrgCambiaCuenta.Show
End Sub

Private Sub ccprvfecha_Click()
    PrgCcprvFecha.Show
End Sub

Private Sub Centro_Click()
    PrgProyecto.Show
End Sub

Private Sub ComparaI_Click()
    PrgComparaGas.Show
End Sub

Private Sub Chequera_Click()
    PrgChequera.Show
End Sub

Private Sub Ciutcuito_Click()
    PrgCircuito.Show
End Sub

Private Sub Cierre_Click()
    PrgCierre.Show
End Sub

Private Sub ComparaGas_Click()
    PrgComparaGas.Show
End Sub

Private Sub Compras_Click()
    PrgCompras.Show
End Sub

Private Sub Concepto_Click()
    PrgConcepto.Show
End Sub

Private Sub CtaCtePrv_Click()
    PrgCcprv.Show
End Sub

Private Sub CtaCtePrv1_Click()
    PrgCcprv1.Show
End Sub

Private Sub Cuentas_Click()
    PrgCuenta.Show
End Sub

Private Sub Deposito_Click()
    PrgDeposito.Show
End Sub

Private Sub empresaii_Click()
    Rem Empresa.Show
End Sub

Private Sub Escalas_Click()
    PrgEscalas.Show
End Sub

Private Sub Fin_Click()
    Rem Menu.WindowState = 1
    Rem A = Shell("Sistema.exe", 1)
    Close
    End
End Sub

Private Sub Imputa1_Click()
    PrgImputa.Show
End Sub

Private Sub Form_Load()
    Menu.Caption = "Sistema de Control de Gestion - Proveedores - Caja y Bancos : " + WNombreEmpresa
    Rem Open "Licencia.txt" For Input As #1
    Rem Input #1, WLicencia
    Rem Close #1
End Sub

Private Sub Grabaasi_Click()
    PrgGrabaImpcyb.Show
End Sub

Private Sub IngrePtoCue_Click()
    PrgIngrePtoCue.Show
End Sub

Private Sub Ingresosvarios_Click()
    PrgIngresosVarios.Show
End Sub

Private Sub IvaComp_Click()
    PrgIvacomp.Show
End Sub

Private Sub listadocontrol_Click()
    PrgControl.Show
End Sub

Private Sub ListaReteRecibos_Click()
    PrgListareterecibos.Show
End Sub

Private Sub ListCaja_Click()
    PrgMovCaja.Show
End Sub

Private Sub ListCartera_Click()
    PrgValcar.Show
End Sub

Private Sub ListCash_Click()
    PrgCashFlow.Show
End Sub

Private Sub ListCheque_Click()
    PrgCheEmi.Show
End Sub

Private Sub Listcob_Click()
    PrgListreci.Show
End Sub

Private Sub ListComp1_Click()
    PrgCompProy.Show
End Sub

Private Sub ListComp2_Click()
    PrgCompcon.Show
End Sub

Private Sub ListComp3_Click()
    PrgCompprov.Show
End Sub

Private Sub ListGastos_Click()
    PrgGastosProy.Show
End Sub

Private Sub ListImpvar_Click()
    Rem PrgImpcyb.Show
End Sub

Private Sub ListIngre_Click()
    PrgIngresosProy.Show
End Sub

Private Sub ListPosdat_Click()
    PrgPosdat.Show
End Sub

Private Sub ListProyPrv_Click()
    PrgProyPrv.Show
End Sub

Private Sub Listrete1_Click()
    PrgListaRete.Show
End Sub

Private Sub Movban_Click()
    PrgMovban.Show
End Sub

Private Sub Pago_Click()
    Prgpago.Show
End Sub

Private Sub parametro_Click()
    PrgParametro.Show
End Sub

Private Sub pasadatos_Click()
    PrgPasaDatos.Show
End Sub

Private Sub posdatfecha_Click()
    PrgPosdatfecha.Show
End Sub

Private Sub PosiIva_Click()
    PrgPosicion.Show
End Sub

Private Sub Procesoreteganancia_Click()
    Rem PrgProcesoReteGanancia.Show
End Sub

Private Sub Prove_Click()
    PrgProve.Show
End Sub

Private Sub Proyecto_Click()
    PrgProyecto.Show
End Sub

Private Sub Pto_Click()
    PrgIngrePto.Show
End Sub

Private Sub Ptocue_Click()
    PrgIngrePtoCue.Show
End Sub

Private Sub RankConc_Click()
    PrgRankCon.Show
End Sub

Private Sub RankProv_Click()
    PrgRankProv.Show
End Sub

Private Sub Recibo_Click()
    PrgRecibos.Show
End Sub

Private Sub ResMov_Click()
    PrgResumenProy.Show
End Sub

Private Sub Reteiva_Click()
    PrgListaReteIva.Show
End Sub

Private Sub SaldoPrv_Click()
    PrgSalprv.Show
End Sub

Private Sub Solic_Click()
    OPEN_FILE_Solicitud
    OPEN_FILE_Configuracion
    PrgSolicitud.Show
End Sub

Private Sub Tipopro_Click()
    PrgTipoPro.Show
End Sub

Private Sub Form_Activate()
    Menu.Caption = "Sistema de Control de Gestion - Proveedores - Caja y Bancos : " + WNombreEmpresa
    If WEmpresa = "" Then
        Menu.Hide
        Empresa.Show
    End If
End Sub

Private Sub valvarfecha_Click()
    PrgValcarFecha.Show
End Sub
