VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Control de Gestion - Control de Importaciones : "
   ClientHeight    =   8310
   ClientLeft      =   150
   ClientTop       =   600
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
   Icon            =   "MenuImpo.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "MenuImpo.frx":0442
   ScaleHeight     =   8310
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
      Left            =   5520
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
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
      MouseIcon       =   "MenuImpo.frx":0884
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Ventas y Stock"
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Image Logo 
      Height          =   660
      Left            =   8595
      Picture         =   "MenuImpo.frx":0B8E
      Top             =   6765
      Width           =   1950
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   960
      Picture         =   "MenuImpo.frx":192D
      Top             =   240
      Width           =   9600
   End
   Begin VB.Menu Maestros 
      Caption         =   "Maestros"
      Begin VB.Menu Caja 
         Caption         =   "Ingreso de Cajas"
      End
      Begin VB.Menu Clieexpo 
         Caption         =   "Ingreso de Clientes"
      End
      Begin VB.Menu Lugar 
         Caption         =   "Ingreso de Lugares de Entrega"
      End
      Begin VB.Menu Traspo 
         Caption         =   "Ingreso de Broker"
      End
      Begin VB.Menu Arti 
         Caption         =   "Ingreso de Articulos"
      End
      Begin VB.Menu Componente 
         Caption         =   "Ingreso de Componentes"
      End
      Begin VB.Menu Formula 
         Caption         =   "Ingreso de Composicion de Artiulos"
      End
   End
   Begin VB.Menu Nov 
      Caption         =   "Novedades"
      Begin VB.Menu Po 
         Caption         =   "Ingreso de Orden de Importación (Puschase Order)"
      End
      Begin VB.Menu Despacho 
         Caption         =   "Ingreso de Despachos de importacion"
      End
      Begin VB.Menu Docu1 
         Caption         =   "Emision de Documentacion /1"
      End
      Begin VB.Menu Docu2 
         Caption         =   "Emision de Documentacion /2"
      End
      Begin VB.Menu Listacompodespa 
         Caption         =   "Listado de Componentes por Despacho"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu ListaOrdenes 
         Caption         =   "Listado de Ordenes de Importacion"
      End
      Begin VB.Menu ListaDespa 
         Caption         =   "Listado de Despachos de Importacion"
      End
      Begin VB.Menu ListaOrdPen 
         Caption         =   "Listado de Ordenes de Importacion Pendientes de Entrega"
      End
   End
   Begin VB.Menu fgh 
      Caption         =   "Procesos"
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menu"
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

Private Sub Arti_Click()
    OPEN_FILE_ArtiExpo
    PrgArtiExpo.Show
End Sub

Private Sub Caja_Click()
    OPEN_FILE_Envase
    PrgCaja.Show
End Sub

Private Sub Calculadora_Click()
    Calculator.Show
End Sub

Private Sub Clieexpo_Click()
    OPEN_FILE_ClieExpo
    PrgClieExpo.Show
End Sub

Private Sub Cotiza_Click()
    OPEN_FILE_Cotiza
    OPEN_FILE_Configuracion
    PrgCotiza.Show
End Sub

Private Sub Escalas_Click()
    OPEN_FILE_Parametro
    OPEN_FILE_Configuracion
    PrgEscalas.Show
End Sub

Private Sub Cliente_Click()
    OPEN_FILE_Clientes
    prgcliente.Show
End Sub

Private Sub Comiven_Click()
    PrgVentVend.Show
End Sub

Private Sub ComparaVen_Click()
    PrgComparaVen.Show
End Sub

Private Sub CompArt_Click()
    OPEN_FILE_Compras
    PrgCompArt.Show
End Sub

Private Sub Ctacte_Click()
    OPEN_FILE_Clientes
    PrgCtaCte.Show
End Sub

Private Sub ctacte2_Click()
    OPEN_FILE_Ctacte
    OPEN_FILE_Clientes
    PrgCtaCte1.Show
End Sub

Private Sub CtacteVen_Click()
    PrgCtaCteVen.Show
End Sub

Private Sub devol_Click()
    OPEN_FILE_Configuracion
    PrgDevol.Show
End Sub

Private Sub factu_Click()
    OPEN_FILE_Configuracion
    PrgFactura.Show
End Sub

Private Sub FactuExpo_Click()
    OPEN_FILE_Configuracion
    WVarios = 1
    PrgFactuExpo.Show
End Sub

Private Sub Componente_Click()
    OPEN_FILE_Componente
    PrgComponente.Show
End Sub

Private Sub Despacho_Click()
    OPEN_FILE_Despacho
    PrgDespacho.Show
End Sub

Private Sub Docu1_Click()
    PrgDocumento1.Show
End Sub

Private Sub Docu2_Click()
    PrgDocumento2.Show
End Sub

Private Sub Fin_Click()
    Close
    End
End Sub

Private Sub Form_Load()
    da = Date$
    Rem If Right$(Date$, 4) <> "2003" Then
    Rem     m$ = "El tiempo de uso para la verificacion del funcionamiento del sistema a terminado" + Chr$(13) + "Comuniquese con su proveedior para adquirir la version completa del mismo"
    Rem    A% = MsgBox(m$, 0, "Sistema de Control de gestion")
    Rem     Close
    Rem     End
    Rem End If
    WEmpresa = "0001"
    Rem Open "Licencia.txt" For Input As #1
    Rem Input #1, WLicencia
    Rem Close #1
    Rem Logo.Picture = LoadPicture("logo.jpg")
End Sub

Private Sub Prove_Click()
    PrgProve.Show
End Sub

Private Sub IngrePtoVend_Click()
    PrgIngrePtoVend.Show
End Sub

Private Sub Ivaven_Click()
    PrgIvaven.Show
End Sub

Private Sub Lineas_Click()
    OPEN_FILE_Lineas
    PrgLinea.Show
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

Private Sub Formula_Click()
    OPEN_FILE_Formula
    PrgFormula.Show
End Sub

Private Sub Listacompodespa_Click()
    PrgListaCompoDespa.Show
End Sub

Private Sub ListaDespa_Click()
    PrgListDespacho.Show
End Sub

Private Sub ListaOrdenes_Click()
    PrgListOrdenImpo.Show
End Sub

Private Sub ListaOrdPen_Click()
    PrgListOrdenImpoPend.Show
End Sub

Private Sub Lugar_Click()
    OPEN_FILE_Lugar
    PrgLugar.Show
End Sub

Private Sub Minimo_Click()
    PrgMinimo.Show
End Sub

Private Sub MovStk_Click()
    OPEN_FILE_Movstk
    PrgMovstk.Show
End Sub

Private Sub parametro_Click()
    OPEN_FILE_Empresa
    PrgParametro.Show
End Sub

Private Sub Pedido_Click()
    OPEN_FILE_Pedido
    OPEN_FILE_Configuracion
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

Private Sub SaldosCta_Click()
    PrgSaldoCta.Show
End Sub

Private Sub Po_Click()
    OPEN_FILE_OrdenImpo
    PrgOrdenImpo.Show
End Sub

Private Sub Traspo_Click()
    OPEN_FILE_Transporte
    PrgTransporte.Show
End Sub

Private Sub Valua_Click()
    PrgValua.Show
End Sub

Private Sub Varios1_Click()
    OPEN_FILE_Configuracion
    WVarios = 1
    PrgVarios.Show
End Sub

Private Sub Varios2_Click()
    OPEN_FILE_Configuracion
    WVarios = 2
    PrgVarios.Show
End Sub

Private Sub Varios3_Click()
    OPEN_FILE_Configuracion
    WVarios = 3
    PrgVarios.Show
End Sub

Private Sub Vende_Click()
    OPEN_FILE_Vendedor
    PrgVendedor.Show
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

    Rem If Right$(Date$, 4) <> "2002" Then
    Rem     Close
    Rem     End
    Rem End If

    If WEmpresa = "" Then
        Rem Empresa.Show
        Rem Empresa.SetFocus
        WEmpresa = "0001"
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Control de Gestion - Ventas y Stock : " + !Nombre
            End If
        End With
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Control de Gestion - Ventas y Stock : " + !Nombre
            End If
        End With
    End If

End Sub


