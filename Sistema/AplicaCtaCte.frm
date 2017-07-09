VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAplicaCtaCte 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Aplicacion de Comprobantes"
   ClientHeight    =   8250
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   11880
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8250
   ScaleWidth      =   11880
   Begin VB.CommandButton Impresion 
      Caption         =   "Impres. F9"
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
      Left            =   10920
      MouseIcon       =   "AplicaCtaCte.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "AplicaCtaCte.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Impresion"
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu F10"
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
      Left            =   10920
      MouseIcon       =   "AplicaCtaCte.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "AplicaCtaCte.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Menu Principal"
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta F4"
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
      Left            =   10920
      MouseIcon       =   "AplicaCtaCte.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "AplicaCtaCte.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Consulta de Datos"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpia F3"
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
      Left            =   10920
      MouseIcon       =   "AplicaCtaCte.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "AplicaCtaCte.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borra  F2"
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
      Left            =   10920
      MouseIcon       =   "AplicaCtaCte.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "AplicaCtaCte.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Elimina el Registro"
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
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
      Left            =   10920
      MouseIcon       =   "AplicaCtaCte.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "AplicaCtaCte.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Ctacte 
      Caption         =   "Cta.Cte. F5"
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
      Left            =   10920
      MouseIcon       =   "AplicaCtaCte.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "AplicaCtaCte.frx":46D2
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cuenta Corriente de Proveedores"
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6720
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   5280
      Width           =   375
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   13
      Top             =   4800
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   12
      Top             =   5280
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Top             =   4800
      Width           =   375
   End
   Begin Crystal.CrystalReport listado 
      Left            =   6240
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "recibo.rpt"
   End
   Begin VB.TextBox Clientes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   1455
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   6720
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Recibo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8520
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      ItemData        =   "AplicaCtaCte.frx":4F9C
      Left            =   6720
      List            =   "AplicaCtaCte.frx":4FA3
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Top             =   4800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4335
      Left            =   120
      TabIndex        =   22
      Top             =   3120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   7646
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "COMPROBANTES CANCELADOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "  3) Factura    4) N/D   5 N/C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Label Debitos 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Label DesClientes 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod. Cilente"
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
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero de Recibo"
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
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "PrgAplicacTAcTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String
Private Auxi1 As String
Private WSaldo As Double
Private Vector(20, 10) As String
Private Provincia(100) As String
Private m(20) As String
Private Impre1(100) As Single
Private Impre2(100) As Single
Private ImpreTipo(100) As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WPostal As String
Private WProvincia As String
Private WProv As String
Private WCuenta(20) As String
Private Debito As Double
Private Credito As Double
Private ZSaldo As Double
Dim ZMes As String
Dim ZAno As String

Dim ZZZRetOtra As Double
Dim ZZZRetOtraII As Double
Dim ZZZRetOtraIII As Double
Dim ZZZRetOtraIV As Double
Dim ZZAnticipo As Double



Dim ZZRecibo As String
Dim ZZRenglon As String
Dim ZZCliente As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZTipoRec As String
Dim ZZRetGanancias As String
Dim ZZRetIva As String
Dim ZZRetOtra As String
Dim ZZRetOtraII As String
Dim ZZRetOtraIII As String
Dim ZZRetOtraIV As String
Dim ZZNroRetganancias As String
Dim ZZNroRetIva As String
Dim ZZNroRetOtra As String
Dim ZZNroRetOtraII As String
Dim ZZNroRetOtraIII As String
Dim ZZNroRetOtraIV As String
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
Dim ZZObservaciones As String
Dim ZZEmpresa As String
Dim ZZClave As String
Dim ZZImporte As String
Dim ZZCuenta As String
Dim ZZDestino As String
Dim ZZOrden As String
Dim ZZDeposito As String


Dim ZZLetra As String
Dim ZZTipo As String
Dim ZZPunto As String
Dim ZZNumero As String
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
Dim ZZProvincia As String
Dim ZZVendedor As String
Dim ZZCosto As String
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

Dim ZZNumeroCheque As String
Dim ZZVectorI(100, 6) As String
Dim ZZVectorII(100, 6) As String
Dim ZZVectorIII(100, 6) As String

Dim ZZRazon As String
Dim ZZPesosI As String
Dim ZZPesosII As String
Dim ZZFechaI As String
Dim zZNumeroI As String
Dim ZZImporteI As String
Dim ZZBanco As String
Dim ZZSucursal As String
Dim ZZNumeroII As String
Dim ZZFechaII As String
Dim ZZImporteII As String
Dim ZZEstructura As String
Dim ZZImporteIII As String

Dim XTexto2 As String
Dim XTexto1 As String

Dim ZCargaCheque(100) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Dim ZZAlta As Integer

Private Sub Suma_Datos()

    ZZSuma = 0
    Debitos.Caption = ""
    For IRow = 1 To 100
        ZZSuma = ZZSuma + Val(WVector1.TextMatrix(IRow, 5))
    Next IRow
    Debitos.Caption = Str$(ZZSuma)
    Debitos.Caption = Alinea("###,###.##", Debitos.Caption)
End Sub

Private Sub Lee_Datos()

    Call Limpia_Vector

    Renglon = 0
    Debito = 0
    Credito = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Recibo.Text + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Select Case Val(rstRecibos!Tiporeg)
                        Case 1
                            Debito = Debito + 1
                            WVector1.Row = Debito
                            WVector1.Col = 1
                            WVector1.Text = rstRecibos!Tipo1
                            WVector1.Col = 2
                            WVector1.Text = rstRecibos!Letra1
                            WVector1.Col = 3
                            WVector1.Text = rstRecibos!Punto1
                            WVector1.Col = 4
                            WVector1.Text = rstRecibos!Numero1
                            WVector1.Col = 5
                            WVector1.Text = rstRecibos!Importe1
                            WVector1.Text = Alinea("###,###.##", WVector1.Text)
                        Case Else
                    End Select
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstRecibos.Close
    End If
    
     
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Clientes.Text = rstCliente!Cliente
        DesClientes.Caption = rstCliente!Fantasia
        WRazon = rstCliente!Fantasia
        WDireccion = rstCliente!Direccion
        WLocalidad = rstCliente!Localidad
        WPostal = rstCliente!Postal
        WProv = rstCliente!Provincia
        rstCliente.Close
        Call Format_datos
    End If
    
End Sub


Private Sub cmdAdd_Click()

    Rem If WLicencia <> "1234-5678-ABCD-EFGH" And Val(Recibo.Text) > 10 Then
    Rem     WMsg$ = "La version del sistema es para un uso limitado de movimientos." + Chr$(13) + _
    REM          "El objetivo es el de verificar las opciones y el funcionamiento del mismo." + Chr$(13) + _
    REM          "Para poder utilizar el sistema sin limite de movimientos se debe adquirir la version definitiva."
    Rem     aaaaaa% = MsgBox(WMsg$, 0, "Sistema de Control de Gestion")
    Rem     Exit Sub
    Rem End If
    
    
    
    For IRow = 1 To 100
    
        WTipo = WVector1.TextMatrix(IRow, 1)
        WLetra = WVector1.TextMatrix(IRow, 2)
        WPunto = WVector1.TextMatrix(IRow, 3)
        WNumero = WVector1.TextMatrix(IRow, 4)
        WDebitos = Val(WVector1.TextMatrix(IRow, 5))
        
        If WDebitos <> 0 Then
        
            WClave = WLetra + WTipo + WPunto + WNumero + "01"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                If Trim(UCase(rstCtaCte!Cliente)) <> Trim(UCase(Clientes.Text)) Then
                    M1$ = "La factura " + WNumero + " no pertenece a este cliente"
                    aaaaaa% = MsgBox(M1$, 0, "Ingreso de Recibos")
                    Exit Sub
                End If
                rstCtaCte.Close
            End If
        
        End If
        
    Next IRow

    If Recibo.Text <> "" And Fecha.Text <> "" Then
    
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        
        Existe = "N"
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Recibo.Text + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
            rstRecibos.Close
            M1$ = "Recibo ya existente"
            aaaaaa% = MsgBox(M1$, 0, "Ingreso de Recibos")
            Existe = "S"
        End If
    
        If Existe <> "S" Then
    
            Call Suma_Datos
        
            ZSumaUs = 0
            Debito = 0
            Credito = 0
            If Val(Debitos.Caption) <> 0 Then
                Debito = Val(Debitos.Caption)
            End If
        
            ZZDife = Abs(Debito)
        
            If ZZDife < 1 Then
    
                Renglon = 0
                For IRow = 1 To 100
        
                    WRow = IRow
                    WVector1.Col = 5
                    WVector1.Row = IRow
                        
                    If Val(WVector1.Text) <> 0 Then
                    
                        Renglon = Renglon + 1
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        
                        ZZRecibo = Recibo.Text
                        ZZRenglon = Auxi1
                        ZZCliente = Clientes.Text
                        ZZfecha = Fecha.Text
                        ZZFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        ZZTipoRec = "3"
                        ZZRetGanancias = ""
                        ZZRetIva = ""
                        ZZRetOtra = ""
                        ZZRetOtraII = ""
                        ZZRetOtraIII = ""
                        ZZRetOtraIV = ""
                        ZZRetSuss = ""
                        ZZNroRetganancias = ""
                        ZZNroRetIva = ""
                        ZZNroRetOtra = ""
                        ZZNroRetOtraII = ""
                        ZZNroRetOtraIII = ""
                        ZZNroRetOtraIV = ""
                        ZZNroRetSuss = ""
                        ZZRetencion = "0"
                        ZZTipoReg = "1"
                        
                        WVector1.Col = 1
                        ZZTipo1 = WVector1.Text
                        WVector1.Col = 2
                        ZZLetra1 = WVector1.Text
                        WVector1.Col = 3
                        ZZPunto1 = WVector1.Text
                        WVector1.Col = 4
                        ZZNumero1 = WVector1.Text
                        Call Ceros(ZZNumero1, 8)
                        WVector1.Col = 5
                        ZZImporte1 = WVector1.Text
                        ZZTipo2 = ""
                        ZZNumero2 = ""
                        ZZFecha2 = ""
                        ZZFechaOrd2 = ""
                        ZZBanco2 = ""
                        ZZImporte2 = 0
                        ZZEstado2 = ""
                        ZZObservaciones = ""
                        ZZEmpresa = WEmpresa
                        ZZClave = ZZRecibo + ZZRenglon
                        ZZImporte = Str$(Debito)
                        ZZCuenta = "1"
                        ZZDestino = ""
                        ZZOrden = "0"
                        ZZDeposito = "0"
                        
                
                        WLetra = ZZLetra1
                        WTipo = ZZTipo1
                        WPunto = ZZPunto1
                        WNumero = ZZNumero1
                        WImporte = ZZImporte1
                        
                        Auxi$ = Clientes.Text
                        Call Ceros(Auxi$, 6)
                        Claveven$ = Auxi$
                        WClave = WLetra + WTipo + WPunto + WNumero + "01"
                        
                        ZZCodigoBanco = ""
                        ZZSucursalCheque = ""
                        ZZTipoCheque = ""
                        ZZClaseCheque = ""
                        ZZCuit = ""
                        
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
                        ZSql = ZSql + "RetOtraII ,"
                        ZSql = ZSql + "RetOtraIII ,"
                        ZSql = ZSql + "RetOtraIV ,"
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
                        ZSql = ZSql + "ClaseCheque ,"
                        ZSql = ZSql + "Cuit ,"
                        ZSql = ZSql + "NroRetGanancias ,"
                        ZSql = ZSql + "NroRetIva ,"
                        ZSql = ZSql + "NroRetOtra ,"
                        ZSql = ZSql + "NroRetOtraII ,"
                        ZSql = ZSql + "NroRetOtraIII ,"
                        ZSql = ZSql + "NroRetOtraIV ,"
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
                        ZSql = ZSql + "'" + ZZRetOtraII + "',"
                        ZSql = ZSql + "'" + ZZRetOtraIII + "',"
                        ZSql = ZSql + "'" + ZZRetOtraIV + "',"
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
                        ZSql = ZSql + "'" + ZZClaseCheque + "',"
                        ZSql = ZSql + "'" + ZZCuit + "',"
                        ZSql = ZSql + "'" + ZZNroRetganancias + "',"
                        ZSql = ZSql + "'" + ZZNroRetIva + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtra + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtraII + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtraIII + "',"
                        ZSql = ZSql + "'" + ZZNroRetOtraIV + "',"
                        ZSql = ZSql + "'" + ZZRetSuss + "',"
                        ZSql = ZSql + "'" + ZZNroRetSuss + "')"
                            
                        spRecibos = ZSql
                        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)

                
                        WLetra = ZZLetra1
                        WTipo = ZZTipo1
                        WPunto = ZZPunto1
                        WNumero = ZZNumero1
                        WImporte = ZZImporte1
                        
                        Auxi$ = Clientes.Text
                        Call Ceros(Auxi$, 6)
                        Claveven$ = Auxi$
                        WClave = WLetra + WTipo + WPunto + WNumero + "01"
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM CtaCte"
                        ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                        spCtaCte = ZSql
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        If rstCtaCte.RecordCount > 0 Then
                            ZSumaUs = ZSumaUs + rstCtaCte!Totalus
                            rstCtaCte.Close
                        End If
                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE CtaCte SET "
                        ZSql = ZSql + " Saldo = Saldo - " + "'" + ZZImporte1 + "',"
                        ZSql = ZSql + " SaldouS = 0"
                        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                        spCtaCte = ZSql
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                    
                    End If
                Next IRow
    
                mm$ = "Grabacion realizada"
                aaa% = MsgBox(mm$, 0, "Archivo de Recibos")
    
                
                Recibo.SetFocus
                
                    Else
                    
                M1$ = "Los Valores de la Aplicacion  no Balancean"
                aaaaaa% = MsgBox(M1$, 0, "Ingreso de Recibos")
            
            End If
        
        End If
        
    End If
                
    
End Sub

Private Sub CmdDelete_Click()
    If Recibo.Text <> "" Then
    
        T$ = "Ingresos de Recibos"
        M1$ = "Desea Anular la Aplicacion"
        Respuestaaaaaa% = MsgBox(M1$, 32 + 4, T$)
        If Respuestaaaaaa% = 6 Then
                
            For da = 1 To 99
            
                Auxi1 = Str$(da)
                Call Ceros(Auxi1, 2)
                WClave = Recibo.Text + Auxi1
            
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Recibos"
                ZSql = ZSql + " Where Recibos.Clave = " + "'" + WClave + "'"
                spRecibos = ZSql
                Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                If rstRecibos.RecordCount > 0 Then
                
                    WTipoRec = rstRecibos!TipoRec
                    WLetra = rstRecibos!Letra1
                    WTipo = rstRecibos!Tipo1
                    WPunto = rstRecibos!Punto1
                    WNumero = rstRecibos!Numero1
                    WImporte = rstRecibos!Importe1
                    WTipoReg = rstRecibos!Tiporeg
                    
                    rstRecibos.Close
                    
                    If Val(WTipoReg) = 1 Then
                    
                        WClave = WLetra + WTipo + WPunto + WNumero + "01"
                    
                        ZSql = ""
                        ZSql = ZSql + "UPDATE CtaCte SET "
                        ZSql = ZSql + " Saldo = Saldo + " + "'" + Str$(WImporte) + "'"
                        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                        spCtaCte = ZSql
                        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                End If
                    
            Next da
            
            ZSql = ""
            ZSql = ZSql + "DELETE Recibos"
            ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Recibo.Text + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            
            Call CmdLimpiar_Click
                        
        End If
        
    End If
    
    Recibo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Call Limpia_Vector

    Recibo.Text = "900000"
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Recibo.SetFocus
    Debitos.Caption = ""
    ZZAlta = 0
    
    Pantalla.Visible = False
    Opcion.Visible = False
                
    Recibo.Text = "900000"
    ZSql = ""
    ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where TipoRec = 3 "
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        rstRecibos.MoveLast
        ZUltimo = IIf(IsNull(rstRecibos!ReciboMayor), "0", rstRecibos!ReciboMayor)
        Recibo.Text = ZUltimo + 1
        rstRecibos.Close
    End If
        
    If Val(Recibo.Text) <= 900000 Then
        Recibo.Text = "900000"
    End If
                
    
End Sub

Private Sub cmdClose_Click()
    PrgAplicacTAcTE.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Recibo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi1 = Recibo.Text
        Call Ceros(Auxi1, 6)
        Recibo.Text = Auxi1
        Existe = "N"
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.Recibo = " + "'" + Recibo.Text + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then
        
            Existe = "S"
            
            Clientes.Text = rstRecibos!Cliente
            Fecha.Text = rstRecibos!Fecha
            rstRecibos.Close
            
        End If
                
        If Existe = "S" Then
            ZZAlta = 1
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            Fecha.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Recibo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(Fecha.Text)) = 8 Then
            Fecha.Text = Left$(Fecha.Text, 6) + "20" + Right$(Trim(Fecha.Text), 2)
        End If
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Clientes.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Clientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Clientes.Text <> "" Then
        
            Clientes.Text = UCase(Trim(Clientes.Text))
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Clientes.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesClientes.Caption = rstCliente!Fantasia
                WRazon = rstCliente!Fantasia
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WProv = rstCliente!Provincia
                rstCliente.Close
                
                Opcion.Clear
                Opcion.AddItem "Clientes"
                Opcion.AddItem "Cuentas Contables"
                Opcion.AddItem "Cuentas Corrientes"
                Opcion.ListIndex = 2
                Call Opcion_Click
    
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
                
                    Else
                    
                Clientes.SetFocus
                
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        Clientes.Text = ""
        DesClientes.Caption = ""
    End If
End Sub


Private Sub Consulta_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuenta Corriestes"

    Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            
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
                            IngresaItem = !Cliente + " " + !Fantasia
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
            Ayuda.SetFocus
           
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Cliente = " + "'" + Clientes.Text + "'"
            ZSql = ZSql + " Order by CtaCte.Numero"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                With rstCtaCte
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            ZSaldo = rstCtaCte!Saldo
                            Call Redondeo(ZSaldo)
                        
                            If ZSaldo <> 0 Then
                                ZSaldo = rstCtaCte!Saldous
                                Call Redondeo(ZSaldo)
                                Auxi2 = Str$(ZSaldo)
                                Auxi2 = Mascara("###,###.##", Auxi2)
                                ZSaldo = rstCtaCte!Saldo
                                Call Redondeo(ZSaldo)
                                Auxi = Str$(ZSaldo)
                                Auxi = Mascara("###,###.##", Auxi)
                                Auxi1 = Str$(rstCtaCte!Numero)
                                Call Ceros(Auxi1, 6)
                                IngresaItem = rstCtaCte!Impre + " " + Auxi1 + " " + rstCtaCte!Fecha + " " + Auxi
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstCtaCte!Clave
                                WIndice.AddItem IngresaItem
                            End If
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCtaCte.Close
            End If
        
        Case Else
    End Select
            
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Select Case XIndice
        Case 0
            Pantalla.Visible = False
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Clientes.Text = WIndice.List(Indice)
            Clientes.Text = UCase(Trim(Clientes.Text))
            Call Clientes_KeyPress(13)
            
        Case 2
            Entra = "S"
            Indice = Pantalla.ListIndex
            Compara1 = WIndice.List(Indice)
        
            For IRow = 1 To 100
                WVector1.Row = IRow
                WVector1.Col = 2
                Compara2 = WVector1.Text
                WVector1.Col = 1
                Compara2 = Compara2 + WVector1.Text
                WVector1.Col = 3
                Compara2 = Compara2 + WVector1.Text
                WVector1.Col = 4
                Compara2 = Compara2 + WVector1.Text + "01"
                If Compara1 = Compara2 Then
                    Entra = "N"
                    Exit For
                End If
            Next IRow
            
            If Entra = "S" Then
            
                For IRow = 1 To 100
                    WVector1.Row = IRow
                    WVector1.Col = 1
                    If WVector1.Text = "" Then
                        XRow = WVector1.Row
                        Exit For
                    End If
                Next IRow
                
                Indice = Pantalla.ListIndex
                WClave = WIndice.List(Indice)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCte"
                ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
            
                    WVector1.Row = XRow
                    WVector1.Col = 1
                    Auxi = rstCtaCte!Tipo
                    Call Ceros(Auxi, 2)
                    WVector1.Text = Auxi
                    
                    WVector1.Row = XRow
                    WVector1.Col = 2
                    WVector1.Text = rstCtaCte!Letra
                
                    WVector1.Row = XRow
                    WVector1.Col = 3
                    Auxi = rstCtaCte!Punto
                    Call Ceros(Auxi, 4)
                    WVector1.Text = Auxi
                
                    WVector1.Row = XRow
                    WVector1.Col = 4
                    Auxi = rstCtaCte!Numero
                    Call Ceros(Auxi, 8)
                    WVector1.Text = Auxi
                                            
                    WParidad = rstCtaCte!Paridad
                    WSaldo = rstCtaCte!Saldo
                    Rem If wparidad <> 0 Then
                    Rem     WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    Rem End If
                    
                    WVector1.Row = XRow
                    WVector1.Col = 5
                    WVector1.Text = WSaldo
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    
                    rstCtaCte.Close
                    
                    WVector1.Row = XRow
                    WVector1.Col = 5
                    
                End If
                    
                Call Suma_Datos
            
            End If
                
            WVector1.Row = XRow
            WVector1.Col = 1
            Call StartEdit
                
        Case Else
    End Select
    
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
            ZSql = ZSql + " Where Cliente.Fantasia LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Cliente + " " + !Fantasia
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Form_Load()

    Call Limpia_Vector
 
    Recibo.Text = "900000"
    Clientes.Text = ""
    DesClientes.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Debitos.Caption = ""
    ZZAlta = 0
    
    
    Recibo.Text = "900000"
    ZSql = ""
    ZSql = ZSql + "Select Max(Recibo) as [ReciboMayor]"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where TipoRec = 3"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        rstRecibos.MoveLast
        ZUltimo = IIf(IsNull(rstRecibos!ReciboMayor), "0", rstRecibos!ReciboMayor)
        Recibo.Text = ZUltimo + 1
        rstRecibos.Close
    End If
    
    If Val(Recibo.Text) <= 900000 Then
        Recibo.Text = "900000"
    End If
    
            
End Sub

Private Sub Impresion_Click()

    T$ = "Ingresos de Recibos"
    M1$ = "Desea imprimir el comprobante"
    Respuestaaaaaa% = MsgBox(M1$, 32 + 4, T$)
    If Respuestaaaaaa% = 6 Then
        Call Impresion_Recibo
    End If

End Sub


Rem
Rem Controles de la grilla
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

Private Sub TotalRete_DblClick()
    PantaRete.Visible = True
    Retganancias.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
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
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
                Call StartEdit
            End If
    
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
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,f5,F10
        Case 112, 113, 114, 115, 116, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
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
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
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

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
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


Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
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

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 6
    WVector1.FixedRows = 1
    WVector1.Rows = 201
    
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
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 2
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Letra"
                WVector1.ColWidth(Ciclo) = 600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 1
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Punto"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "#,###,###.##"
                
            Case Else
                
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To 5
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

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


Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Val(WVector1.Text) <> 0 Then
                If Val(WVector1.Text) = 1 Or Val(WVector1.Text) = 2 Or Val(WVector1.Text) = 3 Then
                    Auxi$ = Str$(Val(WVector1.Text))
                    Call Ceros(Auxi$, 2)
                    WVector1.Text = Auxi$
                        Else
                    If Val(WVector1.Text) = 99 Then
                        WVector1.Col = 4
                            Else
                        WControl = "N"
                    End If
                End If
            End If
        Case 2, 3
            WVector1.Col = XColumna
        Case 4
            Auxi$ = Str$(Val(WVector1.Text))
            Call Ceros(Auxi$, 8)
            WVector1.Text = Auxi$
            
            WVector1.Col = 2
            Claveven$ = WVector1.Text
            WVector1.Col = 1
            Claveven$ = Claveven$ + WVector1.Text
            WVector1.Col = 3
            Claveven$ = Claveven$ + WVector1.Text
            WVector1.Col = 4
            WClave = Claveven$ + WVector1.Text + "01"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                WVector1.Col = 5
                XRow = WVector1.Row
                If Val(WVector1.Text) = 0 Then
                    WParidad = rstCtaCte!Paridad
                    WSaldo = rstCtaCte!Saldo
                    If WParidad <> 0 Then
                        WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    End If
                    WVector1.Text = WSaldo
                    Call Suma_Datos
                End If
                WVector1.Col = 4
                rstCtaCte.Close
                    Else
                WControl = "N"
            End If
            
            
        Case 5
            WVector1.Col = 2
            Claveven$ = WVector1.Text
            WVector1.Col = 1
            Claveven$ = Claveven$ + WVector1.Text
            ZZTipo = WVector1.Text
            WVector1.Col = 3
            Claveven$ = Claveven$ + WVector1.Text
            WVector1.Col = 4
            WClave = Claveven$ + WVector1.Text + "01"
                    
            If Val(ZZTipo) <> 99 Then
                        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CtaCte"
                ZSql = ZSql + " Where CtaCte.Clave = " + "'" + WClave + "'"
                spCtaCte = ZSql
                Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtaCte.RecordCount > 0 Then
                    WParidad = rstCtaCte!Paridad
                    WSaldo = rstCtaCte!Saldo
                    If WParidad <> 0 Then
                        WSaldo = rstCtaCte!Saldo * rstCtaCte!Paridad
                    End If
                    Saldo = Alinea("###,###.##", Str$(WSaldo))
                    rstCtaCte.Close
                        Else
                    Saldo = 0
                End If
                    
                WVector1.Col = 5
                If Abs(Val(WVector1.Text)) > Abs(Val(Saldo)) Then
                    WVector1.Text = ""
                    WControl = "N"
                        Else
                    WVector1.Text = Alinea("###,###.##", WVector1.Text)
                    Call Suma_Datos
                End If
                    
                    Else
            
                WVector1.Text = Alinea("###,###.##", WVector1.Text)
                Call Suma_Datos
            End If

        
        Case Else
    End Select
End Sub

Private Sub Clientes_DblClick()

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Corrientes"
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub CtaCte_Click()

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Opcion.AddItem "Cuentas Corrientes"
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub




Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Recibo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Clientes_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 113
            Call CmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call CtaCte_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub





Private Sub Impresion_Recibo()

    ZSql = ""
    ZSql = ZSql + "DELETE ImpreRecibo"
    spImpreRecibo = ZSql
    Set rstImpreRecibo = db.OpenRecordset(spImpreRecibo, dbOpenSnapshot, dbSQLPassThrough)

    Erase ZZVectorI
    Erase ZZVectorII
    Erase ZZVectorIII
    
    ZLugarI = 0
    ZLugarII = 0
    ZLugarIII = 0
    
    Call Numtolet
    ZSumaCheque = 0
    
    For IRow = 1 To 100

        WRow = IRow
        
        WVector1.Col = 5
        WVector1.Row = IRow
            
        If Val(WVector1.TextMatrix(IRow, 5)) <> 0 Then
        
            ZLugarI = ZLugarI + 1
            
            ZZVectorI(ZLugarI, 1) = ""
            ZZVectorI(ZLugarI, 2) = WVector1.TextMatrix(IRow, 4)
            ZZVectorI(ZLugarI, 3) = WVector1.TextMatrix(IRow, 5)
            
            ZZTipo1 = WVector1.TextMatrix(IRow, 1)
            ZZLetra1 = WVector1.TextMatrix(IRow, 2)
            ZZPunto1 = WVector1.TextMatrix(IRow, 3)
            ZZNumero1 = WVector1.TextMatrix(IRow, 4)
            
            ZZClave = ZZLetra1 + ZZTipo1 + ZZPunto1 + ZZNumero1 + "01"
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Clave = " + "'" + ZZClave + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                ZZVectorI(ZLugarI, 1) = rstCtaCte!Fecha
                rstCtaCte.Close
            End If
        
        End If
        
        WVector1.Col = 10
        WVector1.Row = IRow
        If Val(WVector1.Text) <> 0 Then
        
            ZZTipo2 = WVector1.TextMatrix(IRow, 6)
            ZZNumero2 = WVector1.TextMatrix(IRow, 7)
            ZZFecha2 = WVector1.TextMatrix(IRow, 8)
            ZZBanco2 = WVector1.TextMatrix(IRow, 9)
            ZZImporte2 = WVector1.TextMatrix(IRow, 10)
            ZZSucursal = WVector1.TextMatrix(IRow, 12)
            
            If Val(ZZTipo2) = 2 Then
            
                ZLugarII = ZLugarII + 1
                
                ZZVectorII(ZLugarII, 1) = ZZBanco2
                ZZVectorII(ZLugarII, 2) = ZZSucursal
                ZZVectorII(ZLugarII, 3) = ZZNumero2
                ZZVectorII(ZLugarII, 4) = ZZFecha2
                ZZVectorII(ZLugarII, 5) = ZZImporte2
                
                ZSumaCheque = ZSumaCheque + Val(ZZImporte2)
                
                    Else
            
                ZLugarIII = ZLugarIII + 1
                
                If Val(ZZTipo2) = 1 Then
                    ZZVectorIII(ZLugarIII, 1) = "Efectivo"
                    ZZVectorIII(ZLugarIII, 2) = ZZImporte2
                        Else
                    ZZVectorIII(ZLugarIII, 1) = "Compensacion"
                    ZZVectorIII(ZLugarIII, 2) = ZZImporte2
                End If
            
            End If

        End If
    
    Next IRow
    

    ZZCanti = ZLugarI
    If ZLugarII > ZZCanti Then
        ZZCanti = ZLugarII
    End If
    If ZLugarIII > ZZCanti Then
        ZZCanti = ZLugarIII
    End If
    
    If ZZCanti < 20 Then
        ZZCanti = 20
            Else
        If ZZCanti < 61 Then
            ZZCanti = 60
                Else
            ZZCanti = 99
        End If
    End If
    
    For ZZCiclo = 1 To ZZCanti
    
        ZZRecibo = Recibo.Text
        ZZRenglon = Str$(ZZCiclo)
        ZZfecha = Fecha.Text
        ZZRazon = DesClientes.Caption
        ZZPesosI = XTexto1
        ZZPesosII = XTexto2
        ZZTotal = Debitos.Caption
        If Val(ZZVectorI(ZZCiclo, 3)) <> 0 Then
            ZZFechaI = ZZVectorI(ZZCiclo, 1)
            zZNumeroI = ZZVectorI(ZZCiclo, 2)
            ZZImporteI = ZZVectorI(ZZCiclo, 3)
                Else
            ZZFechaI = ""
            zZNumeroI = ""
            ZZImporteI = ""
        End If
        If Val(ZZVectorII(ZZCiclo, 5)) <> 0 Then
            ZZBanco = ZZVectorII(ZZCiclo, 1)
            ZZSucursal = ZZVectorII(ZZCiclo, 2)
            ZZNumeroII = ZZVectorII(ZZCiclo, 3)
            ZZFechaII = ZZVectorII(ZZCiclo, 4)
            ZZImporteII = ZZVectorII(ZZCiclo, 5)
                Else
            ZZBanco = ""
            ZZSucursal = ""
            ZZNumeroII = ""
            ZZFechaII = ""
            ZZImporteII = ""
        End If
        If Val(ZZVectorIII(ZZCiclo, 2)) <> 0 Then
            ZZEstructura = ZZVectorIII(ZZCiclo, 1)
            ZZImporteIII = ZZVectorIII(ZZCiclo, 2)
                Else
            ZZEstructura = ""
            ZZImporteIII = ""
        End If
    
        ZZCopia = "1"
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreRecibo ("
        ZSql = ZSql + "Copia ,"
        ZSql = ZSql + "Recibo ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "PesosI ,"
        ZSql = ZSql + "PesosII ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "FechaI ,"
        ZSql = ZSql + "NumeroI ,"
        ZSql = ZSql + "ImporteI ,"
        ZSql = ZSql + "Banco ,"
        ZSql = ZSql + "Sucursal ,"
        ZSql = ZSql + "NumeroII  ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "ImporteII ,"
        ZSql = ZSql + "Estructura ,"
        ZSql = ZSql + "ImporteIII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZCopia + "',"
        ZSql = ZSql + "'" + ZZRecibo + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZRazon + "',"
        ZSql = ZSql + "'" + ZZPesosI + "',"
        ZSql = ZSql + "'" + ZZPesosII + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZFechaI + "',"
        ZSql = ZSql + "'" + zZNumeroI + "',"
        ZSql = ZSql + "'" + ZZImporteI + "',"
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + ZZSucursal + "',"
        ZSql = ZSql + "'" + ZZNumeroII + "',"
        ZSql = ZSql + "'" + ZZFechaII + "',"
        ZSql = ZSql + "'" + ZZImporteII + "',"
        ZSql = ZSql + "'" + ZZEstructura + "',"
        ZSql = ZSql + "'" + ZZImporteIII + "')"
            
        spImpreRecibo = ZSql
        Set rstImpreRecibo = db.OpenRecordset(spImpreRecibo, dbOpenSnapshot, dbSQLPassThrough)
        
        
    
        ZZCopia = "2"
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpreRecibo ("
        ZSql = ZSql + "Copia ,"
        ZSql = ZSql + "Recibo ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "PesosI ,"
        ZSql = ZSql + "PesosII ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "FechaI ,"
        ZSql = ZSql + "NumeroI ,"
        ZSql = ZSql + "ImporteI ,"
        ZSql = ZSql + "Banco ,"
        ZSql = ZSql + "Sucursal ,"
        ZSql = ZSql + "NumeroII  ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "ImporteII ,"
        ZSql = ZSql + "Estructura ,"
        ZSql = ZSql + "ImporteIII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZCopia + "',"
        ZSql = ZSql + "'" + ZZRecibo + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZRazon + "',"
        ZSql = ZSql + "'" + ZZPesosI + "',"
        ZSql = ZSql + "'" + ZZPesosII + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZFechaI + "',"
        ZSql = ZSql + "'" + zZNumeroI + "',"
        ZSql = ZSql + "'" + ZZImporteI + "',"
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + ZZSucursal + "',"
        ZSql = ZSql + "'" + ZZNumeroII + "',"
        ZSql = ZSql + "'" + ZZFechaII + "',"
        ZSql = ZSql + "'" + ZZImporteII + "',"
        ZSql = ZSql + "'" + ZZEstructura + "',"
        ZSql = ZSql + "'" + ZZImporteIII + "')"
            
        spImpreRecibo = ZSql
        Set rstImpreRecibo = db.OpenRecordset(spImpreRecibo, dbOpenSnapshot, dbSQLPassThrough)
        
        
    Next ZZCiclo
    

    Listado.WindowTitle = "Impresion de Recibo"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT ImpreRecibo.Copia, ImpreRecibo.Recibo, ImpreRecibo.Renglon, ImpreRecibo.Fecha, ImpreRecibo.Razon, ImpreRecibo.PesosI, ImpreRecibo.PesosII, ImpreRecibo.Total, ImpreRecibo.FechaI, ImpreRecibo.NumeroI, ImpreRecibo.ImporteI, ImpreRecibo.Banco, ImpreRecibo.Sucursal, ImpreRecibo.NumeroII, ImpreRecibo.FechaII, ImpreRecibo.ImporteII, ImpreRecibo.Estructura, ImpreRecibo.ImporteIII " _
            + "From " _
            + DSQ + ".dbo.ImpreRecibo ImpreRecibo " _
            + "Where  " _
            + "ImpreRecibo.Recibo >= 0 AND " _
            + "ImpreRecibo.Recibo <= 999999"
    
    Listado.Connect = Connect()
    
    Uno = "{ImpreRecibo.Recibo} in " + "0" + " to " + "999999"
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    Listado.Destination = 0
    If ZZCanti = 99 Then
        Listado.ReportFileName = "ImpreReciboII.rpt"
        Listado.Action = 1
            Else
        Listado.ReportFileName = "ImpreRecibo.rpt"
        Listado.Action = 1
    End If

End Sub



Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = ""
    sCentimos = "centavos"
    
    Numero = CStr(Val(Debitos.Caption))
    
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

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then
        WTexto1.Text = ""
        WVector1.TextMatrix(WVector1.Row, 1) = ""
        WVector1.TextMatrix(WVector1.Row, 2) = ""
        WVector1.TextMatrix(WVector1.Row, 3) = ""
        WVector1.TextMatrix(WVector1.Row, 4) = ""
        WVector1.TextMatrix(WVector1.Row, 5) = ""
    End If
    
End Sub

Private Sub WTexto2_DblClick()

    If WVector1.Col = 1 Then
        WTexto2.Text = ""
        WVector1.TextMatrix(WVector1.Row, 1) = ""
        WVector1.TextMatrix(WVector1.Row, 2) = ""
        WVector1.TextMatrix(WVector1.Row, 3) = ""
        WVector1.TextMatrix(WVector1.Row, 4) = ""
        WVector1.TextMatrix(WVector1.Row, 5) = ""
    End If
    
End Sub
