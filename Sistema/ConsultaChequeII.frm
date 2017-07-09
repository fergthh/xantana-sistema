VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaChequeII 
   Caption         =   "Consulta de Cheques Propios"
   ClientHeight    =   7635
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   11880
   Begin VB.Frame DatosCheque 
      Height          =   2175
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   8055
      Begin VB.TextBox CodigoBanco 
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
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   21
         Text            =   " "
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox NumeroCheque 
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
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   16
         Text            =   " "
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox ImporteCheque 
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
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   15
         Text            =   " "
         Top             =   1440
         Width           =   1815
      End
      Begin MSMask.MaskEdBox FechaCheque 
         Height          =   285
         Left            =   2520
         TabIndex        =   17
         Top             =   720
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
      Begin VB.Label DesCodigoBanco 
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
         TabIndex        =   23
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codigo Banco"
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
         Index           =   3
         Left            =   720
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Numero"
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
         Index           =   2
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Fecha Cheque"
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
         Left            =   720
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe"
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
         Index           =   7
         Left            =   720
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
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
      Index           =   7
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3840
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
      Index           =   6
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   5040
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
      Index           =   5
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   4560
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4560
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
      Index           =   3
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4560
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4560
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   4560
      Width           =   375
   End
   Begin VB.Frame BusquedaCheque 
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   7335
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
         Left            =   6120
         MouseIcon       =   "ConsultaChequeII.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaChequeII.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Menu Principal"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Numero 
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
         Left            =   2280
         TabIndex        =   0
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4560
         TabIndex        =   2
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Numero"
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
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5415
      Left            =   0
      TabIndex        =   6
      Top             =   2040
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   9551
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgConsultaChequeII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Debito As Double
Private Credito As Double
Private WImpresion(200, 10) As String
Private WImpre2(100, 10) As String
Private WDebito(100, 2) As String
Private WCredito(100, 4) As String
Private WCuenta(100) As String
Private WCuenta1(100) As String
Private WCuentaBco As String
Private WVectorIva(100, 4) As String

Private WNumero As String
Private WSaldo As Double
Private WRetencion As Double
Private WReteIva As Double
Private WCuatri  As String
Private WEmpNombre As String
Private WEmpDirecion As String
Private WEmpLocalidad As String
Private WEmpCuit As String
Private WPrvDireccion As String
Private WPrvCuit As String
Private WLeyenda(10) As String
Private WTipo As String
Private WTipoprv As Single
Private WTipoiva As Single
Private WTipoReteiva As Single
Private WExepcion As Double
Private WNeto As Double
Private WAnticipo As Double
Private WBruto As Double
Private WIva As Double
Private WRetenido As Double
Private WFecha As String
Private WNroRet As Integer
Private WNroRet1 As Integer
Private XNeto As Double
Private XBruto As Double
Private XIva As Double
Private XTBase As Double
Private XImpor As Double
Private XPara(0 To 10) As Double
Private WTasa1(10) As Double
Private WAuxi As Double
Private Total As Double
Private Auxi As String
Private Auxi11 As String
Private XSaldo As Double
Private Tipocuenta As String
Private AuxiFecha As String
Private WProveedor As String
Private WTipocta As Integer
Dim BajaCheque(100) As String
Private WCtaChequeRecha As String
Dim WMinimo1 As Double
Dim WMinimo2 As Double
Dim WMinimo3 As Double
Dim WMinimo4 As Double
Dim WRetMinima As Double
Dim Existe As String
Dim PorceRIva As Double
Private XNetoIva As Double
Private XBrutoIva As Double
Private XIvaIva As Double
Private XReteIva As Double
Private WPorceFactura As Double
Private NetoParcial As Double
Private IvaParcial As Double
Private NetoTotal As Double
Private IvaTotal As Double
Private FacturaTotal As Double
Private ZSaldo As Double
Private ZLetra As String
Private ZTipo As String
Private ZPunto As String
Private ZNumero As String
Private ZProveedor As String
Dim ZMes As String
Dim ZAno As String

Dim ZZClave As String
Dim ZZOrden As String
Dim ZZRenglon As String
Dim ZZProveedor As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZTipoOrd As String
Dim ZZRetGanancias As String
Dim ZZRetIva As String
Dim ZZRetOtra As String
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
Dim ZZBanco2 As String
Dim ZZImporte2 As String
Dim ZZObservaciones2 As String
Dim ZZConcepto As String
Dim ZZObservaciones As String
Dim ZZImporte As String
Dim ZZFechaOrd2 As String
Dim ZZImpoList As String
Dim ZZCuenta As String
Dim ZZSolicitud As String
Dim ZZClaveCheque As String
Dim ZZNroRet As String
Dim ZZNroRet1 As String
Dim ZZPorceIva As String
Dim ZZPorceRIva As String
Dim ZZExepcion As String
Dim ZZImpo1 As String
Dim ZZImpo2 As String
Dim ZZImpo3 As String
Dim ZZImpo4 As String
Dim ZXBrutoIva As String
Dim ZXNetoIva As String
Dim ZXIvaIva As String

Dim ZZLetra  As String
Dim ZZTipo  As String
Dim ZZPunto  As String
Dim ZZNumero  As String
Dim ZZEstado As String
Dim ZZVencimiento As String
Dim ZZVencimiento1 As String
Dim ZZTotal As String
Dim ZZSaldo As String
Dim ZZOrdFecha As String
Dim ZZOrdVencimiento As String
Dim ZZImpre As String
Dim ZZSaldoList As String
Dim ZZNroInterno As String
Dim ZZLista As String
Dim ZZAcumulado As String
Dim ZZEmpresa As String
Dim ZZCodigoEmpresa As String

Dim ZZFecha1  As String
Dim ZZDescripcion  As String
Dim ZZDia  As String
Dim ZZMes  As String
Dim ZZAno  As String
Dim ZZNombre  As String
Dim ZConcepto  As String
Dim ZDesCuenta  As String
Dim ZZRetib  As String
Dim ZZTasa   As String

Dim WPlazo1 As Integer
Dim WVencimiento As String
Dim WAno As String
Dim ZFecha As String
Dim ZCheque(1000, 10) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Dim ZZNumeroCheque As String

Private Sub Acepta_Click()

    Call Limpia_Vector

    ZLugar = 0
    
    ZZNumero = Numero.Text
    Call Ceros(ZZNumero, 8)

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Tipo2 = " + "'" + "02" + "'"
    ZSql = ZSql + " and Pagos.Numero2 = " + "'" + ZZNumero + "'"
    ZSql = ZSql + " Order by Clave"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZSuma = ZPrueba
                    ZLugar = ZLugar + 1
                    ZCheque(ZLugar, 1) = rstPagos!Numero2
                    ZCheque(ZLugar, 2) = rstPagos!Fecha2
                    ZCheque(ZLugar, 3) = rstPagos!Banco2
                    ZCheque(ZLugar, 4) = Str$(rstPagos!Importe2)
                    ZCheque(ZLugar, 5) = rstPagos!Clave
                    ZCheque(ZLugar, 6) = rstPagos!Proveedor
                    ZCheque(ZLugar, 7) = rstPagos!Observaciones
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPagos.Close
    End If
    
    For Ciclo = 1 To ZLugar
        
        WVector1.Row = Ciclo
        ZZLugar = Ciclo
            
        WVector1.Col = 1
        WVector1.Text = ZCheque(Ciclo, 1)
        
        WVector1.Col = 2
        WVector1.Text = ZCheque(Ciclo, 2)
        
        WVector1.Col = 3
        WVector1.Text = ZCheque(Ciclo, 3)
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + ZCheque(Ciclo, 3) + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            WVector1.Col = 4
            WVector1.Text = rstBanco!Nombre
            rstBanco.Close
                Else
            WVector1.Col = 4
            WVector1.Text = ZCheque(Ciclo, 7)
        End If
        
        WVector1.Col = 5
        WVector1.Text = ZCheque(Ciclo, 4)
        WVector1.Text = Pusing("###,###.##", WVector1.Text)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZCheque(Ciclo, 6) + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            WVector1.Col = 6
            WVector1.Text = rstProveedor!Razon
            rstProveedor.Close
                Else
            WVector1.Col = 6
            WVector1.Text = ZCheque(Ciclo, 7)
        End If
            
        WVector1.Col = 7
        WVector1.Text = ZCheque(Ciclo, 5)
    
    Next Ciclo
        

End Sub

Private Sub CmdClose_Click()
    PrgConsultaChequeII.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    Numero.Text = ""
    
End Sub

Private Sub Cancela_Click()

    Call Limpia_Vector
    Numero.Text = ""
    
End Sub

Private Sub numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Acepta_Click
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
End Sub

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 1001
    
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
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 6
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 4000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 7
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case Else
                
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


Private Sub WVector1_DblClick()

    ZZClave = Trim(WVector1.TextMatrix(WVector1.Row, 7))
    
    If ZZClave <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pagos"
        ZSql = ZSql + " Where Pagos.Clave = " + "'" + ZZClave + "'"
        spPagos = ZSql
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
        
            NumeroCheque.Text = rstPagos!Numero2
            FechaCheque.Text = rstPagos!Fecha2
            CodigoBanco.Text = IIf(IsNull(rstPagos!Banco2), "", rstPagos!Banco2)
            ImporteCheque.Text = Str$(rstPagos!Importe2)
            
            DatosCheque.Visible = True
            NumeroCheque.SetFocus
        
            rstPagos.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.banco = " + "'" + CodigoBanco.Text + "'"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                DesCodigoBanco.Caption = rstBanco!Nombre
                rstBanco.Close
            End If
            
        End If
        
        NumeroCheque.SetFocus
        
    End If
        
End Sub


Private Sub NumeroCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZZNumeroCheque = NumeroCheque.Text
        Call Ceros(ZZNumeroCheque, 8)
        NumeroCheque.Text = ZZNumeroCheque
        FechaCheque.SetFocus
    End If
    If KeyAscii = 27 Then
        NumeroCheque.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub FechaCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(FechaCheque.Text)) = 8 Then
            FechaCheque.Text = Left$(FechaCheque.Text, 6) + "20" + Right$(Trim(FechaCheque.Text), 2)
        End If
        Call Valida_fecha1(FechaCheque.Text, Auxi)
        If Auxi = "S" Then
            CodigoBanco.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaCheque.Text = "  /  /    "
    End If
End Sub

Private Sub CodigoBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + CodigoBanco.Text + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            DesCodigoBanco.Caption = rstBanco!Nombre
            rstBanco.Close
            ImporteCheque.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        CodigoBanco.Text = ""
        DesCodigoBanco.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ImporteCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            
        ZZOrdFecha = Right$(FechaCheque.Text, 4) + Mid$(FechaCheque.Text, 4, 2) + Left$(FechaCheque.Text, 2)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Pagos SET "
        ZSql = ZSql + " Numero2 = " + "'" + NumeroCheque.Text + "',"
        ZSql = ZSql + " Fecha2 = " + "'" + FechaCheque.Text + "',"
        ZSql = ZSql + " FechaOrd2 = " + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + " Banco2 = " + "'" + CodigoBanco + "',"
        ZSql = ZSql + " Importe2 = " + "'" + ImporteCheque.Text + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZZClave + "'"
        spPagos = ZSql
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        
        DatosCheque.Visible = False
        Call numero_Keypress(13)
        
    End If
    If KeyAscii = 27 Then
        ImporteCheque.Text = ""
    End If
End Sub





