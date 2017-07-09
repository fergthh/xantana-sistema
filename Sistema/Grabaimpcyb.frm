VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGrabaImpcyb 
   Caption         =   "Proceso de Grabacion de Imputaciones Contables"
   ClientHeight    =   6540
   ClientLeft      =   2835
   ClientTop       =   720
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6540
   ScaleWidth      =   6210
   Begin VB.CommandButton Cancela 
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
      Left            =   4440
      MouseIcon       =   "Grabaimpcyb.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Grabaimpcyb.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salida"
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Proceso 
      Caption         =   "Confirma F1"
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
      Left            =   4440
      MouseIcon       =   "Grabaimpcyb.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "Grabaimpcyb.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Confirma Proceso de Grabacion"
      Top             =   840
      Width           =   855
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
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   5895
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
      Height          =   2595
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.TextBox EmpresaCon 
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   3135
         Begin VB.CheckBox Tipo5 
            Caption         =   "Ventas"
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
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox Tipo4 
            Caption         =   "Compras"
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
            Left            =   1560
            TabIndex        =   9
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Tipo3 
            Caption         =   "Recibos"
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
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox Tipo2 
            Caption         =   "Depositos"
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
            Left            =   240
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Tipo1 
            Caption         =   "Pagos"
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
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   2160
         TabIndex        =   12
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         Enabled         =   0   'False
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
      Begin VB.Label Label3 
         Caption         =   "Nro.Empresa"
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
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
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
         Left            =   360
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
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
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5760
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Imputa.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Movimietos de Bancos"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgGrabaImpcyb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WCtaRetGan As String
Dim WCtaRetIva As String
Dim WCtaRetOtro As String
Dim WCtaRetSuss As String
Dim WCtaDeudores As String
Dim WCtaEfectivo As String
Dim WCtaCheque As String
Dim WCtaDocumentos As String
Dim WCtaProveedores As String
Dim WCtaIva21 As String
Dim WCtaIva5 As String
Dim WCtaIva27 As String
Dim WCtaIva105 As String
Dim WCtaIb As String
Dim WCtaGanancia As String
Dim WCtaChequeRecha As String
Dim WCtaIvaVen As String
Dim WCtaVentas As String
Dim WImpoIva As Double
Dim WMes As String
Dim WAno As String
Dim Mes1 As String
Dim Ano1 As String
Dim WGraba(100, 20) As String
Dim WAsiento As String
Dim XCuenta As String
Dim XDebito As Double
Dim XCredito As Double
Dim WDia As String
Dim ZImporte As Double

Private Sub Proceso_Click()

    If Right$(DesdeFecha.Text, 7) <> Right$(HastaFecha.Text, 7) Then
        m$ = "Parametros de Fecha Incorrecto"
        A% = MsgBox(m$, 0, "Grabacion de Imputaciones Contables")
        Exit Sub
    End If
    
    If Val(EmpresaCon.Text) = 0 Then
        m$ = "Se debe informar el numero de empresa del sistema de Contabilidad"
        A% = MsgBox(m$, 0, "Grabacion de Imputaciones Contables")
        Exit Sub
    End If
    
    DiaDesde = Val(Left$(DesdeFecha.Text, 2))
    Diahasta = Val(Left$(HastaFecha.Text, 2))
    
    WEmpresaConta = EmpresaCon.Text
    Call Ceros(WEmpresaConta, 4)
    OPEN_FILE_CuentaCon
    
    Do
        WDia = Str$(DiaDesde)
        Call Ceros(WDia, 2)
        Fecha.Text = WDia + Mid$(DesdeFecha.Text, 3, 8)
        Call ProcesoDia_Click
        DiaDesde = DiaDesde + 1
        If DiaDesde > Diahasta Then
            Exit Do
        End If
    Loop
    
    Call Cancela_Click

End Sub

Private Sub ProcesoDia_Click()

    On Error GoTo Error_Programa
    
    With rstEmpreCon
        .Index = "Codigo"
        .Seek "=", Val(EmpresaCon.Text)
        If .NoMatch = False Then
            WDesdefecha = !DesdeFecha
            WHastafecha = !HastaFecha
        End If
    End With
    
    WAno = Right$(WDesdefecha, 4)
    WMes = Mid$(WDesdefecha, 4, 2)
    WDia = Left$(WDesdefecha, 2)
    WDesdefecha = WAno + WMes + WDia
    
    WAno = Right$(WHastafecha, 4)
    WMes = Mid$(WHastafecha, 4, 2)
    WDia = Left$(WHastafecha, 2)
    WHastafecha = WAno + WMes + WDia
    
    WMes = Mid$(Fecha.Text, 4, 2)
    WAno = Right$(Fecha.Text, 4)
    
    Call Ceros(WMes, 2)
    Call Ceros(WAno, 4)
    WPeriodo = WAno + WMes + "31"
    
    If WPeriodo >= WDesdefecha And WPeriodo <= WHastafecha Then
    
        WMes = Mid$(Fecha.Text, 4, 2)
        WAno = Right$(Fecha.Text, 4)
        
        Posicion = 0
        Mes1 = Mid$(WDesdefecha, 5, 2)
        Ano1 = Left$(WDesdefecha, 4)
        
        Do
        
            Posicion = Posicion + 1
            Call Ceros(Mes1, 2)
            Call Ceros(Ano1, 4)
            Compara = Ano1 + Mes1
            
            If Left$(WPeriodo, 6) = Left$(Compara, 6) Then
                Exit Do
            End If
            
            Mes1 = Str$(Val(Mes1) + 1)
            If Val(Mes1) > 12 Then
                Mes1 = 1
                Ano1 = Str$(Val(Ano1) + 1)
            End If
            
        Loop
        
        WPosi = Posicion
        
        OPEN_FILE_Asiento

    End If
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WHasta = WAno + WMes + WDia
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
            WCtaRetGan = !CtaRetGan
            WCtaRetIva = !CtaRetIva
            WCtaRetOtro = !CtaRetOtro
            WCtaRetSuss = !CtaRetSuss
            WCtaDeudores = !CtaDeudores
            WCtaEfectivo = !CtaEfectivo
            WCtaCheque = !CtaCheque
            WCtaDocumentos = !CtaDocumentos
            WCtaProveedores = !CtaProveedores
            WCtaIva21 = !CtaIva21
            WCtaIva5 = !CtaIva5
            WCtaIva27 = !CtaIva27
            WCtaIva105 = !CtaIva105
            WCtaIb = !CtaIb
            WCtaGanancia = !CtaGanancia
            WCtaChequeRecha = !CtaChequeRecha
            WCtaIvaVen = !CtaIvaven
            WCtaVentas = !CtaVentas
        End If
    End With
    
    Rem Procesa los pagos

    If Tipo1.Value = 1 Then
    
        Pasa = 0
        Erase WGraba
        Lugar = 0

        With rstPagos
            .Index = "Clave"
            .MoveFirst
            Do
                If WDesde <= !fechaord And !fechaord <= WHasta Then
                
                Rem If !Solicitud = 2 Then
                
                    If Pasa = 0 Then
                        Pasa = 1
                        Corte = !Orden
                        Erase WGraba
                        Lugar = 0
                    End If
                    
                    If Corte <> !Orden Then
                    
                        With rstAsiento
                            .Index = "Asiento"
                            Claveven$ = "99999999"
                            .Seek "<=", Claveven$
                            If .NoMatch = False Then
                                WAsiento = !Asiento + 1
                                    Else
                                WAsiento = "1"
                            End If
                        End With
                        Call Ceros(WAsiento, 6)
                    
                        For WCiclo = 1 To Lugar
                        
                            WFecha = WGraba(WCiclo, 1)
                            ZZObservaciones = WGraba(WCiclo, 2)
                            WCuenta = WGraba(WCiclo, 3)
                            WDebito = Val(WGraba(WCiclo, 4))
                            WCredito = Val(WGraba(WCiclo, 5))
                            WLeyenda = WGraba(WCiclo, 6)
                            WOrden = WGraba(WCiclo, 7)
                            WTipo = WGraba(WCiclo, 8)
                            WLetra = WGraba(WCiclo, 9)
                            WPunto = WGraba(WCiclo, 10)
                            WNumero = WGraba(WCiclo, 11)
                        
                            With rstAsiento
                                .Index = "Clave"
                                .AddNew
                                !Asiento = Val(WAsiento)
                                Auxi1 = Str$(WCiclo)
                                Call Ceros(Auxi1, 2)
                                !Renglon = WCiclo
                                !Fecha = WFecha
                                !Observaciones = Left$(ZZObservaciones, 50)
                                !Cuenta = WCuenta
                                !Debito = WDebito
                                !Credito = WCredito
                                !Leyenda = Left$(WLeyenda, 50)
                                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Clave = WAsiento + Auxi1
                                .Update
                            
                                XCuenta = WCuenta
                                XDebito = WDebito
                                XCredito = WCredito
                            
                                Call Actualiza_Saldos
                                
                            End With
                        Next WCiclo
                    
                        Corte = !Orden
                        Erase WGraba
                        Lugar = 0
                        
                    End If
                    
                    WOrden = !Orden
                    WFecha = !Fecha
                    WFechaOrd = !fechaord
                    WClave = !Clave
                    WProveedor = !Proveedor
                    
                    XObservaciones = Trim(!Observaciones)
                    WObservaciones = "O.P.:" + Str$(WOrden) + " " + XObservaciones
                    If Val(WProveedor) <> 0 Then
                        With RstProveedor
                            .Index = "Proveedor"
                            .Seek "=", WProveedor
                            If .NoMatch = False Then
                                WObservaciones = "O.P.:" + Str$(WOrden) + " " + Trim(!Nombre) + " " + XObservaciones
                            End If
                        End With
                    End If
                        
                    Select Case Val(!Tiporeg)
                        Case 1
                            If !TipoOrd = "3" Or !TipoOrd = "4" Or !TipoOrd = "5" Then
                                WCuenta = !Cuenta
                                    Else
                                Rem proveedor
                                WCuenta = WCtaProveedores
                                WProveedor = !Proveedor
                                With RstProveedor
                                    .Index = "Proveedor"
                                    .Seek "=", WProveedor
                                    If .NoMatch = False Then
                                        WTipoProveedor = !Tipo
                                        With rstTipopro
                                            .Index = "Codigo"
                                            .Seek "=", WTipoProveedor
                                            If .NoMatch = False Then
                                                WCuenta = !Cuenta
                                            End If
                                        End With
                                    End If
                                End With
                            End If
                            
                            WDebito = !Importe1
                            WCredito = 0
                            WTipo = !Tipo1
                            WLetra = !Letra1
                            WPunto = !Punto1
                            WNumero = !Numero1
                            
                            Select Case Val(!Tipo1)
                                Case 1
                                    WImpre = "FAC"
                                Case 2
                                    WImpre = "N/D"
                                Case 3
                                    WImpre = "N/C"
                                Case Else
                                    WImpre = ""
                            End Select
                            
                            Lugar = Lugar + 1
                            WGraba(Lugar, 1) = WFecha
                            WGraba(Lugar, 2) = "Asiento de Pagos"
                            WGraba(Lugar, 3) = WCuenta
                            WGraba(Lugar, 4) = Str$(WDebito)
                            WGraba(Lugar, 5) = Str$(WCredito)
                            WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                            WGraba(Lugar, 7) = WOrden
                            WGraba(Lugar, 8) = WTipo
                            WGraba(Lugar, 9) = WLetra
                            WGraba(Lugar, 10) = WPunto
                            WGraba(Lugar, 11) = WNumero
                            
                            If Val(!Renglon) = 1 And !Retencion <> 0 Then
                            
                                WCredito = !Retencion
                                WDebito = 0
                                WTipo = 0
                                WLetra = ""
                                WPunto = 0
                                WNumero = 0
                                
                                Lugar = Lugar + 1
                                WGraba(Lugar, 1) = WFecha
                                WGraba(Lugar, 2) = "Asiento de Pagos"
                                WGraba(Lugar, 3) = WCtaGanancia
                                WGraba(Lugar, 4) = Str$(WDebito)
                                WGraba(Lugar, 5) = Str$(WCredito)
                                WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                                WGraba(Lugar, 7) = WOrden
                                WGraba(Lugar, 8) = WTipo
                                WGraba(Lugar, 9) = WLetra
                                WGraba(Lugar, 10) = WPunto
                                WGraba(Lugar, 11) = WNumero
                                
                           End If
                                                            
                        Case Else
                            Select Case Val(!Tipo2)
                                Case 1
                                    Rem caja
                                    WCuenta = WCtaEfectivo
                                    WImpre = "EFTO."
                                Case 2
                                    Rem banco
                                    WImpre = "BCO"
                                    WBanco2 = !Banco2
                                    With rstBanco
                                        .Index = "Banco"
                                        .Seek "=", WBanco2
                                        If .NoMatch = False Then
                                            WCuenta = !Cuenta
                                                Else
                                            WCuenta = ""
                                        End If
                                    End With
                                Case 3
                                    Rem che ter
                                    WImpre = "CH.TER"
                                    WCuenta = WCtaCheque
                                Case Else
                                    Rem documentos
                                    WImpre = "VARIOS."
                                    WCuenta = !Cuenta
                            End Select
                            
                            WDebito = 0
                            WCredito = !Importe2
                            WProveedor = !Proveedor
                            WLetra = ""
                            WPunto = 0
                            WNumero = !Numero2
                            WFecha = !Fecha
                            WFechaOrd = !fechaord
                            
                            Lugar = Lugar + 1
                            WGraba(Lugar, 1) = WFecha
                            WGraba(Lugar, 2) = "Asiento de Pagos"
                            WGraba(Lugar, 3) = WCuenta
                            WGraba(Lugar, 4) = Str$(WDebito)
                            WGraba(Lugar, 5) = Str$(WCredito)
                            WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                            WGraba(Lugar, 7) = WOrden
                            WGraba(Lugar, 8) = WTipo
                            WGraba(Lugar, 9) = WLetra
                            WGraba(Lugar, 10) = WPunto
                            WGraba(Lugar, 11) = WNumero
                            
                            If Val(!Renglon) = 1 And !Retencion <> 0 Then
                            
                                WCredito = !Retencion
                                WDebito = 0
                                WTipo = 0
                                WLetra = ""
                                WPunto = 0
                                WNumero = 0
                            
                                Lugar = Lugar + 1
                                WGraba(Lugar, 1) = WFecha
                                WGraba(Lugar, 2) = "Asiento de Pagos"
                                WGraba(Lugar, 3) = WCtaGanancia
                                WGraba(Lugar, 4) = Str$(WDebito)
                                WGraba(Lugar, 5) = Str$(WCredito)
                                WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                                WGraba(Lugar, 7) = WOrden
                                WGraba(Lugar, 8) = WTipo
                                WGraba(Lugar, 9) = WLetra
                                WGraba(Lugar, 10) = WPunto
                                WGraba(Lugar, 11) = WNumero
                            
                            End If
                            
                    End Select
                    
                Rem End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    
    End If
    
    If Pasa <> 0 Then
                
        With rstAsiento
            .Index = "Asiento"
            Claveven$ = "99999999"
            .Seek "<=", Claveven$
            If .NoMatch = False Then
                WAsiento = !Asiento + 1
                    Else
                WAsiento = "1"
            End If
        End With
        Call Ceros(WAsiento, 6)
                    
        For WCiclo = 1 To Lugar
                        
            WFecha = WGraba(WCiclo, 1)
            ZZObservaciones = WGraba(WCiclo, 2)
            WCuenta = WGraba(WCiclo, 3)
            WDebito = Val(WGraba(WCiclo, 4))
            WCredito = Val(WGraba(WCiclo, 5))
            WLeyenda = WGraba(WCiclo, 6)
            WOrden = WGraba(WCiclo, 7)
            WTipo = WGraba(WCiclo, 8)
            WLetra = WGraba(WCiclo, 9)
            WPunto = WGraba(WCiclo, 10)
            WNumero = WGraba(WCiclo, 11)
                        
            With rstAsiento
                .Index = "Clave"
                .AddNew
                !Asiento = Val(WAsiento)
                Auxi1 = Str$(WCiclo)
                Call Ceros(Auxi1, 2)
                !Renglon = WCiclo
                !Fecha = WFecha
                !Observaciones = Left$(ZZObservaciones, 50)
                !Cuenta = WCuenta
                !Debito = WDebito
                !Credito = WCredito
                !Leyenda = Left$(WLeyenda, 50)
                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                !Clave = WAsiento + Auxi1
                .Update
                            
                XCuenta = WCuenta
                XDebito = WDebito
                XCredito = WCredito
                            
                Call Actualiza_Saldos
                                
            End With
        Next WCiclo
                        
    End If
    
    Rem Procesa los depositos
    
    If Tipo2.Value = 1 Then
    
        Pasa = 0
        Erase WGraba
        Lugar = 0
    
        With rstDepositos
            .Index = "Clave"
            .MoveFirst
            Do
                If WDesde <= !fechaord And !fechaord <= WHasta Then
                
                    If Pasa = 0 Then
                        Pasa = 1
                        Erase WGraba
                        Lugar = 0
                        Corte = !Deposito
                    End If
                    
                    If Corte <> !Deposito Then
                    
                        With rstAsiento
                            .Index = "Asiento"
                            Claveven$ = "99999999"
                            .Seek "<=", Claveven$
                            If .NoMatch = False Then
                                WAsiento = !Asiento + 1
                                    Else
                                WAsiento = "1"
                            End If
                        End With
                        Call Ceros(WAsiento, 6)
                    
                        For WCiclo = 1 To Lugar
                        
                            WFecha = WGraba(WCiclo, 1)
                            ZZObservaciones = WGraba(WCiclo, 2)
                            WCuenta = WGraba(WCiclo, 3)
                            WDebito = Val(WGraba(WCiclo, 4))
                            WCredito = Val(WGraba(WCiclo, 5))
                            WLeyenda = WGraba(WCiclo, 6)
                            WOrden = WGraba(WCiclo, 7)
                            WTipo = WGraba(WCiclo, 8)
                            WLetra = WGraba(WCiclo, 9)
                            WPunto = WGraba(WCiclo, 10)
                            WNumero = WGraba(WCiclo, 11)
                            
                            With rstAsiento
                                .Index = "Clave"
                                .AddNew
                                !Asiento = Val(WAsiento)
                                Auxi1 = Str$(WCiclo)
                                Call Ceros(Auxi1, 2)
                                !Renglon = WCiclo
                                !Fecha = WFecha
                                !Observaciones = Left$(ZZObservaciones, 50)
                                !Cuenta = WCuenta
                                !Debito = WDebito
                                !Credito = WCredito
                                !Leyenda = Left$(WLeyenda, 50)
                                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Clave = WAsiento + Auxi1
                                .Update
                            
                                XCuenta = WCuenta
                                XDebito = WDebito
                                XCredito = WCredito
                            
                                Call Actualiza_Saldos
                                
                            End With
                        Next WCiclo
                    
                        Corte = !Deposito
                        Erase WGraba
                        Lugar = 0
                    
                    End If
                
                    WDeposito = !Deposito
                    WFecha = !Fecha
                    WFechaOrd = !fechaord
                    WClave = !Clave
                    WBanco = !Banco
                    WImporte = !Importe2
                    WTipo = !Tipo2
                    WLetra = ""
                    WPunto = 0
                    WNumero = !Numero2
                    
                    With rstBanco
                        .Index = "Banco"
                        .Seek "=", WBanco
                        If .NoMatch = False Then
                            WCuenta = !Cuenta
                        End If
                    End With
                    
                    If Val(WTipo) = 1 Then
                        WObservaciones = "Deposito Nro.:" + Str$(WDeposito) + " Efectivo"
                        WImpre = ""
                            Else
                        WObservaciones = "Deposito Nro.:" + Str$(WDeposito) + " Cheque:" + WNumero
                        WImpre = "Cheque"
                    End If
                    
                    Lugar = Lugar + 1
                    WGraba(Lugar, 1) = WFecha
                    WGraba(Lugar, 2) = "Asiento de Depositos"
                    WGraba(Lugar, 3) = WCuenta
                    WGraba(Lugar, 4) = Str$(WImporte)
                    WGraba(Lugar, 5) = "0"
                    WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                    WGraba(Lugar, 7) = WDeposito
                    WGraba(Lugar, 8) = WTipo
                    WGraba(Lugar, 9) = WLetra
                    WGraba(Lugar, 10) = WPunto
                    WGraba(Lugar, 11) = WNumero
                        
                    Rem With rstBanco
                    Rem     .Index = "Banco"
                    Rem     .Seek "=", WBanco
                    Rem     If .NoMatch = False Then
                    Rem         wobservaciones = !Nombre
                    Rem     End If
                    Rem End With
                        
                    Select Case Val(WTipo)
                        Case 1
                            Rem EFECTIVO
                            WCuenta = WCtaEfectivo
                        Case Else
                            Rem valores en cartera
                            WCuenta = WCtaCheque
                    End Select
                    
                    Lugar = Lugar + 1
                    WGraba(Lugar, 1) = WFecha
                    WGraba(Lugar, 2) = "Asiento de Depositos"
                    WGraba(Lugar, 3) = WCuenta
                    WGraba(Lugar, 4) = Str$(0)
                    WGraba(Lugar, 5) = Str$(WImporte)
                    WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                    WGraba(Lugar, 7) = WDeposito
                    WGraba(Lugar, 8) = WTipo
                    WGraba(Lugar, 9) = WLetra
                    WGraba(Lugar, 10) = WPunto
                    WGraba(Lugar, 11) = WNumero
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    
    End If
    
    If Pasa <> 0 Then
                
        With rstAsiento
            .Index = "Asiento"
            Claveven$ = "99999999"
            .Seek "<=", Claveven$
            If .NoMatch = False Then
                WAsiento = !Asiento + 1
                    Else
                WAsiento = "1"
            End If
        End With
        Call Ceros(WAsiento, 6)
                    
        For WCiclo = 1 To Lugar
                        
            WFecha = WGraba(WCiclo, 1)
            ZZObservaciones = WGraba(WCiclo, 2)
            WCuenta = WGraba(WCiclo, 3)
            WDebito = Val(WGraba(WCiclo, 4))
            WCredito = Val(WGraba(WCiclo, 5))
            WLeyenda = WGraba(WCiclo, 6)
            WOrden = WGraba(WCiclo, 7)
            WTipo = WGraba(WCiclo, 8)
            WLetra = WGraba(WCiclo, 9)
            WPunto = WGraba(WCiclo, 10)
            WNumero = WGraba(WCiclo, 11)
                        
            With rstAsiento
                .Index = "Clave"
                .AddNew
                !Asiento = Val(WAsiento)
                Auxi1 = Str$(WCiclo)
                Call Ceros(Auxi1, 2)
                !Renglon = WCiclo
                !Fecha = WFecha
                !Observaciones = Left$(ZZObservaciones, 50)
                !Cuenta = WCuenta
                !Debito = WDebito
                !Credito = WCredito
                !Leyenda = Left$(WLeyenda, 50)
                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                !Clave = WAsiento + Auxi1
                .Update
                            
                XCuenta = WCuenta
                XDebito = WDebito
                XCredito = WCredito
                            
                Call Actualiza_Saldos
                                
            End With
        Next WCiclo
                        
    End If
    
    
    
    
    Rem Procesa las Cobranzas
    
    If Tipo3.Value = 1 Then

        Pasa = 0
        Erase WGraba
        Lugar = 0
        
        With rstRecibos
            .Index = "Clave"
            .MoveFirst
            Do
                If WDesde <= !fechaord And !fechaord <= WHasta Then
                
                    If Pasa = 0 Then
                        Pasa = 1
                        Erase WGraba
                        Lugar = 0
                        Corte = !Recibo
                    End If
                    
                    If Corte <> !Recibo Then
                    
                        With rstAsiento
                            .Index = "Asiento"
                            Claveven$ = "99999999"
                            .Seek "<=", Claveven$
                            If .NoMatch = False Then
                                WAsiento = !Asiento + 1
                                    Else
                                WAsiento = "1"
                            End If
                        End With
                        Call Ceros(WAsiento, 6)
                    
                        For WCiclo = 1 To Lugar
                        
                            WFecha = WGraba(WCiclo, 1)
                            ZZObservaciones = WGraba(WCiclo, 2)
                            WCuenta = WGraba(WCiclo, 3)
                            WDebito = Val(WGraba(WCiclo, 4))
                            WCredito = Val(WGraba(WCiclo, 5))
                            WLeyenda = WGraba(WCiclo, 6)
                            WOrden = WGraba(WCiclo, 7)
                            WTipo = WGraba(WCiclo, 8)
                            WLetra = WGraba(WCiclo, 9)
                            WPunto = WGraba(WCiclo, 10)
                            WNumero = WGraba(WCiclo, 11)
                            
                            With rstAsiento
                                .Index = "Clave"
                                .AddNew
                                !Asiento = Val(WAsiento)
                                Auxi1 = Str$(WCiclo)
                                Call Ceros(Auxi1, 2)
                                !Renglon = WCiclo
                                !Fecha = WFecha
                                !Observaciones = Left$(ZZObservaciones, 50)
                                !Cuenta = WCuenta
                                !Debito = WDebito
                                !Credito = WCredito
                                !Leyenda = Left$(WLeyenda, 50)
                                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Clave = WAsiento + Auxi1
                                .Update
                            
                                XCuenta = WCuenta
                                XDebito = WDebito
                                XCredito = WCredito
                            
                                Call Actualiza_Saldos
                                
                            End With
                        Next WCiclo
                    
                        Corte = !Recibo
                        Erase WGraba
                        Lugar = 0
                    
                    End If
                    
                    WClave = !Clave
                    WRecibo = !Recibo
                    WRenglon = !Renglon
                    WFecha = !Fecha
                    WFechaOrd = !fechaord
                    
                    XObservaciones = Trim(!Observaciones)
                    If !TipoRec = "3" Then
                        WObservaciones = "Recibo:" + Str$(WRecibo) + " " + XObservaciones
                            Else
                        Wcliente = !Cliente
                        WObservaciones = "Recibo:" + Str$(WRecibo) + " " + XObservaciones
                        With rstClientes
                            .Index = "Cliente"
                            .Seek "=", Wcliente
                            If .NoMatch = False Then
                                WObservaciones = "Recibo:" + Str$(WRecibo) + " " + Trim(!Razon) + " " + XObservaciones
                            End If
                        End With
                    End If
                            
                    Select Case Val(!Tiporeg)
                        Case 1
                            If !TipoRec = "3" Then
                                WCuenta = !Cuenta
                                    Else
                                Rem clientes
                                WCuenta = WCtaDeudores
                            End If
                            
                            WLetra = !Letra1
                            WTipo = !Tipo1
                            WPunto = !Punto1
                            WNumero = !Numero1
                            WImporte = !Importe1
                            
                            Select Case Val(!Tipo1)
                                Case 3
                                    WImpre = "FAC"
                                Case 4
                                    WImpre = "N/D"
                                Case 5
                                    WImpre = "N/C"
                                Case Else
                                    WImpre = ""
                            End Select
                            
                            Lugar = Lugar + 1
                            WGraba(Lugar, 1) = WFecha
                            WGraba(Lugar, 2) = "Asiento de Recibos"
                            WGraba(Lugar, 3) = WCuenta
                            WGraba(Lugar, 4) = "0"
                            WGraba(Lugar, 5) = Str$(WImporte)
                            WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                            WGraba(Lugar, 7) = WRecibo
                            WGraba(Lugar, 8) = WTipo
                            WGraba(Lugar, 9) = WLetra
                            WGraba(Lugar, 10) = WPunto
                            WGraba(Lugar, 11) = WNumero
                                                            
                        Case Else
                            Select Case Val(!Tipo2)
                                Case 1
                                    Rem caja
                                    WCuenta = WCtaEfectivo
                                    WImpre = "EFTO."
                                Case 2
                                    Rem cheques
                                    WImpre = "CH.TER"
                                    WCuenta = WCtaCheque
                                Case 4
                                    WImpre = ""
                                    WCuenta = !Cuenta
                                Case Else
                                    Rem documentos
                                    WImpre = "DOC."
                                    WCuenta = WCtaDocumentos
                            End Select
                                    
                            WLetra = ""
                            WTipo = !Tipo2
                            WPunto = 0
                            WNumero = !Numero2
                            WImporte = !Importe2
                            
                            Lugar = Lugar + 1
                            WGraba(Lugar, 1) = WFecha
                            WGraba(Lugar, 2) = "Asiento de Recibos"
                            WGraba(Lugar, 3) = WCuenta
                            WGraba(Lugar, 4) = Str$(WImporte)
                            WGraba(Lugar, 5) = "0"
                            WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                            WGraba(Lugar, 7) = WRecibo
                            WGraba(Lugar, 8) = WTipo
                            WGraba(Lugar, 9) = WLetra
                            WGraba(Lugar, 10) = WPunto
                            WGraba(Lugar, 11) = WNumero
                            
                    End Select
                    
                    If Val(!Renglon) = 1 And Val(!Retganancias) <> 0 Then
                    
                        WLetra = ""
                        WTipo = 0
                        WPunto = 0
                        WNumero = 0
                        WImporte = !Retganancias
                        WCuenta = WCtaRetGan
                        
                        Lugar = Lugar + 1
                        WGraba(Lugar, 1) = WFecha
                        WGraba(Lugar, 2) = "Asiento de Recibos"
                        WGraba(Lugar, 3) = WCuenta
                        WGraba(Lugar, 4) = Str$(WImporte)
                        WGraba(Lugar, 5) = "0"
                        WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                        WGraba(Lugar, 7) = WRecibo
                        WGraba(Lugar, 8) = WTipo
                        WGraba(Lugar, 9) = WLetra
                        WGraba(Lugar, 10) = WPunto
                        WGraba(Lugar, 11) = WNumero
                            
                    End If
                
                    If Val(!Renglon) = 1 And Val(!RetIva) <> 0 Then
                    
                        WLetra = ""
                        WTipo = 0
                        WPunto = 0
                        WNumero = 0
                        WImporte = !RetIva
                        WCuenta = WCtaRetIva
                            
                        Lugar = Lugar + 1
                        WGraba(Lugar, 1) = WFecha
                        WGraba(Lugar, 2) = "Asiento de Recibos"
                        WGraba(Lugar, 3) = WCuenta
                        WGraba(Lugar, 4) = Str$(WImporte)
                        WGraba(Lugar, 5) = "0"
                        WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                        WGraba(Lugar, 7) = WRecibo
                        WGraba(Lugar, 8) = WTipo
                        WGraba(Lugar, 9) = WLetra
                        WGraba(Lugar, 10) = WPunto
                        WGraba(Lugar, 11) = WNumero
                        
                    End If
                
                    If Val(!Renglon) = 1 And Val(!RetOtra) <> 0 Then
                    
                        WLetra = ""
                        WTipo = 0
                        WPunto = 0
                        WNumero = 0
                        WImporte = !RetOtra
                        WCuenta = WCtaRetOtro
                                
                        Lugar = Lugar + 1
                        WGraba(Lugar, 1) = WFecha
                        WGraba(Lugar, 2) = "Asiento de Recibos"
                        WGraba(Lugar, 3) = WCuenta
                        WGraba(Lugar, 4) = Str$(WImporte)
                        WGraba(Lugar, 5) = "0"
                        WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                        WGraba(Lugar, 7) = WRecibo
                        WGraba(Lugar, 8) = WTipo
                        WGraba(Lugar, 9) = WLetra
                        WGraba(Lugar, 10) = WPunto
                        WGraba(Lugar, 11) = WNumero
                        
                    End If
                    
                    If Val(!Renglon) = 1 And Val(!RetSuss) <> 0 Then
                    
                        WLetra = ""
                        WTipo = 0
                        WPunto = 0
                        WNumero = 0
                        WImporte = !RetSuss
                        WCuenta = WCtaRetSuss
                                
                        Lugar = Lugar + 1
                        WGraba(Lugar, 1) = WFecha
                        WGraba(Lugar, 2) = "Asiento de Recibos"
                        WGraba(Lugar, 3) = WCuenta
                        WGraba(Lugar, 4) = Str$(WImporte)
                        WGraba(Lugar, 5) = "0"
                        WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                        WGraba(Lugar, 7) = WRecibo
                        WGraba(Lugar, 8) = WTipo
                        WGraba(Lugar, 9) = WLetra
                        WGraba(Lugar, 10) = WPunto
                        WGraba(Lugar, 11) = WNumero
                        
                    End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    
    End If
    
    If Pasa <> 0 Then
                
        With rstAsiento
            .Index = "Asiento"
            Claveven$ = "99999999"
            .Seek "<=", Claveven$
            If .NoMatch = False Then
                WAsiento = !Asiento + 1
                    Else
                WAsiento = "1"
            End If
        End With
        Call Ceros(WAsiento, 6)
                    
        For WCiclo = 1 To Lugar
                        
            WFecha = WGraba(WCiclo, 1)
            ZZObservaciones = WGraba(WCiclo, 2)
            WCuenta = WGraba(WCiclo, 3)
            WDebito = Val(WGraba(WCiclo, 4))
            WCredito = Val(WGraba(WCiclo, 5))
            WLeyenda = WGraba(WCiclo, 6)
            WOrden = WGraba(WCiclo, 7)
            WTipo = WGraba(WCiclo, 8)
            WLetra = WGraba(WCiclo, 9)
            WPunto = WGraba(WCiclo, 10)
            WNumero = WGraba(WCiclo, 11)
                        
            With rstAsiento
                .Index = "Clave"
                .AddNew
                !Asiento = Val(WAsiento)
                Auxi1 = Str$(WCiclo)
                Call Ceros(Auxi1, 2)
                !Renglon = WCiclo
                !Fecha = WFecha
                !Observaciones = Left$(ZZObservaciones, 50)
                !Cuenta = WCuenta
                !Debito = WDebito
                !Credito = WCredito
                !Leyenda = Left$(WLeyenda, 50)
                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                !Clave = WAsiento + Auxi1
                .Update
                            
                XCuenta = WCuenta
                XDebito = WDebito
                XCredito = WCredito
                            
                Call Actualiza_Saldos
                                
            End With
        Next WCiclo
                        
    End If
    
    
    
    
    Rem Procesa las Compras
    
    If Tipo4.Value = 1 Then
    
        Pasa = 0
        Erase WGraba
        Lugar = 0
        
        With rstImpcyb
            .Index = "Clave"
            .MoveFirst
            Do
            
                Rem dada
                WOrdFecha = !ordfecha
                WFecha = !Fecha
            
                WClave = Left$(!Clave, 21)
                With rstIvacomp
                    .Index = "Clave"
                    .Seek "=", WClave
                    If .NoMatch = False Then
                        WAno = Right$(!Periodo, 4)
                        WMes = Mid$(!Periodo, 4, 2)
                        WDia = Left$(!Periodo, 2)
                        WFecha = !Periodo
                        WOrdFecha = WAno + WMes + WDia
                    End If
                End With
            
                If WDesde <= WOrdFecha And WOrdFecha <= WHasta Then
                
                    Rem If !Letra <> "X" Then
                
                        If Pasa = 0 Then
                            Pasa = 1
                            Erase WGraba
                            Lugar = 0
                            Corte = Left$(!Clave, 21)
                        End If
                    
                        If Corte <> Left$(!Clave, 21) Then
                    
                            With rstAsiento
                                .Index = "Asiento"
                                Claveven$ = "99999999"
                                .Seek "<=", Claveven$
                                If .NoMatch = False Then
                                    WAsiento = !Asiento + 1
                                        Else
                                    WAsiento = "1"
                                End If
                            End With
                            Call Ceros(WAsiento, 6)
                    
                            For WCiclo = 1 To Lugar
                        
                                WFecha = WGraba(WCiclo, 1)
                                ZZObservaciones = WGraba(WCiclo, 2)
                                WCuenta = WGraba(WCiclo, 3)
                                WDebito = Val(WGraba(WCiclo, 4))
                                WCredito = Val(WGraba(WCiclo, 5))
                                WLeyenda = WGraba(WCiclo, 6)
                                WOrden = WGraba(WCiclo, 7)
                                WTipo = WGraba(WCiclo, 8)
                                WLetra = WGraba(WCiclo, 9)
                                WPunto = WGraba(WCiclo, 10)
                                WNumero = WGraba(WCiclo, 11)
                            
                                With rstAsiento
                                    .Index = "Clave"
                                    .AddNew
                                    !Asiento = Val(WAsiento)
                                    Auxi1 = Str$(WCiclo)
                                    Call Ceros(Auxi1, 2)
                                    !Renglon = WCiclo
                                    !Fecha = WFecha
                                    !Observaciones = Left$(ZZObservaciones, 50)
                                    !Cuenta = WCuenta
                                    !Debito = WDebito
                                    !Credito = WCredito
                                    !Leyenda = Left$(WLeyenda, 50)
                                    !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    !Clave = WAsiento + Auxi1
                                    .Update
                            
                                    XCuenta = WCuenta
                                    XDebito = WDebito
                                    XCredito = WCredito
                            
                                    Call Actualiza_Saldos
                                
                                End With
                            Next WCiclo
                    
                            Corte = Left$(!Clave, 21)
                            Erase WGraba
                            Lugar = 0
                    
                        End If
                        
                        
                        WProveedor = !Proveedor
                        WObservaciones = "Fc" + Str$(!Numero) + " " + Trim(!Observaciones)
                        XObservaciones = Trim(!Observaciones)
                        XNumero = Str$(!Numero)
                        With RstProveedor
                            .Index = "Proveedor"
                            .Seek "=", WProveedor
                            If .NoMatch = False Then
                                WObservaciones = "Fc" + XNumero + " " + Trim(!Nombre) + " " + XObservaciones
                            End If
                        End With
                        
                        Lugar = Lugar + 1
                        WGraba(Lugar, 1) = !Fecha
                        WGraba(Lugar, 2) = "Asiento de Compras"
                        WGraba(Lugar, 3) = !Cuenta
                        WGraba(Lugar, 4) = Str$(!Debito)
                        WGraba(Lugar, 5) = Str$(!Credito)
                        WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                        WGraba(Lugar, 7) = ""
                        WGraba(Lugar, 8) = Str$(!Tipo)
                        WGraba(Lugar, 9) = !Letra
                        WGraba(Lugar, 10) = Str$(!Punto)
                        WGraba(Lugar, 11) = Str$(!Numero)
                    
                    Rem End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    
    End If
    
    If Pasa <> 0 Then
                
        With rstAsiento
            .Index = "Asiento"
            Claveven$ = "99999999"
            .Seek "<=", Claveven$
            If .NoMatch = False Then
                WAsiento = !Asiento + 1
                    Else
                WAsiento = "1"
            End If
        End With
        Call Ceros(WAsiento, 6)
                    
        For WCiclo = 1 To Lugar
                        
            WFecha = WGraba(WCiclo, 1)
            ZZObservaciones = WGraba(WCiclo, 2)
            WCuenta = WGraba(WCiclo, 3)
            WDebito = Val(WGraba(WCiclo, 4))
            WCredito = Val(WGraba(WCiclo, 5))
            WLeyenda = WGraba(WCiclo, 6)
            WOrden = WGraba(WCiclo, 7)
            WTipo = WGraba(WCiclo, 8)
            WLetra = WGraba(WCiclo, 9)
            WPunto = WGraba(WCiclo, 10)
            WNumero = WGraba(WCiclo, 11)
                        
            With rstAsiento
                .Index = "Clave"
                .AddNew
                !Asiento = Val(WAsiento)
                Auxi1 = Str$(WCiclo)
                Call Ceros(Auxi1, 2)
                !Renglon = WCiclo
                !Fecha = WFecha
                !Observaciones = Left$(ZZObservaciones, 50)
                !Cuenta = WCuenta
                !Debito = WDebito
                !Credito = WCredito
                !Leyenda = Left$(WLeyenda, 50)
                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                !Clave = WAsiento + Auxi1
                .Update
                            
                XCuenta = WCuenta
                XDebito = WDebito
                XCredito = WCredito
                            
                Call Actualiza_Saldos
                                
            End With
        Next WCiclo
                        
    End If
    
    
    
    Rem Procesa las ventas
    
    If Tipo5.Value = 1 Then
    
        With rstConfiguracion
            .Index = "Clave"
            .Seek "=", 1
            If .NoMatch = False Then
                ConfigIva1 = !Iva1
                ConfigIva2 = !Iva2
            End If
        End With
    
        Pasa = 0
        Erase WGraba
        Lugar = 0
    
        With rstCtaCte
            .Index = "Clave"
            .MoveFirst
            Do
                If WDesde <= !ordfecha And !ordfecha <= WHasta Then
                
                    If !Tipo < 6 Then
                    
                        If Pasa = 0 Then
                            Pasa = 1
                            Erase WGraba
                            Lugar = 0
                            Corte = Left$(!Clave, 15)
                        End If
                    
                        If Corte <> Left$(!Clave, 15) Then
                    
                            With rstAsiento
                                .Index = "Asiento"
                                Claveven$ = "99999999"
                                .Seek "<=", Claveven$
                                If .NoMatch = False Then
                                    WAsiento = !Asiento + 1
                                        Else
                                    WAsiento = "1"
                                End If
                            End With
                            Call Ceros(WAsiento, 6)
                    
                            For WCiclo = 1 To Lugar
                        
                                WFecha = WGraba(WCiclo, 1)
                                ZZObservaciones = WGraba(WCiclo, 2)
                                WCuenta = WGraba(WCiclo, 3)
                                WDebito = Val(WGraba(WCiclo, 4))
                                WCredito = Val(WGraba(WCiclo, 5))
                                WLeyenda = WGraba(WCiclo, 6)
                                WOrden = WGraba(WCiclo, 7)
                                WTipo = WGraba(WCiclo, 8)
                                WLetra = WGraba(WCiclo, 9)
                                WPunto = WGraba(WCiclo, 10)
                                WNumero = WGraba(WCiclo, 11)
                            
                                With rstAsiento
                                    .Index = "Clave"
                                    .AddNew
                                    !Asiento = Val(WAsiento)
                                    Auxi1 = Str$(WCiclo)
                                    Call Ceros(Auxi1, 2)
                                    !Renglon = WCiclo
                                    !Fecha = WFecha
                                    !Observaciones = Left$(ZZObservaciones, 50)
                                    !Cuenta = WCuenta
                                    !Debito = WDebito
                                    !Credito = WCredito
                                    !Leyenda = Left$(WLeyenda, 50)
                                    !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    !Clave = WAsiento + Auxi1
                                    .Update
                            
                                    XCuenta = WCuenta
                                    XDebito = WDebito
                                    XCredito = WCredito
                            
                                    Call Actualiza_Saldos
                                
                                End With
                            Next WCiclo
                            
                            Corte = Left$(!Clave, 15)
                            Erase WGraba
                            Lugar = 0
                    
                        End If
                        
                        Select Case !Tipo
                            Case 3
                                WImpre = "FAC"
                            Case 4
                                WImpre = "N/D"
                            Case 5
                                WImpre = "N/C"
                            Case Else
                                WImpre = ""
                        End Select
                
                        Wcliente = !Cliente
                        WLetra = !Letra
                        WTipo = !Tipo
                        WPunto = !Punto
                        WNumero = !Numero
                        WFecha = !Fecha
                        WNeto = !Neto
                        WIva = !Iva1 + !Iva2
                        WTotal = !Total
                        WClave = !Clave
                        WProyecto = !Proyecto
                        
                        Rem With rstProyecto
                        Rem     .Index = "Codigo"
                        Rem     .Seek "=", WProyecto
                        Rem     If .NoMatch = False Then
                        Rem         WCuenta = !Cuenta
                        Rem             Else
                        Rem         WCuenta = WCtaVentas
                        Rem     End If
                        Rem End With
                        
                        With rstClientes
                            .Index = "Cliente"
                            .Seek "=", Wcliente
                            If .NoMatch = False Then
                                WObservaciones = "Fc:" + WNumero + " " + Trim(!Razon)
                            End If
                        End With
                        
                        ZRenglon = 0
                        For ZRenglon = 1 To 30
    
                            With rstDesccomp
    
                                Auxi2 = WPunto
                                Call Ceros(Auxi2, 4)
        
                                Auxi = WNumero
                                Call Ceros(Auxi, 8)
        
                                Auxi1 = ZRenglon
                                Call Ceros(Auxi1, 2)
        
                                .Index = "Clave"
                                .Seek "=", WLetra + WTipo + Auxi2 + Auxi + Auxi1
                                If .NoMatch = False Then
        
                                    WImporte = !Importe
                                    ZCuenta = IIf(IsNull(!Cuenta), "", !Cuenta)
                                    If Val(ZCuenta) = 0 Then
                                        ZCuenta = WCtaVentas
                                    End If
                                    
                                    If WImporte <> 0 Then
                                    
                                        If WLetra = "B" Then
                                            WImpoIva = WImporte / (1 + (ConfigIva1 / 100))
                                            Call Redondeo(WImpoIva)
                                            WImporte = WImpoIva
                                        End If
                                    
                                        If Val(WTipo) = 5 Then
                                            WDebito = Abs(WImporte)
                                            WCredito = 0
                                                Else
                                            WDebito = 0
                                            WCredito = Abs(WImporte)
                                        End If
                        
                                        Lugar = Lugar + 1
                                        WGraba(Lugar, 1) = WFecha
                                        WGraba(Lugar, 2) = "Asiento de Ventas"
                                        WGraba(Lugar, 3) = ZCuenta
                                        WGraba(Lugar, 4) = Str$(WDebito)
                                        WGraba(Lugar, 5) = Str$(WCredito)
                                        WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                                        WGraba(Lugar, 7) = ""
                                        WGraba(Lugar, 8) = WTipo
                                        WGraba(Lugar, 9) = WLetra
                                        WGraba(Lugar, 10) = WPunto
                                        WGraba(Lugar, 11) = WNumero
                                        
                                    End If
                
                                End If
        
                            End With
    
                        Next ZRenglon
                        
                        For ZRenglon = 1 To 50
    
                            With rstEstadistica
    
                                Auxi = WNumero
                                Call Ceros(Auxi, 8)
            
                                Auxi1 = ZRenglon
                                Call Ceros(Auxi1, 2)
        
                                .Index = "Clave"
                                .Seek "=", "01" + Auxi + Auxi1
                                If .NoMatch = False Then
        
                                    ZCantidad = !Cantidad
                                    ZPrecio = !Precio
                                    ZImporte = !Cantidad * !Precio
                                    Call Redondeo(ZImporte)
                                    ZCuenta = !Cuenta
                                    If Val(ZCuenta) = 0 Then
                                        ZCuenta = WCtaVentas
                                    End If
                                    
                                    If Val(WTipo) = 2 Then
                                        WDebito = Abs(ZImporte)
                                        WCredito = 0
                                            Else
                                        WDebito = 0
                                        WCredito = Abs(ZImporte)
                                    End If
                                    
                                    If ZImporte <> 0 Then
                                    
                                        Lugar = Lugar + 1
                                        WGraba(Lugar, 1) = WFecha
                                        WGraba(Lugar, 2) = "Asiento de Ventas"
                                        WGraba(Lugar, 3) = ZCuenta
                                        WGraba(Lugar, 4) = Str$(WDebito)
                                        WGraba(Lugar, 5) = Str$(WCredito)
                                        WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                                        WGraba(Lugar, 7) = ""
                                        WGraba(Lugar, 8) = WTipo
                                        WGraba(Lugar, 9) = WLetra
                                        WGraba(Lugar, 10) = WPunto
                                        WGraba(Lugar, 11) = WNumero
                                        
                                    End If
                
                                End If
        
                            End With
    
                        Next ZRenglon
                    
                        Rem If WNeto <> 0 Then
                        Rem
                        Rem     If WNeto < 0 Then
                        Rem         WDebito = Abs(WNeto)
                        Rem         WCredito = 0
                        Rem             Else
                        Rem         WDebito = 0
                        Rem         WCredito = Abs(WNeto)
                        Rem     End If
                        Rem
                        Rem     Lugar = Lugar + 1
                        Rem     WGraba(Lugar, 1) = WFecha
                        Rem     WGraba(Lugar, 2) = "Asiento de Ventas"
                        Rem     WGraba(Lugar, 3) = WCuenta
                        Rem     WGraba(Lugar, 4) = Str$(WDebito)
                        Rem     WGraba(Lugar, 5) = Str$(WCredito)
                        Rem     WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                        Rem     WGraba(Lugar, 7) = ""
                        Rem     WGraba(Lugar, 8) = WTipo
                        Rem     WGraba(Lugar, 9) = WLetra
                        Rem     WGraba(Lugar, 10) = WPunto
                        Rem     WGraba(Lugar, 11) = WNumero
                        Rem
                        Rem End If
                    
                        If WIva <> 0 Then
                        
                            If WIva < 0 Then
                                WDebito = Abs(WIva)
                                WCredito = 0
                                    Else
                                WDebito = 0
                                WCredito = Abs(WIva)
                            End If
                        
                            Lugar = Lugar + 1
                            WGraba(Lugar, 1) = WFecha
                            WGraba(Lugar, 2) = "Asiento de Ventas"
                            WGraba(Lugar, 3) = WCtaIvaVen
                            WGraba(Lugar, 4) = Str$(WDebito)
                            WGraba(Lugar, 5) = Str$(WCredito)
                            WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                            WGraba(Lugar, 7) = ""
                            WGraba(Lugar, 8) = WTipo
                            WGraba(Lugar, 9) = WLetra
                            WGraba(Lugar, 10) = WPunto
                            WGraba(Lugar, 11) = WNumero
                        
                        End If
                        
                        If WTotal <> 0 Then
                        
                            WCuenta = WCtaDeudores
                            With rstClientes
                                .Index = "Cliente"
                                .Seek "=", Wcliente
                                If .NoMatch = False Then
                                    If Val(!Cuenta) <> 0 Then
                                        WCuenta = !Cuenta
                                    End If
                                End If
                            End With
                        
                            If WTotal < 0 Then
                                WDebito = 0
                                WCredito = Abs(WTotal)
                                    Else
                                WDebito = Abs(WTotal)
                                WCredito = 0
                            End If
                        
                            Lugar = Lugar + 1
                            WGraba(Lugar, 1) = WFecha
                            WGraba(Lugar, 2) = "Asiento de Ventas"
                            WGraba(Lugar, 3) = WCuenta
                            WGraba(Lugar, 4) = Str$(WDebito)
                            WGraba(Lugar, 5) = Str$(WCredito)
                            WGraba(Lugar, 6) = Left$(WObservaciones, 50)
                            WGraba(Lugar, 7) = ""
                            WGraba(Lugar, 8) = WTipo
                            WGraba(Lugar, 9) = WLetra
                            WGraba(Lugar, 10) = WPunto
                            WGraba(Lugar, 11) = WNumero
                        
                        End If
                    
                    End If
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    
    End If
    
    If Pasa <> 0 Then
                
        With rstAsiento
            .Index = "Asiento"
            Claveven$ = "99999999"
            .Seek "<=", Claveven$
            If .NoMatch = False Then
                WAsiento = !Asiento + 1
                    Else
                WAsiento = "1"
            End If
        End With
        Call Ceros(WAsiento, 6)
                    
        For WCiclo = 1 To Lugar
                        
            WFecha = WGraba(WCiclo, 1)
            ZZObservaciones = WGraba(WCiclo, 2)
            WCuenta = WGraba(WCiclo, 3)
            WDebito = Val(WGraba(WCiclo, 4))
            WCredito = Val(WGraba(WCiclo, 5))
            WLeyenda = WGraba(WCiclo, 6)
            WOrden = WGraba(WCiclo, 7)
            WTipo = WGraba(WCiclo, 8)
            WLetra = WGraba(WCiclo, 9)
            WPunto = WGraba(WCiclo, 10)
            WNumero = WGraba(WCiclo, 11)
                        
            With rstAsiento
                .Index = "Clave"
                .AddNew
                !Asiento = Val(WAsiento)
                Auxi1 = Str$(WCiclo)
                Call Ceros(Auxi1, 2)
                !Renglon = WCiclo
                !Fecha = WFecha
                !Observaciones = Left$(ZZObservaciones, 50)
                !Cuenta = WCuenta
                !Debito = WDebito
                !Credito = WCredito
                !Leyenda = Left$(WLeyenda, 50)
                !fechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                !Clave = WAsiento + Auxi1
                .Update
                            
                XCuenta = WCuenta
                XDebito = WDebito
                XCredito = WCredito
                            
                Call Actualiza_Saldos
                                
            End With
        Next WCiclo
                        
    End If
    
    Exit Sub
    
Error_Programa:
     Rem coderr = Err
     Rem Call Errores(coderr, "Error en el sistema", "Se produjo el error " + Str$(coderr))
     Resume Next
    
End Sub

Private Sub Cancela_Click()
    With rstPagos
        .Close
    End With
    With rstDepositos
        .Close
    End With
    With rstRecibos
        .Close
    End With
    With rstImpcyb
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    With RstProveedor
        .Close
    End With
    With rstBanco
        .Close
    End With
    With rstTipopro
        .Close
    End With
    With rstProyecto
        .Close
    End With
    With rstClientes
        .Close
    End With
    With rstDesccomp
        .Close
    End With
    With rstConfiguracion
        .Close
    End With
    
    DbsAdminis.Close
    DesdeFecha.SetFocus
    PrgGrabaImpcyb.Hide
    Unload Me
    Menu.SetFocus
End Sub

Sub Form_Load()
    Tipo1.Value = False
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    
    Fecha.Text = "  /  /    "
    Frame2.Visible = True
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_EmpreCon
    OPEN_FILE_Pagos
    OPEN_FILE_Depositos
    OPEN_FILE_Recibos
    OPEN_FILE_Impcyb
    OPEN_FILE_Ctacte
    OPEN_FILE_Proveedor
    OPEN_FILE_Banco
    OPEN_FILE_TipoPro
    OPEN_FILE_Proyecto
    OPEN_FILE_Clientes
    OPEN_FILE_Ivacomp
    OPEN_FILE_DescComp
    OPEN_FILE_Estadistica
    OPEN_FILE_Configuracion
End Sub

Private Sub Actualiza_Saldos()

    With rstCuentaCon
        .Index = "Cuenta"
        .Seek "=", XCuenta
        If .NoMatch = False Then
            .Edit
            Select Case WPosi
                Case 1
                    !Debito1 = !Debito1 + XDebito
                    !Credito1 = !Credito1 + XCredito
                Case 2
                    !Debito2 = !Debito2 + XDebito
                    !Credito2 = !Credito2 + XCredito
                Case 3
                    !Debito3 = !Debito3 + XDebito
                    !Credito3 = !Credito3 + XCredito
                Case 4
                    !Debito4 = !Debito4 + XDebito
                    !Credito4 = !Credito4 + XCredito
                Case 5
                    !Debito5 = !Debito5 + XDebito
                    !Credito5 = !Credito5 + XCredito
                Case 6
                    !Debito6 = !Debito6 + XDebito
                    !Credito6 = !Credito6 + XCredito
                Case 7
                    !Debito7 = !Debito7 + XDebito
                    !Credito7 = !Credito7 + XCredito
                Case 8
                    !Debito8 = !Debito8 + XDebito
                    !Credito8 = !Credito8 + XCredito
                Case 9
                    !Debito9 = !Debito9 + XDebito
                    !Credito9 = !Credito9 + XCredito
                Case 10
                    !Debito10 = !Debito10 + XDebito
                    !Credito10 = !Credito10 + XCredito
                Case 11
                    !Debito11 = !Debito11 + XDebito
                    !Credito11 = !Credito11 + XCredito
                Case 12
                    !Debito12 = !Debito12 + XDebito
                    !Credito12 = !Credito12 + XCredito
                Case Else
            End Select
            .Update
        End If
    End With

End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
                Else
            DesdeFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFecha.Text = "  /  /    "
    End If
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            EmpresaCon.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
    End If
End Sub

Private Sub Empresacon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFecha.SetFocus
    End If
    If KeyAscii = 27 Then
        EmpresaCon.Text = ""
    End If
End Sub

Private Sub DesdeFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub EmpresaCon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Proceso_Click
        Case 121
            Call Cancela_Click
        Case Else
    End Select
End Sub







