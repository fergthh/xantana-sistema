VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsultaPago 
   Caption         =   "Consulta de Pagos"
   ClientHeight    =   7635
   ClientLeft      =   1080
   ClientTop       =   465
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   8160
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
      TabIndex        =   22
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
      TabIndex        =   21
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
      TabIndex        =   20
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
      TabIndex        =   19
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
      TabIndex        =   18
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
      TabIndex        =   17
      Top             =   4560
      Width           =   375
   End
   Begin VB.Frame BusquedaCheque 
      Height          =   2655
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   6135
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
         Left            =   4680
         MouseIcon       =   "consultapago.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "consultapago.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Menu Principal"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Pago 
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
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Efectivo 
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
         TabIndex        =   6
         Text            =   " "
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Compensacion 
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
         TabIndex        =   5
         Text            =   " "
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox Maximo 
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
         TabIndex        =   4
         Text            =   " "
         Top             =   1440
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
      Begin MSMask.MaskEdBox PlazoMaximo 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
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
      Begin MSMask.MaskEdBox PlazoInicial 
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label1 
         Caption         =   "Importe a Pagar"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Efectivo"
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
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Compensacion"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Maximo Cheque"
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
         TabIndex        =   11
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Plazo Maximo"
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
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "Plazo Inicial"
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
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4575
      Left            =   240
      TabIndex        =   16
      Top             =   2880
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8070
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgConsultaPago"
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

Private Numero As String
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
Dim ZCheque(100, 5) As String

Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub Acepta_Click()

    Call Limpia_Vector

    OrdInicial = Right$(PlazoInicial.Text, 4) + Mid$(PlazoInicial.Text, 4, 2) + Left$(PlazoInicial.Text, 2)
    OrdMaximo = Right$(PlazoMaximo.Text, 4) + Mid$(PlazoMaximo.Text, 4, 2) + Left$(PlazoMaximo.Text, 2)
    
    ZSuma = Val(Compensacion.Text)
    ZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.TipoReg = '2'"
    ZSql = ZSql + " and Recibos.Tipo2 = '02'"
    ZSql = ZSql + " and Recibos.Estado2 <> 'X'"
    ZSql = ZSql + " Order by Recibos.Importe2 desc"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstRecibos!FechaOrd2 >= OrdInicial And rstRecibos!FechaOrd2 <= OrdMaximo Then
                        If rstRecibos!Importe2 <> 0 Then
                            If Val(Maximo.Text) = 0 Or rstRecibos!Importe2 <= Val(Maximo.Text) Then
                            
                                ZPrueba = ZSuma + rstRecibos!Importe2
                                If ZPrueba < Val(Pago.Text) Then
                                    ZSuma = ZPrueba
                                    ZLugar = ZLugar + 1
                                    ZCheque(ZLugar, 1) = rstRecibos!Numero2
                                    ZCheque(ZLugar, 2) = rstRecibos!Fecha2
                                    ZCheque(ZLugar, 3) = rstRecibos!Banco2
                                    ZCheque(ZLugar, 4) = Str$(rstRecibos!Importe2)
                                    ZCheque(ZLugar, 5) = rstRecibos!Clave
                                End If
                            
                            End If
                        End If
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    If ZSuma + Val(Efectivo.Text) >= Val(Pago.Text) Then
    
        For Ciclo = 1 To ZLugar
            
            WVector1.Row = Ciclo
            ZZLugar = Ciclo
                
            WVector1.Col = 2
            WVector1.Text = ZCheque(Ciclo, 1)
            
            WVector1.Col = 3
            WVector1.Text = ZCheque(Ciclo, 2)
            
            WVector1.Col = 4
            WVector1.Text = ""
                
            WVector1.Col = 5
            WVector1.Text = ZCheque(Ciclo, 3)
            
            WVector1.Col = 6
            WVector1.Text = ZCheque(Ciclo, 4)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        
            WVector1.Col = 1
            WVector1.Text = "03"
                
            BajaCheque(WVector1.Row) = ZCheque(Ciclo, 5)
        
        Next Ciclo
        
        ZDife = Val(Pago.Text) - ZSuma
        
        If Val(Efectivo.Text) <> 0 Then
        
            If Val(Efectivo.Text) > ZDife Then
            
                ZZLugar = ZZLugar + 1
            
                WVector1.Row = ZZLugar
                    
                WVector1.Col = 2
                WVector1.Text = ""
                
                WVector1.Col = 3
                WVector1.Text = ""
                
                WVector1.Col = 4
                WVector1.Text = ""
                    
                WVector1.Col = 5
                WVector1.Text = ""
                
                WVector1.Col = 6
                WVector1.Text = Str$(ZDife)
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            
                WVector1.Col = 1
                WVector1.Text = "01"
                    
                BajaCheque(ZZCiclo) = ""
                
                ZDife = 0
                    
                    Else
            
                ZZLugar = ZZLugar + 1
            
                WVector1.Row = ZZLugar
                    
                WVector1.Col = 2
                WVector1.Text = ""
                
                WVector1.Col = 3
                WVector1.Text = ""
                
                WVector1.Col = 4
                WVector1.Text = ""
                    
                WVector1.Col = 5
                WVector1.Text = ""
                
                WVector1.Col = 6
                WVector1.Text = Efectivo.Text
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                            
                WVector1.Col = 1
                WVector1.Text = "01"
                    
                BajaCheque(ZZCiclo) = ""
                
                ZDife = ZDife - Val(Efectivo.Text)
                
            End If
            
        End If
        
        If Val(Compensacion.Text) <> 0 Then
        
            ZZLugar = ZZLugar + 1
        
            WVector1.Row = ZZLugar
                
            WVector1.Col = 2
            WVector1.Text = ""
            
            WVector1.Col = 3
            WVector1.Text = ""
            
            WVector1.Col = 4
            WVector1.Text = ""
                
            WVector1.Col = 5
            WVector1.Text = ""
            
            WVector1.Col = 6
            WVector1.Text = Compensacion.Text
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                        
            WVector1.Col = 1
            WVector1.Text = "04"
                
            BajaCheque(ZZCiclo) = ""
                
            ZDife = 0
            
        End If

            Else
            
        m$ = "No hay valores que cumplan con las condiciones requeridas"
        A% = MsgBox(m$, 0, "Busqueda de Cheques en cartera")
        
    End If
    
    Rem BusquedaCheque.Visible = False

End Sub

Private Sub CmdClose_Click()
    PrgConsultaPago.Hide
    Unload Me
    MenuA2.Show
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    WAno = Mid$(WFecha, 7, 4)
    Ano = Val(WAno)
    WMes = Mid$(WFecha, 4, 2)
    Mes = Val(WMes)
    WDia = Mid$(WFecha, 1, 2)
    Dia = Val(WDia)
    
    WAno = Str$(Val(WAno) - 1)
    Call Ceros(WAno, 4)

    ZFecha = Mid$(WFecha, 1, 6) + WAno
    
    Dife = DateDiff("d", ZFecha, WFecha)

    WPlazo1 = Dife - 29
    Call Calcula_vencimiento(ZFecha, WPlazo1, WVencimiento)
    
    Pago.Text = ""
    Efectivo.Text = ""
    Compensacion.Text = ""
    Maximo.Text = ""
    PlazoMaximo.Text = "  /  /    "
    PlazoInicial.Text = WVencimiento

    BusquedaCheque.Visible = True
    
End Sub

Private Sub Cancela_Click()

    Call Limpia_Vector

    WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    WAno = Mid$(WFecha, 7, 4)
    Ano = Val(WAno)
    WMes = Mid$(WFecha, 4, 2)
    Mes = Val(WMes)
    WDia = Mid$(WFecha, 1, 2)
    Dia = Val(WDia)
    
    WAno = Str$(Val(WAno) - 1)
    Call Ceros(WAno, 4)

    ZFecha = Mid$(WFecha, 1, 6) + WAno
    
    Dife = DateDiff("d", ZFecha, WFecha)

    WPlazo1 = Dife - 29
    Call Calcula_vencimiento(ZFecha, WPlazo1, WVencimiento)
    Pago.Text = ""
    Efectivo.Text = ""
    Compensacion.Text = ""
    Maximo.Text = ""
    PlazoMaximo.Text = "  /  /    "
    PlazoInicial.Text = WVencimiento
End Sub

Private Sub Pago_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Efectivo.SetFocus
    End If
    If KeyAscii = 27 Then
        Pago.Text = ""
    End If
End Sub

Private Sub Efectivo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Compensacion.SetFocus
    End If
    If KeyAscii = 27 Then
        Efectivo.Text = ""
    End If
End Sub

Private Sub Compensacion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Maximo.SetFocus
    End If
    If KeyAscii = 27 Then
        Compensacion.Text = ""
    End If
End Sub


Private Sub Maximo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PlazoMaximo.SetFocus
    End If
    If KeyAscii = 27 Then
        Maximo.Text = ""
    End If
End Sub

Private Sub PlazoMaximo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(PlazoMaximo.Text)) = 8 Then
            PlazoMaximo.Text = Left$(PlazoMaximo.Text, 6) + "20" + Right$(Trim(PlazoMaximo.Text), 2)
        End If
        Call Valida_fecha1(PlazoMaximo.Text, Auxi)
        If Auxi = "S" Then
            PlazoInicial.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        PlazoMaximo.Text = "  /  /    "
    End If
End Sub

Private Sub PlazoInicial_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(PlazoInicial.Text)) = 8 Then
            PlazoInicial.Text = Left$(PlazoInicial.Text, 6) + "20" + Right$(Trim(PlazoInicial.Text), 2)
        End If
        Call Valida_fecha1(PlazoInicial.Text, Auxi)
        If Auxi = "S" Then
            Rem Efectivo.SetFocus
            Call Acepta_Click
        End If
    End If
    If KeyAscii = 27 Then
        PlazoInicial.Text = "  /  /    "
    End If
End Sub



Private Sub Efectivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Pago_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Compensacion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Maximo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PlazoMaximo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub PlazoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
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
    WVector1.Cols = 7
    WVector1.FixedRows = 1
    WVector1.Rows = 51
    
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
                WVector1.Text = "Numero"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 8
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1200
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 2
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Banco"
                WVector1.ColWidth(Ciclo) = 700
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 1150
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 30
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
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


