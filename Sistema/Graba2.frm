VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGraba2 
   Caption         =   "Grabacion de Imputaciones Contables"
   ClientHeight    =   5415
   ClientLeft      =   2895
   ClientTop       =   1020
   ClientWidth     =   6615
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5415
   ScaleWidth      =   6615
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   1320
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
         Begin VB.OptionButton Tipo4 
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
            Height          =   255
            Left            =   1680
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Tipo3 
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
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Tipo2 
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
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Tipo1 
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
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Left            =   4200
         TabIndex        =   6
         Top             =   720
         Width           =   975
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
         Height          =   255
         Left            =   4200
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label Label3 
         Caption         =   "Fecha Grabacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1575
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
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
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgGraba2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XTipopro(100) As String
Private XBanco(1000) As String
Private XProveedor(10000) As Integer
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstTipoPro As Recordset
Dim spTipoPro As String
Dim RstPagos As Recordset
Dim spPagos As String
Dim rstDepositos As Recordset
Dim spDepositos As String
Dim rstBanco As Recordset
Dim spBanco As String
Dim rstRecibos As Recordset
Dim spRecibos As String
Dim RstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim WConfac As String
Dim WUnidad As String
Dim Vector(10000, 15) As String
Dim Impre(100, 15) As String
Dim WMes As String
Dim WAno As String
Dim Mes1 As String
Dim Ano1 As String
Dim Compara As String
Dim WAsiento As String
Dim Auxi1 As String
Dim WCuenta As String
Dim WLeyenda As String
Dim WDebito As Double
Dim WCredito As Double
Dim XCuenta As String
Dim XDebito As Double
Dim XCredito As Double

Private Sub Acepta_Click()

    On Error GoTo Error_Programa
    
    With rstEmpre
        .Index = "Codigo"
        .Seek "=", 1
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
    WFecha = WAno + WMes + WDia
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHAsta = WAno + WMes + WDia
    
    Erase Vector
    Lugar = 0
    
    spBanco = "ListaBancos"
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Do
                If .EOF = False Then
                    XLugar = XLugar + 1
                    XBanco(rstBanco!Banco) = rstBanco!Cuenta
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBanco.Close
    End If
    
    XBanco(0) = "111001"
    
    spProveedor = "ListaProveedores"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    XLugar = XLugar + 1
                    XProveedor(RstProveedor!Proveedor) = RstProveedor!TipoPro
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
    
    spTipoPro = "ListaTipopro"
    Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoPro.RecordCount > 0 Then
        With rstTipoPro
            .MoveFirst
            Do
                If .EOF = False Then
                    XLugar = XLugar + 1
                    XTipopro(rstTipoPro!Codigo) = rstTipoPro!Cuenta
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTipoPro.Close
    End If
    
    If Tipo1.Value = True Then
    
    WProceso = "P"

    spPagos = "ListaPagos"
    Set RstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If RstPagos.RecordCount > 0 Then

    With RstPagos
            .MoveFirst
            Do
                If WDesde <= !FechaOrd And !FechaOrd <= WHAsta Then
                
                    WMarca = IIf(IsNull(RstPagos!Marca), "", RstPagos!Marca)
                    If WMarca <> "X" Then
                
                        Select Case Val(!Tiporeg)
                            Case 1
                                Select Case !TipoOrd
                                    Case 3
                                        WCuenta = !Cuenta
                                    Case 4
                                        WCuenta = XBanco(!Banco2)
                                    Case 5
                                        WCuenta = "111"
                                    Case Else
                                        Rem proveedor
                                        WCuenta = XTipopro(XProveedor(!Proveedor))
                                End Select
                            
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Orden
                                Vector(Lugar, 2) = WCuenta
                                Vector(Lugar, 3) = Str$(!Importe1)
                                Vector(Lugar, 4) = 0
                                Vector(Lugar, 5) = !Proveedor
                                Vector(Lugar, 6) = !Letra1
                                Vector(Lugar, 7) = !Tipo1
                                Vector(Lugar, 8) = !Punto1
                                Vector(Lugar, 9) = !Numero1
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = !Observaciones
    
                                If Val(!Renglon) = 1 And !Retencion <> 0 Then
                            
                                    Lugar = Lugar + 1
                    
                                    Vector(Lugar, 1) = !Orden
                                    Vector(Lugar, 2) = "2134"
                                    Vector(Lugar, 3) = Str$(!Retencion)
                                    Vector(Lugar, 4) = 0
                                    Vector(Lugar, 5) = !Proveedor
                                    Vector(Lugar, 6) = ""
                                    Vector(Lugar, 7) = ""
                                    Vector(Lugar, 8) = ""
                                    Vector(Lugar, 9) = ""
                                    Vector(Lugar, 10) = !Fecha
                                    Vector(Lugar, 11) = !Observaciones
                                    
                                End If
                                
                                If Val(!Renglon) = 1 And !RetOtra <> 0 Then
                            
                                    Lugar = Lugar + 1
                    
                                    Vector(Lugar, 1) = !Orden
                                    Vector(Lugar, 2) = "21399"
                                    Vector(Lugar, 3) = Str$(!RetOtra)
                                    Vector(Lugar, 4) = 0
                                    Vector(Lugar, 5) = !Proveedor
                                    Vector(Lugar, 6) = ""
                                    Vector(Lugar, 7) = ""
                                    Vector(Lugar, 8) = ""
                                    Vector(Lugar, 9) = ""
                                    Vector(Lugar, 10) = !Fecha
                                    Vector(Lugar, 11) = !Observaciones
                                    
                                End If
                                
                                                            
                            Case Else
                                Select Case Val(!Tipo2)
                                    Case 1
                                        Rem caja
                                        WCuenta = "111001"
                                    Case 2
                                        Rem banco
                                        WCuenta = "999999"
                                        WBanco2 = !Banco2
                                        WCuenta = XBanco(WBanco2)
                                    Case 3
                                        Rem cheque terceros
                                        WCuenta = "112201"
                                    Case Else
                                        Rem documentos
                                        WCuenta = "111001"
                                End Select
                                        
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Orden
                                Vector(Lugar, 2) = WCuenta
                                Vector(Lugar, 3) = 0
                                Vector(Lugar, 4) = Str$(!Importe2)
                                Vector(Lugar, 5) = !Proveedor
                                Vector(Lugar, 6) = ""
                                Vector(Lugar, 7) = !Tipo2
                                Vector(Lugar, 8) = ""
                                Vector(Lugar, 9) = !numero2
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = !Observaciones
                            
                                If Val(!Renglon) = 1 And !Retencion <> 0 Then
                            
                                    Lugar = Lugar + 1
                  
                                    Vector(Lugar, 1) = !Orden
                                    Vector(Lugar, 2) = "2134"
                                    Vector(Lugar, 3) = Str$(!Retencion)
                                    Vector(Lugar, 4) = 0
                                    Vector(Lugar, 5) = !Proveedor
                                    Vector(Lugar, 6) = ""
                                    Vector(Lugar, 7) = ""
                                    Vector(Lugar, 8) = ""
                                    Vector(Lugar, 9) = ""
                                    Vector(Lugar, 10) = !Fecha
                                    Vector(Lugar, 11) = !Observaciones
                                    
                                End If
                                
                                If Val(!Renglon) = 1 And !RetOtra <> 0 Then
                            
                                    Lugar = Lugar + 1
                  
                                    Vector(Lugar, 1) = !Orden
                                    Vector(Lugar, 2) = "213399"
                                    Vector(Lugar, 3) = Str$(!RetOtra)
                                    Vector(Lugar, 4) = 0
                                    Vector(Lugar, 5) = !Proveedor
                                    Vector(Lugar, 6) = ""
                                    Vector(Lugar, 7) = ""
                                    Vector(Lugar, 8) = ""
                                    Vector(Lugar, 9) = ""
                                    Vector(Lugar, 10) = !Fecha
                                    Vector(Lugar, 11) = !Observaciones
                                    
                                End If
                                
                            
                            End Select
                        End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    RstPagos.Close
    
    End If
    
    End If
    
    If Tipo2.Value = True Then
    
    WProceso = "D"
    
    spDepositos = "ListaDepositos"
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
    
    With rstDepositos
            .MoveFirst
            Do
                If WDesde <= !FechaOrd And !FechaOrd <= WHAsta Then
                
                    WMarca = IIf(IsNull(rstDepositos!Marca), "", rstDepositos!Marca)
                    If WMarca <> "X" Then
                
                        Lugar = Lugar + 1
              
                        Vector(Lugar, 1) = !Deposito
                        Vector(Lugar, 2) = XBanco(!Banco)
                        Vector(Lugar, 3) = Str$(!Importe2)
                        Vector(Lugar, 4) = 0
                        Vector(Lugar, 5) = !Banco
                        Vector(Lugar, 6) = ""
                        Vector(Lugar, 7) = !Tipo2
                        Vector(Lugar, 8) = ""
                        Vector(Lugar, 9) = !numero2
                        Vector(Lugar, 10) = !Fecha
                        Vector(Lugar, 11) = ""
                        
                        If !Tipo2 = 1 Then
                            Rem efectivo
                            WCuenta = "111001"
                                Else
                            Rem cheque terceros
                            WCuenta = "112201"
                        End If
                        
                        Lugar = Lugar + 1
              
                        Vector(Lugar, 1) = !Deposito
                        Vector(Lugar, 2) = WCuenta
                        Vector(Lugar, 3) = 0
                        Vector(Lugar, 4) = Str$(!Importe2)
                        Vector(Lugar, 5) = !Banco
                        Vector(Lugar, 6) = ""
                        Vector(Lugar, 7) = !Tipo2
                        Vector(Lugar, 8) = ""
                        Vector(Lugar, 9) = !numero2
                        Vector(Lugar, 10) = !Fecha
                        Vector(Lugar, 11) = ""
                    
                    End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    rstDepositos.Close
    
    End If
    
    End If
    
    If Tipo3.Value = True Then
    
    WProceso = "R"

    spRecibos = "ListaRecibos"
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then

    With rstRecibos
            .MoveFirst
            Do
                If WDesde <= !FechaOrd And !FechaOrd <= WHAsta Then
                
                    WMarca = IIf(IsNull(rstRecibos!Marca), "", rstRecibos!Marca)
                    If WMarca <> "X" Then
                
                        Select Case Val(!Tiporeg)
                            Case 1
                                If !TipoRec = "3" Then
                                    WCuenta = !Cuenta
                                        Else
                                    WCuenta = "1141"
                                End If
                            
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Recibo
                                Vector(Lugar, 2) = WCuenta
                                If !Importe1 > 0 Then
                                    Vector(Lugar, 3) = "0"
                                    Vector(Lugar, 4) = Str$(!Importe1)
                                        Else
                                    Vector(Lugar, 4) = "0"
                                    Vector(Lugar, 3) = Str$(Abs(!Importe1))
                                End If
                                Vector(Lugar, 5) = !Cliente
                                Vector(Lugar, 6) = !Letra1
                                Vector(Lugar, 7) = !Tipo1
                                Vector(Lugar, 8) = !Punto1
                                Vector(Lugar, 9) = !Numero1
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = ""
                                                            
                            Case Else
                                Select Case Val(!Tipo2)
                                    Case 1
                                        Rem caja
                                        WCuenta = "111001"
                                    Case 2
                                        Rem cheque terceros
                                        WCuenta = "112201"
                                    Case 4
                                        WCuenta = !Cuenta
                                    Case Else
                                        Rem documentos
                                        WCuenta = "111001"
                                End Select
                                        
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Recibo
                                Vector(Lugar, 2) = WCuenta
                                Vector(Lugar, 3) = Str$(!Importe2)
                                Vector(Lugar, 4) = 0
                                Vector(Lugar, 5) = !Cliente
                                Vector(Lugar, 6) = ""
                                Vector(Lugar, 7) = !Tipo2
                                Vector(Lugar, 8) = ""
                                Vector(Lugar, 9) = !numero2
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = ""
                            
                        End Select
                
                        If Val(!Renglon) = 1 And Val(!Retganancias) <> 0 Then
                    
                            Lugar = Lugar + 1
              
                            Vector(Lugar, 1) = !Recibo
                            Vector(Lugar, 2) = "116102"
                            Vector(Lugar, 3) = Str$(!Retganancias)
                            Vector(Lugar, 4) = 0
                            Vector(Lugar, 5) = !Cliente
                            Vector(Lugar, 6) = ""
                            Vector(Lugar, 7) = ""
                            Vector(Lugar, 8) = ""
                            Vector(Lugar, 9) = ""
                            Vector(Lugar, 10) = !Fecha
                            Vector(Lugar, 11) = ""
                            
                        End If
                                
                
                        If Val(!Renglon) = 1 And Val(!RetIva) <> 0 Then
                        
                            Lugar = Lugar + 1
              
                            Vector(Lugar, 1) = !Recibo
                            Vector(Lugar, 2) = "116202"
                            Vector(Lugar, 3) = Str$(!RetIva)
                            Vector(Lugar, 4) = 0
                            Vector(Lugar, 5) = !Cliente
                            Vector(Lugar, 6) = ""
                            Vector(Lugar, 7) = ""
                            Vector(Lugar, 8) = ""
                            Vector(Lugar, 9) = ""
                            Vector(Lugar, 10) = !Fecha
                            Vector(Lugar, 11) = ""
                        
                        End If
                
                        If Val(!Renglon) = 1 And Val(!RetOtra) <> 0 Then
                        
                            Lugar = Lugar + 1
              
                            Vector(Lugar, 1) = !Recibo
                            Vector(Lugar, 2) = "116302"
                            Vector(Lugar, 3) = Str$(!RetOtra)
                            Vector(Lugar, 4) = 0
                            Vector(Lugar, 5) = !Cliente
                            Vector(Lugar, 6) = ""
                            Vector(Lugar, 7) = ""
                            Vector(Lugar, 8) = ""
                            Vector(Lugar, 9) = ""
                            Vector(Lugar, 10) = !Fecha
                            Vector(Lugar, 11) = ""
                        
                        End If
                
                        If Val(!Renglon) = 1 And Val(!Retencion) <> 0 Then
                        
                            Lugar = Lugar + 1
              
                            Vector(Lugar, 1) = !Recibo
                            Vector(Lugar, 2) = "116999"
                            Vector(Lugar, 3) = Str$(!Retencion)
                            Vector(Lugar, 4) = 0
                            Vector(Lugar, 5) = !Cliente
                            Vector(Lugar, 6) = ""
                            Vector(Lugar, 7) = ""
                            Vector(Lugar, 8) = ""
                            Vector(Lugar, 9) = ""
                            Vector(Lugar, 10) = !Fecha
                            Vector(Lugar, 11) = ""
                        
                        End If
                        
                    End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    rstRecibos.Close
    
    End If
    
    End If
    
    
    If Tipo4.Value = True Then
    
    WProceso = "V"
    
    spCtacte = "ListaCtacte"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
    With rstCtacte
            .MoveFirst
            Do
                If WDesde <= !ordfecha And !ordfecha <= WHAsta Then
                
                    If Val(!Tipo) = 3 Or Val(!Tipo) = 4 Or Val(!Tipo) = 5 Or Val(!Tipo) = 13 Or Val(!Tipo) = 14 Or Val(!Tipo) = 15 Then
                    
                        WMarca = IIf(IsNull(rstCtacte!Marca), "", rstCtacte!Marca)
                        If WMarca <> "X" Then
                        
                            XTotal = !Neto + !Iva1 + !Iva2 + !ImpoIb
                            If XTotal <> 0 Then
                            
                            If XTotal <> 0 Then
                    
                                WConfac = !Rubro
                                WUnidad = !Unidad
                                Call Ceros(WConfac, 3)
                                Call Ceros(WUnidad, 3)
                            
                                Lugar = Lugar + 1
                  
                                Vector(Lugar, 1) = !Clave
                                Vector(Lugar, 2) = "1141"
                                If Val(!Tipo) <> 5 And Val(!Tipo) <> 15 Then
                                    Vector(Lugar, 3) = Str$(!Neto + !Iva1 + !Iva2)
                                    Vector(Lugar, 4) = "0"
                                        Else
                                    Vector(Lugar, 3) = "0"
                                    Vector(Lugar, 4) = Str$(Abs(!Neto + !Iva1 + !Iva2))
                                End If
                                Vector(Lugar, 5) = !Cliente
                                Vector(Lugar, 6) = ""
                                Vector(Lugar, 7) = !Tipo
                                Vector(Lugar, 8) = ""
                                Vector(Lugar, 9) = !Numero
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = ""
                            End If
                    
                            If !Neto <> 0 Then
                            
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Clave
                                Vector(Lugar, 2) = "46" + WUnidad + WConfac
                                If Val(!Tipo) <> 5 And Val(!Tipo) <> 15 Then
                                    Vector(Lugar, 3) = "0"
                                    Vector(Lugar, 4) = Str$(!Neto)
                                        Else
                                    Vector(Lugar, 3) = Str$(Abs(!Neto))
                                    Vector(Lugar, 4) = "0"
                                End If
                                Vector(Lugar, 5) = !Cliente
                                Vector(Lugar, 6) = ""
                                Vector(Lugar, 7) = !Tipo
                                Vector(Lugar, 8) = ""
                                Vector(Lugar, 9) = !Numero
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = ""
                                
                            End If
                            
                            If !Iva1 <> 0 Then
                            
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Clave
                                Vector(Lugar, 2) = "2133"
                                If Val(!Tipo) <> 5 And Val(!Tipo) <> 15 Then
                                    Vector(Lugar, 3) = "0"
                                    Vector(Lugar, 4) = Str$(!Iva1)
                                        Else
                                    Vector(Lugar, 3) = Str$(Abs(!Iva1))
                                    Vector(Lugar, 4) = "0"
                                End If
                                Vector(Lugar, 5) = !Cliente
                                Vector(Lugar, 6) = ""
                                Vector(Lugar, 7) = !Tipo
                                Vector(Lugar, 8) = ""
                                Vector(Lugar, 9) = !Numero
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = ""
                                
                            End If
                            
                            If !Iva2 <> 0 Then
                            
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Clave
                                Vector(Lugar, 2) = "2133"
                                If Val(!Tipo) <> 5 And Val(!Tipo) <> 15 Then
                                    Vector(Lugar, 3) = "0"
                                    Vector(Lugar, 4) = Str$(!Iva2)
                                        Else
                                    Vector(Lugar, 3) = Str$(Abs(!Iva2))
                                    Vector(Lugar, 4) = "0"
                                End If
                                Vector(Lugar, 5) = !Cliente
                                Vector(Lugar, 6) = ""
                                Vector(Lugar, 7) = !Tipo
                                Vector(Lugar, 8) = ""
                                Vector(Lugar, 9) = !Numero
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = ""
                               
                            End If
                            
                            If !ImpoIb <> 0 Then
                            
                                Lugar = Lugar + 1
              
                                Vector(Lugar, 1) = !Clave
                                Vector(Lugar, 2) = "21399"
                                If Val(!Tipo) <> 5 And Val(!Tipo) <> 15 Then
                                    Vector(Lugar, 3) = "0"
                                    Vector(Lugar, 4) = Str$(!ImpoIb)
                                        Else
                                    Vector(Lugar, 3) = Str$(Abs(!ImpoIb))
                                    Vector(Lugar, 4) = "0"
                                End If
                                Vector(Lugar, 5) = !Cliente
                                Vector(Lugar, 6) = ""
                                Vector(Lugar, 7) = !Tipo
                                Vector(Lugar, 8) = ""
                                Vector(Lugar, 9) = !Numero
                                Vector(Lugar, 10) = !Fecha
                                Vector(Lugar, 11) = ""
                                
                            End If
                            
                            
                        End If
                        
                        End If
                    
                    End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    rstCtacte.Close
    
    End If
    
    End If
    
    WLugar = Lugar
    
    Erase Impre
    Pasa = 0
    Lugar = 0
    
    For x = 1 To WLugar
    
        If Pasa = 0 Then
        
            Pasa = 1
    
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
            
            Select Case WProceso
                Case "P"
                    Corte = Vector(x, 1)
                    Worden = Vector(x, 1)
                    WProveedor = Vector(x, 5)
                    WLetra = Vector(x, 6)
                    WTipo = Vector(x, 7)
                    WPunto = Vector(x, 8)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
                    Wobservaciones = Vector(x, 11)
                    
                Case "D"
                    Corte = Vector(x, 1)
                    WDeposito = Vector(x, 1)
                    WBanco = Vector(x, 5)
                    WTipo = Vector(x, 7)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
                    
                Case "R"
                    Corte = Vector(x, 1)
                    WRecibo = Vector(x, 1)
                    WCliente = Vector(x, 5)
                    WLetra = Vector(x, 6)
                    WTipo = Vector(x, 7)
                    WPunto = Vector(x, 8)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
                    
                Case "V"
                    Corte = Vector(x, 1)
                    WFactura = Vector(x, 1)
                    WCliente = Vector(x, 5)
                    WTipo = Vector(x, 7)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
                
                Case Else
            End Select

            Erase Impre
            Lugar = 0
            
        End If
        
        If Corte <> Vector(x, 1) Then
        
            Select Case WProceso
                Case "P"
                    DesProveedor = Wobservaciones
                    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        With RstProveedor
                            DesProveedor = IIf(IsNull(!Nombre), "", !Nombre)
                        End With
                        RstProveedor.Close
                    End If
            
                    With rstAsiento
        
                        Renglon = 0
                        .Index = "Clave"
                                        
                        For a = 1 To 100
                        
                            If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                                WCuenta = Impre(a, 1)
                                WDebito = Val(Impre(a, 2))
                                WCredito = Val(Impre(a, 3))
                                WLeyenda = "Orden de Pago Nro. " + Worden + " " + DesProveedor
                                WLeyenda = Left$(WLeyenda, 50)

                                .AddNew
                                !Asiento = Val(WAsiento)
                                Auxi1 = Str$(a)
                                Call Ceros(Auxi1, 2)
                                !Renglon = a
                                !Fecha = Fecha.Text
                                !Observaciones = "Asientos de Ordenes de Pago"
                                !Cuenta = WCuenta
                                !Debito = WDebito
                                !Credito = WCredito
                                !Leyenda = Left$(WLeyenda, 50)
                                !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                                !Clave = WAsiento + Auxi1
                                .Update
                            
                                XCuenta = WCuenta
                                XDebito = WDebito
                                XCredito = WCredito
                            
                                Call Actualiza_Saldos
                                
                            End If
                            
                        Next a
                
                    End With
                    
                    WMarca = "X"
                    XParam = "'" + Worden + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                    spPagos = "ActualizaPagosAsiento " + XParam
                    Set RstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    
                Case "D"
                    DesBanco = ""
                    spBanco = "ConsultaBanco " + "'" + WBanco + "'"
                    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                    If rstBanco.RecordCount > 0 Then
                        With rstBanco
                            DesBanco = IIf(IsNull(!Nombre), "", !Nombre)
                        End With
                        rstBanco.Close
                    End If
            
                    With rstAsiento
        
                        Renglon = 0
                        .Index = "Clave"
                                        
                        For a = 1 To 100
                        
                            If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                                WCuenta = Impre(a, 1)
                                WDebito = Val(Impre(a, 2))
                                WCredito = Val(Impre(a, 3))
                                WLeyenda = "Deposito Nro. " + WDeposito + " " + DesBanco
                                WLeyenda = Left$(WLeyenda, 50)

                                .AddNew
                                !Asiento = Val(WAsiento)
                                Auxi1 = Str$(a)
                                Call Ceros(Auxi1, 2)
                                !Renglon = a
                                !Fecha = Fecha.Text
                                !Observaciones = "Asientos de Depositos"
                                !Cuenta = WCuenta
                                !Debito = WDebito
                                !Credito = WCredito
                                !Leyenda = Left$(WLeyenda, 50)
                                !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                                !Clave = WAsiento + Auxi1
                                .Update
                            
                                XCuenta = WCuenta
                                XDebito = WDebito
                                XCredito = WCredito
                            
                                Call Actualiza_Saldos
                                
                            End If
                            
                        Next a
                
                    End With
                    
                    WMarca = "X"
                    XParam = "'" + WDeposito + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                    spDepositos = "ActualizaDepositosAsiento " + XParam
                    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
                    
                Case "R"
                    DesCliente = ""
                    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                    Set RstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If RstCliente.RecordCount > 0 Then
                        With RstCliente
                            DesCliente = IIf(IsNull(!Razon), "", !Razon)
                        End With
                        RstCliente.Close
                    End If
            
                    With rstAsiento
        
                        Renglon = 0
                        .Index = "Clave"
                                        
                        For a = 1 To 100
                        
                            If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                                WCuenta = Impre(a, 1)
                                WDebito = Val(Impre(a, 2))
                                WCredito = Val(Impre(a, 3))
                                WLeyenda = "Recibo Nro. " + WRecibo + " " + DesCliente
                                WLeyenda = Left$(WLeyenda, 50)

                                .AddNew
                                !Asiento = Val(WAsiento)
                                Auxi1 = Str$(a)
                                Call Ceros(Auxi1, 2)
                                !Renglon = a
                                !Fecha = Fecha.Text
                                !Observaciones = "Asientos de Cobranzas"
                                !Cuenta = WCuenta
                                !Debito = WDebito
                                !Credito = WCredito
                                !Leyenda = Left$(WLeyenda, 50)
                                !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                                !Clave = WAsiento + Auxi1
                                .Update
                            
                                XCuenta = WCuenta
                                XDebito = WDebito
                                XCredito = WCredito
                            
                                Call Actualiza_Saldos
                                
                            End If
                            
                        Next a
                
                    End With
                    
                    WMarca = "X"
                    XParam = "'" + WRecibo + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                    spRecibo = "ActualizaRecibosAsiento " + XParam
                    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
                    
                Case "V"
                    DesCliente = ""
                    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                    Set RstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If RstCliente.RecordCount > 0 Then
                        With RstCliente
                            DesCliente = IIf(IsNull(!Razon), "", !Razon)
                        End With
                        RstCliente.Close
                    End If
            
                    With rstAsiento
        
                        Renglon = 0
                        .Index = "Clave"
                                        
                        For a = 1 To 100
                        
                            If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                                WCuenta = Impre(a, 1)
                                WDebito = Val(Impre(a, 2))
                                WCredito = Val(Impre(a, 3))
                                WLeyenda = "Factura Nro. " + Mid$(WFactura, 7, 8) + " " + DesCliente
                                WLeyenda = Left$(WLeyenda, 50)

                                .AddNew
                                !Asiento = Val(WAsiento)
                                Auxi1 = Str$(a)
                                Call Ceros(Auxi1, 2)
                                !Renglon = a
                                !Fecha = Fecha.Text
                                !Observaciones = "Asientos de Ventas"
                                !Cuenta = WCuenta
                                !Debito = WDebito
                                !Credito = WCredito
                                !Leyenda = Left$(WLeyenda, 50)
                                !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                                !Clave = WAsiento + Auxi1
                                .Update
                            
                                XCuenta = WCuenta
                                XDebito = WDebito
                                XCredito = WCredito
                            
                                Call Actualiza_Saldos
                                
                            End If
                            
                        Next a
                
                    End With
                    
                    WMarca = "X"
                    XParam = "'" + WFactura + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                    spCtacte = "ActualizaCtacteAsiento " + XParam
                    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                    
                Case Else
            
            End Select
                
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
            
            Select Case WProceso
                Case "P"
                    Corte = Vector(x, 1)
                    Worden = Vector(x, 1)
                    WProveedor = Vector(x, 5)
                    WLetra = Vector(x, 6)
                    WTipo = Vector(x, 7)
                    WPunto = Vector(x, 8)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
                    Wobservaciones = Vector(x, 11)
                    
                Case "D"
                    Corte = Vector(x, 1)
                    WDeposito = Vector(x, 1)
                    WBanco = Vector(x, 5)
                    WTipo = Vector(x, 7)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
                    
                Case "R"
                    Corte = Vector(x, 1)
                    WRecibo = Vector(x, 1)
                    WCliente = Vector(x, 5)
                    WLetra = Vector(x, 6)
                    WTipo = Vector(x, 7)
                    WPunto = Vector(x, 8)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
                    
                Case "V"
                    Corte = Vector(x, 1)
                    WFactura = Vector(x, 1)
                    WCliente = Vector(x, 5)
                    WTipo = Vector(x, 7)
                    WNumero = Vector(x, 9)
                    WFecha = Vector(x, 10)
               
                Case Else
                
            End Select
            
            Erase Impre
            Lugar = 0
            
        End If
        
        Lugar = Lugar + 1
        
        Impre(Lugar, 1) = Vector(x, 2)
        Impre(Lugar, 2) = Vector(x, 3)
        Impre(Lugar, 3) = Vector(x, 4)
    
    Next x
    
    If Pasa <> 0 Then
    
        Select Case WProceso
            Case "P"
                DesProveedor = Wobservaciones
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    With RstProveedor
                        DesProveedor = IIf(IsNull(!Nombre), "", !Nombre)
                    End With
                    RstProveedor.Close
                End If
            
                With rstAsiento
        
                    Renglon = 0
                    .Index = "Clave"
                                        
                    For a = 1 To 100
                        
                        If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                            WCuenta = Impre(a, 1)
                            WDebito = Val(Impre(a, 2))
                            WCredito = Val(Impre(a, 3))
                            WLeyenda = "Orden de Pago Nro. " + Worden + " " + DesProveedor
                            WLeyenda = Left$(WLeyenda, 50)

                            .AddNew
                            !Asiento = Val(WAsiento)
                            Auxi1 = Str$(a)
                            Call Ceros(Auxi1, 2)
                            !Renglon = a
                            !Fecha = Fecha.Text
                            !Observaciones = "Asientos de Ordenes de Pago"
                            !Cuenta = WCuenta
                            !Debito = WDebito
                            !Credito = WCredito
                            !Leyenda = Left$(WLeyenda, 50)
                            !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            !Clave = WAsiento + Auxi1
                            .Update
                            
                            XCuenta = WCuenta
                            XDebito = WDebito
                            XCredito = WCredito
                            
                            Call Actualiza_Saldos
                                
                        End If
                            
                    Next a
                
                End With
                    
                WMarca = "X"
                XParam = "'" + Worden + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                spPagos = "ActualizaPagosAsiento " + XParam
                Set RstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                
            Case "D"
                DesBanco = ""
                spBanco = "ConsultaBanco " + "'" + WBanco + "'"
                Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                If rstBanco.RecordCount > 0 Then
                    With rstBanco
                        DesBanco = IIf(IsNull(!Nombre), "", !Nombre)
                    End With
                    rstBanco.Close
                End If
            
                With rstAsiento
        
                    Renglon = 0
                    .Index = "Clave"
                                        
                    For a = 1 To 100
                        
                        If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                            WCuenta = Impre(a, 1)
                            WDebito = Val(Impre(a, 2))
                            WCredito = Val(Impre(a, 3))
                            WLeyenda = "Deposito Nro. " + WDeposito + " " + DesBanco
                            WLeyenda = Left$(WLeyenda, 50)

                            .AddNew
                            !Asiento = Val(WAsiento)
                            Auxi1 = Str$(a)
                            Call Ceros(Auxi1, 2)
                            !Renglon = a
                            !Fecha = Fecha.Text
                            !Observaciones = "Asientos de Depositos"
                            !Cuenta = WCuenta
                            !Debito = WDebito
                            !Credito = WCredito
                            !Leyenda = Left$(WLeyenda, 50)
                            !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            !Clave = WAsiento + Auxi1
                            .Update
                            
                            XCuenta = WCuenta
                            XDebito = WDebito
                            XCredito = WCredito
                            
                            Call Actualiza_Saldos
                                
                        End If
                            
                    Next a
                
                End With
                    
                WMarca = "X"
                XParam = "'" + WDeposito + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                spDepositos = "ActualizaDepositosAsiento " + XParam
                Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
                
            Case "R"
                DesCliente = ""
                spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                Set RstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If RstCliente.RecordCount > 0 Then
                    With RstCliente
                        DesCliente = IIf(IsNull(!Razon), "", !Razon)
                    End With
                    RstCliente.Close
                End If
            
                With rstAsiento
        
                    Renglon = 0
                    .Index = "Clave"
                                        
                    For a = 1 To 100
                        
                        If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                            WCuenta = Impre(a, 1)
                            WDebito = Val(Impre(a, 2))
                            WCredito = Val(Impre(a, 3))
                            WLeyenda = "Recibo Nro. " + WRecibo + " " + DesCliente
                            WLeyenda = Left$(WLeyenda, 50)

                            .AddNew
                            !Asiento = Val(WAsiento)
                            Auxi1 = Str$(a)
                            Call Ceros(Auxi1, 2)
                            !Renglon = a
                            !Fecha = Fecha.Text
                            !Observaciones = "Asientos de Cobranzas"
                            !Cuenta = WCuenta
                            !Debito = WDebito
                            !Credito = WCredito
                            !Leyenda = Left$(WLeyenda, 50)
                            !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            !Clave = WAsiento + Auxi1
                            .Update
                            
                            XCuenta = WCuenta
                            XDebito = WDebito
                            XCredito = WCredito
                            
                            Call Actualiza_Saldos
                                
                        End If
                            
                    Next a
                
                End With
                    
                WMarca = "X"
                XParam = "'" + WRecibo + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                spRecibo = "ActualizaRecibosAsiento " + XParam
                Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
                
            Case "V"
                DesCliente = ""
                spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                Set RstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If RstCliente.RecordCount > 0 Then
                    With RstCliente
                        DesCliente = IIf(IsNull(!Razon), "", !Razon)
                    End With
                    RstCliente.Close
                End If
            
                With rstAsiento
        
                    Renglon = 0
                    .Index = "Clave"
                                    
                    For a = 1 To 100
                        
                        If Val(Impre(a, 2)) <> 0 Or Val(Impre(a, 3)) <> 0 Then
        
                            WCuenta = Impre(a, 1)
                            WDebito = Val(Impre(a, 2))
                            WCredito = Val(Impre(a, 3))
                            WLeyenda = "Factura Nro. " + Mid$(WFactura, 7, 8) + " " + DesCliente
                            WLeyenda = Left$(WLeyenda, 50)

                            .AddNew
                            !Asiento = Val(WAsiento)
                            Auxi1 = Str$(a)
                            Call Ceros(Auxi1, 2)
                            !Renglon = a
                            !Fecha = Fecha.Text
                            !Observaciones = "Asientos de Ventas"
                            !Cuenta = WCuenta
                            !Debito = WDebito
                            !Credito = WCredito
                            !Leyenda = Left$(WLeyenda, 50)
                            !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                            !Clave = WAsiento + Auxi1
                            .Update
                            
                            XCuenta = WCuenta
                            XDebito = WDebito
                            XCredito = WCredito
                            
                            Call Actualiza_Saldos
                                
                        End If
                            
                    Next a
                
                End With
                    
                WMarca = "X"
                XParam = "'" + WFactura + "','" _
                        + WMarca + "','" _
                        + WAsiento + "'"
                
                spCtacte = "ActualizaCtacteAsiento " + XParam
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                
            Case Else
            
        End Select
                
    End If
    
    Call Cancela_click
    
    Exit Sub
    
Error_Programa:
     Rem coderr = Err
     Rem Call Errores(coderr, "Error en el sistema", "Se produjo el error " + Str$(coderr))
     Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstImpcyb
        .Close
    End With
    With rstEmpre
        .Close
    End With
    With rstCue
        .Close
    End With
    With rstAsiento
        .Close
    End With
    Fecha.SetFocus
    PrgGraba2.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Fecha.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Impcyb
    OPEN_FILE_Empre
    OPEN_FILE_Cue
    OPEN_FILE_Asiento
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
End Sub

Private Sub Actualiza_Saldos()

    With rstCue
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


