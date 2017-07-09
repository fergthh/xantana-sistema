VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovCaja 
   Caption         =   "Listado de SubDiario de Caja Diaria"
   ClientHeight    =   3480
   ClientLeft      =   2280
   ClientTop       =   1305
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3480
   ScaleWidth      =   7110
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton Panta 
         Caption         =   "Pantalla F1"
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
         Left            =   600
         MouseIcon       =   "movcaja.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "movcaja.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Impre 
         Caption         =   "Listado F9"
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
         Left            =   1920
         MouseIcon       =   "movcaja.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "movcaja.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1440
         Width           =   855
      End
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
         Left            =   3240
         MouseIcon       =   "movcaja.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "movcaja.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salida"
         Top             =   1440
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   2280
         TabIndex        =   4
         Top             =   840
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   480
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
         Left            =   720
         TabIndex        =   3
         Top             =   840
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6480
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "MovCaja.rpt"
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
Attribute VB_Name = "PrgMovCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WInicial As Double
Dim ZVector(10000, 25) As String
Dim ZLugar As Integer

Dim ZZBanco As String
Dim ZZNumero As String
Dim ZZComprobante As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZAcredita As String
Dim ZZAcreditaOrd As String
Dim ZZObservaciones As String
Dim ZZDebito As String
Dim ZZCredito As String
Dim ZZEmpresa As String
Dim ZZTipoComp As String
Dim ZZDesEmpresa As String
Dim ZZPeriodo As String

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    WInicial = 0
    ZOrden = 0
    
    If Val(WEmpresa) = 2 Then
        WInicial = -2088.69 - 1650
    End If

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    WTitulo = "Del " + Desde.Text + " al " + Hasta.Text
        
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Actividad = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Rem dada
    Rem Borra los movimientos anteriores
    Rem dada
    
    ZSql = ""
    ZSql = ZSql + "DELETE MovBan"
    spMovBan = ZSql
    Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    










    
    
    
    
    
    
    
    
    Rem dada
    Rem lee lAS TRANSFERENCIAS
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Transferencia"
    ZSql = ZSql + " Where Transferencia.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Transferencia.TipoI = '" + "1" + "'"
    spTransferencia = ZSql
    Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    If rstTransferencia.RecordCount > 0 Then
        With rstTransferencia
            .MoveFirst
            Do
        
                If WDesde <= rstTransferencia!ordfecha Then
            
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = Str$(rstTransferencia!Codigo)
                    ZVector(ZLugar, 2) = rstTransferencia!Fecha
                    ZVector(ZLugar, 3) = rstTransferencia!BancoI
                    ZVector(ZLugar, 4) = Str$(rstTransferencia!Importe)
                    ZVector(ZLugar, 5) = rstTransferencia!ordfecha
                    ZVector(ZLugar, 6) = rstTransferencia!Observaciones
                    ZVector(ZLugar, 7) = Str$(rstTransferencia!TipoI)
                    ZVector(ZLugar, 8) = "1"
                    
                End If
            
                If WDesde > rstTransferencia!ordfecha Then
                    WImporte = rstTransferencia!Importe
                    WInicial = WInicial - WImporte
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            
            Loop
        End With
        rstTransferencia.Close
    
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Transferencia"
    ZSql = ZSql + " Where Transferencia.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Transferencia.TipoII = '" + "1" + "'"
    spTransferencia = ZSql
    Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    If rstTransferencia.RecordCount > 0 Then
        With rstTransferencia
            .MoveFirst
            Do
        
                If WDesde <= rstTransferencia!ordfecha Then
            
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = Str$(rstTransferencia!Codigo)
                    ZVector(ZLugar, 2) = rstTransferencia!Fecha
                    ZVector(ZLugar, 3) = rstTransferencia!BancoII
                    ZVector(ZLugar, 4) = Str$(rstTransferencia!Importe)
                    ZVector(ZLugar, 5) = rstTransferencia!ordfecha
                    ZVector(ZLugar, 6) = rstTransferencia!Observaciones
                    ZVector(ZLugar, 7) = Str$(rstTransferencia!TipoII)
                    ZVector(ZLugar, 8) = "2"
                    
                End If
            
                If WDesde > rstTransferencia!ordfecha Then
                    WImporte = rstTransferencia!Importe
                    WInicial = WInicial + WImporte
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            
            Loop
        End With
        rstTransferencia.Close
    
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZOrden = ZOrden + 1
        WWOrden = Str$(ZOrden)
        WCodigo = ZVector(Ciclo, 1)
        WFecha = ZVector(Ciclo, 2)
        WBanco = ZVector(Ciclo, 3)
        WImporte = ZVector(Ciclo, 4)
        WOrdFecha = ZVector(Ciclo, 5)
        WObservaciones = ZVector(Ciclo, 6)
        WTipo = ZVector(Ciclo, 7)
        WTipoII = ZVector(Ciclo, 8)
        WComprobante = ""
                    
        ZZBanco = WBanco
        ZZNumero = WComprobante
        ZZComprobante = Right$(Trim(WCodigo), 6)
        ZZfecha = WFecha
        ZZFechaOrd = WOrdFecha
        ZZAcredita = WFecha
        ZZAcreditaOrd = WOrdFecha
        ZZObservaciones = Left$(WObservaciones, 30)
        If Val(WTipoII) = 1 Then
            ZZDebito = "0"
            ZZCredito = WImporte
                Else
            ZZDebito = WImporte
            ZZCredito = "0"
        End If
        ZZEmpresa = WEmpresa
        ZZTipoComp = "Tran"
        ZZDesEmpresa = WNombreEmpresa
        ZZPeriodo = WTitulo
        ZZDesProveedor = WDesProveedor
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO MovBan ("
        ZSql = ZSql + "banco ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "fechaord ,"
        ZSql = ZSql + "Acredita ,"
        ZSql = ZSql + "AcreditaOrd ,"
        ZSql = ZSql + "observaciones ,"
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "numero ,"
        ZSql = ZSql + "debito ,"
        ZSql = ZSql + "credito ,"
        ZSql = ZSql + "comprobante ,"
        ZSql = ZSql + "empresa ,"
        ZSql = ZSql + "Tipocomp ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "DesEmpresa ,"
        ZSql = ZSql + "Periodo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZAcredita + "',"
        ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "',"
        ZSql = ZSql + "'" + ZZDesProveedor + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZDebito + "',"
        ZSql = ZSql + "'" + ZZCredito + "',"
        ZSql = ZSql + "'" + ZZComprobante + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZTipoComp + "',"
        ZSql = ZSql + "'" + ZZSaldo + "',"
        ZSql = ZSql + "'" + WNombreEmpresa + "',"
        ZSql = ZSql + "'" + WTitulo + "')"
        spMovBan = ZSql
        Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    
        
    Next Ciclo
                    
    
    
    
    













    
    
    
    
    Rem dada
    Rem lee las facturas de pagadas de contado
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    ZSql = ZSql + " Where IvaComp.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and IvaComp.Contado = 1"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveFirst
            Do
        
                WFechaOrd = IIf(IsNull(!ordfecha), "00000000", !ordfecha)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
            
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = rstIvaComp!Proveedor
                    ZVector(ZLugar, 2) = rstIvaComp!Tipo
                    ZVector(ZLugar, 3) = rstIvaComp!Letra
                    ZVector(ZLugar, 4) = rstIvaComp!Punto
                    ZVector(ZLugar, 5) = rstIvaComp!Numero
                    ZVector(ZLugar, 6) = rstIvaComp!Fecha
                    ZVector(ZLugar, 7) = rstIvaComp!Vencimiento
                    ZVector(ZLugar, 8) = rstIvaComp!Periodo
                    ZVector(ZLugar, 9) = Str$(rstIvaComp!Neto)
                    ZVector(ZLugar, 10) = Str$(rstIvaComp!Iva21)
                    ZVector(ZLugar, 11) = Str$(rstIvaComp!Iva5)
                    ZVector(ZLugar, 12) = Str$(rstIvaComp!Iva27)
                    ZVector(ZLugar, 13) = Str$(rstIvaComp!Ib)
                    ZVector(ZLugar, 14) = Str$(rstIvaComp!ImpInterno)
                    ZVector(ZLugar, 15) = Str$(rstIvaComp!ImpCombustible)
                    ZVector(ZLugar, 16) = Str$(rstIvaComp!Iva105)
                    ZVector(ZLugar, 17) = Str$(rstIvaComp!Exento)
                    ZVector(ZLugar, 18) = rstIvaComp!Impre
                    ZVector(ZLugar, 19) = rstIvaComp!ordfecha
                    ZVector(ZLugar, 20) = rstIvaComp!Observaciones
                    ZVector(ZLugar, 21) = Str$(rstIvaComp!Neto + rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva27 + rstIvaComp!Ib + rstIvaComp!Iva105 + rstIvaComp!Exento + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible)
                    
                End If
                
                If WDesde > WFechaOrd Then
                    WImporte = rstIvaComp!Neto + rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva27 + rstIvaComp!Ib + rstIvaComp!Iva105 + rstIvaComp!Exento + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible
                    WInicial = WInicial - WImporte
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            
            Loop
        End With
        rstIvaComp.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZOrden = ZOrden + 1
        WWOrden = Str$(ZOrden)
        WProveedor = ZVector(Ciclo, 1)
        WTipo = ZVector(Ciclo, 2)
        WLetra = ZVector(Ciclo, 3)
        WPunto = ZVector(Ciclo, 4)
        WNumero = ZVector(Ciclo, 5)
        WFecha = ZVector(Ciclo, 6)
        WVencimiento = ZVector(Ciclo, 7)
        WPeriodo = ZVector(Ciclo, 8)
        WNeto = ZVector(Ciclo, 9)
        WIva21 = ZVector(Ciclo, 10)
        WIva5 = ZVector(Ciclo, 11)
        WIva27 = ZVector(Ciclo, 12)
        WIb = ZVector(Ciclo, 13)
        WImpInterno = ZVector(Ciclo, 14)
        WImpCombustible = ZVector(Ciclo, 15)
        WIva105 = ZVector(Ciclo, 16)
        WExento = ZVector(Ciclo, 17)
        WImpre = ZVector(Ciclo, 18)
        WOrdFecha = ZVector(Ciclo, 19)
        WObservaciones = ZVector(Ciclo, 20)
        WImporte = ZVector(Ciclo, 21)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            WObservaciones = rstProveedor!Nombre
            rstProveedor.Close
        End If
                    
        ZZBanco = WBanco
        ZZNumero = ""
        ZZComprobante = WNumero
        ZZfecha = WFecha
        ZZFechaOrd = WOrdFecha
        ZZAcredita = WFecha
        ZZAcreditaOrd = WOrdFecha
        ZZObservaciones = Left$(WObservaciones, 30)
        ZZDebito = "0"
        ZZCredito = WImporte
        ZZEmpresa = WEmpresa
        ZZTipoComp = "FAC."
        ZZDesEmpresa = WNombreEmpresa
        ZZPeriodo = WTitulo
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO MovBan ("
        ZSql = ZSql + "banco ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "fechaord ,"
        ZSql = ZSql + "Acredita ,"
        ZSql = ZSql + "AcreditaOrd ,"
        ZSql = ZSql + "observaciones ,"
        ZSql = ZSql + "numero ,"
        ZSql = ZSql + "debito ,"
        ZSql = ZSql + "credito ,"
        ZSql = ZSql + "comprobante ,"
        ZSql = ZSql + "empresa ,"
        ZSql = ZSql + "Tipocomp ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "DesEmpresa ,"
        ZSql = ZSql + "Periodo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZAcredita + "',"
        ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZDebito + "',"
        ZSql = ZSql + "'" + ZZCredito + "',"
        ZSql = ZSql + "'" + ZZComprobante + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZTipoComp + "',"
        ZSql = ZSql + "'" + ZZSaldo + "',"
        ZSql = ZSql + "'" + WNombreEmpresa + "',"
        ZSql = ZSql + "'" + WTitulo + "')"
        spMovBan = ZSql
        Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    
        
    Next Ciclo
                    

    
    
    
    
    
    
    
    

    
    
    
    
    Rem dada
    Rem lee las facturas de ventas cobradas de contado
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and CtaCte.Contado = 1"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        With rstCtaCte
            .MoveFirst
            Do
        
                WFechaOrd = IIf(IsNull(!ordfecha), "00000000", !ordfecha)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
            
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = rstCtaCte!Cliente
                    ZVector(ZLugar, 2) = rstCtaCte!Tipo
                    ZVector(ZLugar, 3) = rstCtaCte!Letra
                    ZVector(ZLugar, 4) = rstCtaCte!Punto
                    ZVector(ZLugar, 5) = rstCtaCte!Numero
                    ZVector(ZLugar, 6) = rstCtaCte!Fecha
                    ZVector(ZLugar, 7) = Str$(rstCtaCte!Total)
                    ZVector(ZLugar, 8) = rstCtaCte!ordfecha
                    
                End If
                
                If WDesde > WFechaOrd Then
                    WImporte = rstCtaCte!Total
                    WInicial = WInicial + WImporte
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            
            Loop
        End With
        rstCtaCte.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZOrden = ZOrden + 1
        WWOrden = Str$(ZOrden)
        WCliente = ZVector(Ciclo, 1)
        WTipo = ZVector(Ciclo, 2)
        WLetra = ZVector(Ciclo, 3)
        WPunto = ZVector(Ciclo, 4)
        WNumero = ZVector(Ciclo, 5)
        WFecha = ZVector(Ciclo, 6)
        WTotal = ZVector(Ciclo, 7)
        WOrdFecha = ZVector(Ciclo, 8)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WObservaciones = rstCliente!Razon
            rstCliente.Close
        End If
                    
        ZZBanco = "0"
        ZZNumero = ""
        ZZComprobante = WNumero
        ZZfecha = WFecha
        ZZFechaOrd = WOrdFecha
        ZZAcredita = WFecha
        ZZAcreditaOrd = WOrdFecha
        ZZObservaciones = Left$(WObservaciones, 30)
        ZZDebito = WTotal
        ZZCredito = "0"
        ZZEmpresa = WEmpresa
        ZZTipoComp = "FAC."
        ZZDesEmpresa = WNombreEmpresa
        ZZPeriodo = WTitulo
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO MovBan ("
        ZSql = ZSql + "banco ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "fechaord ,"
        ZSql = ZSql + "Acredita ,"
        ZSql = ZSql + "AcreditaOrd ,"
        ZSql = ZSql + "observaciones ,"
        ZSql = ZSql + "numero ,"
        ZSql = ZSql + "debito ,"
        ZSql = ZSql + "credito ,"
        ZSql = ZSql + "comprobante ,"
        ZSql = ZSql + "empresa ,"
        ZSql = ZSql + "Tipocomp ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "DesEmpresa ,"
        ZSql = ZSql + "Periodo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZAcredita + "',"
        ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZDebito + "',"
        ZSql = ZSql + "'" + ZZCredito + "',"
        ZSql = ZSql + "'" + ZZComprobante + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZTipoComp + "',"
        ZSql = ZSql + "'" + ZZSaldo + "',"
        ZSql = ZSql + "'" + WNombreEmpresa + "',"
        ZSql = ZSql + "'" + WTitulo + "')"
        spMovBan = ZSql
        Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    
        
    Next Ciclo
                    

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem lee las ordenes de pago
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Tipo2 = '01'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
            
                WFechaOrd = IIf(IsNull(!fechaord), "00000000", !fechaord)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
                If rstPagos!Importe2 <> 0 Then
            
                    WBanco = rstPagos!Banco2
                    WFecha = rstPagos!Fecha
                    WFechaOrd = rstPagos!fechaord
                    WAcredita = rstPagos!Fecha
                    WAcreditaOrd = rstPagos!fechaord
                    WObservaciones = ""
                    WObservaciones = rstPagos!Observaciones
                    WNumero = rstPagos!Numero2
                    WImporte = rstPagos!Importe2
                    WOrden = rstPagos!Orden
                    WProveedor = rstPagos!Proveedor
                        
                    ZLugar = ZLugar + 1
                            
                    ZVector(ZLugar, 1) = Str$(WBanco)
                    ZVector(ZLugar, 2) = WFecha
                    ZVector(ZLugar, 3) = WFechaOrd
                    ZVector(ZLugar, 4) = WAcredita
                    ZVector(ZLugar, 5) = WAcreditaOrd
                    ZVector(ZLugar, 6) = Left$(WObservaciones, 30)
                    ZVector(ZLugar, 7) = WNumero
                    ZVector(ZLugar, 8) = "0"
                    ZVector(ZLugar, 9) = Str$(WImporte)
                    ZVector(ZLugar, 10) = WOrden
                    ZVector(ZLugar, 11) = WProveedor
                    
                End If
                End If
                    
                If WDesde > WFechaOrd Then
                    WInicial = WInicial - rstPagos!Importe2
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstPagos.Close
    End If
        
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.TipoOrd = '3'"
    ZSql = ZSql + " and Pagos.Tiporeg = '1'"
    Rem ZSql = ZSql + " and Pagos.Cuenta = '" + WCtaEfectivo + "'"
    ZSql = ZSql + " and Pagos.Cuenta = '" + "9999999999999" + "'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
            
                WFechaOrd = IIf(IsNull(!fechaord), "00000000", !fechaord)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
            
                    WBanco = rstPagos!Banco2
                    WFecha = rstPagos!Fecha
                    WFechaOrd = rstPagos!fechaord
                    WAcredita = rstPagos!Fecha
                    WAcreditaOrd = rstPagos!fechaord
                    WObservaciones = ""
                    WObservaciones = rstPagos!Observaciones
                    WNumero = rstPagos!Numero2
                    WImporte = rstPagos!Importe2
                    WOrden = rstPagos!Orden
                    WProveedor = rstPagos!Proveedor
                        
                    ZLugar = ZLugar + 1
                            
                    ZVector(ZLugar, 1) = Str$(WBanco)
                    ZVector(ZLugar, 2) = WFecha
                    ZVector(ZLugar, 3) = WFechaOrd
                    ZVector(ZLugar, 4) = WAcredita
                    ZVector(ZLugar, 5) = WAcreditaOrd
                    ZVector(ZLugar, 6) = Left$(WObservaciones, 30)
                    ZVector(ZLugar, 7) = WNumero
                    ZVector(ZLugar, 8) = "0"
                    ZVector(ZLugar, 9) = Str$(WImporte)
                    ZVector(ZLugar, 10) = WOrden
                    ZVector(ZLugar, 11) = WProveedor
                    
                End If
                    
                If WDesde > WFechaOrd Then
                    WInicial = WInicial - rstPagos!Importe2
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstPagos.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZOrden = ZOrden + 1
        WWOrden = Str$(ZOrden)
        WBanco = ZVector(Ciclo, 1)
        WFecha = ZVector(Ciclo, 2)
        WFechaOrd = ZVector(Ciclo, 3)
        WFecha2 = ZVector(Ciclo, 4)
        WFechaOrd2 = ZVector(Ciclo, 5)
        WObservaciones = ZVector(Ciclo, 6)
        WNumero = ZVector(Ciclo, 7)
        WDebito = ZVector(Ciclo, 8)
        WCredito = ZVector(Ciclo, 9)
        WOrden = ZVector(Ciclo, 10)
        WProveedor = ZVector(Ciclo, 11)
        
        ZZBanco = WBanco
        ZZNumero = WNumero
        ZZComprobante = WOrden
        ZZfecha = WFecha
        ZZFechaOrd = WFechaOrd
        ZZAcredita = WFecha2
        ZZAcreditaOrd = WFechaOrd2
        ZZObservaciones = Left$(WObservaciones, 30)
        ZZDebito = WDebito
        ZZCredito = WCredito
        ZZTipoComp = "O.P."
        ZZEmpresa = "1"
        ZZDesEmpresa = WNombreEmpresa
        ZZPeriodo = WTitulo
        
        
        If Trim(WProveedor) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                If Trim(ZZObservaciones) <> "" Then
                    ZZObservaciones = Trim(rstProveedor!Nombre) + "  (" + Trim(ZZObservaciones) + ")"
                        Else
                    ZZObservaciones = rstProveedor!Nombre
                End If
                rstProveedor.Close
            End If
        End If
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO MovBan ("
        ZSql = ZSql + "banco ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "fechaord ,"
        ZSql = ZSql + "Acredita ,"
        ZSql = ZSql + "AcreditaOrd ,"
        ZSql = ZSql + "observaciones ,"
        ZSql = ZSql + "numero ,"
        ZSql = ZSql + "debito ,"
        ZSql = ZSql + "credito ,"
        ZSql = ZSql + "comprobante ,"
        ZSql = ZSql + "empresa ,"
        ZSql = ZSql + "Tipocomp ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "DesEmpresa ,"
        ZSql = ZSql + "Periodo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZAcredita + "',"
        ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
        ZSql = ZSql + "'" + Left$(ZZObservaciones, 30) + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZDebito + "',"
        ZSql = ZSql + "'" + ZZCredito + "',"
        ZSql = ZSql + "'" + ZZComprobante + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZTipoComp + "',"
        ZSql = ZSql + "'" + ZZSaldo + "',"
        ZSql = ZSql + "'" + WNombreEmpresa + "',"
        ZSql = ZSql + "'" + WTitulo + "')"
        spMovBan = ZSql
        Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem lee los depositos
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Depositos"
    ZSql = ZSql + " Where Depositos.Renglon = '01'"
    ZSql = ZSql + " and Depositos.Tipo2 = '01'"
    spDepositos = ZSql
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then

        With rstDepositos
            .MoveFirst
            Do
                WFechaOrd = IIf(IsNull(!fechaord), "00000000", !fechaord)
                If Trim(WDesde) <= Trim(WFechaOrd) And Trim(WFechaOrd) <= Trim(WHasta) Then
                
                    WBanco = 0
                    WFecha = rstDepositos!Fecha
                    WFechaOrd = rstDepositos!fechaord
                    WAcredita = rstDepositos!Acredita
                    WAcreditaOrd = rstDepositos!AcreditaOrd
                    WObservaciones = "DEPOSITO "
                    WNumero = "00" + rstDepositos!Deposito
                    WImporte = rstDepositos!Importe2
                    WDeposito = "00" + rstDepositos!Deposito
                    XBanco = rstDepositos!Banco
                    
                    ZLugar = ZLugar + 1
                                
                    ZVector(ZLugar, 1) = Str$(WBanco)
                    ZVector(ZLugar, 2) = WFecha
                    ZVector(ZLugar, 3) = WFechaOrd
                    ZVector(ZLugar, 4) = WAcredita
                    ZVector(ZLugar, 5) = WAcreditaOrd
                    ZVector(ZLugar, 6) = Left$(WObservaciones, 30)
                    ZVector(ZLugar, 7) = WDeposito
                    ZVector(ZLugar, 8) = "0"
                    ZVector(ZLugar, 9) = Str$(WImporte)
                    ZVector(ZLugar, 10) = WNumero
                    
                End If
                
                If WDesde > WFechaOrd Then
                    WInicial = WInicial - !Importe2
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstDepositos.Close
    End If
        
        
    
    For Ciclo = 1 To ZLugar
    
        ZOrden = ZOrden + 1
        WWOrden = Str$(ZOrden)
        WBanco = ZVector(Ciclo, 1)
        WFecha = ZVector(Ciclo, 2)
        WFechaOrd = ZVector(Ciclo, 3)
        WFecha2 = ZVector(Ciclo, 4)
        WFechaOrd2 = ZVector(Ciclo, 5)
        WObservaciones = ZVector(Ciclo, 6)
        WNumero = ZVector(Ciclo, 7)
        WDebito = ZVector(Ciclo, 8)
        WCredito = ZVector(Ciclo, 9)
        WOrden = ZVector(Ciclo, 10)
                    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + ZZBanco + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            WObservaciones = "Deposito Banco: " + rstBanco!Nombre
            rstBanco.Close
        End If
        ZZBanco = ""
                    
        ZZBanco = WBanco
        ZZNumero = ""
        ZZComprobante = WNumero
        ZZfecha = WFecha
        ZZFechaOrd = WFechaOrd
        ZZAcredita = WFecha2
        ZZAcreditaOrd = WFechaOrd2
        ZZObservaciones = Left$(WObservaciones, 30)
        ZZDebito = WDebito
        ZZCredito = WCredito
        ZZTipoComp = "Dep."
        ZZEmpresa = "1"
        ZZDesEmpresa = WNombreEmpresa
        ZZPeriodo = WTitulo
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO MovBan ("
        ZSql = ZSql + "banco ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "fechaord ,"
        ZSql = ZSql + "Acredita ,"
        ZSql = ZSql + "AcreditaOrd ,"
        ZSql = ZSql + "observaciones ,"
        ZSql = ZSql + "numero ,"
        ZSql = ZSql + "debito ,"
        ZSql = ZSql + "credito ,"
        ZSql = ZSql + "comprobante ,"
        ZSql = ZSql + "empresa ,"
        ZSql = ZSql + "Tipocomp ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "DesEmpresa ,"
        ZSql = ZSql + "Periodo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZAcredita + "',"
        ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZDebito + "',"
        ZSql = ZSql + "'" + ZZCredito + "',"
        ZSql = ZSql + "'" + ZZComprobante + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZTipoComp + "',"
        ZSql = ZSql + "'" + ZZSaldo + "',"
        ZSql = ZSql + "'" + WNombreEmpresa + "',"
        ZSql = ZSql + "'" + WTitulo + "')"
        spMovBan = ZSql
        Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    
        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem lee los recibos
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Tiporeg = '2'"
    ZSql = ZSql + " and Recibos.Tipo2 = '01'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
            .MoveFirst
            
                
            Do
            
            
                WFechaOrd = IIf(IsNull(!fechaord), "00000000", !fechaord)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
            
                    
                    WClave = !Clave
                    WRecibo = !recibo
                    WRenglon = !Renglon
                    WFecha = !Fecha
                    WFechaOrd = !fechaord
                    WObservaciones = !Observaciones
                    WCliente = !Cliente
                            
                    WImpre = ""
                    WCuenta = !Cuenta
                                    
                    WLetra = ""
                    WTipo = !Tipo2
                    WPunto = 0
                    WNumero = !Numero2
                    WImporte = !Importe2
                            
                    ZLugar = ZLugar + 1
                                
                    ZVector(ZLugar, 1) = WBanco
                    ZVector(ZLugar, 2) = WFecha
                    ZVector(ZLugar, 3) = WFechaOrd
                    ZVector(ZLugar, 4) = WFecha
                    ZVector(ZLugar, 5) = WFechaOrd
                    ZVector(ZLugar, 6) = Left$(WObservaciones, 30)
                    ZVector(ZLugar, 7) = WNumero
                    ZVector(ZLugar, 8) = Str$(WImporte)
                    ZVector(ZLugar, 9) = "0"
                    ZVector(ZLugar, 10) = WRecibo
                    ZVector(ZLugar, 11) = WCliente
                    
                End If
                    
                If WDesde > WFechaOrd Then
                    WInicial = WInicial + !Importe2
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    
    For Ciclo = 1 To ZLugar
    
        ZOrden = ZOrden + 1
        WWOrden = Str$(ZOrden)
        WBanco = ZVector(Ciclo, 1)
        WFecha = ZVector(Ciclo, 2)
        WFechaOrd = ZVector(Ciclo, 3)
        WFecha2 = ZVector(Ciclo, 4)
        WFechaOrd2 = ZVector(Ciclo, 5)
        WObservaciones = ZVector(Ciclo, 6)
        WNumero = ZVector(Ciclo, 7)
        WDebito = ZVector(Ciclo, 8)
        WCredito = ZVector(Ciclo, 9)
        WRecibo = ZVector(Ciclo, 10)
        WCliente = ZVector(Ciclo, 11)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WObservaciones = rstCliente!Razon
            rstCliente.Close
                Else
            WObservaciones = WCliente
        End If
                    
        ZZBanco = WBanco
        ZZNumero = WNumero
        ZZComprobante = WRecibo
        ZZfecha = WFecha
        ZZFechaOrd = WFechaOrd
        ZZAcredita = WFecha2
        ZZAcreditaOrd = WFechaOrd2
        ZZObservaciones = Left$(WObservaciones, 30)
        ZZDebito = WDebito
        ZZCredito = WCredito
        ZZTipoComp = "Rec."
        ZZEmpresa = "1"
        ZZDesEmpresa = WNombreEmpresa
        ZZPeriodo = WTitulo
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO MovBan ("
        ZSql = ZSql + "banco ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "fechaord ,"
        ZSql = ZSql + "Acredita ,"
        ZSql = ZSql + "AcreditaOrd ,"
        ZSql = ZSql + "observaciones ,"
        ZSql = ZSql + "numero ,"
        ZSql = ZSql + "debito ,"
        ZSql = ZSql + "credito ,"
        ZSql = ZSql + "comprobante ,"
        ZSql = ZSql + "empresa ,"
        ZSql = ZSql + "Tipocomp ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "DesEmpresa ,"
        ZSql = ZSql + "Periodo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZBanco + "',"
        ZSql = ZSql + "'" + WWOrden + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZFechaOrd + "',"
        ZSql = ZSql + "'" + ZZAcredita + "',"
        ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZDebito + "',"
        ZSql = ZSql + "'" + ZZCredito + "',"
        ZSql = ZSql + "'" + ZZComprobante + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZTipoComp + "',"
        ZSql = ZSql + "'" + ZZSaldo + "',"
        ZSql = ZSql + "'" + WNombreEmpresa + "',"
        ZSql = ZSql + "'" + WTitulo + "')"
        spMovBan = ZSql
        Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    
        
    Next Ciclo
    
    
    
    
    
    
    ZOrden = ZOrden + 1
    WWOrden = Str$(ZOrden)
    ZZBanco = Str$(XDa)
    ZZNumero = ""
    ZZComprobante = ""
    ZZfecha = ""
    ZZFechaOrd = "00000000"
    ZZAcredita = ""
    ZZAcreditaOrd = "00000000"
    ZZObservaciones = "Saldo Inicial"
    If WInicial > 0 Then
        ZZCredito = "0"
        ZZDebito = Str$(WInicial)
            Else
        ZZCredito = Str$(Abs(WInicial))
        ZZDebito = "0"
    End If
    ZZTipoComp = ""
    ZZEmpresa = "1"
    ZZDesEmpresa = WNombreEmpresa
    ZZPeriodo = WTitulo
        
    
    ZSql = ""
    ZSql = ZSql + "INSERT INTO MovBan ("
    ZSql = ZSql + "banco ,"
    ZSql = ZSql + "Orden ,"
    ZSql = ZSql + "fecha ,"
    ZSql = ZSql + "fechaord ,"
    ZSql = ZSql + "Acredita ,"
    ZSql = ZSql + "AcreditaOrd ,"
    ZSql = ZSql + "observaciones ,"
    ZSql = ZSql + "numero ,"
    ZSql = ZSql + "debito ,"
    ZSql = ZSql + "credito ,"
    ZSql = ZSql + "comprobante ,"
    ZSql = ZSql + "empresa ,"
    ZSql = ZSql + "Tipocomp ,"
    ZSql = ZSql + "Saldo ,"
    ZSql = ZSql + "DesEmpresa ,"
    ZSql = ZSql + "Periodo )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + ZZBanco + "',"
    ZSql = ZSql + "'" + WWOrden + "',"
    ZSql = ZSql + "'" + ZZfecha + "',"
    ZSql = ZSql + "'" + ZZFechaOrd + "',"
    ZSql = ZSql + "'" + ZZAcredita + "',"
    ZSql = ZSql + "'" + ZZAcreditaOrd + "',"
    ZSql = ZSql + "'" + ZZObservaciones + "',"
    ZSql = ZSql + "'" + ZZNumero + "',"
    ZSql = ZSql + "'" + ZZDebito + "',"
    ZSql = ZSql + "'" + ZZCredito + "',"
    ZSql = ZSql + "'" + ZZComprobante + "',"
    ZSql = ZSql + "'" + ZZEmpresa + "',"
    ZSql = ZSql + "'" + ZZTipoComp + "',"
    ZSql = ZSql + "'" + ZZSaldo + "',"
    ZSql = ZSql + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + "'" + WTitulo + "')"
    spMovBan = ZSql
    Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem calcula los saldos
    Rem dada
    
    ZLugar = 0
    ZPasa = 0
    ZSaldo = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Movban"
    ZSql = ZSql + " Order by Movban.Banco, Movban.FechaOrd"
    spMovBan = ZSql
    Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovBan.RecordCount > 0 Then
    
        With rstMovBan
            .MoveFirst
            
                
            Do
            
            
                If ZPasa = 0 Then
                    ZCorte = rstMovBan!Banco
                    ZPasa = 1
                End If
                
                If ZCorte <> rstMovBan!Banco Then
                    ZSaldo = 0
                    ZCorte = rstMovBan!Banco
                End If
                
                ZSaldo = ZSaldo + rstMovBan!Debito - rstMovBan!Credito

                ZLugar = ZLugar + 1
                ZVector(ZLugar, 1) = rstMovBan!Orden
                ZVector(ZLugar, 2) = Str$(ZSaldo)
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstMovBan.Close
    End If
    
    
    
    For Ciclo = 1 To ZLugar
    
        WWOrden = ZVector(Ciclo, 1)
        WSaldo = ZVector(Ciclo, 2)
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Movban SET "
        ZSql = ZSql + " Saldo = " + "'" + WSaldo + "',"
        ZSql = ZSql + " OrdenII = " + "'" + Str$(Ciclo) + "'"
        ZSql = ZSql + " Where Orden = " + "'" + WWOrden + "'"
        spMovBan = ZSql
        Set rstMovBan = db.OpenRecordset(spMovBan, dbOpenSnapshot, dbSQLPassThrough)

        
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT movban.banco, movban.fecha, movban.fechaord, movban.AcreditaOrd, movban.observaciones, movban.numero, movban.debito, movban.credito, movban.comprobante, movban.Tipocomp, movban.DesEmpresa, movban.Periodo, movban.Orden, movban.OrdenII, movban.Saldo " _
                + "From " _
                + DSQ + ".dbo.movban movban " _
                + "Where " _
                + "movban.banco >= 0 AND " _
                + "movban.banco <= 9999"
    
    Listado.Connect = Connect()
    
    Rem Uno = "{IvaComp.Letra} <> " + Chr$(34) + "X" + Chr$(34)
    Rem Dos = " and {IvaComp.OrdPeriodo} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    
    Rem Listado.GroupSelectionFormula = Uno + Dos
    Rem Listado.SelectionFormula = Uno + Dos
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgMovCaja.Hide
    Unload Me
    MenuAdminis.SetFocus
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
    If KeyAscii = 27 Then
        Desde.Text = "  /  /    "
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Panta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_Click
        Case Else
    End Select
End Sub
















