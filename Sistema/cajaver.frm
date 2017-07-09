VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovCajaOtro 
   Caption         =   "Listado de Caja"
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
         MouseIcon       =   "cajaver.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "cajaver.frx":030A
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
         MouseIcon       =   "cajaver.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "cajaver.frx":0E56
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
         MouseIcon       =   "cajaver.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "cajaver.frx":19A2
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
      ReportFileName  =   "MovCajaOtro.rpt"
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
Attribute VB_Name = "PrgMovCajaOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WInicial As Double
Dim ZVector(10000, 25) As String
Dim ZVectorII(10000, 25) As String
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
    ZSql = ZSql + " and Transferencia.TipoI = '" + "3" + "'"
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
                    ZVector(ZLugar, 4) = Str$(rstTransferencia!IMPORTE)
                    ZVector(ZLugar, 5) = rstTransferencia!ordfecha
                    ZVector(ZLugar, 6) = rstTransferencia!Observaciones
                    ZVector(ZLugar, 7) = Str$(rstTransferencia!TipoI)
                    ZVector(ZLugar, 8) = "1"
                    
                End If
            
                If WDesde > rstTransferencia!ordfecha Then
                    WImporte = rstTransferencia!IMPORTE
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
    ZSql = ZSql + " and Transferencia.TipoII = '" + "3" + "'"
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
                    ZVector(ZLugar, 4) = Str$(rstTransferencia!IMPORTE)
                    ZVector(ZLugar, 5) = rstTransferencia!ordfecha
                    ZVector(ZLugar, 6) = rstTransferencia!Observaciones
                    ZVector(ZLugar, 7) = Str$(rstTransferencia!TipoII)
                    ZVector(ZLugar, 8) = "2"
                    
                End If
            
                If WDesde > rstTransferencia!ordfecha Then
                    WImporte = rstTransferencia!IMPORTE
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
    Rem lee las ordenes de pago
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    ZLugarII = 0
    Erase ZVectorII
    
    
    txtUserName = "SA"
    txtPassword = "Sw58125812"
    txtOdbc = "Fragancias"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Tipo2 = '05'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
            
                WFechaOrd = IIf(IsNull(rstPagos!fechaord), "00000000", rstPagos!fechaord)
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
                    If WImporte >= 0 Then
                        ZVector(ZLugar, 8) = "0"
                        ZVector(ZLugar, 9) = Str$(WImporte)
                            Else
                        ZVector(ZLugar, 9) = "0"
                        ZVector(ZLugar, 8) = Str$(WImporte * -1)
                    End If
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
    ZSql = ZSql + " FROM GastosCaja"
    spGastosCaja = ZSql
    Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
    If rstGastosCaja.RecordCount > 0 Then
        With rstGastosCaja
            .MoveFirst
            Do
            
                WFechaOrd = IIf(IsNull(rstGastosCaja!ordfecha), "00000000", rstGastosCaja!ordfecha)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
                If rstGastosCaja!IMPORTE <> 0 Then
            
                    WBanco = 0
                    WFecha = rstGastosCaja!Fecha
                    WFechaOrd = rstGastosCaja!ordfecha
                    WAcredita = rstGastosCaja!Fecha
                    WAcreditaOrd = rstGastosCaja!ordfecha
                    WObservaciones = ""
                    WObservaciones = rstGastosCaja!Observaciones
                    WNumero = rstGastosCaja!Codigo
                    WImporte = rstGastosCaja!IMPORTE
                    WOrden = rstGastosCaja!Codigo
                    WProveedor = rstGastosCaja!Concepto
                        
                    ZLugarII = ZLugarII + 1
                            
                    ZVectorII(ZLugarII, 1) = Str$(WBanco)
                    ZVectorII(ZLugarII, 2) = WFecha
                    ZVectorII(ZLugarII, 3) = WFechaOrd
                    ZVectorII(ZLugarII, 4) = WAcredita
                    ZVectorII(ZLugarII, 5) = WAcreditaOrd
                    ZVectorII(ZLugarII, 6) = Left$(WObservaciones, 30)
                    ZVectorII(ZLugarII, 7) = WNumero
                    If WImporte >= 0 Then
                        ZVectorII(ZLugarII, 8) = "0"
                        ZVectorII(ZLugarII, 9) = Str$(WImporte)
                            Else
                        ZVectorII(ZLugarII, 9) = "0"
                        ZVectorII(ZLugarII, 8) = Str$(WImporte * -1)
                    End If
                    ZVectorII(ZLugarII, 10) = WOrden
                    ZVectorII(ZLugarII, 11) = WProveedor
                    
                End If
                End If
                    
                If WDesde > WFechaOrd Then
                    WInicial = WInicial - rstPagos!IMPORTE
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstGastosCaja.Close
    End If
        
    
    
    txtUserName = "SA"
    txtPassword = "Sw58125812"
    txtOdbc = "FraganciasII"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Tipo2 = '05'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
            
                WFechaOrd = IIf(IsNull(rstPagos!fechaord), "00000000", rstPagos!fechaord)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
                If rstGastosCaja!IMPORTE <> 0 Then
            
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
                    If WImporte >= 0 Then
                        ZVector(ZLugar, 8) = "0"
                        ZVector(ZLugar, 9) = Str$(WImporte)
                            Else
                        ZVector(ZLugar, 9) = "0"
                        ZVector(ZLugar, 8) = Str$(WImporte * -1)
                    End If
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
    ZSql = ZSql + " FROM GastosCaja"
    spGastosCaja = ZSql
    Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
    If rstGastosCaja.RecordCount > 0 Then
        With rstGastosCaja
            .MoveFirst
            Do
            
                WFechaOrd = IIf(IsNull(rstGastosCaja!ordfecha), "00000000", rstGastosCaja!ordfecha)
                If WDesde <= WFechaOrd And WFechaOrd <= WHasta Then
                If rstPagos!Importe2 <> 0 Then
            
                    WBanco = 0
                    WFecha = rstGastosCaja!Fecha
                    WFechaOrd = rstGastosCaja!ordfecha
                    WAcredita = rstGastosCaja!Fecha
                    WAcreditaOrd = rstGastosCaja!ordfecha
                    WObservaciones = ""
                    WObservaciones = rstGastosCaja!Observaciones
                    WNumero = rstGastosCaja!Codigo
                    WImporte = rstGastosCaja!IMPORTE
                    WOrden = rstGastosCaja!Codigo
                    WProveedor = rstGastosCaja!Concepto
                        
                    ZLugarII = ZLugarII + 1
                            
                    ZVectorII(ZLugarII, 1) = Str$(WBanco)
                    ZVectorII(ZLugarII, 2) = WFecha
                    ZVectorII(ZLugarII, 3) = WFechaOrd
                    ZVectorII(ZLugarII, 4) = WAcredita
                    ZVectorII(ZLugarII, 5) = WAcreditaOrd
                    ZVectorII(ZLugarII, 6) = Left$(WObservaciones, 30)
                    ZVectorII(ZLugarII, 7) = WNumero
                    If WImporte >= 0 Then
                        ZVectorII(ZLugarII, 8) = "0"
                        ZVectorII(ZLugarII, 9) = Str$(WImporte)
                            Else
                        ZVectorII(ZLugarII, 9) = "0"
                        ZVectorII(ZLugarII, 8) = Str$(WImporte * -1)
                    End If
                    ZVectorII(ZLugarII, 10) = WOrden
                    ZVectorII(ZLugarII, 11) = WProveedor
                    
                End If
                End If
                    
                If WDesde > WFechaOrd Then
                    WInicial = WInicial - rstPagos!IMPORTE
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstGastosCaja.Close
    End If
        
    
    If ZZNivel = 1 Then
        txtUserName = "SA"
        txtPassword = "Sw58125812"
        txtOdbc = "FraganciasII"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Else
        txtUserName = "SA"
        txtPassword = "Sw58125812"
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
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
    
    
    
    
    
    For Ciclo = 1 To ZLugarII
    
        ZOrden = ZOrden + 1
        WWOrden = Str$(ZOrden)
        WBanco = ZVectorII(Ciclo, 1)
        WFecha = ZVectorII(Ciclo, 2)
        WFechaOrd = ZVectorII(Ciclo, 3)
        WFecha2 = ZVectorII(Ciclo, 4)
        WFechaOrd2 = ZVectorII(Ciclo, 5)
        WObservaciones = ZVectorII(Ciclo, 6)
        WNumero = ZVectorII(Ciclo, 7)
        WDebito = ZVectorII(Ciclo, 8)
        WCredito = ZVectorII(Ciclo, 9)
        WOrden = ZVectorII(Ciclo, 10)
        WProveedor = ZVectorII(Ciclo, 11)
        
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
        ZZTipoComp = "Gasto"
        ZZEmpresa = "1"
        ZZDesEmpresa = WNombreEmpresa
        ZZPeriodo = WTitulo
        
        
        If Trim(WProveedor) <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + WProveedor + "'"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                If Trim(ZZObservaciones) <> "" Then
                    ZZObservaciones = Trim(rstConceptos!Nombre) + "  (" + Trim(ZZObservaciones) + ")"
                        Else
                    ZZObservaciones = rstConceptos!Nombre
                End If
                rstConceptos.Close
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
    PrgMovCajaOtro.Hide
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
















