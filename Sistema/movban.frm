VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovban 
   Caption         =   "Listado de Movimientos de Bancos"
   ClientHeight    =   7470
   ClientLeft      =   2280
   ClientTop       =   825
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   7110
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
      Top             =   4200
      Visible         =   0   'False
      Width           =   6855
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
      Height          =   2790
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   5535
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
         Left            =   360
         MouseIcon       =   "movban.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "movban.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2760
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
         Left            =   3000
         MouseIcon       =   "movban.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "movban.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Consulta de Datos"
         Top             =   2760
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
         Left            =   1680
         MouseIcon       =   "movban.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "movban.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2760
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
         Left            =   4320
         MouseIcon       =   "movban.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "movban.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salida"
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox HastaBanco 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   9
         Text            =   " "
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox DesdeBanco 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   8
         Text            =   " "
         Top             =   1200
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   720
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
         Left            =   1680
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label Label5 
         Caption         =   "Tipo Listado"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Banco"
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
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Banco"
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
         Top             =   1200
         Width           =   1335
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6600
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "movban.rpt"
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
Attribute VB_Name = "PrgMovban"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WInicial(100) As Double

Dim ZBanco(100, 2) As String
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

    Rem On Error GoTo WError
    WTitulo = "Del " + Desde.Text + " al " + Hasta.Text
    ZZSaldo = "0"
    ZOrden = 0
    
    If Val(Desdebanco.Text) = 0 Then
        Desdebanco.Text = "0"
    End If
    If Val(HastaBanco.Text) = 0 Then
        HastaBanco.Text = "0"
    End If

    For XDa = 1 To 100
        WInicial(XDa) = 0
    Next XDa

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia


    Rem dada
    Rem lee los bancos
    Rem dada

    Lugarbanco = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Banco"
    ZSql = ZSql + " Order by Banco.Banco"
    spBanco = ZSql
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Do
                If .EOF = False Then
                    Lugarbanco = Lugarbanco + 1
                    ZBanco(Lugarbanco, 1) = rstBanco!Banco
                    ZBanco(Lugarbanco, 2) = rstBanco!Cuenta
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBanco.Close
    End If
    
    
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
    ZSql = ZSql + " and Transferencia.BancoI >= '" + Desdebanco.Text + "'"
    ZSql = ZSql + " and Transferencia.BancoI <= '" + HastaBanco.Text + "'"
    spTransferencia = ZSql
    Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    If rstTransferencia.RecordCount > 0 Then
        With rstTransferencia
            .MoveFirst
            Do
        
                If WDesde <= rstTransferencia!ordfecha And rstTransferencia!BancoI <> 0 Then
            
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = Str$(rstTransferencia!Codigo)
                    ZVector(ZLugar, 2) = rstTransferencia!Fecha
                    ZVector(ZLugar, 3) = rstTransferencia!BancoI
                    ZVector(ZLugar, 4) = Str$(rstTransferencia!Importe)
                    ZVector(ZLugar, 5) = rstTransferencia!ordfecha
                    ZVector(ZLugar, 6) = rstTransferencia!Observaciones
                    ZVector(ZLugar, 7) = Str$(rstTransferencia!TipoI)
                    ZVector(ZLugar, 8) = "1"
                    ZVector(ZLugar, 9) = rstTransferencia!NroCheque
                    
                End If
            
                If WDesde > rstTransferencia!ordfecha Then
                    WImporte = rstTransferencia!Importe
                    WInicial(rstTransferencia!BancoI) = WInicial(rstTransferencia!BancoI) - WImporte
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
    ZSql = ZSql + " and Transferencia.BancoII >= '" + Desdebanco.Text + "'"
    ZSql = ZSql + " and Transferencia.BancoII <= '" + HastaBanco.Text + "'"
    spTransferencia = ZSql
    Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    If rstTransferencia.RecordCount > 0 Then
        With rstTransferencia
            .MoveFirst
            Do
        
                If WDesde <= rstTransferencia!ordfecha And rstTransferencia!BancoII <> 0 Then
            
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = Str$(rstTransferencia!Codigo)
                    ZVector(ZLugar, 2) = rstTransferencia!Fecha
                    ZVector(ZLugar, 3) = rstTransferencia!BancoII
                    ZVector(ZLugar, 4) = Str$(rstTransferencia!Importe)
                    ZVector(ZLugar, 5) = rstTransferencia!ordfecha
                    ZVector(ZLugar, 6) = rstTransferencia!Observaciones
                    ZVector(ZLugar, 7) = Str$(rstTransferencia!TipoII)
                    ZVector(ZLugar, 8) = "2"
                    ZVector(ZLugar, 9) = rstTransferencia!NroCheque
                    
                End If
            
                If WDesde > rstTransferencia!ordfecha Then
                    WImporte = rstTransferencia!Importe
                    WInicial(rstTransferencia!BancoII) = WInicial(rstTransferencia!BancoII) + WImporte
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
        WComprobante = ZVector(Ciclo, 9)
                    
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
    Rem lee los debitos y creditos banarios
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM GastosBancarios"
    ZSql = ZSql + " Where GastosBancarios.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and GastosBancarios.Banco >= '" + Desdebanco.Text + "'"
    ZSql = ZSql + " and GastosBancarios.Banco <= '" + HastaBanco.Text + "'"
    spGastosBancarios = ZSql
    Set rstGastosBancarios = db.OpenRecordset(spGastosBancarios, dbOpenSnapshot, dbSQLPassThrough)
    If rstGastosBancarios.RecordCount > 0 Then
        With rstGastosBancarios
            .MoveFirst
            Do
        
                If WDesde <= rstGastosBancarios!ordfecha Then
            
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = Str$(rstGastosBancarios!Codigo)
                    ZVector(ZLugar, 2) = rstGastosBancarios!Fecha
                    ZVector(ZLugar, 3) = rstGastosBancarios!Banco
                    ZVector(ZLugar, 4) = Str$(rstGastosBancarios!Importe)
                    ZVector(ZLugar, 5) = rstGastosBancarios!ordfecha
                    ZVector(ZLugar, 6) = rstGastosBancarios!Observaciones
                    ZVector(ZLugar, 7) = Str$(rstGastosBancarios!TipoMovimiento)
                    ZVector(ZLugar, 8) = rstGastosBancarios!Comprobante
                    
                End If
            
                If WDesde > rstGastosBancarios!ordfecha Then
                    If rstGastosBancarios!TipoMovimiento = 0 Then
                        WImporte = rstGastosBancarios!Importe
                        WInicial(rstGastosBancarios!Banco) = WInicial(rstGastosBancarios!Banco) - WImporte
                            Else
                        WImporte = rstGastosBancarios!Importe
                        WInicial(rstGastosBancarios!Banco) = WInicial(rstGastosBancarios!Banco) + WImporte
                    End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            
            Loop
        End With
        rstGastosBancarios.Close
    
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
        WTipoMOvimiento = ZVector(Ciclo, 7)
        WComprobante = ZVector(Ciclo, 8)
        WDesProveedor = ""
                    
        ZZBanco = WBanco
        ZZNumero = WComprobante
        ZZComprobante = Right$(Trim(WCodigo), 6)
        ZZfecha = WFecha
        ZZFechaOrd = WOrdFecha
        ZZAcredita = WFecha
        ZZAcreditaOrd = WOrdFecha
        ZZObservaciones = Left$(WObservaciones, 30)
        If Val(WTipoMOvimiento) = 0 Then
            ZZDebito = "0"
            ZZCredito = WImporte
                Else
            ZZDebito = WImporte
            ZZCredito = "0"
        End If
        ZZEmpresa = WEmpresa
        ZZTipoComp = "G.B."
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
    Rem lee las facturas
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    ZSql = ZSql + " Where IvaComp.Contado = 4"
    ZSql = ZSql + " and IvaComp.Banco >= '" + Desdebanco.Text + "'"
    ZSql = ZSql + " and IvaComp.Banco <= '" + HastaBanco.Text + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveFirst
            Do
            
                If WDesde <= rstIvaComp!ordfecha And rstIvaComp!ordfecha <= WHasta Then
            
                    WBanco = rstIvaComp!Banco
                    WOrden = rstIvaComp!Numero
                    WFecha = rstIvaComp!Fecha
                    WFechaOrd = rstIvaComp!ordfecha
                    WAcredita = rstIvaComp!Fecha
                    WAcreditaOrd = rstIvaComp!ordfecha
                    WObservaciones = ""
                    WObservaciones = rstIvaComp!Observaciones
                    WNumero = rstIvaComp!Numero
                    WImporte = rstIvaComp!Neto + rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva27 + rstIvaComp!Ib + rstIvaComp!Exento + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible
                    WOrden = rstIvaComp!Numero
                    WProveedor = rstIvaComp!Proveedor
                        
                    ZLugar = ZLugar + 1
                            
                    ZVector(ZLugar, 1) = Str$(WBanco)
                    ZVector(ZLugar, 2) = WFecha
                    ZVector(ZLugar, 3) = WFechaOrd
                    ZVector(ZLugar, 4) = WAcredita
                    ZVector(ZLugar, 5) = WAcreditaOrd
                    ZVector(ZLugar, 6) = Left$(WObservaciones, 30)
                    ZVector(ZLugar, 7) = WNumero
                    If WImporte > 0 Then
                        ZVector(ZLugar, 8) = "0"
                        ZVector(ZLugar, 9) = Str$(WImporte)
                            Else
                        ZVector(ZLugar, 8) = Str$(Abs(WImporte))
                        ZVector(ZLugar, 9) = "0"
                    End If
                    ZVector(ZLugar, 10) = WOrden
                    ZVector(ZLugar, 11) = WProveedor
                    
                End If
                
                If WDesde > rstIvaComp!ordfecha Then
                    Rem If Val(rstIvaComp!Tipo2) = 2 Then
                        If Val(rstIvaComp!Banco) >= Val(Desdebanco.Text) And Val(rstIvaComp!Banco) <= Val(HastaBanco.Text) Then
                            WImporte = rstIvaComp!Neto + rstIvaComp!Iva21 + rstIvaComp!Iva5 + rstIvaComp!Iva27 + rstIvaComp!Ib + rstIvaComp!Exento + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible
                            WInicial(rstIvaComp!Banco) = WInicial(rstIvaComp!Banco) - WImporte
                        End If
                    Rem End If
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
        
        WDesProveedor = ""
        If WProveedor <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                WDesProveedor = Trim(rstProveedor!Nombre) + "  " + WObservaciones
                rstProveedor.Close
            End If
        End If
        
                    
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
        ZZTipoComp = "Fac"
        ZZEmpresa = "1"
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
        ZSql = ZSql + "'" + Left$(ZZObservaciones, 30) + "',"
        ZSql = ZSql + "'" + Left$(ZZDesProveedor, 50) + "',"
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
    ZSql = ZSql + " Where Pagos.Tipo2 = '02'"
    ZSql = ZSql + " and Pagos.Banco2 >= '" + Desdebanco.Text + "'"
    ZSql = ZSql + " and Pagos.Banco2 <= '" + HastaBanco.Text + "'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
            
                If Val(rstPagos!FechaOrd2) = 0 Then
                    WFechaOrd = rstPagos!fechaord
                    WFecha = rstPagos!Fecha
                    WFechaOrd2 = rstPagos!fechaord
                    WFecha2 = rstPagos!Fecha
                        Else
                    WFechaOrd = rstPagos!fechaord
                    WFecha = rstPagos!Fecha
                    WFechaOrd2 = Trim(rstPagos!FechaOrd2)
                    WFecha2 = rstPagos!Fecha2
                End If
            
                If Tipo.ListIndex = 0 Then
                    FechaCompara = WFechaOrd2
                        Else
                    FechaCompara = WFechaOrd
                    WFechaOrd2 = rstPagos!fechaord
                    WFecha2 = rstPagos!Fecha
                End If
            
                If WDesde <= FechaCompara And FechaCompara <= WHasta Then
            
                    WBanco = rstPagos!Banco2
                    WOrden = rstPagos!Orden
                    WFecha = WFecha
                    WFechaOrd = FechaCompara
                    WAcredita = WFecha2
                    WAcreditaOrd = WFechaOrd2
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
                    If WImporte > 0 Then
                        ZVector(ZLugar, 8) = "0"
                        ZVector(ZLugar, 9) = Str$(WImporte)
                            Else
                        ZVector(ZLugar, 8) = Str$(Abs(WImporte))
                        ZVector(ZLugar, 9) = "0"
                    End If
                    ZVector(ZLugar, 10) = WOrden
                    ZVector(ZLugar, 11) = WProveedor
                    
                End If
                
                If WDesde > FechaCompara Then
                    If Val(rstPagos!Tipo2) = 2 Then
                        If Val(rstPagos!Banco2) >= Val(Desdebanco.Text) And Val(rstPagos!Banco2) <= Val(HastaBanco.Text) Then
                            aa = rstPagos!Importe2
                            WInicial(rstPagos!Banco2) = WInicial(rstPagos!Banco2) - rstPagos!Importe2
                        End If
                    End If
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
    ZSql = ZSql + " and Pagos.Cuenta <> '" + "" + "'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
        
                If Val(rstPagos!FechaOrd2) = 0 Then
                    WFechaOrd = rstPagos!fechaord
                    WFecha = rstPagos!Fecha
                    WFechaOrd2 = rstPagos!fechaord
                    WFecha2 = rstPagos!Fecha
                        Else
                    WFechaOrd = rstPagos!fechaord
                    WFecha = rstPagos!Fecha
                    WFechaOrd2 = rstPagos!FechaOrd2
                    WFecha2 = rstPagos!Fecha2
                End If
            
                If Tipo.ListIndex = 0 Then
                    FechaCompara = WFechaOrd2
                        Else
                    FechaCompara = WFechaOrd
                    WFechaOrd2 = rstPagos!fechaord
                    WFecha2 = rstPagos!Fecha
                End If
            
                If WDesde <= FechaCompara And FechaCompara <= WHasta Then
            
                    For Ciclo = 1 To Lugarbanco
                    
                        If rstPagos!Cuenta = ZBanco(Ciclo, 2) Then
                    
                            WClave = rstPagos!Clave
                            WOrden = rstPagos!Orden
                            WRenglon = rstPagos!Renglon
                            WFecha = rstPagos!Fecha
                            WFechaOrd = FechaCompara
                            WObservaciones = rstPagos!Observaciones
                            WProveedor = rstPagos!Proveedor
                            
                            WImpre = ""
                            WCuenta = rstPagos!Cuenta
                                        
                            WLetra = ""
                            WTipo = ""
                            WPunto = 0
                            WNumero = ""
                            WImporte = rstPagos!Importe1
                                
                            ZLugar = ZLugar + 1
                                
                            ZVector(ZLugar, 1) = ZBanco(Ciclo, 1)
                            ZVector(ZLugar, 2) = WFecha
                            ZVector(ZLugar, 3) = WFechaOrd
                            ZVector(ZLugar, 4) = WFecha
                            ZVector(ZLugar, 5) = WFechaOrd
                            ZVector(ZLugar, 6) = Left$(WObservaciones, 30)
                            ZVector(ZLugar, 7) = WOrden
                            ZVector(ZLugar, 8) = "0"
                            If WImporte > 0 Then
                                ZVector(ZLugar, 8) = "0"
                                ZVector(ZLugar, 9) = Str$(WImporte)
                                    Else
                                ZVector(ZLugar, 8) = Str$(Abs(WImporte))
                                ZVector(ZLugar, 9) = "0"
                            End If
                            ZVector(ZLugar, 10) = WOrden
                            ZVector(ZLugar, 11) = WProveedor
                                
                        End If
                        
                    Next Ciclo
                    
                End If
                    
                If WDesde > FechaCompara Then
                    If Val(rstPagos!TipoOrd) = 3 And Val(rstPagos!Tiporeg) = 1 Then
                        For Ciclo = 1 To Lugarbanco
                            If rstPagos!Cuenta = ZBanco(Ciclo, 2) Then
                                XBanco = Val(ZBanco(Ciclo, 1))
                                WInicial(XBanco) = WInicial(XBanco) - rstPagos!Importe1
                            End If
                        Next Ciclo
                    End If
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
        
        WDesProveedor = ""
        If WProveedor <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                WDesProveedor = Trim(rstProveedor!Nombre) + "  " + WObservaciones
                rstProveedor.Close
            End If
        End If
        
                    
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
        ZSql = ZSql + "'" + Left$(ZZObservaciones, 30) + "',"
        ZSql = ZSql + "'" + Left$(ZZDesProveedor, 50) + "',"
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
    ZSql = ZSql + " Where Depositos.Banco >= '" + Desdebanco.Text + "'"
    ZSql = ZSql + " and Depositos.Banco <= '" + HastaBanco.Text + "'"
    ZSql = ZSql + " and Depositos.Renglon = '01'"
    spDepositos = ZSql
    Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
    If rstDepositos.RecordCount > 0 Then
    
        With rstDepositos
            .MoveFirst
            Do
        
                WFechaOrd = rstDepositos!fechaord
                WFecha = rstDepositos!Fecha
                WFechaOrd2 = rstDepositos!AcreditaOrd
                WFecha2 = rstDepositos!Acredita
            
                If Tipo.ListIndex = 0 Then
                    FechaCompara = WFechaOrd2
                        Else
                    FechaCompara = WFechaOrd
                    WFechaOrd2 = !fechaord
                    WFecha2 = !Fecha
                End If
        
                If WDesde <= FechaCompara And FechaCompara <= WHasta Then
                
                    WBanco = !Banco
                    WFecha = !Fecha
                    WFechaOrd = FechaCompara
                    WAcredita = !Acredita
                    WAcreditaOrd = !AcreditaOrd
                    WObservaciones = "Deposito"
                    WNumero = !Deposito
                    WImporte = !Importe
                    WDeposito = !Deposito
                        
                    ZLugar = ZLugar + 1
                                
                    ZVector(ZLugar, 1) = Str$(WBanco)
                    ZVector(ZLugar, 2) = WFecha
                    ZVector(ZLugar, 3) = WFechaOrd
                    ZVector(ZLugar, 4) = WAcredita
                    ZVector(ZLugar, 5) = WAcreditaOrd
                    ZVector(ZLugar, 6) = Left$(WObservaciones, 30)
                    ZVector(ZLugar, 7) = WDeposito
                    ZVector(ZLugar, 8) = Str$(WImporte)
                    ZVector(ZLugar, 9) = "0"
                    ZVector(ZLugar, 10) = WNumero
                    
                End If
                
                If WDesde > FechaCompara Then
                    WInicial(!Banco) = WInicial(!Banco) + !Importe
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
        ZZDesProveedor = ""
        
    
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
    Rem lee los recibos
    Rem dada
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Tiporeg = '2'"
    ZSql = ZSql + " and Recibos.Tipo2 = '03'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
            .MoveFirst
            Do
            
                If Val(Desdebanco.Text) <= Val(!Banco2) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                
                
                    If WDesde <= !FechaOrd2 And !FechaOrd2 <= WHasta Then
                    
                        WClave = !Clave
                        WBanco = !Banco2
                        WRecibo = !recibo
                        WRenglon = !Renglon
                        WFecha = !Fecha
                        WFechaOrd = !fechaord
                        If !TipoRec = "3" Then
                            WObservaciones = !Observaciones
                                Else
                            WObservaciones = !Cliente
                        End If
                        
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
                        ZVector(ZLugar, 9) = ""
                        ZVector(ZLugar, 10) = WRecibo
                    
                    End If
                
                    If WDesde > !FechaOrd2 Then
                        WImporte = !Importe2
                        WInicial(Val(!Banco2)) = WInicial(Val(!Banco2)) + WImporte
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
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WObservaciones + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WObservaciones = rstCliente!Razon
            rstCliente.Close
        End If
                    
        ZZBanco = Trim(WBanco)
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
        ZZDesProveedor = ""
        
    
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
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    For XDa = 1 To 100
    
        If WInicial(XDa) <> 0 Then
        
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
            If WInicial(XDa) > 0 Then
                ZZCredito = Str$(WInicial(XDa))
                ZZDebito = "0"
                    Else
                ZZCredito = "0"
                ZZDebito = Str$(Abs(WInicial(XDa)))
            End If
            ZZTipoComp = ""
            ZZEmpresa = "1"
            ZZDesEmpresa = WNombreEmpresa
            ZZPeriodo = WTitulo
            ZZDesProveedor = ""
        
    
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
        
        End If
        
    Next XDa
    
    
    
    
    
    
    
    
    
    
    
    
    
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
    
    
    
    
    
    
    
    
    
    
    
    
    
        
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Varios = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT movban.banco, movban.Fecha, movban.fechaord, movban.Acredita, movban.AcreditaOrd, movban.observaciones, movban.numero, movban.debito, movban.credito, movban.comprobante, movban.Tipocomp, movban.DesEmpresa, movban.Periodo, movban.Proveedor, movban.Orden, movban.OrdenII, movban.Saldo,  " _
                + "Banco.Nombre " _
                + "From " _
                + DSQ + ".dbo.Movban movban, " _
                + DSQ + ".dbo.Banco Banco " _
                + "Where " _
                + "movban.banco = Banco.Banco AND " _
                + "movban.banco >= " + Desdebanco.Text + " AND " _
                + "movban.banco <= " + HastaBanco.Text
    
    Listado.Connect = Connect()
    
    Uno = "{movban.Banco} in " + Desdebanco.Text + " to " + HastaBanco.Text
    Rem Dos = " and {IvaComp.OrdPeriodo} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    
    Rem Listado.GroupSelectionFormula = Uno + Dos
    Rem Listado.SelectionFormula = Uno + Dos
    
    Listado.Action = 1

    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgMovban.Hide
    Unload Me
    MenuAdminis.Show
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
            Desdebanco.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Private Sub DesdeBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaBanco.SetFocus
    End If
    If KeyAscii = 27 Then
        Desdebanco.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaBanco.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    
    Tipo.Clear
    
    Tipo.AddItem "Por Vencimiento"
    Tipo.AddItem "Por Fecha de Emision"
    
    Tipo.ListIndex = 0

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Desdebanco.Text = ""
    HastaBanco.Text = ""
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

Private Sub DesdeBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaBanco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Panta_Click
        Case 115
            Call Consulta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_Click
        Case Else
    End Select
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Banco"
    ZSql = ZSql + " Order by Banco.Banco"
    spBanco = ZSql
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
                Do
                If .EOF = False Then
                    IngresaItem = Str$(!Banco) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Banco
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBanco.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Indice = Pantalla.ListIndex
    WCodigo = WIndice.List(Indice)
    Desdebanco.Text = WCodigo
    HastaBanco.Text = WCodigo
    Desdebanco.SetFocus
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
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Banco"
    ZSql = ZSql + " Where Banco.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
    ZSql = ZSql + " Order by Banco.Banco"
    spBanco = ZSql
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        With rstBanco
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(!Banco) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Banco
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstBanco.Close
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub
