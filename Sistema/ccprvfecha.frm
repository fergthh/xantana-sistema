VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCcprvFecha 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Proveedores a Fecha"
   ClientHeight    =   7770
   ClientLeft      =   1455
   ClientTop       =   615
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   9120
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
      Height          =   300
      Left            =   480
      TabIndex        =   2
      Text            =   " "
      Top             =   3600
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   1680
      TabIndex        =   6
      Top             =   240
      Width           =   5655
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
         Left            =   4200
         MouseIcon       =   "ccprvfecha.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ccprvfecha.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salida"
         Top             =   1920
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
         Left            =   1800
         MouseIcon       =   "ccprvfecha.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ccprvfecha.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1920
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
         MouseIcon       =   "ccprvfecha.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ccprvfecha.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consulta de Datos"
         Top             =   1920
         Width           =   855
      End
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
         MouseIcon       =   "ccprvfecha.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ccprvfecha.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Hasta 
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
         MaxLength       =   8
         TabIndex        =   3
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Desde 
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
         MaxLength       =   8
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   480
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
      Begin VB.Label Label3 
         Caption         =   "Fecha Emision"
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
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
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
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   480
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ccprv.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Proveedores"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   360
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
      Height          =   3570
      ItemData        =   "ccprvfecha.frx":2D30
      Left            =   480
      List            =   "ccprvfecha.frx":2D37
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   7695
   End
End
Attribute VB_Name = "PrgCcprvFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Acumula As Double
Private Pasa As Single
Private WSaldo As Double
Dim Importe1 As Double
Dim Importe2 As Double
Dim Importe3 As Double
Dim BajaPago(10000) As String
Dim ZVector(10000, 30) As String
Dim ZLugar As Integer
Dim LugarPago As Integer

Dim ZZTipo As String
Dim ZZLetra As String
Dim ZZPunto  As String
Dim ZZImpre  As String
Dim ZZNumero  As String
Dim ZZRenglon  As String
Dim ZZCliente  As String
Dim ZZfecha  As String
Dim ZZEstado  As String
Dim ZZVencimiento  As String
Dim ZZTotal  As String
Dim ZZSaldo  As String
Dim ZZNeto  As String
Dim ZZIva1  As String
Dim ZZIva2  As String
Dim ZZOrdFecha  As String
Dim ZZOrdVencimiento  As String
Dim ZZPedido  As String
Dim ZZRemito  As String
Dim ZZOrden  As String
Dim ZZProvincia  As String
Dim ZZCosto  As String
Dim ZZImporte1  As String
Dim ZZImporte2  As String
Dim ZZImporte3  As String
Dim ZZImporte4  As String
Dim ZZImporte5  As String
Dim ZZImporte6  As String
Dim ZZImporte7  As String
Dim ZZClave  As String


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
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFechaOrd = WAno + WMes + WDia

    ZZDesEmpresa = WNombreEmpresa
    ZZPeriodo = Fecha.Text

    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    
    Rem dada
    Rem Borra los movimientos anteriores
    Rem dada
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpCtaCte"
    spImpCtaCte = ZSql
    Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCtePrv"
    ZSql = ZSql + " Where CtaCtePrv.Proveedor >= '" + Desde.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.Proveedor <= '" + Hasta.Text + "'"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCtePrv.RecordCount > 0 Then
    
        With rstCtaCtePrv
            .MoveFirst
            Do
                
                WImporte1 = 0
                WImporte2 = 0
                WImporte3 = 0
                    
                If !Total > 0 Then
                    WImporte1 = !Total
                    WImporte2 = 0
                        Else
                    WImporte1 = 0
                    WImporte2 = !Total
                End If
                
                WImporte3 = !Saldo
                Importe3 = WImporte3
                Call Redondeo(Importe3)
                WImporte3 = Importe3
                    
                WTipo = !Tipo
                WImpre = !Impre
                WNumero = !Numero
                WRenglon = 0
                WCliente = !Proveedor
                WFecha = !Fecha
                WEstado = !Estado
                WVencimiento = !Vencimiento
                WTotal = !Total
                WSaldo = !Saldo
                WOrdFecha = !ordfecha
                WOrdvencimiento = !OrdVencimiento
                WImporte4 = 0
                WIMporte5 = 0
                WIMporte6 = 0
                WIMporte7 = 0
                WClave = !Clave
                WLetra = !Letra
                WPunto = !Punto
                    
                ZZTipo = WTipo
                ZZLetra = WLetra
                ZZPunto = WPunto
                ZZImpre = WImpre
                ZZNumero = WNumero
                ZZRenglon = Str$(WRenglon)
                ZZCliente = WCliente
                ZZfecha = WFecha
                ZZEstado = WEstado
                ZZVencimiento = WVencimiento
                ZZTotal = Str$(WTotal)
                ZZSaldo = Str$(WSaldo)
                ZZOrdFecha = WOrdFecha
                ZZOrdVencimiento = WOrdvencimiento
                ZZImporte1 = Str$(WImporte1)
                ZZImporte2 = Str$(WImporte2)
                ZZImporte3 = Str$(WImporte3)
                ZZImporte4 = Str$(WImporte4)
                ZZImporte5 = Str$(WIMporte5)
                ZZImporte6 = Str$(WIMporte6)
                ZZImporte7 = Str$(WIMporte7)
                
                WProveedor = Trim(!Proveedor)
                Call Ceros(WPunto, 3)
                ZZClave = WProveedor + WTipo + WPunto + Right$(WNumero, 6)
                
                ZLugar = ZLugar + 1
                
                ZVector(ZLugar, 1) = ZZTipo
                ZVector(ZLugar, 2) = ZZLetra
                ZVector(ZLugar, 3) = ZZPunto
                ZVector(ZLugar, 4) = ZZImpre
                ZVector(ZLugar, 5) = ZZNumero
                ZVector(ZLugar, 6) = ZZRenglon
                ZVector(ZLugar, 7) = ZZCliente
                ZVector(ZLugar, 8) = ZZfecha
                ZVector(ZLugar, 9) = ZZEstado
                ZVector(ZLugar, 10) = ZZVencimiento
                ZVector(ZLugar, 11) = ZZTotal
                ZVector(ZLugar, 12) = ZZSaldo
                ZVector(ZLugar, 13) = ZZOrdFecha
                ZVector(ZLugar, 14) = ZZOrdVencimiento
                ZVector(ZLugar, 15) = ZZImporte1
                ZVector(ZLugar, 16) = ZZImporte2
                ZVector(ZLugar, 17) = ZZImporte3
                ZVector(ZLugar, 18) = ZZImporte4
                ZVector(ZLugar, 19) = ZZImporte5
                ZVector(ZLugar, 20) = ZZImporte6
                ZVector(ZLugar, 21) = ZZImporte7
                ZVector(ZLugar, 22) = ZZClave
                    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstCtaCtePrv.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZZTipo = ZVector(Ciclo, 1)
        ZZLetra = ZVector(Ciclo, 2)
        ZZPunto = ZVector(Ciclo, 3)
        ZZImpre = ZVector(Ciclo, 4)
        ZZNumero = ZVector(Ciclo, 5)
        ZZRenglon = ZVector(Ciclo, 6)
        ZZCliente = ZVector(Ciclo, 7)
        ZZfecha = ZVector(Ciclo, 8)
        ZZEstado = ZVector(Ciclo, 9)
        ZZVencimiento = ZVector(Ciclo, 10)
        ZZTotal = ZVector(Ciclo, 11)
        ZZSaldo = ZVector(Ciclo, 12)
        ZZOrdFecha = ZVector(Ciclo, 13)
        ZZOrdVencimiento = ZVector(Ciclo, 14)
        ZZImporte1 = ZVector(Ciclo, 15)
        ZZImporte2 = ZVector(Ciclo, 16)
        ZZImporte3 = ZVector(Ciclo, 17)
        ZZImporte4 = ZVector(Ciclo, 18)
        ZZImporte5 = ZVector(Ciclo, 19)
        ZZImporte6 = ZVector(Ciclo, 20)
        ZZImporte7 = ZVector(Ciclo, 21)
        ZZClave = ZVector(Ciclo, 22)
    
    
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ImpCtaCte ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "fecha ,"
        ZSql = ZSql + "Estado ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Total ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "OrdVencimiento ,"
        ZSql = ZSql + "Impre ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Importe3 ,"
        ZSql = ZSql + "Importe4 ,"
        ZSql = ZSql + "Importe5 ,"
        ZSql = ZSql + "Importe6 ,"
        ZSql = ZSql + "Importe7 ,"
        ZSql = ZSql + "Periodo ,"
        ZSql = ZSql + "DesEmpresa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZEstado + "',"
        ZSql = ZSql + "'" + ZZVencimiento + "',"
        ZSql = ZSql + "'" + ZZTotal + "',"
        ZSql = ZSql + "'" + ZZSaldo + "',"
        ZSql = ZSql + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
        ZSql = ZSql + "'" + ZZImpre + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZImporte3 + "',"
        ZSql = ZSql + "'" + ZZImporte4 + "',"
        ZSql = ZSql + "'" + ZZImporte5 + "',"
        ZSql = ZSql + "'" + ZZImporte6 + "',"
        ZSql = ZSql + "'" + ZZImporte7 + "',"
        ZSql = ZSql + "'" + ZZPeriodo + "',"
        ZSql = ZSql + "'" + ZZDesEmpresa + "')"
        
        spImpCtaCte = ZSql
        Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    
    
    
    Erase BajaPago
    LugarPago = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Proveedor >= '" + Desde.Text + "'"
    ZSql = ZSql + " and Pagos.Proveedor <= '" + Hasta.Text + "'"
    ZSql = ZSql + " and Pagos.FechaOrd > '" + WFechaOrd + "'"
    ZSql = ZSql + " and Pagos.Renglon = '" + "01" + "'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
        With rstPagos
            .MoveFirst
            Do
                LugarPago = LugarPago + 1
                BajaPago(LugarPago) = !Orden
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstPagos.Close
    End If
    
    For Cicla = 1 To LugarPago
    
        WOrden = BajaPago(Cicla)
        WTipoOrd = 0
        
        For da = 1 To 99
            Auxi1 = Str$(da)
            Call Ceros(Auxi1, 2)
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pagos"
            ZSql = ZSql + " Where Pagos.Orden = " + "'" + WOrden + "'"
            ZSql = ZSql + " and Pagos.Renglon = " + "'" + Auxi1 + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            If rstPagos.RecordCount > 0 Then
            
                WLetra = rstPagos!Letra1
                WTipo = rstPagos!Tipo1
                WPunto = rstPagos!Punto1
                WNumero = rstPagos!Numero1
                WImporte = rstPagos!Importe1
                WClaveCheque = rstPagos!ClaveCheque
                WTipoOrd = rstPagos!TipoOrd
                WTipoReg = rstPagos!Tiporeg
                WProveedor = rstPagos!Proveedor
                
                rstPagos.Close
            
                If Val(WTipoReg) = 1 Then
                
                    If Val(WTipoOrd) = 1 Then
                    
                        WProveedor = Trim(WProveedor)
                        Call Ceros(WPunto, 3)
                        Claveven$ = WProveedor + WTipo + WPunto + Right$(WNumero, 6)
                        
                        ZSql = ""
                        ZSql = ZSql + "UPDATE ImpCtaCte SET "
                        ZSql = ZSql + " Importe3 = Importe3 + " + "'" + Str$(WImporte) + "'"
                        ZSql = ZSql + " Where Clave = " + "'" + Claveven$ + "'"
                        spImpCtaCte = ZSql
                        Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                        
                End If
                    
                        Else
                    
                Exit For
                
            End If
            
        Next da
        
        If WTipoOrd = 1 Then
        
            WLetra = "A"
            WTipo = "04"
            WPunto = "000"
            WNumero = WOrden
        
            Call Ceros(WNumero, 6)
            Call Ceros(WProveedor, 6)
            
            WClave = WProveedor + WTipo + WPunto + WNumero
        
            ZSql = ""
            ZSql = ZSql + "DELETE ImpCtaCte"
            ZSql = ZSql + " Where ImpCtaCte.Clave = " + "'" + WClave + "'"
            spImpCtaCte = ZSql
            Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        If WTipoOrd = 2 Then
        
            WLetra = "A"
            WTipo = "05"
            WPunto = "000"
            WNumero = WOrden
        
            Call Ceros(WNumero, 6)
            Call Ceros(WProveedor, 6)
        
            WClave = WProveedor + WTipo + WPunto + WNumero
        
            ZSql = ""
            ZSql = ZSql + "DELETE ImpCtaCte"
            ZSql = ZSql + " Where ImpCtaCte.Clave = " + "'" + WClave + "'"
            spImpCtaCte = ZSql
            Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Cicla
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpCtaCte"
    ZSql = ZSql + " Where ImpCtaCte.OrdFecha > " + "'" + WFechaOrd + "'"
    ZSql = ZSql + " or ImpCtaCte.Importe3 = 0"
    spImpCtaCte = ZSql
    Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.ReportFileName = "CcPrvFecha.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Impctacte.Numero, Impctacte.Cliente, Impctacte.fecha, Impctacte.Ordfecha, Impctacte.Impre, Impctacte.Importe1, Impctacte.Importe2, Impctacte.Importe3, Impctacte.Periodo, Impctacte.DesEmpresa, " _
                + "Proveedor.Nombre " _
                + "From " _
                + DSQ + ".dbo.Impctacte Impctacte, " _
                + DSQ + ".dbo.Proveedor Proveedor " _
                + "Where " _
                + "Impctacte.Cliente = Proveedor.Proveedor AND " _
                + "Impctacte.Cliente >= '' AND " _
                + "Impctacte.Cliente <= 'ZZZZZZ'"
    
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
    PrgCcprvFecha.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Order by Proveedor.Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        With rstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = !Proveedor + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProveedor.Close
    End If
            
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Indice = Pantalla.ListIndex
    Desde.Text = WIndice.List(Indice)
    Hasta.Text = WIndice.List(Indice)
    Desde.SetFocus
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
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Frame2.Visible = True
    Rem Call Consulta_Click
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
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
    ZSql = ZSql + " Order by Proveedor.Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        With rstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = !Proveedor + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProveedor.Close
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub


Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
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





