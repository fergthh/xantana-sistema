VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCashFlow 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cash Flow"
   ClientHeight    =   6225
   ClientLeft      =   2880
   ClientTop       =   585
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
   ScaleWidth      =   6135
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   720
      TabIndex        =   3
      Top             =   120
      Width           =   4815
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
         MouseIcon       =   "CashFlow.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "CashFlow.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   4680
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
         Left            =   3480
         MouseIcon       =   "CashFlow.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "CashFlow.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salida"
         Top             =   4680
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
         Left            =   2040
         MouseIcon       =   "CashFlow.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "CashFlow.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Impresion x Impresora"
         Top             =   4680
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
         Left            =   1560
         TabIndex        =   9
         Top             =   3960
         Width           =   2295
      End
      Begin MSMask.MaskEdBox Vence4 
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence3 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence2 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence1 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Emision 
         Height          =   300
         Left            =   2520
         TabIndex        =   0
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
         Caption         =   "Fecha de Emision"
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
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Parametros de Fechas"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5640
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "CashFlow.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Saldos de Cuenta Corriente de Proveedores"
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
      Left            =   5160
      TabIndex        =   2
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WSaldo As Double
Dim ZVector(10000, 25) As String
Dim ZLugar As Integer

Dim ZZTipo As String
Dim ZZDescripcion  As String
Dim ZZImporte1  As String
Dim ZZImporte2  As String
Dim ZZImporte3  As String
Dim ZZImporte4  As String
Dim ZZImporte5  As String
Dim ZZVencimiento  As String
Dim ZZProveedor  As String
Dim ZZCliente As String
Dim ZZLetra  As String
Dim ZZImpre  As String
Dim ZZPunto  As String
Dim ZZNumero  As String
Dim ZZObservaciones As String
Dim ZZCodigoEmpresa As String

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
    
    WActividad = " al " + Emision.Text
        
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Auxi1 = " + "'" + Vence1.Text + "',"
    ZSql = ZSql + " Auxi2 = " + "'" + Vence2.Text + "',"
    ZSql = ZSql + " Auxi3 = " + "'" + Vence3.Text + "',"
    ZSql = ZSql + " Auxi4 = " + "'" + Vence4.Text + "',"
    ZSql = ZSql + " Actividad = " + "'" + WActividad + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    

    Listado.WindowTitle = "Listado de Cash Flow"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WEmision = Right$(Emision.Text, 4) + Mid$(Emision.Text, 4, 2) + Left$(Emision.Text, 2)
    
    Fecha1 = Right$(Vence1.Text, 4) + Mid$(Vence1.Text, 4, 2) + Left$(Vence1.Text, 2)
    Fecha2 = Right$(Vence2.Text, 4) + Mid$(Vence2.Text, 4, 2) + Left$(Vence2.Text, 2)
    Fecha3 = Right$(Vence3.Text, 4) + Mid$(Vence3.Text, 4, 2) + Left$(Vence3.Text, 2)
    Fecha4 = Right$(Vence4.Text, 4) + Mid$(Vence4.Text, 4, 2) + Left$(Vence4.Text, 2)
    
    
    
    Rem dada
    Rem Borra los movimientos anteriores
    Rem dada
    
    ZSql = ""
    ZSql = ZSql + "DELETE Cash"
    spCash = ZSql
    Set rstCash = db.OpenRecordset(spCash, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    Rem dada
    Rem Graba las deudas de los proveedords
    Rem dada
    
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCtePrv"
    ZSql = ZSql + " Where CtaCtePrv.Saldo <> 0"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCtePrv.RecordCount > 0 Then
    
        With rstCtaCtePrv
            .MoveFirst
            Do
                
                WSaldo = !Saldo
                Call Redondeo(WSaldo)
                
                If WSaldo <> 0 Then
                
                    WSaldo = !Saldo
                    WVencimiento = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)
                    WProveedor = !Proveedor
                    WDescripcion = !Observaciones
                    WLetra = !Letra
                    WImpre = !Impre
                    WPunto = !Punto
                    WNumero = !Numero
                    
                    WImporte1 = 0
                    WImporte2 = 0
                    WImporte3 = 0
                    WImporte4 = 0
                    WIMporte5 = 0
                    
                    If WVencimiento <= Fecha1 Then
                        WImporte1 = WSaldo
                            Else
                        If WVencimiento <= Fecha2 Then
                            WImporte2 = WSaldo
                                Else
                            If WVencimiento <= Fecha3 Then
                                WImporte3 = WSaldo
                                    Else
                                If WVencimiento <= Fecha4 Then
                                    WImporte4 = WSaldo
                                        Else
                                    WIMporte5 = WSaldo
                                End If
                            End If
                        End If
                    End If
                    
                    ZZTipo = "1"
                    ZZDescripcion = Left$(WDescripcion, 30) + "  " + WImpre + " " + WLetra + " " + WPunto + " " + WNumero
                    ZZImporte1 = Str$(WImporte1 * -1)
                    ZZImporte2 = Str$(WImporte2 * -1)
                    ZZImporte3 = Str$(WImporte3 * -1)
                    ZZImporte4 = Str$(WImporte4 * -1)
                    ZZImporte5 = Str$(WIMporte5 * -1)
                    ZZVencimiento = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)
                    ZZProveedor = WProveedor
                    ZZLetra = WLetra
                    ZZImpre = WImpre
                    ZZPunto = WPunto
                    ZZNumero = WNumero
                    
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = ZZTipo
                    ZVector(ZLugar, 2) = ZZDescripcion
                    ZVector(ZLugar, 3) = ZZImporte1
                    ZVector(ZLugar, 4) = ZZImporte2
                    ZVector(ZLugar, 5) = ZZImporte3
                    ZVector(ZLugar, 6) = ZZImporte4
                    ZVector(ZLugar, 7) = ZZImporte5
                    ZVector(ZLugar, 8) = ZZVencimiento
                    ZVector(ZLugar, 9) = ZZProveedor
                    ZVector(ZLugar, 10) = ZZLetra
                    ZVector(ZLugar, 11) = ZZImpre
                    ZVector(ZLugar, 12) = ZZPunto
                    ZVector(ZLugar, 13) = ZZNumero
                    
                End If
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
        ZZDescripcion = ZVector(Ciclo, 2)
        ZZImporte1 = ZVector(Ciclo, 3)
        ZZImporte2 = ZVector(Ciclo, 4)
        ZZImporte3 = ZVector(Ciclo, 5)
        ZZImporte4 = ZVector(Ciclo, 6)
        ZZImporte5 = ZVector(Ciclo, 7)
        ZZVencimiento = ZVector(Ciclo, 8)
        ZZProveedor = ZVector(Ciclo, 9)
        ZZLetra = ZVector(Ciclo, 10)
        ZZImpre = ZVector(Ciclo, 11)
        ZZPunto = ZVector(Ciclo, 12)
        ZZNumero = ZVector(Ciclo, 13)
        ZZObservaciones = ""
        ZZCodigoEmpresa = "1"

        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            ZZDescripcion = rstProveedor!Nombre
            rstProveedor.Close
        End If
                    
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Cash ("
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Importe3 ,"
        ZSql = ZSql + "Importe4 ,"
        ZSql = ZSql + "Importe5 ,"
        ZSql = ZSql + "Importe6 ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Observacones ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Impre ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "CodigoEmpresa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZImporte3 + "',"
        ZSql = ZSql + "'" + ZZImporte4 + "',"
        ZSql = ZSql + "'" + ZZImporte5 + "',"
        ZSql = ZSql + "'" + ZZImporte6 + "',"
        ZSql = ZSql + "'" + ZZVencimiento + "',"
        ZSql = ZSql + "'" + ZZObservacones + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZImpre + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZCodigoEmpresa + "')"
        
        spCash = ZSql
        Set rstCash = db.OpenRecordset(spCash, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
                    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem Graba las deudas de los clientes
    Rem dada
    
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.Saldo <> 0"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
    
        With rstCtaCte
            .MoveFirst
            Do
                WSaldo = !Saldo
                Call Redondeo(WSaldo)
                
                If WSaldo <> 0 Then
                
                    WSaldo = !Saldo
                    WVencimiento = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)
                    WCliente = !Cliente
                    WDescripcion = ""
                    WLetra = !Letra
                    WImpre = !Impre
                    WPunto = !Punto
                    WNumero = !Numero
                    
                    WImporte1 = 0
                    WImporte2 = 0
                    WImporte3 = 0
                    WImporte4 = 0
                    WIMporte5 = 0
                    
                    If WVencimiento <= Fecha1 Then
                        WImporte1 = WSaldo
                            Else
                        If WVencimiento <= Fecha2 Then
                            WImporte2 = WSaldo
                                Else
                            If WVencimiento <= Fecha3 Then
                                WImporte3 = WSaldo
                                    Else
                                If WVencimiento <= Fecha4 Then
                                    WImporte4 = WSaldo
                                        Else
                                    WIMporte5 = WSaldo
                                End If
                            End If
                        End If
                    End If
                    
                    ZZTipo = "2"
                    ZZDescripcion = Left$(WDescripcion, 30) + "  " + WImpre + " " + WLetra + " " + WPunto + " " + WNumero
                    ZZImporte1 = Str$(WImporte1 * -1)
                    ZZImporte2 = Str$(WImporte2 * -1)
                    ZZImporte3 = Str$(WImporte3 * -1)
                    ZZImporte4 = Str$(WImporte4 * -1)
                    ZZImporte5 = Str$(WIMporte5 * -1)
                    ZZVencimiento = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)
                    ZZCliente = WCliente
                    ZZLetra = WLetra
                    ZZImpre = WImpre
                    ZZPunto = WPunto
                    ZZNumero = WNumero
                    
                    ZLugar = ZLugar + 1
                
                    ZVector(ZLugar, 1) = ZZTipo
                    ZVector(ZLugar, 2) = ZZDescripcion
                    ZVector(ZLugar, 3) = ZZImporte1
                    ZVector(ZLugar, 4) = ZZImporte2
                    ZVector(ZLugar, 5) = ZZImporte3
                    ZVector(ZLugar, 6) = ZZImporte4
                    ZVector(ZLugar, 7) = ZZImporte5
                    ZVector(ZLugar, 8) = ZZVencimiento
                    ZVector(ZLugar, 9) = ZZCliente
                    ZVector(ZLugar, 10) = ZZLetra
                    ZVector(ZLugar, 11) = ZZImpre
                    ZVector(ZLugar, 12) = ZZPunto
                    ZVector(ZLugar, 13) = ZZNumero
                    
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
    
        ZZTipo = ZVector(Ciclo, 1)
        ZZDescripcion = ZVector(Ciclo, 2)
        ZZImporte1 = ZVector(Ciclo, 3)
        ZZImporte2 = ZVector(Ciclo, 4)
        ZZImporte3 = ZVector(Ciclo, 5)
        ZZImporte4 = ZVector(Ciclo, 6)
        ZZImporte5 = ZVector(Ciclo, 7)
        ZZVencimiento = ZVector(Ciclo, 8)
        ZZCliente = ZVector(Ciclo, 9)
        ZZLetra = ZVector(Ciclo, 10)
        ZZImpre = ZVector(Ciclo, 11)
        ZZPunto = ZVector(Ciclo, 12)
        ZZNumero = ZVector(Ciclo, 13)
        ZZObservaciones = ""
        ZZCodigoEmpresa = "1"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.razon = " + "'" + ZZCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZDescripcoin = rstCliente!Razon
            rstCliente.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Cash ("
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Importe3 ,"
        ZSql = ZSql + "Importe4 ,"
        ZSql = ZSql + "Importe5 ,"
        ZSql = ZSql + "Importe6 ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Observacones ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Impre ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "CodigoEmpresa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZImporte3 + "',"
        ZSql = ZSql + "'" + ZZImporte4 + "',"
        ZSql = ZSql + "'" + ZZImporte5 + "',"
        ZSql = ZSql + "'" + ZZImporte6 + "',"
        ZSql = ZSql + "'" + ZZVencimiento + "',"
        ZSql = ZSql + "'" + ZZObservacones + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZImpre + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZCodigoEmpresa + "')"
        
        spCash = ZSql
        Set rstCash = db.OpenRecordset(spCash, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
                    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem Graba los cheques posdatados
    Rem dada
    
    
    ZLugar = 0
    Erase ZVector
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pagos"
    ZSql = ZSql + " Where Pagos.Tipo2 = '02'"
    ZSql = ZSql + " and Pagos.Tiporeg = '2'"
    ZSql = ZSql + " and Pagos.FechaOrd2 >= '" + WEmision + "'"
    spPagos = ZSql
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
    
        With rstPagos
            .MoveFirst
            Do
            
                WVencimiento = !FechaOrd2
                    
                WImporte = !Importe2
                WBanco = !Banco2
                WDescripcion = !Observaciones2
                WNumero = !Numero2
                    
                WImporte1 = 0
                WImporte2 = 0
                WImporte3 = 0
                WImporte4 = 0
                WIMporte5 = 0
                    
                If WVencimiento <= Fecha1 Then
                    WImporte1 = WImporte
                        Else
                    If WVencimiento <= Fecha2 Then
                        WImporte2 = WImporte
                            Else
                        If WVencimiento <= Fecha3 Then
                            WImporte3 = WImporte
                                Else
                            If WVencimiento <= Fecha4 Then
                                WImporte4 = WImporte
                                    Else
                                WIMporte5 = WImporte
                            End If
                        End If
                    End If
                End If
                
                ZZTipo = "3"
                ZZDescripcion = Left$(WDescripcion, 30)
                ZZImporte1 = Str$(WImporte1 * -1)
                ZZImporte2 = Str$(WImporte2 * -1)
                ZZImporte3 = Str$(WImporte3 * -1)
                ZZImporte4 = Str$(WImporte4 * -1)
                ZZImporte5 = Str$(WIMporte5 * -1)
                ZZVencimiento = WVencimiento
                ZZBanco = WBanco
                ZZLetra = ""
                ZZImpre = ""
                ZZPunto = ""
                ZZNumero = WNumero
                    
                ZLugar = ZLugar + 1
                
                ZVector(ZLugar, 1) = ZZTipo
                ZVector(ZLugar, 2) = ZZDescripcion
                ZVector(ZLugar, 3) = ZZImporte1
                ZVector(ZLugar, 4) = ZZImporte2
                ZVector(ZLugar, 5) = ZZImporte3
                ZVector(ZLugar, 6) = ZZImporte4
                ZVector(ZLugar, 7) = ZZImporte5
                ZVector(ZLugar, 8) = ZZVencimiento
                ZVector(ZLugar, 9) = ZZBanco
                ZVector(ZLugar, 10) = ZZLetra
                ZVector(ZLugar, 11) = ZZImpre
                ZVector(ZLugar, 12) = ZZPunto
                ZVector(ZLugar, 13) = ZZNumero
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
        End With
        rstPagos.Close
    End If
    
    
    
    
    For Ciclo = 1 To ZLugar
    
        ZZTipo = ZVector(Ciclo, 1)
        ZZDescripcion = ZVector(Ciclo, 2)
        ZZImporte1 = ZVector(Ciclo, 3)
        ZZImporte2 = ZVector(Ciclo, 4)
        ZZImporte3 = ZVector(Ciclo, 5)
        ZZImporte4 = ZVector(Ciclo, 6)
        ZZImporte5 = ZVector(Ciclo, 7)
        ZZVencimiento = ZVector(Ciclo, 8)
        ZZBanco = ZVector(Ciclo, 9)
        ZZLetra = ZVector(Ciclo, 10)
        ZZImpre = ZVector(Ciclo, 11)
        ZZPunto = ZVector(Ciclo, 12)
        ZZNumero = ZVector(Ciclo, 13)
        ZZObservaciones = ""
        ZZCodigoEmpresa = "1"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + ZZBanco + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            ZZDescripcion = rstBanco!Nombre
            rstBanco.Close
        End If
        
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Cash ("
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Importe3 ,"
        ZSql = ZSql + "Importe4 ,"
        ZSql = ZSql + "Importe5 ,"
        ZSql = ZSql + "Importe6 ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Observacones ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Impre ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "CodigoEmpresa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZImporte3 + "',"
        ZSql = ZSql + "'" + ZZImporte4 + "',"
        ZSql = ZSql + "'" + ZZImporte5 + "',"
        ZSql = ZSql + "'" + ZZImporte6 + "',"
        ZSql = ZSql + "'" + ZZVencimiento + "',"
        ZSql = ZSql + "'" + ZZObservacones + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZImpre + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZCodigoEmpresa + "')"
        
        spCash = ZSql
        Set rstCash = db.OpenRecordset(spCash, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
                    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem dada
    Rem Graba los cheques en cartera
    Rem dada
    
    
    ZLugar = 0
    Erase ZVector
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.Tiporeg = '2'"
    ZSql = ZSql + " and Recibos.Tipo2 = '02'"
    ZSql = ZSql + " and Recibos.Estado2 <> 'X'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
            .MoveFirst
            Do
                
                WVencimiento = Right(!Fecha2, 4) + Mid(!Fecha2, 4, 2) + Left(!Fecha2, 2)
                WImporte = !Importe2
                WDescripcion = !Banco2
                WNumero = !Numero2
                    
                WImporte1 = 0
                WImporte2 = 0
                WImporte3 = 0
                WImporte4 = 0
                WIMporte5 = 0
                    
                If WVencimiento <= Fecha1 Then
                    WImporte1 = WImporte
                        Else
                    If WVencimiento <= Fecha2 Then
                        WImporte2 = WImporte
                            Else
                        If WVencimiento <= Fecha3 Then
                            WImporte3 = WImporte
                                Else
                            If WVencimiento <= Fecha4 Then
                                WImporte4 = WImporte
                                    Else
                                WIMporte5 = WImporte
                            End If
                        End If
                    End If
                End If
                
                ZZTipo = "4"
                ZZDescripcion = Left$(WDescripcion, 30)
                ZZImporte1 = Str$(WImporte1 * -1)
                ZZImporte2 = Str$(WImporte2 * -1)
                ZZImporte3 = Str$(WImporte3 * -1)
                ZZImporte4 = Str$(WImporte4 * -1)
                ZZImporte5 = Str$(WIMporte5 * -1)
                ZZVencimiento = WVencimiento
                ZZBanco = ""
                ZZLetra = ""
                ZZImpre = ""
                ZZPunto = ""
                ZZNumero = WNumero
                    
                ZLugar = ZLugar + 1
                
                ZVector(ZLugar, 1) = ZZTipo
                ZVector(ZLugar, 2) = ZZDescripcion
                ZVector(ZLugar, 3) = ZZImporte1
                ZVector(ZLugar, 4) = ZZImporte2
                ZVector(ZLugar, 5) = ZZImporte3
                ZVector(ZLugar, 6) = ZZImporte4
                ZVector(ZLugar, 7) = ZZImporte5
                ZVector(ZLugar, 8) = ZZVencimiento
                ZVector(ZLugar, 9) = ZZBanco
                ZVector(ZLugar, 10) = ZZLetra
                ZVector(ZLugar, 11) = ZZImpre
                ZVector(ZLugar, 12) = ZZPunto
                ZVector(ZLugar, 13) = ZZNumero
                    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZZTipo = ZVector(Ciclo, 1)
        ZZDescripcion = ZVector(Ciclo, 2)
        ZZImporte1 = ZVector(Ciclo, 3)
        ZZImporte2 = ZVector(Ciclo, 4)
        ZZImporte3 = ZVector(Ciclo, 5)
        ZZImporte4 = ZVector(Ciclo, 6)
        ZZImporte5 = ZVector(Ciclo, 7)
        ZZVencimiento = ZVector(Ciclo, 8)
        ZZLetra = ZVector(Ciclo, 10)
        ZZImpre = ZVector(Ciclo, 11)
        ZZPunto = ZVector(Ciclo, 12)
        ZZNumero = ZVector(Ciclo, 13)
        ZZObservaciones = ""
        ZZCodigoEmpresa = "1"
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Cash ("
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Descripcion ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Importe3 ,"
        ZSql = ZSql + "Importe4 ,"
        ZSql = ZSql + "Importe5 ,"
        ZSql = ZSql + "Importe6 ,"
        ZSql = ZSql + "Vencimiento ,"
        ZSql = ZSql + "Observacones ,"
        ZSql = ZSql + "Letra ,"
        ZSql = ZSql + "Impre ,"
        ZSql = ZSql + "Punto ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "CodigoEmpresa )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZTipo + "',"
        ZSql = ZSql + "'" + ZZDescripcion + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZImporte3 + "',"
        ZSql = ZSql + "'" + ZZImporte4 + "',"
        ZSql = ZSql + "'" + ZZImporte5 + "',"
        ZSql = ZSql + "'" + ZZImporte6 + "',"
        ZSql = ZSql + "'" + ZZVencimiento + "',"
        ZSql = ZSql + "'" + ZZObservacones + "',"
        ZSql = ZSql + "'" + ZZLetra + "',"
        ZSql = ZSql + "'" + ZZImpre + "',"
        ZSql = ZSql + "'" + ZZPunto + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZCodigoEmpresa + "')"
        
        spCash = ZSql
        Set rstCash = db.OpenRecordset(spCash, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "CashFlow.rpt"
            Else
        Listado.ReportFileName = "CashFlowcons.rpt"
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Cash.Tipo, Cash.Descripcion, Cash.Importe1, Cash.Importe2, Cash.Importe3, Cash.Importe4, Cash.Importe5, Cash.Letra, Cash.Impre, Cash.Numero, " _
                    + "Auxiliar.Nombre, Auxiliar.Actividad, Auxiliar.Auxi1, Auxiliar.Auxi2, Auxiliar.Auxi3, Auxiliar.Auxi4 " _
                    + "From " _
                    + DSQ + ".dbo.Cash Cash, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar " _
                    + "Where " _
                    + "Cash.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Cash.Tipo >= 0 AND " _
                    + "Cash.Tipo <= 9999"
    
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgCashFlow.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Emision_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Emision.Text, Auxi)
        If Auxi = "S" Then
            Vence1.SetFocus
                Else
            Emision.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Emision.Text = "  /  /    "
    End If
End Sub

Private Sub Vence1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Then
            Vence2.SetFocus
                Else
            Vence1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence1.Text = "  /  /    "
    End If
End Sub

Private Sub Vence2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence2.Text, Auxi)
        If Auxi = "S" Then
            Vence3.SetFocus
                Else
            Vence2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence2.Text = "  /  /    "
    End If
End Sub

Private Sub Vence3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence3.Text, Auxi)
        If Auxi = "S" Then
            Vence4.SetFocus
                Else
            Vence3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence3.Text = "  /  /    "
    End If
End Sub

Private Sub Vence4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence4.Text, Auxi)
        If Auxi = "S" Then
            Vence1.SetFocus
                Else
            Vence4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence4.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Emision.Text = "  /  /    "
    Vence1.Text = "  /  /    "
    Vence2.Text = "  /  /    "
    Vence3.Text = "  /  /    "
    Vence4.Text = "  /  /    "
    Frame2.Visible = True
    
    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub Emision_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
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

