VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAnalizaPedidos 
   AutoRedraw      =   -1  'True
   Caption         =   "Analisis de Pedidos Semanal"
   ClientHeight    =   8790
   ClientLeft      =   2910
   ClientTop       =   525
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   8790
   ScaleWidth      =   6135
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
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   600
      TabIndex        =   8
      Top             =   0
      Width           =   4695
      Begin VB.CheckBox TipoV 
         Caption         =   "Exporta"
         Height          =   375
         Left            =   3480
         TabIndex        =   20
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox Sector 
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
         Left            =   1200
         MaxLength       =   8
         TabIndex        =   16
         Top             =   5160
         Width           =   855
      End
      Begin VB.ComboBox TipoII 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   4560
         Width           =   2055
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
         Left            =   1200
         TabIndex        =   5
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CommandButton Panta 
         Caption         =   "Panta F1"
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
         Left            =   3600
         MouseIcon       =   "AnalizaPedidos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "AnalizaPedidos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   360
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
         Left            =   3600
         MouseIcon       =   "AnalizaPedidos.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "AnalizaPedidos.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1920
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
         Left            =   3600
         MouseIcon       =   "AnalizaPedidos.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "AnalizaPedidos.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salida"
         Top             =   3360
         Width           =   855
      End
      Begin MSMask.MaskEdBox Vence4 
         Height          =   375
         Left            =   960
         TabIndex        =   3
         Top             =   2520
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
         Left            =   960
         TabIndex        =   2
         Top             =   2040
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
         Left            =   960
         TabIndex        =   1
         Top             =   1560
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
         Left            =   960
         TabIndex        =   0
         Top             =   1080
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
      Begin MSMask.MaskEdBox Vence5 
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   3000
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
      Begin MSMask.MaskEdBox Vence6 
         Height          =   375
         Left            =   960
         TabIndex        =   19
         Top             =   3480
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
      Begin VB.Label DesSector 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   2280
         TabIndex        =   18
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Sector"
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
         TabIndex        =   17
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
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
         Left            =   360
         TabIndex        =   14
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Parametros de Fechas"
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
         Left            =   600
         TabIndex        =   9
         Top             =   480
         Width           =   2415
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5640
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ProyPrv.rpt"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   2640
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
      Height          =   2595
      ItemData        =   "AnalizaPedidos.frx":21E4
      Left            =   120
      List            =   "AnalizaPedidos.frx":21EB
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   5895
   End
End
Attribute VB_Name = "PrgAnalizaPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WTrabajo(10000, 13) As String
Dim Pasa As Integer
Dim Lugar As Integer
Dim Impo1 As Double
Dim Impo2 As Double
Dim Impo3 As Double
Dim Impo4 As Double
Dim Impo5 As Double
Dim Impo6 As Double
Dim Impo7 As Double
Dim WImpoDto As Double


Dim WWLinea As String
Dim WWTipo As String
Dim WWFragancia As String
Dim WWCalidad As String
Dim WWTamano As String


Dim ZZArticulo As String
Dim ZZCliente As String
Dim ZZCantidad As Double
Dim ZZPrecio As Double
Dim ZZfecha As String
Dim WWMoneda As Integer
Dim WWParidad As Double

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
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + "1" + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        WWParidad = rstDolar!Paridad
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE AnalizaPedido"
    spAnalizaPedido = ZSql
    Set rstAnalizaPedido = db.OpenRecordset(spAnalizaPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Analisis de pedidos mensual"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Fecha1 = Right$(Vence1.Text, 4) + Mid$(Vence1.Text, 4, 2) + Left$(Vence1.Text, 2)
    Fecha2 = Right$(Vence2.Text, 4) + Mid$(Vence2.Text, 4, 2) + Left$(Vence2.Text, 2)
    Fecha3 = Right$(Vence3.Text, 4) + Mid$(Vence3.Text, 4, 2) + Left$(Vence3.Text, 2)
    Fecha4 = Right$(Vence4.Text, 4) + Mid$(Vence4.Text, 4, 2) + Left$(Vence4.Text, 2)
    Fecha5 = Right$(Vence5.Text, 4) + Mid$(Vence5.Text, 4, 2) + Left$(Vence5.Text, 2)
    Fecha6 = Right$(Vence6.Text, 4) + Mid$(Vence6.Text, 4, 2) + Left$(Vence6.Text, 2)
    
    If Trim(Fecha2) <> "" And Fecha2 <> "00000000" Then
        ZZFechaHasta = Fecha2
            Else
        Fecha2 = "99999999"
    End If
    If Trim(Fecha3) <> "" And Fecha3 <> "00000000" Then
        ZZFechaHasta = Fecha3
            Else
        Fecha3 = "99999999"
    End If
    If Trim(Fecha4) <> "" And Fecha4 <> "00000000" Then
        ZZFechaHasta = Fecha4
            Else
        Fecha4 = "99999999"
    End If
    If Trim(Fecha5) <> "" And Fecha5 <> "00000000" Then
        ZZFechaHasta = Fecha5
            Else
        Fecha5 = "99999999"
    End If
    If Trim(Fecha6) <> "" And Fecha6 <> "00000000" Then
        ZZFechaHasta = Fecha6
            Else
        Fecha6 = "99999999"
    End If
    
    Pasa = 0
    Lugar = 0
    Impo1 = 0
    Impo2 = 0
    Impo3 = 0
    Impo4 = 0
    Impo5 = 0
    Impo6 = 0
    Erase WTrabajo

    ZSql = ""
    ZSql = ZSql + "Select *, Articulo.Sector as [WSector]"
    ZSql = ZSql + " FROM Pedido, Articulo "
    ZSql = ZSql + " Where Pedido.Articulo = Articulo.Codigo"
    ZSql = ZSql + " and Pedido.OrdFecha >= " + "'" + Fecha1 + "'"
    ZSql = ZSql + " and Pedido.OrdFecha <= " + "'" + ZZFechaHasta + "'"
    ZSql = ZSql + " Order by Pedido.Clave"
        
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZPasa = "N"
                    Select Case Tipo.ListIndex
                        Case 0
                            ZPasa = "S"
                        Case 1
                            If Val(rstPedido!WSector) <> 9 Then
                                ZPasa = "S"
                            End If
                        Case 2
                            If Val(rstPedido!WSector) = 9 Then
                                ZPasa = "S"
                            End If
                        Case Else
                            If Val(rstPedido!WSector) = Val(Sector.Text) Then
                                ZPasa = "S"
                            End If
                    End Select
                
                
                    If ZPasa = "S" Then
                    
                        Impo1 = 0
                        Impo2 = 0
                        Impo3 = 0
                        Impo4 = 0
                        
                        ZZArticulo = rstPedido!articulo
                        ZZCliente = rstPedido!Cliente
                        ZZCantidad = rstPedido!Cantidad
                        ZZFacturado = rstPedido!facturado
                        ZZAjuste = rstPedido!Ajuste
                        ZZfecha = rstPedido!Fecha
                        ZZOrdFecha = rstPedido!OrdFecha
                        ZZPedido = rstPedido!Numero
                        ZZPrecio = rstPedido!Precio
                        WWLinea = rstPedido!Linea
                        WWTipo = rstPedido!Tipo
                        WWFragancia = rstPedido!Fragancia
                        WWCalidad = rstPedido!Calidad
                        WWTamano = rstPedido!Tamano
                        WWSector = rstPedido!WSector
                        
                        
                        Select Case TipoII.ListIndex
                            Case 0
                                ZZImpo = ZZCantidad
                            Case 1
                                ZZImpo = ZZFacturado
                            Case Else
                                ZZImpo = ZZCantidad - ZZFacturado - ZZAjuste
                                If ZZImpo < 0 Then
                                    ZZImpo = 0
                                End If
                        End Select
                        
                        If ZZImpo > 0 Then
                            
                            Lugar = Lugar + 1
                            
                            WTrabajo(Lugar, 1) = ZZArticulo
                            WTrabajo(Lugar, 2) = ZZCliente
                            WTrabajo(Lugar, 3) = Str$(ZZImpo)
                            WTrabajo(Lugar, 4) = ZZfecha
                            WTrabajo(Lugar, 5) = ZZPedido
                            WTrabajo(Lugar, 6) = WWLinea
                            WTrabajo(Lugar, 7) = WWTipo
                            WTrabajo(Lugar, 8) = WWFragancia
                            WTrabajo(Lugar, 9) = WWCalidad
                            WTrabajo(Lugar, 10) = WWTamano
                            WTrabajo(Lugar, 11) = ZZOrdFecha
                            WTrabajo(Lugar, 12) = Str$(ZZPrecio)
                            WTrabajo(Lugar, 13) = WWSector
                        
                        End If
                        
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstPedido.Close
    End If
    
    
    For Ciclo = 1 To Lugar
    
                    
        ZZArticulo = WTrabajo(Ciclo, 1)
        ZZCliente = WTrabajo(Ciclo, 2)
        ZZCantidad = Val(WTrabajo(Ciclo, 3))
        ZZfecha = WTrabajo(Ciclo, 4)
        ZZPedido = WTrabajo(Ciclo, 5)
        WWLinea = WTrabajo(Ciclo, 6)
        WWTipo = WTrabajo(Ciclo, 7)
        WWFragancia = WTrabajo(Ciclo, 8)
        WWCalidad = WTrabajo(Ciclo, 9)
        WWTamano = WTrabajo(Ciclo, 10)
        WWOrdFecha = WTrabajo(Ciclo, 11)
        ZZPrecio = Val(WTrabajo(Ciclo, 12))
        WWSector = WTrabajo(Ciclo, 13)
                
        If ZZPrecio = 0 Then
            Call Calcula_Costo
        End If
            
        ZImpo1 = 0
        ZImpo2 = 0
        ZImpo3 = 0
        ZImpo4 = 0
        ZImpo5 = 0
        If WWOrdFecha >= Fecha1 And WWOrdFecha <= Fecha2 Then
            ZImpo1 = ZZPrecio * ZZCantidad
                Else
            If WWOrdFecha > Fecha2 And WWOrdFecha <= Fecha3 Then
                ZImpo2 = ZZPrecio * ZZCantidad
                    Else
                If WWOrdFecha > Fecha3 And WWOrdFecha <= Fecha4 Then
                    ZImpo3 = ZZPrecio * ZZCantidad
                        Else
                    If WWOrdFecha > Fecha4 And WWOrdFecha <= Fecha5 Then
                        ZImpo4 = ZZPrecio * ZZCantidad
                            Else
                        If WWOrdFecha > Fecha5 Then
                            ZImpo5 = ZZPrecio * ZZCantidad
                        End If
                    End If
                End If
            End If
        End If
                
                
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZZCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZZRazon = rstCliente!Razon
            rstCliente.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Sector"
        ZSql = ZSql + " Where Sector.Codigo = " + "'" + WWSector + "'"
        spSector = ZSql
        Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        If rstSector.RecordCount > 0 Then
            WWDesSector = rstSector!Descripcion
            rstSector.Close
        End If
        
        ZImpo6 = ZImpo1 + ZImpo2 + ZImpo3 + ZImpo4 + ZImpo5
                    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO AnalizaPedido ("
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Pedido ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Impo1 ,"
        ZSql = ZSql + "Impo2 ,"
        ZSql = ZSql + "Impo3 ,"
        ZSql = ZSql + "Impo4 ,"
        ZSql = ZSql + "Impo5 ,"
        ZSql = ZSql + "Impo6 ,"
        ZSql = ZSql + "Impre1 ,"
        ZSql = ZSql + "Impre2 ,"
        ZSql = ZSql + "Impre3 ,"
        ZSql = ZSql + "Impre4 ,"
        ZSql = ZSql + "Impre5 ,"
        ZSql = ZSql + "Impre6 ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "Sector ,"
        ZSql = ZSql + "DesSector )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZPedido + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + Str$(ZImpo1) + "',"
        ZSql = ZSql + "'" + Str$(ZImpo2) + "',"
        ZSql = ZSql + "'" + Str$(ZImpo3) + "',"
        ZSql = ZSql + "'" + Str$(ZImpo4) + "',"
        ZSql = ZSql + "'" + Str$(ZImpo5) + "',"
        ZSql = ZSql + "'" + Str$(ZImpo6) + "',"
        ZSql = ZSql + "'" + Vence2.Text + "',"
        ZSql = ZSql + "'" + Vence3.Text + "',"
        ZSql = ZSql + "'" + Vence4.Text + "',"
        ZSql = ZSql + "'" + Vence5.Text + "',"
        ZSql = ZSql + "'" + Vence1.Text + "',"
        ZSql = ZSql + "'" + Vence6.Text + "',"
        ZSql = ZSql + "'" + ZZRazon + "',"
        ZSql = ZSql + "'" + WWSector + "',"
        ZSql = ZSql + "'" + WWDesSector + "')"
        spAnalizaPedido = ZSql
        Set rstAnalizaPedido = db.OpenRecordset(spAnalizaPedido, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
                    
    
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT AnalizaPedido.Cliente, AnalizaPedido.Pedido, AnalizaPedido.Fecha, AnalizaPedido.Impo1, AnalizaPedido.Impo2, AnalizaPedido.Impo3, AnalizaPedido.Impo4, AnalizaPedido.Impo5, AnalizaPedido.Impo6, AnalizaPedido.Impre1, AnalizaPedido.Impre2, AnalizaPedido.Impre3, AnalizaPedido.Impre4, AnalizaPedido.Impre5, AnalizaPedido.Impre6, AnalizaPedido.Razon " _
            + "From " _
            + DSQ + ".dbo.AnalizaPedido AnalizaPedido " _
            + "Where " _
            + "AnalizaPedido.Pedido >= 0 AND " _
            + "AnalizaPedido.Pedido <= 999999"
    
    Listado.Connect = Connect()
    
    Uno = "{AnalizaPedido.Pedido} in 0 to 999999"
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    If TipoV.Value = 1 Then
        Listado.ReportFileName = "analisispedidosExporta.rpt"
            Else
        Listado.ReportFileName = "analisispedidos.rpt"
    End If
    Listado.Action = 1
    
    Exit Sub
    
WError:
    
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgAnalizaPedidos.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Sector_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Sector"
        ZSql = ZSql + " Where Sector.Codigo = " + "'" + Sector.Text + "'"
        spSector = ZSql
        Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        If rstSector.RecordCount > 0 Then
            DesSector.Caption = rstSector!Descripcion
                Else
            DesSector.Caption = ""
            Sector.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Sector.Text = ""
        DesSector.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            Vence5.SetFocus
                Else
            Vence4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence4.Text = "  /  /    "
    End If
End Sub

Private Sub Vence5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence5.Text, Auxi)
        If Auxi = "S" Then
            Vence6.SetFocus
                Else
            Vence5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence5.Text = "  /  /    "
    End If
End Sub

Private Sub Vence6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence5.Text, Auxi)
        If Auxi = "S" Then
            Vence1.SetFocus
                Else
            Vence6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence6.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Total"
    Tipo.AddItem "Mc"
    Tipo.AddItem "Kiama"
    Tipo.AddItem "Sector"
    
    Tipo.ListIndex = 0

    TipoII.Clear
    
    TipoII.AddItem "Pedidos"
    TipoII.AddItem "Entregados"
    TipoII.AddItem "Pendientes"
    
    TipoII.ListIndex = 0

    Vence1.Text = "  /  /    "
    Vence2.Text = "  /  /    "
    Vence3.Text = "  /  /    "
    Vence4.Text = "  /  /    "
    Vence5.Text = "  /  /    "
    Vence6.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Vence1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence6_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub



Private Sub Calcula_Costo()


    ZZOrdFecha = Right$(ZZfecha, 4) + Mid$(ZZfecha, 4, 2) + Left$(ZZfecha, 2)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZZCliente + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WWNroLista = Str$(rstCliente!NroLista)
        rstCliente.Close
    End If


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ClienteLista"
    ZSql = ZSql + " Where ClienteLista.Cliente = " + "'" + ZZCliente + "'"
    ZSql = ZSql + " and ClienteLista.LInea = " + "'" + WWLinea + "'"
    ZSql = ZSql + " and ClienteLista.Tipo = " + "'" + WWTipo + "'"
        
    spClienteLista = ZSql
    Set rstClientelista = db.OpenRecordset(spClienteLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientelista.RecordCount > 0 Then
        WWNroLista = Str$(rstClientelista!Lista)
        rstClientelista.Close
    End If



    WWTope1 = 0
    WWValor1 = 0
    WWTope2 = 0
    WWValor2 = 0
    WWTope3 = 0
    WWValor3 = 0
    WWTope4 = 0
    WWValor4 = 0
    WWDesde = "00/00/0000"
    WWHasta = "00/00/0000"
    WWOrdDesde = "00000000"
    WWOrdHasta = "00000000"
    WWMoneda = 0


    WWNroLista = Trim(WWNroLista)
    
    ZZLee = "S"

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Lista = " + "'" + WWNroLista + "'"
    ZSql = ZSql + " and Precios.LInea = " + "'" + WWLinea + "'"
    ZSql = ZSql + " and Precios.Tipo = " + "'" + WWTipo + "'"
    ZSql = ZSql + " and Precios.fragancia = " + "'" + WWFragancia + "'"
    ZSql = ZSql + " and Precios.Calidad = " + "'" + WWCalidad + "'"
    ZSql = ZSql + " and Precios.Tamano = " + "'" + WWTamano + "'"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        WWTope1 = rstPrecios!Tope1
        WWValor1 = rstPrecios!Valor1
        WWTope2 = rstPrecios!Tope2
        WWValor2 = rstPrecios!Valor2
        WWTope3 = rstPrecios!Tope3
        WWValor3 = rstPrecios!Valor3
        WWTope4 = rstPrecios!Tope4
        WWValor4 = rstPrecios!Valor4
        WWDesde = rstPrecios!Desde
        WWHasta = rstPrecios!Hasta
        WWOrdDesde = rstPrecios!OrdDesde
        WWOrdHasta = rstPrecios!OrdHasta
        WWMoneda = rstPrecios!Moneda
        rstPrecios.Close
        
        ZZLee = "N"
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.LInea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and Articulo.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and Articulo.fragancia = " + "'" + WWFragancia + "'"
        ZSql = ZSql + " and Articulo.Calidad = " + "'" + WWCalidad + "'"
        ZSql = ZSql + " and Articulo.Tamano = " + "'" + WWTamano + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZZActivo = rstArticulo!Activo
            rstArticulo.Close
        End If
        
        If (WWValor1 = 0 And WWValor2 = 0 And WWValor3 = 0 And WWValor4 = 0) Or ZZActivo = 1 Then
            ZZLee = "S"
        End If
        
        If ZZOrdFecha > WWOrdHasta Then
            ZZLee = "S"
        End If
                
    End If
    
    If ZZLee = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + WWNroLista + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + WWCalidad + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + WWTamano + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
        
            WWTope1 = rstPrecios!Tope1
            WWValor1 = rstPrecios!Valor1
            WWTope2 = rstPrecios!Tope2
            WWValor2 = rstPrecios!Valor2
            WWTope3 = rstPrecios!Tope3
            WWValor3 = rstPrecios!Valor3
            WWTope4 = rstPrecios!Tope4
            WWValor4 = rstPrecios!Valor4
            WWDesde = rstPrecios!Desde
            WWHasta = rstPrecios!Hasta
            WWOrdDesde = rstPrecios!OrdDesde
            WWOrdHasta = rstPrecios!OrdHasta
            WWMoneda = rstPrecios!Moneda
            rstPrecios.Close
            
            ZZLee = "N"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.LInea = " + "'" + WWLinea + "'"
            ZSql = ZSql + " and Articulo.Tipo = " + "'" + WWTipo + "'"
            ZSql = ZSql + " and Articulo.fragancia = " + "'" + WWFragancia + "'"
            ZSql = ZSql + " and Articulo.Calidad = " + "'" + WWCalidad + "'"
            ZSql = ZSql + " and Articulo.Tamano = " + "'" + WWTamano + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZActivo = rstArticulo!Activo
                rstArticulo.Close
            End If
            
            If (WWValor1 = 0 And WWValor2 = 0 And WWValor3 = 0 And WWValor4 = 0) Or ZZActivo = 1 Then
                ZZLee = "S"
            End If
            
            If ZZOrdFecha > WWOrdHasta Then
                ZZLee = "S"
            End If
            
        End If
    End If
                
    If ZZLee = "S" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + WWNroLista + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + "" + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + WWTamano + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
        
            WWTope1 = rstPrecios!Tope1
            WWValor1 = rstPrecios!Valor1
            WWTope2 = rstPrecios!Tope2
            WWValor2 = rstPrecios!Valor2
            WWTope3 = rstPrecios!Tope3
            WWValor3 = rstPrecios!Valor3
            WWTope4 = rstPrecios!Tope4
            WWValor4 = rstPrecios!Valor4
            WWDesde = rstPrecios!Desde
            WWHasta = rstPrecios!Hasta
            WWOrdDesde = rstPrecios!OrdDesde
            WWOrdHasta = rstPrecios!OrdHasta
            WWMoneda = rstPrecios!Moneda
            rstPrecios.Close
            
            ZZLee = "N"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.LInea = " + "'" + WWLinea + "'"
            ZSql = ZSql + " and Articulo.Tipo = " + "'" + WWTipo + "'"
            ZSql = ZSql + " and Articulo.fragancia = " + "'" + WWFragancia + "'"
            ZSql = ZSql + " and Articulo.Calidad = " + "'" + WWCalidad + "'"
            ZSql = ZSql + " and Articulo.Tamano = " + "'" + WWTamano + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZZActivo = rstArticulo!Activo
                rstArticulo.Close
            End If
            
            If (WWValor1 = 0 And WWValor2 = 0 And WWValor3 = 0 And WWValor4 = 0) Or ZZActivo = 1 Then
                ZZLee = "S"
            End If
            
            If ZZOrdFecha > WWOrdHasta Then
                ZZLee = "S"
            End If
            
        End If
    End If
                

    
    If WWCanti < WWTope1 Then
        WWPrecio = WWValor1
            Else
        If WWCanti < WWTope2 Then
            WWPrecio = WWValor2
                Else
            If WWCanti < WWTope3 Then
                WWPrecio = WWValor3
                    Else
                WWPrecio = WWValor4
            End If
        End If
    End If


    ZZEntra = "N"

    WWWWTope1 = 0
    WWWWValor1 = 0
    WWWWTope2 = 0
    WWWWValor2 = 0
    WWWWTope3 = 0
    WWWWValor3 = 0
    WWWWTope4 = 0
    WWWWValor4 = 0
    WWWWDesde = 0
    WWWWHasta = 0
    WWWWOrdDesde = "000000000"
    WWWWOrdHasta = "000000000"
    
    WWDto = 0



    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ClienteBonifica"
    ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + ZZCliente + "'"
    ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
    ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
    ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + WWFragancia + "'"
    ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + WWCalidad + "'"
    ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + WWTamano + "'"
    ZSql = ZSql + " Order by ClienteBonifica.orddesde"
    spClienteBonifica = ZSql
    Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstClienteBonifica.RecordCount > 0 Then
        With rstClienteBonifica
            .MoveLast
            ZZComparaI = rstClienteBonifica!OrdHasta
            ZZComparaII = rstClienteBonifica!OrdHasta
            Rem If ZZComparaI <> "" And ZZComparaII = "" Then
            Rem     ZZComparaII = "20991231"
            Rem End If
            If ZZComparaII > ZZOrdFecha Then
                ZZEntra = "S"
                WWWWTope1 = rstClienteBonifica!Tope1
                WWWWValor1 = rstClienteBonifica!Valor1
                WWWWTope2 = rstClienteBonifica!Tope2
                WWWWValor2 = rstClienteBonifica!Valor2
                WWWWTope3 = rstClienteBonifica!Tope3
                WWWWValor3 = rstClienteBonifica!Valor3
                WWWWTope4 = rstClienteBonifica!Tope4
                WWWWValor4 = rstClienteBonifica!Valor4
                WWWWDesde = rstClienteBonifica!Desde
                WWWWHasta = rstClienteBonifica!Hasta
                WWWWOrdDesde = rstClienteBonifica!OrdDesde
                WWWWOrdHasta = rstClienteBonifica!OrdHasta
            End If
            rstClienteBonifica.Close
        End With
        
    End If

    
    If ZZEntra = "N" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + ZZCliente + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + WWCalidad + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + WWTamano + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
        
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + ZZCliente + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + WWTamano + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
        
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + ZZCliente + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + WWTipo + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
        
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + ZZCliente + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + WWLinea + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + "" + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
            
        End If
    End If
    
    If ZZEntra = "N" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteBonifica"
        ZSql = ZSql + " Where ClienteBonifica.Cliente = " + "'" + ZZCliente + "'"
        ZSql = ZSql + " and ClienteBonifica.Linea = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Tipo = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Fragancia = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.Calidad = " + "'" + "" + "'"
        ZSql = ZSql + " and ClienteBonifica.TAmano = " + "'" + "" + "'"
        ZSql = ZSql + " Order by ClienteBonifica.orddesde"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteBonifica.RecordCount > 0 Then
            With rstClienteBonifica
                .MoveLast
                ZZComparaI = rstClienteBonifica!OrdHasta
                ZZComparaII = rstClienteBonifica!OrdHasta
                Rem If ZZComparaI <> "" And ZZComparaII = "" Then
                Rem     ZZComparaII = "20991231"
                Rem End If
                If ZZComparaII > ZZOrdFecha Then
                    ZZEntra = "S"
                    WWWWTope1 = rstClienteBonifica!Tope1
                    WWWWValor1 = rstClienteBonifica!Valor1
                    WWWWTope2 = rstClienteBonifica!Tope2
                    WWWWValor2 = rstClienteBonifica!Valor2
                    WWWWTope3 = rstClienteBonifica!Tope3
                    WWWWValor3 = rstClienteBonifica!Valor3
                    WWWWTope4 = rstClienteBonifica!Tope4
                    WWWWValor4 = rstClienteBonifica!Valor4
                    WWWWDesde = rstClienteBonifica!Desde
                    WWWWHasta = rstClienteBonifica!Hasta
                    WWWWOrdDesde = rstClienteBonifica!OrdDesde
                    WWWWOrdHasta = rstClienteBonifica!OrdHasta
                End If
                rstClienteBonifica.Close
            End With
            
        End If
    End If


    If ZZOrdFecha > WWOrdDesde Or ZZOrdFecha < WWOrdHasta Then
        
        If WWCanti < WWWWTope1 Then
            WWDto = WWWWValor1
                Else
            If WWCanti < WWWWTope2 Then
                WWDto = WWWWValor2
                    Else
                If WWDto < WWWWTope3 Then
                    WWDto = WWWWValor3
                        Else
                    WWDto = WWWWValor4
                End If
            End If
        End If
        
    End If

    If WWDto <> 0 Then
        WWPrecio = WWValor1
        If WWValor2 > WWPrecio Then
            WWPrecio = WWValor2
        End If
        If WWValor3 > WWPrecio Then
            WWPrecio = WWValor3
        End If
        If WWValor4 > WWPrecio Then
            WWPrecio = WWValor4
        End If
    End If

    WImpoDto = 0
    WDescuento = WWDto
    If WDescuento <> 0 Then
        WImpoDto = WWPrecio * WDescuento / 100
        Call Redondeo(WImpoDto)
        WWPrecio = WWPrecio - WImpoDto
    End If

    ZZPrecio = WWPrecio

    If WWMoneda = 1 Then
        ZZPrecio = ZZPrecio * WWParidad
    End If

End Sub

