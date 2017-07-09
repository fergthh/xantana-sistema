VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaComisiones 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Comisiones por Ventas"
   ClientHeight    =   7755
   ClientLeft      =   1935
   ClientTop       =   750
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   7755
   ScaleWidth      =   8145
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
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   4935
      Begin VB.TextBox Comision 
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
         MaxLength       =   4
         TabIndex        =   17
         Text            =   " "
         Top             =   840
         Width           =   1215
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
         Left            =   2280
         TabIndex        =   15
         Top             =   2640
         Width           =   2175
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
         Left            =   3840
         MouseIcon       =   "ListaComisiones.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaComisiones.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salida"
         Top             =   3240
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
         Left            =   1440
         MouseIcon       =   "ListaComisiones.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaComisiones.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Impresion x Impresora"
         Top             =   3240
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
         Left            =   2640
         MouseIcon       =   "ListaComisiones.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaComisiones.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consulta de Datos"
         Top             =   3240
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
         Left            =   240
         MouseIcon       =   "ListaComisiones.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ListaComisiones.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   3240
         Width           =   855
      End
      Begin MSMask.MaskEdBox HastaFec 
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox DesdeFec 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.TextBox Desde 
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
         MaxLength       =   4
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Comision"
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
         Left            =   720
         TabIndex        =   16
         Top             =   960
         Width           =   1455
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
         Left            =   720
         TabIndex        =   14
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Vendedor"
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
         Left            =   720
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Listcomi.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
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
      Left            =   6960
      TabIndex        =   2
      Top             =   1200
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
      Height          =   2400
      ItemData        =   "ListaComisiones.frx":2D30
      Left            =   120
      List            =   "ListaComisiones.frx":2D37
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListaComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPlanilla(10000, 7) As String
Dim ZComision As Double
Dim ZImporte As Double
Dim ZTipoComision As Integer
Dim ZDescuento As Integer

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
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WTitulo = "del " + DesdeFec.Text + " al " + HastaFec.Text
    If ZZNivel = 0 Then
        WTituloII = "       4.8.5. Listado de Comisiones       "
            Else
        WTituloII = "*****  4.8.5. Listado de Comisiones ******"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Recibos SET "
    ZSql = ZSql + " Lista = " + "'" + "" + "',"
    ZSql = ZSql + " TituloII = " + "'" + WTituloII + "',"
    ZSql = ZSql + " Titulo = " + "'" + WTitulo + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Configuracion"
    ZSql = ZSql + " Where Configuracion.Clave = 1"
    spConfiguracion = ZSql
    Set rstConfiguracion = db.OpenRecordset(spConfiguracion, dbOpenSnapshot, dbSQLPassThrough)
    If rstConfiguracion.RecordCount > 0 Then
        ConfigIva1 = rstConfiguracion!Iva1
        ConfigIva2 = rstConfiguracion!Iva2
        ConfigPercepcion = rstConfiguracion!Percepcion
        ConfigPunto = rstConfiguracion!Punto
        CantiFac = rstConfiguracion!CantiFac
        CantiRem = rstConfiguracion!CantiRem
        CantiArti = rstConfiguracion!CantiArti
        rstConfiguracion.Close
    End If
    
    
    ZLugar = 0
    Erase WPlanilla
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.FechaOrd <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Recibos.Importe1 <> 0"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    WPlanilla(ZLugar, 1) = rstRecibos!Tipo1
                    WPlanilla(ZLugar, 2) = rstRecibos!Letra1
                    WPlanilla(ZLugar, 3) = rstRecibos!Punto1
                    WPlanilla(ZLugar, 4) = rstRecibos!Numero1
                    WPlanilla(ZLugar, 5) = rstRecibos!Cliente
                    WPlanilla(ZLugar, 6) = rstRecibos!Clave
                    WPlanilla(ZLugar, 7) = rstRecibos!Importe1
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZTipo = WPlanilla(Ciclo, 1)
        ZLetra = WPlanilla(Ciclo, 2)
        ZPunto = WPlanilla(Ciclo, 3)
        ZNumero = WPlanilla(Ciclo, 4)
        ZCliente = WPlanilla(Ciclo, 5)
        ZClave = WPlanilla(Ciclo, 6)
        ZImporte = Val(WPlanilla(Ciclo, 7))
        
        ZVendedor = 0
        ZTipoComision = 0
        ZPartida = ""
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZVendedor = rstCliente!Vendedor
            rstCliente.Close
        End If
    
        If Val(Desde.Text) = ZVendedor Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Tipo = " + "'" + ZTipo + "'"
            ZSql = ZSql + " and CtaCte.Letra = " + "'" + ZLetra + "'"
            ZSql = ZSql + " and CtaCte.Punto = " + "'" + ZPunto + "'"
            ZSql = ZSql + " and CtaCte.Numero = " + "'" + ZNumero + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                ZTipoComision = IIf(IsNull(rstCtaCte!Comision), "0", rstCtaCte!Comision)
                ZDescuento = IIf(IsNull(rstCtaCte!Descuento), "0", rstCtaCte!Descuento)
                ZPartida = rstCtaCte!Partida
                ZTipoIva = rstCtaCte!TipoIva
                rstCtaCte.Close
            End If
            
            If ZTipoComision = 0 Then
                If ZZNivel = 0 Then
                    ZPorce = Val(Comision.Text)
                        Else
                    ZPorce = Val(Comision.Text)
                End If
                If ZVendedor = 9 Then
                    ZPorce = 10
                End If
                    Else
                If ZZNivel = 0 Then
                    ZPorce = Val(Comision.Text) / 2
                        Else
                    ZPorce = Val(Comision.Text) / 2
                End If
                If ZVendedor = 9 Then
                    ZPorce = 5
                End If
            End If
                    
            If ZTipoIva = 0 Then
                ZImporte = ZImporte / (1 + (ConfigIva1) / 100)
                    Else
                ZImporte = ZImporte / (1 + (ConfigIva2) / 100)
            End If
            Call Redondeo(ZImporte)
                    
            ZNeto = ZImporte
            
            If ZZNivel > 0 Then
                Select Case Trim(UCase(ZPartida))
                    Case "S"
                        ZNeto = 0
                    Case "/"
                        ZNeto = ZImporte
                    Case "?"
                        ZNeto = ZImporte * 11
                    Case Else
                        ZNeto = 0
                End Select
            End If
            
            If ZNeto > 0 Then
            
                ZComision = ZNeto * (ZPorce / 100)
                Call Redondeo(ZComision)
                Auxi = Str$(ZComision)
                Auxi = Pusing("###,###.##", Auxi)
                
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Recibos SET "
                ZSql = ZSql + " Vendedor = " + "'" + Str$(ZVendedor) + "',"
                ZSql = ZSql + " Lista = " + "'" + "S" + "',"
                ZSql = ZSql + " PorceDto = " + "'" + Str$(ZDescuento) + "',"
                ZSql = ZSql + " TipoComi = " + "'" + Str$(ZTipoComision) + "',"
                ZSql = ZSql + " Neto = " + "'" + Str$(ZNeto) + "',"
                ZSql = ZSql + " Porce = " + "'" + Str$(ZPorce) + "',"
                ZSql = ZSql + " Comision = " + "'" + Auxi + "'"
                ZSql = ZSql + " Where Recibos.Clave = " + "'" + ZClave + "'"
                spRecibos = ZSql
                Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
            
        End If
        
    Next Ciclo
    
    Listado.WindowTitle = "Listado de Comisiones por Venta"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Val(Desde.Text) = 0 Then
        Desde.Text = "0"
    End If

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Renglon, Recibos.Cliente, Recibos.Fecha, Recibos.Fechaord, Recibos.Tipo1, Recibos.Numero1, Recibos.Vendedor, Recibos.Lista, Recibos.Neto, Recibos.Porce, Recibos.Comision, Recibos.Titulo, Recibos.TituloII, Recibos.PorceDto, Recibos.TipoComi, " _
            + "Cliente.Razon, " _
            + "Vendedor.Nombre " _
            + "From " _
            + DSQ + ".dbo.Recibos Recibos, " _
            + DSQ + ".dbo.Cliente Cliente, " _
            + DSQ + ".dbo.Vendedor Vendedor " _
            + "Where " _
            + "Recibos.Cliente = Cliente.Cliente AND " _
            + "Cliente.Vendedor = Vendedor.Codigo AND " _
            + "Recibos.Fechaord >= '" + WDesde + "' AND " _
            + "Recibos.Fechaord <= '" + WHasta + "' AND " _
            + "Recibos.Lista = 'S'"
    
    Uno = "{Recibos.FechaOrd} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "ListaComisionesCobra.rpt"
            Else
        Listado.ReportFileName = "ListaComisionesCobraII.rpt"
    End If
    
    Listado.Action = 1
    
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaComisiones.Hide
    Unload Me
    Menu41.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comision.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Comision_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFec.SetFocus
    End If
    If KeyAscii = 27 Then
        Comision.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFec.Text = "  /  /    "
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0

    Desde.Text = ""
    Comision.Text = ""
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Ayuda.Visible = True
    Ayuda.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Vendedor"
    ZSql = ZSql + " Order by Vendedor.Codigo"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(!Codigo) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Codigo
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Indice = Pantalla.ListIndex
    DesdeVend.Text = WIndice.List(Indice)
    HastaVend.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    DesdeVend.SetFocus
    
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
    
    XIndice = 0
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Vendedor.Codigo"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstVendedor!Codigo) + " " + rstVendedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstVendedor!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

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

Private Sub DesdeFec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFec_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call Cancela_click
        Case Else
    End Select
End Sub













