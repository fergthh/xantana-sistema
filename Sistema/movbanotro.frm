VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgMovbanOtro 
   Caption         =   "Listado de Movimientos de Bancos"
   ClientHeight    =   6645
   ClientLeft      =   2280
   ClientTop       =   825
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   6645
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
      Top             =   3240
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
      Top             =   3600
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6000
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   4695
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   9
         Text            =   " "
         Top             =   2160
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
         Left            =   2160
         MaxLength       =   4
         TabIndex        =   8
         Text            =   " "
         Top             =   1800
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   960
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
         Left            =   1800
         TabIndex        =   0
         Top             =   600
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
      Begin VB.Image Panta 
         Height          =   480
         Left            =   3720
         MouseIcon       =   "movbanotro.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "movbanotro.frx":030A
         ToolTipText     =   "Emision por Pantalla"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Consulta 
         Height          =   480
         Left            =   3720
         MouseIcon       =   "movbanotro.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "movbanotro.frx":0E56
         ToolTipText     =   "Consulta de Datos"
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image Impre 
         Height          =   480
         Left            =   3720
         MouseIcon       =   "movbanotro.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "movbanotro.frx":19A2
         ToolTipText     =   "Emision por Impresora"
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   3720
         MouseIcon       =   "movbanotro.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "movbanotro.frx":24EE
         ToolTipText     =   "Menu Principal"
         Top             =   2160
         Width           =   480
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
         Top             =   2160
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
         Top             =   1800
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
         Top             =   960
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
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.PictureBox Listado 
      Height          =   480
      Left            =   6120
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   12
      Top             =   1320
      Width           =   1200
   End
End
Attribute VB_Name = "PrgMovbanOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WInicial() As Variant ' Matriz de 2 dimensiones que contiene registros
Dim ZBanco(100, 2) As String

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
    
    If Val(DesdeBanco.Text) = 0 Then
        DesdeBanco.Text = "0"
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

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Actividad = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
            .Update
        End If
    End With
    
    Lugarbanco = 0
    
    With rstBanco
        .Index = "Banco"
        .MoveFirst
        Do
            If .EOF = False Then
                Lugarbanco = Lugarbanco + 1
                ZBanco(Lugarbanco, 1) = !Banco
                ZBanco(Lugarbanco, 2) = !Cuenta
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    da = 0
    With rstMovban
        .Index = "Clave"
        .Seek "=", da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
            
    With rstPagos
        .Index = "Clave"
        .MoveFirst
        Do
            If Val(!FechaOrd2) = 0 Then
                WFechaord2 = !fechaord
                WFecha2 = !Fecha
                    Else
                WFechaord2 = !FechaOrd2
                WFecha2 = !Fecha2
            End If
            If WDesde <= !fechaord And !fechaord <= WHasta Then
                If Val(!Tipo2) = 2 Then
                    If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                        WBanco = !Banco2
                        WOrden = !Orden
                        WFecha = !Fecha
                        WFechaord = !fechaord
                        WAcredita = WFecha2
                        WAcreditaOrd = WFechaord2
                        WObservaciones = ""
                        WObservaciones = !Observaciones
                        WNumero = !Numero2
                        WImporte = !Importe2
                        WOrden = !Orden
                
                        With rstMovban
                            .AddNew
                            !da = 0
                            !Banco = WBanco
                            !Fecha = WFecha
                            !fechaord = WFechaord
                            !Acredita = WAcredita
                            !AcreditaOrd = WAcreditaOrd
                            !Observaciones = Left$(WObservaciones, 30)
                            !Numero = WNumero
                            !Debito = 0
                            !Credito = WImporte
                            !Comprobante = WOrden
                            !Empresa = 1
                            .Update
                        End With
                    End If
                End If
                    
                If Val(!Tiporeg) = 1 And Val(!Banco2) <> 0 Then
                    If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                        WBanco = !Banco2
                        WOrden = !Orden
                        WFecha = !Fecha
                        WFechaord = !fechaord
                        WAcredita = !Fecha
                        WAcreditaOrd = !fechaord
                        WObservaciones = ""
                        WObservaciones = !Observaciones
                        WNumero = ""
                        WImporte = !Importe1
                        WOrden = !Orden
                
                        With rstMovban
                            .AddNew
                            !da = 0
                            !Banco = WBanco
                            !Fecha = WFecha
                            !fechaord = WFechaord
                            !Acredita = WAcredita
                            !AcreditaOrd = WAcreditaOrd
                            !Observaciones = WObservaciones
                            !Numero = WNumero
                            !Debito = WImporte
                            !Credito = 0
                            !Comprobante = WOrden
                            !Empresa = 1
                            .Update
                        End With
                    End If
                End If
                    
            End If
                
            If WDesde > !fechaord Then
                If Val(!Tipo2) = 2 Then
                    If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                        WInicial(!Banco2) = WInicial(!Banco2) - !Importe2
                    End If
                End If
                If Val(!Tiporeg) = 1 And !Banco2 <> 0 Then
                    If Val(!Banco2) >= Val(DesdeBanco.Text) And Val(!Banco2) <= Val(HastaBanco.Text) Then
                        WInicial(!Banco2) = WInicial(!Banco2) + !Importe1
                    End If
                End If
            End If
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
            
    With rstDepositos
        .Index = "Clave"
        .MoveFirst
        Do
            If WDesde <= !fechaord And !fechaord <= WHasta Then
                If Val(!Banco) >= Val(DesdeBanco.Text) And Val(!Banco) <= Val(HastaBanco.Text) Then
                    If Val(!Renglon) = 1 Then
                        WBanco = !Banco
                        WFecha = !Fecha
                        WFechaord = !fechaord
                        WAcredita = !Acredita
                        WAcreditaOrd = !AcreditaOrd
                        WObservaciones = "Deposito"
                        WNumero = !Deposito
                        WImporte = !Importe
                        WDeposito = !Deposito
                    
                        With rstMovban
                            .AddNew
                            !Banco = WBanco
                            !Fecha = WFecha
                            !fechaord = WFechaord
                            !Acredita = WAcredita
                            !AcreditaOrd = WAcreditaOrd
                            !Observaciones = WObservaciones
                            !Numero = WNumero
                            !Credito = 0
                            !Debito = WImporte
                            !Comprobante = WDeposito
                            !Empresa = 1
                            .Update
                        End With
                    End If
                End If
            End If
                
            If WDesde > !fechaord Then
                If Val(!Banco) >= Val(DesdeBanco.Text) And Val(!Banco) <= Val(HastaBanco.Text) Then
                    If Val(!Renglon) = 1 Then
                        WInicial(!Banco) = WInicial(!Banco) + !Importe
                    End If
                End If
            End If
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    With rstRecibos
        .Index = "Clave"
        .MoveFirst
        Do
            If Val(!Tiporeg) = 2 And Val(!Tipo2) = 4 Then
            
                If WDesde <= !fechaord And !fechaord <= WHasta Then
                
                    For Ciclo = 1 To Lugarbanco
                    
                        If !Cuenta = ZBanco(Ciclo, 2) Then
                    
                            WClave = !Clave
                            WRecibo = !Recibo
                            WRenglon = !Renglon
                            WFecha = !Fecha
                            WFechaord = !fechaord
                            If !TipoRec = "3" Then
                                WObservaciones = !Observaciones
                                    Else
                                WCliente = !Cliente
                                With rstClientes
                                    .Index = "Cliente"
                                    .Seek "=", WCliente
                                    If .NoMatch = False Then
                                        WObservaciones = !Razon
                                    End If
                                End With
                            End If
                            
                            WImpre = ""
                            WCuenta = !Cuenta
                                    
                            WLetra = ""
                            WTipo = !Tipo2
                            WPunto = 0
                            WNumero = !Numero2
                            WImporte = !Importe2
                            
                            With rstMovban
                                .AddNew
                                !Banco = ZBanco(Lugarbanco, 1)
                                !Fecha = WFecha
                                !fechaord = WFechaord
                                !Acredita = WFecha
                                !AcreditaOrd = WFechaord
                                !Observaciones = WObservaciones
                                !Numero = WRecibo
                                !Credito = 0
                                !Debito = WImporte
                                !Comprobante = WRecibo
                                !Empresa = 1
                                .Update
                            End With
                            
                        End If
                        
                    Next Ciclo
                
                End If
                
                If WDesde > !fechaord Then
                
                    For Ciclo = 1 To Lugarbanco
                        If !Cuenta = ZBanco(Ciclo, 2) Then
                            WImporte = !Importe2
                            WInicial(Val(ZBanco(Ciclo, 1))) = WInicial(Val(ZBanco(Ciclo, 1))) + WImporte
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
    
    For XDa = 1 To 100
    
        If WInicial(XDa) <> 0 Then
    
            With rstMovban
                .AddNew
                !Banco = XDa
                !Fecha = "00/00/0000"
                !fechaord = "00000000"
                !Acredita = "00/00/0000"
                !AcreditaOrd = "00000000"
                !Observaciones = "Saldo Inicial"
                !Numero = 0
                If WInicial(XDa) > 0 Then
                    !Credito = 0
                    !Debito = WInicial(XDa)
                        Else
                    !Credito = Abs(WInicial(XDa))
                    !Debito = 0
                End If
                !Comprobante = "000000"
                !Empresa = 1
                .Update
            End With
        End If
        
    Next XDa

    Listado.GroupSelectionFormula = "{Movban.banco} in " + DesdeBanco.Text + " to " + HastaBanco.Text
    Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    WDestino = Listado.Destination
    Listado.PrintFileName = "dada.txt"
    Listado.Destination = 2
    Listado.Action = 1
    Listado.Destination = WDestino
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstBanco
        .Close
    End With
    With rstDepositos
        .Close
    End With
    With RstProveedor
        .Close
    End With
    With rstPagos
        .Close
    End With
    With rstMovban
        .Close
    End With
    With rstRecibos
        .Close
    End With
    
    DbsAdminis.Close
    
    Desde.SetFocus
    PrgMovban.Hide
    Unload Me
    Menu.SetFocus
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

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Banco
    OPEN_FILE_Depositos
    OPEN_FILE_Proveedor
    OPEN_FILE_Recibos
    OPEN_FILE_Pagos
    OPEN_FILE_Movban
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            DesdeBanco.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub DesdeBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaBanco.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    With rstBanco
        .Index = "Banco"
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
    With rstBanco
        Indice = Pantalla.ListIndex
        Claveven$ = WIndice.List(Indice)
        DesdeBanco.Text = Claveven$
        .Index = "Banco"
        Claveven$ = DesdeBanco.Text
        .Seek "=", Claveven$
        If .NoMatch = False Then
            DesdeBanco.Text = !Banco
            HastaBanco.Text = !Banco
                Else
            DesdeBanco.Text = Claveven$
            HastaBanco.Text = Claveven$
        End If
    End With
    DesdeBanco.SetFocus
End Sub

Sub Form_Load()
    ReDim WInicial(1 To 100)

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeBanco.Text = ""
    HastaBanco.Text = ""
    Frame2.Visible = True
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    With rstBanco
        .Index = "Banco"
        .MoveFirst
        Do
            If .EOF = False Then
                da = Len(!Nombre) - WEspacios
                For aa = 1 To da + 1
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        IngresaItem = Str$(!Banco) + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Banco
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next aa
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

