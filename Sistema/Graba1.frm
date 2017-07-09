VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGraba1 
   AutoRedraw      =   -1  'True
   Caption         =   "Grabacion de Imputaciones Contables (Compras)"
   ClientHeight    =   2820
   ClientLeft      =   2490
   ClientTop       =   2055
   ClientWidth     =   7020
   LinkTopic       =   "Form2"
   ScaleHeight     =   2820
   ScaleWidth      =   7020
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5640
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2175
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   5055
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   1440
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   1080
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
         Left            =   3720
         TabIndex        =   8
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
         Left            =   3720
         TabIndex        =   3
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   480
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
         Left            =   240
         TabIndex        =   10
         Top             =   480
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
         Left            =   240
         TabIndex        =   7
         Top             =   1440
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
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5640
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgGraba1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstImpu As Recordset
Dim spImpu As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim rstIvacomp As Recordset
Dim spIvacomp As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
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
Dim XProveedor As String
Dim XLetra As String
Dim XTipocomp As String
Dim XPuntocomp As String
Dim XNroComp As String
Dim Compro(100) As String
Dim WProveedor As String
Dim WLetra As String
Dim WTipocomp As String
Dim WPuntocomp As String
Dim WNroComp As String

Private Sub Acepta_Click()

    Compro(1) = "FC"
    Compro(2) = "ND"
    Compro(3) = "NC"

    With rstEmpre
        .Index = "Codigo"
        .Seek "=", 1
        If .NoMatch = False Then
            WDesdefecha = !Desdefecha
            WHastafecha = !Hastafecha
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
    WHasta = WAno + WMes + WDia
    
    WTipo = 2
    Erase Vector
    Lugar = 0
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
                 
    spImpu = "ListaImputacDesdeHastaGrabacion " + XParam
    Set RstImpu = db.OpenRecordset(spImpu, dbOpenSnapshot, dbSQLPassThrough)
    If RstImpu.RecordCount > 0 Then
    
    With RstImpu
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                Lugar = Lugar + 1
                
                Vector(Lugar, 1) = Left$(!Clave, 27)
                Vector(Lugar, 2) = !Cuenta
                Vector(Lugar, 3) = Str$(!Debito)
                Vector(Lugar, 4) = Str$(!Credito)
                Vector(Lugar, 5) = !Proveedor
                Vector(Lugar, 6) = !LetraComp
                Vector(Lugar, 7) = !TipoComp
                Vector(Lugar, 8) = !PuntoComp
                Vector(Lugar, 9) = !NroComp
                Vector(Lugar, 10) = !Fecha
                Vector(Lugar, 11) = !Observaciones
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    RstImpu.Close
    End If
    
    WLugar = Lugar
    
    Erase Impre
    Pasa = 0
    Lugar = 0
    
    For x = 1 To WLugar
    
        WMarca = ""
            
        WProveedor = Vector(x, 5)
        Call Ceros(WProveedor, 11)
                
        WLetra = Vector(x, 6)
                
        WTipo = Vector(x, 7)
        Call Ceros(WTipo, 2)
                
        WPunto = Vector(x, 8)
        Call Ceros(WPunto, 4)
                
        WNumero = Vector(x, 9)
        Call Ceros(WNumero, 8)
                
        ClaveIvacomp = WProveedor + WLetra + WTipo + WPunto + WNumero
        spIvacomp = "ConsultaIvacomp " + "'" + ClaveIvacomp + "'"
        Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
        If rstIvacomp.RecordCount > 0 Then
            WMarca = IIf(IsNull(rstIvacomp!Marca), "", rstIvacomp!Marca)
            rstIvacomp.Close
        End If
        
        If WMarca <> "X" Then

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
            
                Corte = Vector(x, 1)
            
                XProveedor = Vector(x, 5)
                XLetracomp = Vector(x, 6)
                XTipocomp = Vector(x, 7)
                XPuntocomp = Vector(x, 8)
                XNroComp = Vector(x, 9)
                XFecha = Vector(x, 10)
                XObservaciones = Vector(x, 11)
            
                Erase Impre
                Lugar = 0
            
            End If
        
            If Corte <> Vector(x, 1) Then
        
                spProveedor = "ConsultaProveedores " + "'" + XProveedor + "'"
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
                            WLeyenda = Compro(Val(XTipocomp)) + " " + XLetracomp + " " + XPuntocomp + " " + XNroComp + " " + DesProveedor

                            .AddNew
                            !Asiento = Val(WAsiento)
                            Auxi1 = Str$(a)
                            Call Ceros(Auxi1, 2)
                            !Renglon = a
                            !Fecha = Fecha.Text
                            !Observaciones = "Asientos de Compras"
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
                
                    WProveedor = XProveedor
                    Call Ceros(WProveedor, 11)
                
                    WLetra = XLetracomp
                
                    WTipo = XTipocomp
                    Call Ceros(WTipo, 2)
                
                    WPunto = XPuntocomp
                    Call Ceros(WPunto, 4)
                
                    WNumero = XNroComp
                    Call Ceros(WNumero, 8)
                
                    ClaveIvacomp = WProveedor + WLetra + WTipo + WPunto + WNumero
                    WMarca = "X"
                
                    XParam = "'" + ClaveIvacomp + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
                    spIvacomp = "ActualizaIvacompAsiento " + XParam
                    Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
                
                End With
            
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
            
                Corte = Vector(x, 1)
            
                XProveedor = Vector(x, 5)
                XLetracomp = Vector(x, 6)
                XTipocomp = Vector(x, 7)
                XPuntocomp = Vector(x, 8)
                XNroComp = Vector(x, 9)
                XFecha = Vector(x, 10)
                XObservaciones = Vector(x, 11)
            
                Erase Impre
                Lugar = 0
            
            End If
        
            Lugar = Lugar + 1
        
            Impre(Lugar, 1) = Vector(x, 2)
            Impre(Lugar, 2) = Vector(x, 3)
            Impre(Lugar, 3) = Vector(x, 4)
            
        End If
    
    Next x
    
    If Pasa <> 0 Then
    
        spProveedor = "ConsultaProveedores " + "'" + XProveedor + "'"
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
                    WLeyenda = Compro(Val(XTipocomp)) + " " + XLetracomp + " " + XPuntocomp + " " + XNroComp + " " + DesProveedor
                
                    .AddNew
                    !Asiento = Val(WAsiento)
                    Auxi1 = Str$(a)
                    Call Ceros(Auxi1, 2)
                    !Renglon = a
                    !Fecha = Fecha.Text
                    !Observaciones = "Asientos de Compras"
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
            
            WProveedor = XProveedor
            Call Ceros(WProveedor, 11)
                
            WLetra = XLetracomp
                
            WTipo = XTipocomp
            Call Ceros(WTipo, 2)
                
            WPunto = XPuntocomp
            Call Ceros(WPunto, 4)
                
            WNumero = XNroComp
            Call Ceros(WNumero, 8)
                
            ClaveIvacomp = WProveedor + WLetra + WTipo + WPunto + WNumero
            WMarca = "X"
                
            XParam = "'" + ClaveIvacomp + "','" _
                            + WMarca + "','" _
                            + WAsiento + "'"
                
            spIvacomp = "ActualizaIvacompAsiento " + XParam
            Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
                
        End With
                
    End If
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstImputac
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
    PrgGraba1.Hide
    Unload Me
    Menu.Show
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
    OPEN_FILE_Empresa
    OPEN_FILE_Imputac
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



