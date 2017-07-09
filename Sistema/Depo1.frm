VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDepo1 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingresos de Depositos"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   1065
   ClientWidth     =   11880
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   11880
   Begin VB.CommandButton LimpiaLinea 
      Caption         =   "Limpia Linea"
      Height          =   300
      Left            =   3480
      TabIndex        =   23
      Top             =   2040
      Width           =   1215
   End
   Begin VB.ListBox WVector 
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Deposito 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   855
   End
   Begin MSMask.MaskEdBox Acredita 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox Importe 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   15
      TabIndex        =   4
      Text            =   " "
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Banco 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   2
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   5880
      TabIndex        =   16
      Top             =   360
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSDBGrid.DBGrid DbGrid1 
      Height          =   3375
      Left            =   120
      OleObjectBlob   =   "Depo1.frx":0000
      TabIndex        =   5
      Top             =   2400
      Width           =   5295
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6300
      ItemData        =   "Depo1.frx":09C6
      Left            =   5640
      List            =   "Depo1.frx":09CD
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   2400
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpar"
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1320
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Fec.Acreditacion"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Importe"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Creditos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3600
      TabIndex        =   19
      Top             =   5880
      Width           =   1815
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tipo de Doc. : 1) Ef.    3) Ch. Terc."
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   5880
      Width           =   3255
   End
   Begin VB.Label DesBanco 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   2520
      TabIndex        =   17
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Banco"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro. Deposito"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "PrgDepo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' Número máximo de campos del conjunto de registros.
Private Auxi As String
Private dada As String
Private Vector(10, 6) As String
Private Numero As String


Private Sub Suma_Datos()
    Creditos.Caption = ""
    
    For iRow = 0 To 9
        DbGrid1.Col = 4
        DbGrid1.Row = iRow
        Auxi = DbGrid1.Text
        Call Conver(Auxi, dada)
        If Val(Auxi) <> 0 Then
            Creditos.Caption = Str$(Val(Creditos.Caption) + Val(Auxi))
        End If
    Next iRow
    Creditos.Caption = Pusing("###,###.##", Creditos.Caption)
    DbGrid1.Col = 0
    DbGrid1.Row = 0
End Sub

Private Sub Lee_Datos()
    Renglon = 0
    Debito = 0
    Credito = 0
    Do
        With rstDepositos
            .Index = "Clave"
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            .Seek "=", Deposito.Text + Auxi1
             If .NoMatch = False Then
                Credito = Credito + 1
                DbGrid1.Row = Credito - 1
                DbGrid1.Col = 0
                DbGrid1.Text = !Tipo2
                DbGrid1.Col = 1
                DbGrid1.Text = !Numero2
                DbGrid1.Col = 2
                DbGrid1.Text = !Fecha2
                DbGrid1.Col = 3
                If !Observaciones2 <> "" Then
                    DbGrid1.Text = !Observaciones2
                End If
                DbGrid1.Col = 4
                DbGrid1.Text = !Importe2
                DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    Else
                Exit Do
            End If
        End With
    Loop
End Sub
Sub Verifica_datos()
    If Importe.Text = 0 Then
        Importe.Text = "0"
    End If
End Sub
Sub Format_datos()
    Importe.Text = Pusing("###,###.##", Importe.Text)
End Sub

Sub Imprime_Datos()
    With rstBanco
        .Index = "Banco"
        .Seek "=", Banco.Text
        If .NoMatch = False Then
            Banco.Text = !Banco
            DesBanco.Caption = !Nombre
            Call Format_datos
        End If
    End With
End Sub

Private Sub cmdAdd_Click()

    If Deposito.Text <> "" And Fecha.Text <> "" And Banco.Text <> "" Then
    
    If Existe <> "S" Then
    
        Call Suma_Datos
        
        Debito = 0
        Credito = 0
        
        If Val(Importe.Text) <> 0 Then
            Debito = Val(Importe.Text)
        End If
        
        If Val(Creditos.Caption) <> 0 Then
            Credito = Val(Creditos.Caption)
        End If
        
        If Debito = Credito Then
    
        With rstDepositos
            Renglon = 0
            .Index = "Clave"
            For iRow = 0 To 9
                WRow = iRow
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                Auxi = DbGrid1.Text
                Call Conver(Auxi, dada)
                If Val(Auxi) <> 0 Then
                    .AddNew
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    Auxi2 = Str$(Val(Deposito.Text))
                    Call Ceros(Auxi2, 6)
                    !Deposito = Auxi2
                    !Renglon = Auxi1
                    !Banco = Banco.Text
                    !Importe = Val(Importe.Text)
                    !Fecha = Fecha.Text
                    !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Acredita = Acredita.Text
                    !AcreditaOrd = Right$(Acredita.Text, 4) + Mid$(Acredita.Text, 4, 2) + Left$(Acredita.Text, 2)
                    DbGrid1.Col = 0
                    !Tipo2 = DbGrid1.Text
                    DbGrid1.Col = 1
                    !Numero2 = DbGrid1.Text
                    DbGrid1.Col = 2
                    !Fecha2 = DbGrid1.Text
                    DbGrid1.Col = 3
                    !Observaciones2 = DbGrid1.Text
                    DbGrid1.Col = 4
                    !Importe2 = Val(Auxi)
                    !Empresa = 1
                    !Clave = !Deposito + !Renglon
                    .Update
                    .Bookmark = .LastModified
                    
                    With rstRecibos
                        .Index = "Clave"
                        DbGrid1.Col = 5
                        Claveven$ = DbGrid1.Text
                        .Seek "=", Claveven$
                        If .NoMatch = False Then
                            .Edit
                            !Estado2 = "X"
                            !Destino = "Deposito Nro : " + Str$(Deposito.Text) + " Banco : " + DesBanco.Caption
                            .Update
                        End If
                    End With
                    
                End If
                
            Next iRow
        End With
        
        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = "1"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                WCtaEfectivo = !CtaEfectivo
                WCtaCheques = !CtaCheque
            End If
        End With
        
        With rstImputac
            Renglon = 0
            .Index = "Clave"
            
            .AddNew
            !Tipomovi = "4"
            !Proveedor = "000000"
            !TipoComp = "01"
            !LetraComp = "A"
            !PuntoComp = "0000"
            !NroComp = Deposito.Text
            Renglon = Renglon + 1
            Auxi1 = Str$(Renglon)
            Call Ceros(Auxi1, 2)
            !Renglon = Auxi1$
            !Fecha = Fecha.Text
            !Debito = Val(Importe.Text)
            !Credito = 0
            !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            !Titulo = "Depositos"
            !Empresa = 1
            With rstBanco
                .Index = "Banco"
                .Seek "=", Val(Banco.Text)
                If .NoMatch = False Then
                    WCuenta = !Cuenta
                End If
            End With
            !Cuenta = WCuenta
            !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
            .Update
            
            For iRow = 0 To 9
                WRow = iRow
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                Auxi = DbGrid1.Text
                Call Conver(Auxi, dada)
                If Val(Auxi) <> 0 Then
                    .AddNew
                    !Tipomovi = "4"
                    !Proveedor = "000000"
                    !TipoComp = "01"
                    !LetraComp = "A"
                    !PuntoComp = "0000"
                    !NroComp = Deposito.Text
                    Renglon = Renglon + 1
                    Auxi1 = Str$(Renglon)
                    Call Ceros(Auxi1, 2)
                    !Renglon = Auxi1$
                    !Fecha = Fecha.Text
                    !Credito = Val(Auxi)
                    !Debito = 0
                    !FechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                    !Titulo = "Depositos"
                    !Empresa = 1
                    DbGrid1.Col = 0
                    Select Case Val(DbGrid1.Text)
                        Case 1
                            !Cuenta = WCtaEfectivo
                        Case Else
                            !Cuenta = WCtaCheques
                    End Select
                    !Clave = !Tipomovi + !TipoComp + !LetraComp + !PuntoComp + !NroComp + !Renglon
                    .Update
                End If
                
            Next iRow
        End With
        
        Call ImpreDeposito


        Call CmdLimpiar_Click
        Deposito.SetFocus
        
        End If
        
        End If
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Deposito.Text <> "" Then
            Rem Borro los datos anteriores
            Rem For iRow = 0 To 20
            Rem     Auxi1 = Str$(iRow)
            Rem     Call Ceros(Auxi1, 2)
            Rem     .Seek "=", Deposito.text + Auxi1
            Rem     If .NoMatch = False Then
            Rem         .Delete
            Rem     End If
            Rem Next iRow
    End If
    Deposito.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Pantalla.Visible = False

    For iCol = 0 To 5
        For iRow = 0 To 9
            DbGrid1.Col = iCol
            DbGrid1.Row = iRow
            DbGrid1.Text = ""
        Next iRow
    Next iCol
    Deposito.Text = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Importe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Acredita.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Creditos.Caption = ""
    Deposito.SetFocus

    With rstDepositos
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Deposito.Text = !Deposito + 1
                Else
            Deposito.Text = ""
        End If
    End With

End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstBanco
        .Close
    End With
    With rstImputac
        .Close
    End With
    With rstDepositos
        .Close
    End With
    DbsAdminis.Close
    Deposito.SetFocus
    PrgDeposito.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub


Private Sub Deposito_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Auxi1 = Deposito.Text
        Call Ceros(Auxi1, 6)
        Deposito.Text = Auxi1
        
        With rstDepositos
            Existe = "N"
            .Index = "Clave"
            Claveven$ = Deposito.Text + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Existe = "S"
                If !Banco <> "" Then
                    Banco.Text = !Banco
                End If
                If !Importe <> "" Then
                    Importe.Text = !Importe
                End If
                Fecha.Text = !Fecha
                Acredita.Text = !Acredita
            End If
        End With
        If Existe = "S" Then
            Call Imprime_Datos
            Call Lee_Datos
            Call Suma_Datos
            DbGrid1.Col = 0
            DbGrid1.Row = 0
            DbGrid1.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Banco.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Banco.Text) <> 0 Then
            With rstBanco
                .Index = "Banco"
                Claveven$ = Banco.Text
                .Seek "=", Banco.Text
                If .NoMatch Then
                    Banco.Text = Claveven$
                    Banco.SetFocus
                        Else
                    Banco.Text = !Banco
                    DesBanco.Caption = !Nombre
                    Rem Call Imprime_Datos
                    Acredita.SetFocus
                End If
            End With
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Acredita_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Acredita.Text, Auxi)
        If Auxi = "S" Then
            Importe.SetFocus
                Else
            Acredita.SetFocus
        End If
    End If
End Sub

Private Sub Importe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Importe.Text = Pusing("###,###.##", Importe.Text)
        DbGrid1.Col = 0
        DbGrid1.Row = 0
        DbGrid1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub Consulta_Click()

     XRow = DbGrid1.Row
     XCol = DbGrid1.Col

     Opcion.Clear

     Opcion.AddItem "Bancos"
     Opcion.AddItem "Cheques terceros"

     Opcion.Visible = True
     
End Sub

Private Sub LimpiaLinea_Click()

    Rem Dada

    NoToma = DbGrid1.Row
    Erase Vector
    Lugar = 0
    
    For iRow = 0 To 9
        If NoToma <> iRow Then
            Lugar = Lugar + 1
            DbGrid1.Row = iRow
            DbGrid1.Col = 0
            Vector(Lugar, 1) = DbGrid1.Text
            DbGrid1.Col = 1
            Vector(Lugar, 2) = DbGrid1.Text
            DbGrid1.Col = 2
            Vector(Lugar, 3) = DbGrid1.Text
            DbGrid1.Col = 3
            Vector(Lugar, 4) = DbGrid1.Text
            DbGrid1.Col = 4
            Vector(Lugar, 5) = DbGrid1.Text
            DbGrid1.Col = 5
            Vector(Lugar, 6) = DbGrid1.Text
        End If
        
        DbGrid1.Col = 0
        DbGrid1.Text = ""
        DbGrid1.Text = ""
        DbGrid1.Col = 1
        DbGrid1.Text = ""
        DbGrid1.Col = 2
        DbGrid1.Text = ""
        DbGrid1.Col = 3
        DbGrid1.Text = ""
        DbGrid1.Col = 4
        DbGrid1.Text = ""
        DbGrid1.Col = 5
        DbGrid1.Text = ""
        
    Next iRow
    
    For da = 1 To Lugar
        
        DbGrid1.Row = da - 1
        DbGrid1.Col = 0
        DbGrid1.Text = Vector(da, 1)
        DbGrid1.Col = 1
        DbGrid1.Text = Vector(da, 2)
        DbGrid1.Col = 2
        DbGrid1.Text = Vector(da, 3)
        DbGrid1.Col = 3
        DbGrid1.Text = Vector(da, 4)
        DbGrid1.Col = 4
        DbGrid1.Text = Vector(da, 5)
        DbGrid1.Col = 5
        DbGrid1.Text = Vector(da, 6)
    
    Next da

End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    WVector.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
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
                
        Case 1
            With rstRecibos
                .Index = "Fecha2"
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Val(!Tiporeg) = 2 Then
                            If Val(!Tipo2) = 2 And !Estado2 <> "X" Then
                                Auxi$ = Str$(!Importe2)
                                Auxi$ = Mascara("###,###.##", Auxi$)
                                Numero = Str$(Val(!Numero2))
                                Call Ceros(Numero, 6)
                                IngresaItem = Numero + "    " + !Fecha2 + "      " + Auxi$ + "      " + !Banco2
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Clave
                                WIndice.AddItem IngresaItem
                                IngresaItem = ""
                                WVector.AddItem IngresaItem
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
     
        Case Else
    End Select
    
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_DblClick()

    Select Case XIndice
        Case 0
            With rstBanco
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Banco.Text = Claveven$
                .Index = "Banco"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DesBanco.Caption = !Nombre
                            Else
                    Banco.Text = ""
                End If
            End With
                
            Pantalla.Visible = False
            Banco.SetFocus
            
        Case 1
        
            Indice = Pantalla.ListIndex
            Auxi = WVector.List(Indice)
            If Auxi <> "X" Then
            
            For iRow = 0 To 9
                DbGrid1.Col = 4
                DbGrid1.Row = iRow
                Auxi = DbGrid1.Text
                Call Conver(Auxi, dada)
                If Val(Auxi) = 0 Then
                    Exit For
                End If
            Next iRow
            
            With rstRecibos
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Clave"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                
                    DbGrid1.Col = 0
                    DbGrid1.Text = 3
                    
                    DbGrid1.Col = 1
                    DbGrid1.Text = !Numero2
                
                    DbGrid1.Col = 2
                    DbGrid1.Text = !Fecha2
                
                    DbGrid1.Col = 3
                    DbGrid1.Text = !Banco2
                
                    DbGrid1.Col = 4
                    DbGrid1.Text = !Importe2
                    DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    
                    DbGrid1.Col = 5
                    DbGrid1.Text = Claveven$
                    
                    Call Suma_Datos
                    
                    DbGrid1.Row = XRow
                    DbGrid1.Col = 0
                    
                    Auxi = "X"
                    WVector.List(Indice) = Auxi
                    
                    Pantalla.List(Indice) = ""
                    
                End If
                If DbGrid1.Row < 10 Then
                    DbGrid1.Row = DbGrid1.Row + 1
                    DbGrid1.Col = 0
                    KeyCode = 0
                            Else
                    DbGrid1.Col = 0
                    KeyCode = 0
                End If
            End With
            
            End If
                
        Case Else
    End Select
    
End Sub
Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DbGrid1.Col
            Case 0
                If KeyCode = 13 Then
                    If Val(DbGrid1.Text) = 1 Or Val(DbGrid1.Text) = 3 Then
                        Auxi$ = Str$(Val(DbGrid1.Text))
                        Call Ceros(Auxi$, 2)
                        DbGrid1.Text = Auxi$
                        
                        Select Case Val(DbGrid1.Text)
                            Case 1
                                DbGrid1.Col = 1
                                DbGrid1.Text = ""
                                DbGrid1.Col = 2
                                DbGrid1.Text = ""
                                DbGrid1.Col = 3
                                DbGrid1.Text = ""
                                DbGrid1.Col = 4
                                DbGrid1.Text = Importe.Text
                                DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                                Call Suma_Datos
                                DbGrid1.Row = iRow
                                If DbGrid1.Row < 10 Then
                                    DbGrid1.Row = DbGrid1.Row + 1
                                    DbGrid1.Col = 0
                                    KeyCode = 0
                                        Else
                                    DbGrid1.Col = 0
                                    KeyCode = 0
                                End If
                                
                            Case 2
                                Call Consulta_Click
                                
                            Case Else
                                DbGrid1.Col = 1
                                KeyCode = 0
                                
                        End Select
                        
                            Else
                            
                        DbGrid1.Col = 0
                        KeyCode = 0
                        
                    End If
                End If
                
            Case 1
                If KeyCode = 13 Then
                    DbGrid1.Col = 0
                    If Val(DbGrid1.Text) = 3 Then
                        DbGrid1.Col = 1
                        KeyCode = 0
                    End If
                End If
                
            Case 4
                If KeyCode = 13 Then
                    iRow = DbGrid1.Row
                    DbGrid1.Col = 4
                    DbGrid1.Text = Pusing("###,###.##", DbGrid1.Text)
                    Call Suma_Datos
                    DbGrid1.Row = iRow
                    If DbGrid1.Row < 10 Then
                        DbGrid1.Row = DbGrid1.Row + 1
                        DbGrid1.Col = 0
                        KeyCode = 0
                            Else
                        DbGrid1.Col = 0
                        KeyCode = 0
                    End If
                End If

            Case Else
                
    End Select
    
End Sub

' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DbGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub


Private Sub Form_Load()

ReDim UserData(0 To 5, 0 To 9)

mTotalRows& = 10

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DbGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DbGrid1.Columns.Count - 1 To 0 Step -1
     DbGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 5
    DbGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DbGrid1.Columns(newcnt).Caption = "Tipo"
             DbGrid1.Columns(newcnt).Width = 400
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = False
         Case 1
             DbGrid1.Columns(newcnt).Caption = "Numero"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 2
             DbGrid1.Columns(newcnt).Caption = "Fecha"
             DbGrid1.Columns(newcnt).Width = 1150
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 3
             DbGrid1.Columns(newcnt).Caption = "Nombre"
             DbGrid1.Columns(newcnt).Width = 1000
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case 4
             DbGrid1.Columns(newcnt).Caption = "Importe"
             DbGrid1.Columns(newcnt).Width = 1200
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Alignment = 1
             DbGrid1.Columns(newcnt).Locked = True
         Case 5
             DbGrid1.Columns(newcnt).Caption = ""
             DbGrid1.Columns(newcnt).Width = 1
             DbGrid1.Columns(newcnt).AllowSizing = False
             DbGrid1.Columns(newcnt).Locked = True
         Case Else

     End Select
     DbGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
 Next i
 
    Pantalla.Visible = False
    
    Deposito.Text = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Importe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Acredita.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Creditos.Caption = ""
    
    With rstDepositos
        .Index = "Clave"
        Claveven$ = "99999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Deposito.Text = !Deposito + 1
                Else
            Deposito.Text = ""
        End If
    End With
     
End Sub


Private Sub ImpreDeposito()

        Printer.Font = "Times New Roman"
        Printer.FontSize = "10"
        Printer.Print Tab(1); ""
        Printer.FontSize = "9"

        With rstEmpresa
            .Index = "Empresa"
            Claveven$ = WEmpresa
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Impretit = !Nombre
                    Else
                Impretit = ""
            End If
        End With
        
        Printer.Print Tab(1); Impretit
        Printer.Print Tab(1); "DEPOSITO"
        Printer.Print Tab(1); ""
        Printer.Print Tab(1); "Deposito Nro.";
        Printer.Print Tab(20); Deposito.Text
        Printer.Print Tab(1); "Banco";
        Printer.Print Tab(20); Banco.Text;
        Printer.Print Tab(30); DesBanco.Caption
        Printer.Print Tab(1); "Fecha";
        Printer.Print Tab(20); Fecha.Text
        Printer.Print Tab(1); "Total";
        Printer.Print Tab(20); Pusing("###,###.##", Importe.Text)
        Printer.Print Tab(1); ""
        Printer.Print Tab(1); ""

        Printer.Print Tab(1); "Tipo";
        Printer.Print Tab(20); "Numero";
        Printer.Print Tab(40); "Banco";
        Printer.Print Tab(80); "Importe"
        Printer.Print Tab(1); ""
        
        For iRow = 0 To 9
            DbGrid1.Col = 4
            DbGrid1.Row = iRow
            Auxi = DbGrid1.Text
            Call Conver(Auxi, dada)
            If Val(Auxi) <> 0 Then
                DbGrid1.Col = 0
                If Val(DbGrid1.Text) = 1 Then
                    Printer.Print Tab(1); "Efectivo";
                    DbGrid1.Col = 4
                    Printer.Print Tab(80); DbGrid1.Text
                        Else
                    Printer.Print Tab(1); "Cheque";
                    DbGrid1.Col = 1
                    Printer.Print Tab(20); DbGrid1.Text;
                    DbGrid1.Col = 3
                    Printer.Print Tab(40); DbGrid1.Text;
                    DbGrid1.Col = 4
                    Printer.Print Tab(80); DbGrid1.Text
                End If
            End If
        Next iRow
        
        Printer.EndDoc

End Sub
