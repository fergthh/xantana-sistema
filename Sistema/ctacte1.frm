VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCtaCte1 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Clientes"
   ClientHeight    =   8280
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8280
   ScaleWidth      =   11685
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   9600
      TabIndex        =   22
      Top             =   840
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   8
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   4440
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   7
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4440
      Width           =   375
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu F10"
      Height          =   975
      Left            =   6000
      MouseIcon       =   "ctacte1.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ctacte1.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Salida"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton DatosCli 
      Caption         =   "Cliente F4"
      Height          =   975
      Left            =   4920
      MouseIcon       =   "ctacte1.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "ctacte1.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Salida"
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   6
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   5
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   4
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   3
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3840
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3480
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   2
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3840
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   5295
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9340
      _Version        =   327680
   End
   Begin VB.TextBox Ayuda 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4335
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Datos"
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.OptionButton Todos 
         Caption         =   "Total"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton Pendiente 
         Caption         =   "Pendiente"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.TextBox Cliente 
      Height          =   300
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ListBox Pantalla 
      Height          =   1620
      ItemData        =   "ctacte1.frx":1B50
      Left            =   120
      List            =   "ctacte1.frx":1B57
      TabIndex        =   1
      Top             =   480
      Width           =   4335
   End
   Begin VB.ListBox WIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Saldo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   7920
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Saldo"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   6960
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   2760
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cliente"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCtaCte1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private Importe1 As Double
Private Importe2 As Double
Private Importe3 As Double
Private WTipo As Integer
Private WSaldo As Double

Private Sub cmdClose_Click()
    PrgCtaCte1.Hide
    Unload Me
    MenuVen.Show
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
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Order by Cliente.Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = !Cliente + " " + !Fantasia
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Cliente
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
    End If
            
    Rem Pantalla.Visible = True
    Rem Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Command1_Click()
    Call Pasactacte
End Sub

Private Sub datoscli_Click()
    PCliente = Cliente.Text
    prgcli.Show
End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Cliente.Text = WIndice.List(Indice)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Fantasia
        rstCliente.Close
        Call Proceso_Click
    End If
    
    Cliente.SetFocus
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
 
    Cliente.Text = ""
    DesCliente.Caption = ""

    WVector1.Col = 1
    WVector1.Row = 1
    
    Pendiente.Value = True
    Call Consulta_Click
    Cliente.Text = PCliente
    Rem Cliente.SetFocus
    
End Sub

Private Sub Proceso_Click()

    WSalida = "N"
    
    Call Limpia_Vector
    
    Renglon = 0
    WSaldo = 0
    ZZFechaActual = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    If Todos.Value = True Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " Order by CtaCte.Cliente,CtaCte.OrdFecha,CtaCte.Impre,CtaCte.Numero"
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCte"
        ZSql = ZSql + " Where CtaCte.Cliente = " + "'" + Cliente.Text + "'"
        ZSql = ZSql + " and CtaCte.Saldo <> 0 "
        ZSql = ZSql + " Order by CtaCte.Cliente,CtaCte.OrdFecha,CtaCte.Impre,CtaCte.Numero"
            
    End If
    
        
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
    
        With rstCtaCte
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstCtaCte!Total > 0 Then
                        Importe1 = rstCtaCte!Total
                        Importe2 = 0
                            Else
                        Importe1 = 0
                        Importe2 = rstCtaCte!Total
                    End If
                    Importe3 = rstCtaCte!Saldo
                    Call Redondeo(Importe3)
                
                    If Importe3 <> 0 Or Todos.Value = True Then
                    
                        Renglon = Renglon + 1
                            
                        WVector1.Row = Renglon
                        

                        If rstCtaCte!Contado = 1 Then
                            Importe2 = Importe1
                        End If
                                                
                        
                        Select Case rstCtaCte!Tipo
                            Case 1
                                WVector1.Col = 1
                                WVector1.Text = "Fac"
                            Case 2
                                WVector1.Col = 1
                                WVector1.Text = "Dev"
                            Case 3
                                WVector1.Col = 1
                                WVector1.Text = "Fac"
                            Case 4
                                WVector1.Col = 1
                                WVector1.Text = "N/D"
                            Case 5
                                WVector1.Col = 1
                                WVector1.Text = "N/C"
                            Case 6
                                WVector1.Col = 1
                                WVector1.Text = "Rec"
                            Case 7
                                WVector1.Col = 1
                                WVector1.Text = "Ant"
                            Case 50
                                WVector1.Col = 1
                                WVector1.Text = "Doc"
                            Case Else
                        End Select
                            
                        WVector1.Col = 2
                        WVector1.Text = Pusing("######", Str$(rstCtaCte!Numero))
                
                        WVector1.Col = 3
                        WVector1.Text = rstCtaCte!Fecha
                
                        If Importe1 <> 0 Then
                            WVector1.Col = 4
                            WVector1.Text = Pusing("###,###,###.##", Str$(Importe1))
                                Else
                            WVector1.Col = 4
                            WVector1.Text = ""
                        End If
                    
                        If Importe2 <> 0 Then
                            WVector1.Col = 5
                            WVector1.Text = Pusing("###,###,###.##", Str$(Importe2))
                                Else
                            WVector1.Col = 5
                            WVector1.Text = ""
                        End If
                
                        If Importe3 <> 0 Then
                            WVector1.Col = 6
                            WVector1.Text = Pusing("###,###,###.##", Str$(Importe3))
                                Else
                            WVector1.Col = 6
                            WVector1.Text = ""
                        End If
                        
                        WSaldo = WSaldo + Importe3
                
                        WVector1.Col = 7
                        WVector1.Text = IIf(IsNull(rstCtaCte!Partida), "", rstCtaCte!Partida)
                        
                        ZZfecha = rstCtaCte!Fecha
                        ZDias = 0
                        ZDias = DateDiff("d", ZZfecha, ZZFechaActual)
                        If ZDias < 0 Then
                            ZDias = 0
                        End If
                
                        WVector1.Col = 8
                        WVector1.Text = Str$(ZDias)
                        
                        
                        
                        
                        
                    End If
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCte.Close
    End If
    
    Saldo.Caption = Pusing("###,###,###.##", Str$(WSaldo))
    WVector1.Col = 0
    WVector1.Row = 0

End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WCliente = Cliente.Text
        Cliente.Text = WCliente
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Fantasia
            rstCliente.Close
            Call Proceso_Click
            WVector1.TopRow = 1
            WVector1.Col = 1
            WVector1.Row = 1
            Cliente.SetFocus
                Else
            Cliente.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 9
    WVector1.FixedRows = 1
    WVector1.Rows = 10000
    
    WVector1.ColWidth(0) = 200
    
    WVector1.Row = 0
    
    WVector1.Col = 1
    WVector1.Text = "Tipo"
    WVector1.ColWidth(1) = 600
    WVector1.ColAlignment(1) = flexAlignLeftCenter
    
    WVector1.Col = 2
    WVector1.Text = "Numero"
    WVector1.ColWidth(2) = 1300
    WVector1.ColAlignment(2) = flexAlignRightCenter
    
    WVector1.Col = 3
    WVector1.Text = "Fecha"
    WVector1.ColWidth(3) = 1300
    WVector1.ColAlignment(3) = flexAlignRightCenter
    
    WVector1.Col = 4
    WVector1.Text = "Debito"
    WVector1.ColWidth(4) = 1300
    WVector1.ColAlignment(4) = flexAlignRightCenter
    
    WVector1.Col = 5
    WVector1.Text = "Credito"
    WVector1.ColWidth(5) = 1300
    WVector1.ColAlignment(5) = flexAlignRightCenter
    
    WVector1.Col = 6
    WVector1.Text = "Saldo"
    WVector1.ColWidth(6) = 1300
    WVector1.ColAlignment(6) = flexAlignRightCenter
    
    WVector1.Col = 7
    WVector1.Text = "Partida"
    WVector1.ColWidth(7) = 1300
    WVector1.ColAlignment(7) = flexAlignRightCenter
    
    WVector1.Col = 8
    WVector1.Text = "Dias"
    WVector1.ColWidth(8) = 1300
    WVector1.ColAlignment(8) = flexAlignRightCenter
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Pendiente_Click()
    Call Proceso_Click
End Sub

Private Sub Todos_Click()
    Call Proceso_Click
End Sub

Private Sub WVector1_DblClick()

    WVector1.Col = 1
    Tipo = WVector1.Text
    
    If Tipo = "Rec" Then
        WVector1.Col = 2
        WRecibo = WVector1.Text
        Rem PrgRec.Show
    End If
    If Tipo = "Fac" Then
        WVector1.Col = 2
        ZZPasaNumero = WVector1.Text
        ZZPasaCliente = Cliente.Text
        If ZZNivel = 0 Then
            PrgFacturaConsulta.Show
                Else
            PrgFacturaRemitoMenuConsulta.Show
        End If
    End If
    
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
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Fantasia LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Fantasia
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM SubRepresentante"
            ZSql = ZSql + " Where SubRepresentante.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by SubRepresentante.Codigo"
            spSubRepresentante = ZSql
            Set rstSubRepresentante = db.OpenRecordset(spSubRepresentante, dbOpenSnapshot, dbSQLPassThrough)
            If rstSubRepresentante.RecordCount > 0 Then
                With rstSubRepresentante
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstSubRepresentante!Codigo) + " " + rstSubRepresentante!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstSubRepresentante!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSubRepresentante.Close
            End If
            
        Case 2
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Representante"
            ZSql = ZSql + " Where Representante.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Representante.Codigo"
            spRepresentante = ZSql
            Set rstRepresentante = db.OpenRecordset(spRepresentante, dbOpenSnapshot, dbSQLPassThrough)
            If rstRepresentante.RecordCount > 0 Then
                With rstRepresentante
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstRepresentante!Codigo) + " " + rstRepresentante!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstRepresentante!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstRepresentante.Close
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

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Pendiente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Todos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub WVector1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 115
            Call datoscli_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub







Private Sub Pasactacte()

    Dim ZZPasa(1000, 6) As String
    
    Erase ZZPasa
    ZZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM pasactacte"
    sppasaCtaCte = ZSql
    Set rstpasactacte = db.OpenRecordset(sppasaCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstpasactacte.RecordCount > 0 Then
    
        With rstpasactacte
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZLugar = ZZLugar + 1
                    
                    ZZPasa(ZZLugar, 1) = rstpasactacte!Tipo
                    ZZPasa(ZZLugar, 2) = rstpasactacte!Numero
                        ZZPasa(ZZLugar, 3) = rstpasactacte!Cliente
                    ZZPasa(ZZLugar, 4) = rstpasactacte!Razon
                    ZZPasa(ZZLugar, 5) = Str$(rstpasactacte!Saldo)
                    ZZPasa(ZZLugar, 6) = rstpasactacte!Fecha
        
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstpasactacte.Close
    End If
    
    
    For Ciclo = 1 To ZZLugar
    
        ZZZZTipo = ZZPasa(Ciclo, 1)
        ZZZZNumero = ZZPasa(Ciclo, 2)
        ZZZZCliente = ZZPasa(Ciclo, 3)
        ZZZZRazon = ZZPasa(Ciclo, 4)
        ZZZZSaldo = Val(ZZPasa(Ciclo, 5))
        ZZZZFecha = ZZPasa(Ciclo, 6)
    
            
        Auxi = ZZZZNumero
        Call Ceros(Auxi, 8)
                
        WPunto = "0002"
        
        If Val(ZZZZTipo) = 1 Then
            ZZTipo = "01"
            ZZImpre = "FC"
            ZZTotal = Str$(ZZZZSaldo)
            ZZSaldo = Str$(ZZZZSaldo)
                Else
            ZZTipo = "02"
            ZZImpre = "NC"
            ZZTotal = Str$(ZZZZSaldo * -1)
            ZZSaldo = Str$(ZZZZSaldo * -1)
        End If
                
        ZZPunto = WPunto
        ZZLetra = "A"
        ZZNumero = Auxi
        ZZRenglon = "01"
        ZZCliente = ZZZZCliente
        ZZfecha = ZZZZFecha
        ZZEstado = "0"
        ZZVencimiento = ZZZZFecha
        ZZNeto = "0"
        ZZIva1 = "0"
        ZZIva2 = "0"
        ZZExento = "0"
        ZZOrdFecha = Right$(ZZZZFecha, 4) + Mid$(ZZZZFecha, 4, 2) + Left$(ZZZZFecha, 2)
        ZZOrdVencimiento = Right$(ZZZZFecha, 4) + Mid$(ZZZZFecha, 4, 2) + Left$(ZZZZFecha, 2)
        ZZPedido = ""
        ZZRemito = ""
        ZZOrden = ""
        ZZProvincia = ""
        ZZVendedor = ""
        ZZCosto = "0"
        ZZImporte1 = "0"
        ZZImporte2 = "0"
        ZZImporte3 = "0"
        ZZImporte4 = "0"
        ZZImporte5 = "0"
        ZZImporte6 = "0"
        ZZImporte7 = "0"
        ZZTipoventa = "0"
        ZZProyecto = ""
        ZZParidad = "0"
        ZZRemito1 = ""
        ZZRemito2 = ""
        ZZBusqueda = ZZLetra + WPunto + Auxi
        
        ZZDescuento = ""
        ZZPago = ""
        ZZPartida = ""
        ZZExpreso = ""
        ZZTipoIva = ""
        ZZComision = ""
        ZZRemito = ""
        
        ZZContado = "0"
        ZZEntregada = "1"
        
        ZZClave = ZZLetra + ZZTipo + WPunto + Auxi + "01"
        
        ZZLinea = ""
        
        ZZNetoTotal = ZZNeto
        ZZTotalUs = ""
        ZZSaldoUs = ""
        ZZCae = ""
        ZZVtoCae = ""
        ZZCodigoEmpresa = "2"
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ctacte ("
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
        ZSql = ZSql + "OrdFecha  ,"
        ZSql = ZSql + "OrdVencimiento ,"
        ZSql = ZSql + "Impre ,"
        ZSql = ZSql + "Neto ,"
        ZSql = ZSql + "NetoTotal ,"
        ZSql = ZSql + "Iva1 ,"
        ZSql = ZSql + "Iva2 ,"
        ZSql = ZSql + "Exento ,"
        ZSql = ZSql + "Pedido ,"
        ZSql = ZSql + "Remito ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "Vendedor ,"
        ZSql = ZSql + "Costo ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Importe3 ,"
        ZSql = ZSql + "Importe4 ,"
        ZSql = ZSql + "Importe5 ,"
        ZSql = ZSql + "Importe6 ,"
        ZSql = ZSql + "Importe7 ,"
        ZSql = ZSql + "Tipoventa ,"
        ZSql = ZSql + "Proyecto ,"
        ZSql = ZSql + "Paridad ,"
        ZSql = ZSql + "TotalUs ,"
        ZSql = ZSql + "SaldoUs ,"
        ZSql = ZSql + "Remito1 ,"
        ZSql = ZSql + "Remito2 ,"
        ZSql = ZSql + "Descuento ,"
        ZSql = ZSql + "Partida ,"
        ZSql = ZSql + "Pago ,"
        ZSql = ZSql + "Linea ,"
        ZSql = ZSql + "Expreso ,"
        ZSql = ZSql + "TipoIva ,"
        ZSql = ZSql + "Comision ,"
        ZSql = ZSql + "NroRemito ,"
        ZSql = ZSql + "Cae ,"
        ZSql = ZSql + "VtoCae ,"
        ZSql = ZSql + "Contado ,"
        ZSql = ZSql + "Entregada ,"
        ZSql = ZSql + "CodigoEmpresa ,"
        ZSql = ZSql + "Busqueda )"
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
        ZSql = ZSql + "'" + ZZNeto + "',"
        ZSql = ZSql + "'" + ZZNetoTotal + "',"
        ZSql = ZSql + "'" + ZZIva1 + "',"
        ZSql = ZSql + "'" + ZZIva2 + "',"
        ZSql = ZSql + "'" + ZZExento + "',"
        ZSql = ZSql + "'" + ZZPedido + "',"
        ZSql = ZSql + "'" + ZZRemito + "',"
        ZSql = ZSql + "'" + ZZOrden + "',"
        ZSql = ZSql + "'" + ZZProvincia + "',"
        ZSql = ZSql + "'" + ZZVendedor + "',"
        ZSql = ZSql + "'" + ZZCosto + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZImporte3 + "',"
        ZSql = ZSql + "'" + ZZImporte4 + "',"
        ZSql = ZSql + "'" + ZZImporte5 + "',"
        ZSql = ZSql + "'" + ZZImporte6 + "',"
        ZSql = ZSql + "'" + ZZImporte7 + "',"
        ZSql = ZSql + "'" + ZZTipoventa + "',"
        ZSql = ZSql + "'" + ZZProyecto + "',"
        ZSql = ZSql + "'" + ZZParidad + "',"
        ZSql = ZSql + "'" + ZZTotalUs + "',"
        ZSql = ZSql + "'" + ZZSaldoUs + "',"
        ZSql = ZSql + "'" + ZZRemito1 + "',"
        ZSql = ZSql + "'" + ZZRemito2 + "',"
        ZSql = ZSql + "'" + ZZDescuento + "',"
        ZSql = ZSql + "'" + ZZPartida + "',"
        ZSql = ZSql + "'" + ZZPago + "',"
        ZSql = ZSql + "'" + ZZLinea + "',"
        ZSql = ZSql + "'" + ZZExpreso + "',"
        ZSql = ZSql + "'" + ZZTipoIva + "',"
        ZSql = ZSql + "'" + ZZComision + "',"
        ZSql = ZSql + "'" + ZZRemito + "',"
        ZSql = ZSql + "'" + ZZCae + "',"
        ZSql = ZSql + "'" + ZZVtoCae + "',"
        ZSql = ZSql + "'" + ZZContado + "',"
        ZSql = ZSql + "'" + ZZEntregada + "',"
        ZSql = ZSql + "'" + ZZCodigoEmpresa + "',"
        ZSql = ZSql + "'" + ZZBusqueda + "')"
                                
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo

End Sub
