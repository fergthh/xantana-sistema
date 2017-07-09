VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgNotaEnvio 
   AutoRedraw      =   -1  'True
   Caption         =   "Nota de Envio"
   ClientHeight    =   8115
   ClientLeft      =   1005
   ClientTop       =   420
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   9900
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9495
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   1920
         TabIndex        =   24
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox Postal 
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
         MaxLength       =   50
         TabIndex        =   22
         Text            =   " "
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox Copia 
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
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   20
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpia F3"
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
         Left            =   2760
         MouseIcon       =   "NotaEnvio.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "NotaEnvio.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Limpia la pantalla"
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Proceso 
         Caption         =   "Graba"
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
         Left            =   1200
         MouseIcon       =   "NotaEnvio.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "NotaEnvio.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Provincia 
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
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   2520
         Width           =   5535
      End
      Begin VB.TextBox Cliente 
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
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   975
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
         Left            =   4440
         MouseIcon       =   "NotaEnvio.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "NotaEnvio.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Consulta de Datos"
         Top             =   3480
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
         Left            =   6120
         MouseIcon       =   "NotaEnvio.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "NotaEnvio.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Salida"
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Bultos 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   8
         Text            =   " "
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Razon 
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
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   720
         Width           =   5655
      End
      Begin VB.TextBox Direccion 
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
         MaxLength       =   50
         TabIndex        =   6
         Text            =   " "
         Top             =   1440
         Width           =   5535
      End
      Begin VB.TextBox Localidad 
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
         MaxLength       =   50
         TabIndex        =   5
         Text            =   " "
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Etiqueta"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Postal"
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
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Copias"
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
         Left            =   4080
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social"
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
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Bultos"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Localidad"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   1575
      End
   End
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
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   7815
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ventclie.rpt"
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
      Left            =   6840
      TabIndex        =   2
      Top             =   5160
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
      ItemData        =   "NotaEnvio.frx":2D30
      Left            =   120
      List            =   "NotaEnvio.frx":2D37
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgNotaEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZProvincia(100) As String

Private Sub CmdLimpiar_Click()

    Cliente.Text = ""
    Razon.Text = ""
    Bultos.Text = ""
    Copia.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Provincia.Text = ""
    Postal.Text = ""
    
    Cliente.SetFocus

End Sub


Private Sub Proceso_Click()

    If Tipo.ListIndex = 0 Then
    
        Open "lpt1" For Output As #1
        Rem Open "dada.txt" For Output As #1
    
        For Ciclo = 1 To Val(Copia.Text)
    
            
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
            Print #1, Tab(7); Left$(Razon.Text, 20); " "; Cliente.Text;
            Print #1, Tab(52); Left$(Razon.Text, 20); " "; Cliente.Text
            
            Print #1, ""
            Print #1, ""
            
            Print #1, Tab(7); Left$(Direccion.Text, 30);
            Print #1, Tab(52); Left$(Direccion.Text, 30)
            
            Print #1, ""
            Print #1, ""
            
            ZZLocalidad = Trim(Postal.Text) + " - " + Trim(Localidad.Text)
            Print #1, Tab(7); Left$(ZZLocalidad, 30);
            Print #1, Tab(52); Left$(ZZLocalidad, 30)
            
            Print #1, ""
            Print #1, ""
            
            Print #1, Tab(7); Left$(Provincia.Text, 20); Tab(32); Bultos.Text;
            Print #1, Tab(52); Left$(Provincia.Text, 20); Tab(77); Bultos.Text
            
            Print #1, ""
            Print #1, ""
            Print #1, ""
            Print #1, ""
            
        Next Ciclo
        
        Close #1
    
            Else
            
        ZSql = ""
        ZSql = ZSql + "UPDATE Cliente SET "
        ZSql = ZSql + " ImpreRazon = " + "'" + Razon.Text + "',"
        ZSql = ZSql + " ImpreDireccion = " + "'" + Direccion.Text + "',"
        ZSql = ZSql + " ImpreLocalidad = " + "'" + Localidad.Text + "',"
        ZSql = ZSql + " ImpreProvincia = " + "'" + Provincia.Text + "',"
        ZSql = ZSql + " ImpreBultos = " + "'" + Bultos.Text + "'"
        ZSql = ZSql + " Where Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
                            
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        
        Listado.SQLQuery = "SELECT Cliente.Cliente, Cliente.ImpreRazon, Cliente.ImpreDireccion, Cliente.ImpreLocalidad, Cliente.ImpreProvincia, Cliente.ImpreBultos " _
                + "From " _
                + DSQ + ".dbo.Cliente Cliente " _
                + "Where " _
                + "Cliente.Cliente >= '" + Cliente.Text + "' AND " _
                + "Cliente.Cliente <= '" + Cliente.Text + "'"
        
        Uno = "{Cliente.Cliente} in " + Chr$(34) + Cliente.Text + Chr$(34) + " to " + Chr$(34) + Cliente.Text + Chr$(34)
        
        Listado.GroupSelectionFormula = Uno
        Listado.SelectionFormula = Uno
                
                
        ZZSuma = 0
        ZZCopia = Val(Copia.Text)
        Do
            If ZZCopia >= 3 Then
                ZZSuma = ZZSuma + 1
                ZZCopia = ZZCopia - 4
                    Else
                Exit Do
            End If
        Loop
                
                
        Listado.ReportFileName = "NotaEnvio.rpt"
        Listado.CopiesToPrinter = ZZSuma
        Listado.Destination = 1
        Listado.Destination = 0
        
        Listado.Action = 1
        
        
        If ZZCopia > 0 Then
            Listado.ReportFileName = "NotaEnvioII.rpt"
            Listado.CopiesToPrinter = 1
            Listado.Destination = 1
            Listado.Destination = 0
            
            Listado.Action = 1
        End If
        
    
    End If
    
    Call CmdLimpiar_Click
    a = s
    
End Sub

Private Sub Cancela_click()
    PrgNotaEnvio.Hide
    Unload Me
    Menu4.Show
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Trim(Cliente.Text) <> "" Then
            Auxi = UCase(Left$(Cliente.Text, 1))
            Auxi1 = Mid$(Cliente.Text, 2, 5)
            Call Ceros(Auxi1, 3)
            Cliente.Text = Auxi + "-" + Auxi1
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Razon.Text = Trim(rstCliente!Razon)
            Postal.Text = rstCliente!Postal
            ZZProvincia = rstCliente!Provincia
            Provincia.Text = ZProvincia(ZZProvincia)
            Direccion.Text = Trim(rstCliente!Direccion)
            Localidad.Text = Trim(rstCliente!Localidad)
            rstCliente.Close
            Bultos.SetFocus
        End If
        
    End If
    
    If KeyAscii = 27 Then
        Cliente.Text = ""
    End If
    
End Sub

Private Sub Bultos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Copia.SetFocus
    End If
    If KeyAscii = 27 Then
        Bultos.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Copia_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Bultos.SetFocus
    End If
    If KeyAscii = 27 Then
        Copia.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    ZProvincia(0) = "Capital Federal"
    ZProvincia(1) = "Buenos Aires"
    ZProvincia(2) = "Catamarca"
    ZProvincia(3) = "Cordoba"
    ZProvincia(4) = "Corrientes"
    ZProvincia(5) = "Chaco"
    ZProvincia(6) = "Chubut"
    ZProvincia(7) = "Entre Rios"
    ZProvincia(8) = "Formosa"
    ZProvincia(9) = "Jujuy"
    ZProvincia(10) = "La Pampa"
    ZProvincia(11) = "La Rioja"
    ZProvincia(12) = "Mendoza"
    ZProvincia(13) = "Misiones"
    ZProvincia(14) = "Neuquen"
    ZProvincia(15) = "Rio Negro"
    ZProvincia(16) = "Salta"
    ZProvincia(17) = "San Juan"
    ZProvincia(18) = "San Luis"
    ZProvincia(19) = "Santa Cruz"
    ZProvincia(20) = "Santa Fe"
    ZProvincia(21) = "Santiago del Estero"
    ZProvincia(22) = "Tucuman"
    ZProvincia(23) = "Tierra del Fuego"
    ZProvincia(24) = "Exterior"
    ZProvincia(25) = ""

    Cliente.Text = ""
    Razon.Text = ""
    Bultos.Text = ""
    Copia.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Provincia.Text = ""
    Postal.Text = ""
    
    Tipo.Clear
    
    Tipo.AddItem "Formulario"
    Tipo.AddItem "Laser"
    
    Tipo.ListIndex = 0
 
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
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Order by Cliente.Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = !Cliente + " " + !Razon
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
            
    Pantalla.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Indice = Pantalla.ListIndex
    Cliente.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    Call Cliente_KeyPress(13)
    
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
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Cliente + " " + !Razon
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

Private Sub Razon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Bultos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Copia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Direccion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Localidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Provincia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Postal_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 115
            Call Consulta_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub













