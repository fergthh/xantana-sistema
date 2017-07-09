VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCtaCteCampana 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Clientes por Campaña"
   ClientHeight    =   6840
   ClientLeft      =   2490
   ClientTop       =   585
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
   ScaleWidth      =   7215
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
      Top             =   3960
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Frame Frame2 
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
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
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
         Left            =   4920
         MouseIcon       =   "CtaCteCampana.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteCampana.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salida"
         Top             =   2400
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
         Left            =   2520
         MouseIcon       =   "CtaCteCampana.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteCampana.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2400
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
         Left            =   3720
         MouseIcon       =   "CtaCteCampana.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteCampana.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Consulta de Datos"
         Top             =   2400
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
         Left            =   1320
         MouseIcon       =   "CtaCteCampana.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteCampana.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox HastaCli 
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
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   2
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox DesdeCli 
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
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   1
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   3480
         TabIndex        =   4
         Top             =   1800
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   285
         Left            =   3480
         TabIndex        =   3
         Top             =   1440
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   3480
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
         Caption         =   "Fecha"
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
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Left            =   1440
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   1440
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
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
         Left            =   1440
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
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
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impctacte.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Clientes"
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
      Left            =   4560
      TabIndex        =   6
      Top             =   5280
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
      Height          =   2205
      ItemData        =   "CtaCteCampana.frx":2D30
      Left            =   120
      List            =   "CtaCteCampana.frx":2D37
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   6735
   End
End
Attribute VB_Name = "PrgCtaCteCampana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WPasa As String
Private WTitulo As String
Private Importe3 As Double

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
    
    If DesdeFecha.Text = "  /  /    " Then
        DesdeFecha.Text = "00/00/0000"
    End If
    If HastaFecha.Text = "  /  /    " Then
        HastaFecha.Text = "99/99/9999"
    End If
    
    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia
    

    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Varios = " + "'" + Fecha.Text + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZBase2 = DateDiff("d", "01/01/2000", Fecha.Text)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "',"
    ZSql = ZSql + " Base2 = " + "'" + Str$(ZZBase2) + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)

    Listado.WindowTitle = "Listado de Cuenta Corriente de Clientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Val(DesdeCli.Text) = 0 Then
        DesdeCli.Text = "0"
    End If
    If Val(HastaCli.Text) = 0 Then
        HastaCli.Text = "0"
    End If

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.Total, CtaCte.Saldo, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Partida, CtaCte.Campana, CtaCte.Base1, CtaCte.Base2, CtaCte.Linea, " _
                + "Cliente.Razon, Cliente.Telefono, Cliente.Vendedor, " _
                + "Auxiliar.Nombre, Auxiliar.Varios " _
                + "From " _
                + DSQ + ".dbo.CtaCte CtaCte, " _
                + DSQ + ".dbo.Cliente Cliente, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "CtaCte.Cliente = Cliente.Cliente AND " _
                + "CtaCte.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "CtaCte.Cliente >= " + DesdeCli.Text + " AND " _
                + "CtaCte.Cliente <= " + HastaCli.Text + " AND " _
                + "CtaCte.Saldo <> 0 AND " _
                + "CtaCte.OrdFecha >= '" + WDesde + "' AND " _
                + "CtaCte.OrdFecha <= '" + WHasta + "' AND " _
                + "CtaCte.Campana <> ''"
        
    Uno = "{CtaCte.Saldo} <> 0 "
    Dos = " and {CtaCte.Cliente} in " + DesdeCli.Text + " to " + HastaCli.Text
    Tres = " and {CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
        
    Listado.ReportFileName = "CtactePendCampana.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgCtaCteCampana.Hide
    Unload Me
    Menu.Show
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
                    IngresaItem = Str$(!Cliente) + " " + !Razon
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
    DesdeCli.Text = WIndice.List(Indice)
    HastaCli.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    DesdeCli.SetFocus
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            DesdeCli.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub DesdeCli_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCli.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeCli.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaCli_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Hastaclie.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
                Else
            DesdeFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFecha.Text = "  /  /    "
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Fecha.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Fecha.Text = "  /  /    "
    DesdeCli.Text = ""
    HastaCli.Text = ""
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Frame2.Visible = True
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
                            IngresaItem = Str$(rstCliente!Cliente) + " " + rstCliente!Razon
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
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
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

Private Sub DesdeCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaCli_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdeFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFecha_KeyDown(KeyCode As Integer, Shift As Integer)
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



