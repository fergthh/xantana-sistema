VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form PrgTipoClie 
   AutoRedraw      =   -1  'True
   Caption         =   "Tipo de Clientes"
   ClientHeight    =   5175
   ClientLeft      =   1050
   ClientTop       =   690
   ClientWidth     =   11415
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   11415
   Begin VB.TextBox Descripcion 
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
      Left            =   5520
      MaxLength       =   50
      TabIndex        =   24
      Top             =   345
      Width           =   4575
   End
   Begin VB.TextBox Codigo 
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
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   23
      Text            =   " "
      Top             =   345
      Width           =   975
   End
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo (F8)"
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
      Left            =   7920
      MouseIcon       =   "tipoclie.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":030A
      TabIndex        =   20
      ToolTipText     =   "Salida"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Siguiente 
      Caption         =   "Siguiente (F7)"
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
      Left            =   6840
      MouseIcon       =   "tipoclie.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":0A56
      TabIndex        =   19
      ToolTipText     =   "Registro Siguiente"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "Anterior (F6)"
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
      Left            =   5880
      MouseIcon       =   "tipoclie.frx":0E98
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":11A2
      TabIndex        =   18
      ToolTipText     =   "Registro Anterior"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primer (F5)"
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
      MouseIcon       =   "tipoclie.frx":15E4
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":18EE
      TabIndex        =   17
      ToolTipText     =   "Primer Registro"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu (F10)"
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
      Left            =   9840
      MouseIcon       =   "tipoclie.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":203A
      TabIndex        =   16
      ToolTipText     =   "Salida"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Lista 
      Caption         =   "Listado (F9)"
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
      Left            =   8880
      MouseIcon       =   "tipoclie.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":2B86
      TabIndex        =   15
      ToolTipText     =   "Impresion "
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consultar (F4)"
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
      MouseIcon       =   "tipoclie.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":36D2
      TabIndex        =   14
      ToolTipText     =   "Consulta de Datos"
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar (F3)"
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
      Left            =   2880
      MouseIcon       =   "tipoclie.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":421E
      TabIndex        =   13
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borrar (F2)"
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
      Left            =   1920
      MouseIcon       =   "tipoclie.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":4D6A
      TabIndex        =   12
      ToolTipText     =   "Elimina el Registro"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Grabar (F1)"
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
      Left            =   960
      MouseIcon       =   "tipoclie.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "tipoclie.frx":58B6
      TabIndex        =   11
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1080
      Width           =   855
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
      Left            =   2340
      TabIndex        =   10
      Top             =   2400
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
      Height          =   2055
      Left            =   3120
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela F12"
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
         Left            =   4080
         MouseIcon       =   "tipoclie.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "tipoclie.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Confirma F11"
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
         MouseIcon       =   "tipoclie.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "tipoclie.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   8
         Text            =   " "
         Top             =   720
         Width           =   855
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   7
         Text            =   " "
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   4
         Top             =   720
         Width           =   1335
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
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10320
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Listado de Vendedor"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9240
      TabIndex        =   1
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
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
      Height          =   1980
      Left            =   3900
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   3615
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
      Height          =   2220
      ItemData        =   "tipoclie.frx":6F90
      Left            =   2340
      List            =   "tipoclie.frx":6F97
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo "
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
      Index           =   0
      Left            =   1320
      TabIndex        =   26
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion "
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
      Index           =   1
      Left            =   3600
      TabIndex        =   25
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "PrgTipoClie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxi As String
Private Const WMaxHeight = 5690
Private Const WMinHeight = 2790
Private WDatosValidos As Boolean

Sub Imprime_Nombre()
End Sub

Sub Verifica_datos()
    ' Comprobamos que hayan datos que guardar
    If Trim(Codigo.Text) = "" Or Trim(Descripcion.Text) = "" Then
    
        MsgBox "Todos los datos son obligatorios.", vbInformation
        WDatosValidos = False
        Codigo.SetFocus
        Exit Sub
    End If
End Sub

Sub Format_datos()
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoClie"
    ZSql = ZSql + " Where TipoClie.Codigo = " + "'" + Codigo.Text + "'"
    spTipoClie = ZSql
    Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoClie.RecordCount > 0 Then
        Descripcion.Text = Trim(rstTipoClie!Descripcion)
        rstTipoClie.Close
        'Call Format_datos
        'Call Imprime_Nombre
    End If
End Sub

Private Sub Acepta_Click()

    ZZDesde = Desde.Text
    ZZHasta = Hasta.Text

    If Trim(ZZDesde) = "" Then
         ZZDesde = ""
    End If
    
    If Trim(ZZHasta) = "" Then
         ZZHasta = "ZZZZZZ"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE TipoClie SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spTipoClie = ZSql
    Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
        
    Listado.WindowTitle = "Listado de Tipo de Clientes"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.ReportFileName = App.Path + "\tipocliente.rpt"
    
    Listado.SQLQuery = "SELECT TipoClie.Codigo, TipoClie.Descripcion " _
                + "From " _
                + "TipoClie " _
                + "Where " _
                + "TipoClie.Codigo >= '" + ZZDesde + "' AND " _
                + "TipoClie.Codigo <= '" + ZZHasta + "'"
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{TipoClie.Codigo} in " + Chr$(34) + ZZDesde + Chr$(34) + " to " + Chr(34) + ZZHasta + Chr$(34)
    Listado.SelectionFormula = "{TipoClie.Codigo} in " + Chr$(34) + ZZDesde + Chr$(34) + " to " + Chr(34) + ZZHasta + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    Call ContraerFormulario
    
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
    Codigo.SetFocus
End Sub

Private Sub cmdAdd_Click()
        
    Call Verifica_datos
    
    Dim XCodigo, XDescripcion As String
    XCodigo = Trim(Codigo.Text)
    XDescripcion = Trim(Descripcion.Text)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoClie"
        ZSql = ZSql + " Where TipoClie.Codigo = " + "'" + XCodigo + "'"
        spTipoClie = ZSql
        Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoClie.RecordCount > 0 Then
            rstTipoClie.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE TipoClie SET "
            ZSql = ZSql + " Descripcion = " + "'" + XDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + XCodigo + "'"
            spTipoClie = ZSql
            Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO TipoClie ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + XCodigo + "',"
            ZSql = ZSql + "'" + XDescripcion + "')"
            spTipoClie = ZSql
            Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Tipo de Clientes")
        
        Call CmdLimpiar_Click
    
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoClie"
        ZSql = ZSql + " Where TipoClie.Codigo = " + "'" + Codigo.Text + "'"
        spTipoClie = ZSql
        Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoClie.RecordCount > 0 Then
            rstTipoClie.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE TipoClie"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spTipoClie = ZSql
                Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
                    
                Call CmdLimpiar_Click
                
            End If
        End If
    
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    On Error GoTo WError
    
    
    WDatosValidos = True
    
    Codigo.Text = ""
    Descripcion.Text = ""
    
    Call ContraerFormulario
    
    Codigo.SetFocus
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Linea) as [LineaMayor]"
    Rem ZSql = ZSql + " FROM TipoClie"
    Rem spTipoClie = ZSql
    Rem Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstTipoClie.RecordCount > 0 Then
    Rem     rstTipoClie.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstTipoClie!CodigoMayor), "0", rstTipoClie!CodigoMayor)
    Rem     codigo.text = ZUltimo + 1
    Rem     rstTipoClie.Close
    Rem End If
    
    Exit Sub
    
WError:

    Resume Next
        
    
End Sub

Private Sub CmdClose_Click()
    PrgTipoClie.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Descripcion_DblClick()
    ' Abrimos la consulta tambien cuando de haga doble click sobre la descripcion independientemente si tiene o no contenido.
    Call AbrirConsulta
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
    Call ExpandirFormulario
End Sub

Private Sub Descripcion_KeyPress(KeyAscii As Integer)
    Rem If KeyAscii = 13 Then
    Rem     Cuenta.SetFocus
    Rem End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            Codigo.Text = UCase(Codigo.Text)
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoClie"
            ZSql = ZSql + " Where TipoClie.Codigo = " + "'" + Codigo.Text + "'"
            spTipoClie = ZSql
            Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoClie.RecordCount > 0 Then
                rstTipoClie.Close
                Call Imprime_Datos
                    Else
                WTipoClie = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WTipoClie
            End If
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Tipo de Clientes"

     Rem Opcion.Visible = True
     
     Opcion.ListIndex = 0
     Call Opcion_Click
     
     Call ExpandirFormulario
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select Codigo, Descripcion"
            ZSql = ZSql + " FROM TipoClie"
            ZSql = ZSql + " Order by TipoClie.Descripcion"
            spTipoClie = ZSql
            Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoClie.RecordCount > 0 Then
                With rstTipoClie
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Trim(!Codigo) + " " + Trim(!Descripcion)
                            Pantalla.AddItem IngresaItem
                            IngresaItem = Trim(!Codigo)
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoClie.Close
            End If
            
        Case Else
    End Select
            
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
    Select Case XIndice
        Case 0
            indice = Pantalla.ListIndex
            Codigo.Text = Trim(WIndice.List(indice))
            Call Codigo_KeyPress(13)
            Call ContraerFormulario
        Case Else
    End Select
    
End Sub

Private Sub ContraerFormulario()
    Frame2.Visible = False
    Me.Height = WMinHeight
End Sub

Private Sub ExpandirFormulario()
    Me.Height = WMaxHeight
End Sub

Sub Form_Load()

    On Error GoTo WError
    
    Call CmdLimpiar_Click
    
     ZSql = ""
     ZSql = ZSql + "Select Codigo"
     ZSql = ZSql + " FROM TipoClie"
     spTipoClie = ZSql
     Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
     If rstTipoClie.RecordCount > 0 Then
         rstTipoClie.MoveLast
         ZUltimo = IIf(IsNull(rstTipoClie!Codigo), "0", rstTipoClie!Codigo)
         Codigo.Text = ZUltimo + 1
         rstTipoClie.Close
     End If
    
    Exit Sub
    
WError:

    Resume Next
        
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
            ZSql = ZSql + " FROM TipoClie"
            ZSql = ZSql + " Where TipoClie.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by TipoClie.Descripcion"
            spTipoClie = ZSql
            Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoClie.RecordCount > 0 Then
                With rstTipoClie
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoClie.Close
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

Private Sub Codigo_DblClick()

    Call AbrirConsulta

End Sub

Private Sub AbrirConsulta()
    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Tipo de Clientes"

    Opcion.ListIndex = 0
    
    Call Opcion_Click
End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Panta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impresora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 113
            Call cmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call Primer_Click
        Case 117
            Call Anterior_Click
        Case 118
            Call Siguiente_Click
        Case 119
            Call Ultimo_Click
        Case 120
            Call Lista_Click
        Case 121
            Call CmdClose_Click
        Case 122
            Call Acepta_Click
        Case 123
            Call Cancela_Click
        Case Else
    End Select
End Sub


Private Sub Anterior_Click()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoClie"
    ZSql = ZSql + " Where TipoClie.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by TipoClie.Codigo"
    spTipoClie = ZSql
    Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoClie.RecordCount > 0 Then
        With rstTipoClie
            .MoveLast
            Codigo.Text = rstTipoClie!Codigo
        End With
        rstTipoClie.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Tipo de Clientes")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM TipoClie"
    spTipoClie = ZSql
    Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoClie.RecordCount > 0 Then
        rstTipoClie.MoveFirst
        ZUltimo = IIf(IsNull(rstTipoClie!CodigoMenor), "", rstTipoClie!CodigoMenor)
        Codigo.Text = ZUltimo
        rstTipoClie.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM TipoClie"
    spTipoClie = ZSql
    Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoClie.RecordCount > 0 Then
        rstTipoClie.MoveLast
        ZUltimo = IIf(IsNull(rstTipoClie!CodigoMayor), "", rstTipoClie!CodigoMayor)
        Codigo.Text = ZUltimo
        rstTipoClie.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoClie"
    ZSql = ZSql + " Where TipoClie.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by TipoClie.Codigo"
    spTipoClie = ZSql
    Set rstTipoClie = db.OpenRecordset(spTipoClie, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoClie.RecordCount > 0 Then
        With rstTipoClie
            .MoveFirst
            Codigo.Text = rstTipoClie!Codigo
        End With
        rstTipoClie.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Tipo de Clientes")
    End If
End Sub


































