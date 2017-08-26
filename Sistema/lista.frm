VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form PrgLista 
   AutoRedraw      =   -1  'True
   Caption         =   "Lista de Precio"
   ClientHeight    =   5310
   ClientLeft      =   1050
   ClientTop       =   690
   ClientWidth     =   9960
   LinkTopic       =   "Form2"
   ScaleHeight     =   5310
   ScaleWidth      =   9960
   Visible         =   0   'False
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
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
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
         MouseIcon       =   "lista.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "lista.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
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
         MouseIcon       =   "lista.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "lista.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   25
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ListBox PantallaFiltrada 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "lista.frx":0E98
      Left            =   1560
      List            =   "lista.frx":0E9F
      TabIndex        =   27
      Top             =   2760
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo F8"
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
      Left            =   6960
      MouseIcon       =   "lista.frx":0EAD
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":11B7
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Salida"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Siguiente 
      Caption         =   "Siguien. F7"
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
      Left            =   6000
      MouseIcon       =   "lista.frx":15F9
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":1903
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Registro Siguiente"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "Anterior F6"
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
      Left            =   5040
      MouseIcon       =   "lista.frx":1D45
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":204F
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Registro Anterior"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primer F5"
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
      MouseIcon       =   "lista.frx":2491
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":279B
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Primer Registro"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
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
      Left            =   8880
      MouseIcon       =   "lista.frx":2BDD
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":2EE7
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Salida"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Lista 
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
      Left            =   7920
      MouseIcon       =   "lista.frx":3729
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":3A33
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Impresion "
      Top             =   1080
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
      Left            =   3120
      MouseIcon       =   "lista.frx":4275
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":457F
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Consulta de Datos"
      Top             =   1080
      Width           =   855
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
      Left            =   2160
      MouseIcon       =   "lista.frx":4DC1
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":50CB
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borra  F2"
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
      MouseIcon       =   "lista.frx":590D
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":5C17
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Elimina el Registro"
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
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
      MouseIcon       =   "lista.frx":6459
      MousePointer    =   99  'Custom
      Picture         =   "lista.frx":6763
      Style           =   1  'Graphical
      TabIndex        =   15
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
      Left            =   1560
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   6735
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
      Left            =   2760
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8880
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Fragancia.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Vendedor"
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
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
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
      Left            =   2760
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   4575
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
      Left            =   1800
      TabIndex        =   13
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
      Height          =   2460
      ItemData        =   "lista.frx":6FA5
      Left            =   1560
      List            =   "lista.frx":6FAC
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   6735
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
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2175
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
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PrgLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxi As String
Private Const WMaxHeight = 5970
Private Const WMinHeight = 2775

Sub Imprime_Nombre()
End Sub

Private Function Verifica_datos() As Boolean
    Dim Valido As Boolean
    Valido = True
    
     ' Se pide como minimo el codigo y una descripcion.
    If Trim(Codigo.Text) = "" Then Valido = False
    If Trim(Descripcion.Text) = "" Then Valido = False

    Verifica_datos = Valido
    
End Function

Sub Format_datos()
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Lista"
    ZSql = ZSql + " Where Lista.Codigo = " + "'" + Codigo.Text + "'"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstLista.RecordCount > 0 Then
        Descripcion.Text = Trim(rstLista!Descripcion)
        rstLista.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
End Sub

Private Sub Acepta_Click()

    ZZDesde = UCase(Trim(Desde.Text))
    ZZHasta = UCase(Trim(Hasta.Text))

    Rem If Val(Desde.Text) = 0 Then
    Rem      Desde.Text = "0"
    Rem End If
    If Trim(ZZHasta) = "" Then
         ZZHasta = "ZZZZZZZ"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Lista SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.ReportFileName = App.Path & "/WListadoPrecios.rpt" 'Cambiar el nombre por algo mas descriptivo.
    
    Listado.WindowTitle = "Listado de Listas de Precios"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Lista.Codigo, Lista.Descripcion, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Lista Lista, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Lista.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Lista.Codigo >= '" + ZZDesde + "' AND " _
                + "Lista.Codigo <= '" + ZZHasta + "'"
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Lista.Codigo} in " + Chr$(34) + ZZDesde + Chr$(34) + " to " + Chr(34) + ZZHasta + Chr$(34)
    Listado.SelectionFormula = "{Lista.Codigo} in " + Chr$(34) + ZZDesde + Chr$(34) + " to " + Chr(34) + ZZHasta + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    Me.Height = WMinHeight
    
End Sub

Private Sub Ayuda_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(Ayuda.Text) <> "" Then
        Dim WTextoABuscar, WTexto
        
        WTextoABuscar = Trim(Ayuda.Text)
        
        PantallaFiltrada.Clear
        
        For i = 0 To Pantalla.ListCount
            WTexto = Pantalla.List(i)
            
            If WTexto Like "*" & WTextoABuscar & "*" Or WTexto Like "*" & UCase(WTextoABuscar) & "*" Then
                PantallaFiltrada.AddItem WTexto
            End If
        Next
        
        PantallaFiltrada.Visible = True
    Else
        PantallaFiltrada.Visible = False
    End If
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
    Codigo.SetFocus
    Me.Height = WMinHeight
End Sub

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
    
        If Not Verifica_datos Then
            m$ = "Alguno de los datos no es correcto. Verifique y vuelva a intentarlo."
            aaaaaa% = MsgBox(m$, 0, "Archivo de Lista de Precios")
            Exit Sub
        End If
        
        Dim WCodigo, WDescripcion As String
        
        WCodigo = Left$(Codigo.Text, 4)
        WDescripcion = Left$(Descripcion.Text, 50)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Lista"
        ZSql = ZSql + " Where Lista.Codigo = " + "'" + WCodigo + "'"
        spLista = ZSql
        Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstLista.RecordCount > 0 Then
            rstLista.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Lista SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Lista ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        If ZZNivel = 0 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
   
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Lista"
        ZSql = ZSql + " Where Lista.Codigo = " + "'" + WCodigo + "'"
        spLista = ZSql
        Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstLista.RecordCount > 0 Then
            rstLista.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Lista SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Lista ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
   
        If ZZNivel = 0 Then
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            txtUserName = "SA"
            txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Lista de Precios")
        
        Call CmdLimpiar_Click
    
        Codigo.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Lista"
        ZSql = ZSql + " Where Lista.Codigo = " + "'" + Codigo.Text + "'"
        spLista = ZSql
        Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstLista.RecordCount > 0 Then
            rstLista.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuestaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE Lista"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spLista = ZSql
                Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
                    
                Call CmdLimpiar_Click
                
            End If
        End If
    
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    On Error GoTo WError
    
    Codigo.Text = ""
    Descripcion.Text = ""
    
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = True
    Impresora.Value = False
    Frame2.Visible = False
    
    Me.Height = WMinHeight
    
    Codigo.SetFocus
    
    Exit Sub
    
WError:

    Resume Next
        
    
End Sub

Private Sub CmdClose_Click()
    PrgLista.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = True
    Me.Height = WMaxHeight
    Impresora.Value = False
    
    Ayuda.Visible = False
    Pantalla.Visible = False
    PantallaFiltrada.Visible = False
    
    Frame2.Visible = True
    Desde.SetFocus
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
            ZSql = ZSql + " FROM Lista"
            ZSql = ZSql + " Where Lista.Codigo = " + "'" + Codigo.Text + "'"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
            If rstLista.RecordCount > 0 Then
                rstLista.Close
                Call Imprime_Datos
                    Else
                WLista = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WLista
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

     Opcion.AddItem "Listas"

     Rem Opcion.Visible = True
     
     Opcion.ListIndex = 0
     Call Opcion_Click
     
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
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lista"
            ZSql = ZSql + " Order by Lista.Descripcion"
            spLista = ZSql
            Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
            If rstLista.RecordCount > 0 Then
                With rstLista
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
                rstLista.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    PantallaFiltrada.Visible = False
    Frame2.Visible = False
    Me.Height = WMaxHeight
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
            Codigo.Text = WIndice.List(indice)
            Me.Height = WMinHeight
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    On Error GoTo WError
    
    Call CmdLimpiar_Click
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Linea) as [LineaMayor]"
    Rem ZSql = ZSql + " FROM Lista"
    Rem spLista = ZSql
    Rem Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstLista.RecordCount > 0 Then
    Rem     rstLista.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstLista!CodigoMayor), "0", rstLista!CodigoMayor)
    Rem     codigo.text = ZUltimo + 1
    Rem     rstLista.Close
    Rem End If
    
    Exit Sub
    
WError:

    Resume Next
        
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next
    
    ' No se ejecuta porque estoy probando filtrado dinamico.
    'If KeyAscii > 31 Then
    '    ZAyuda = Ayuda.Text + Chr$(KeyAscii)
    '        Else
    '    Select Case KeyAscii
    '        Case 27
    '           Ayuda.Text = ""
    '           ZAyuda = ""
    '        Case 8
    '            WEspacios = Len(Ayuda.Text)
    '            If WEspacios > 0 Then
    '               WEspacios = WEspacios - 1
    '               ZAyuda = Left$(Ayuda.Text, WEspacios)
    '            End If
    '        Case Else
    '            ZAyuda = Ayuda.Text
    '    End Select
    'End If
    'WEspacios = Len(ZAyuda)
    
    'XIndice = Opcion.ListIndex
    
    'Select Case XIndice
    '    Case 0
    '        ZSql = ""
    '        ZSql = ZSql + "Select *"
    '        ZSql = ZSql + " FROM Lista"
    '        ZSql = ZSql + " Where Lista.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
    '        ZSql = ZSql + " Order by Lista.Descripcion"
    '        spLista = ZSql
    '        Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    '        If rstLista.RecordCount > 0 Then
    '            With rstLista
    '                .MoveFirst
    '                Do
    '                    If .EOF = False Then
    '                        IngresaItem = !Codigo + " " + !Descripcion
    '                        Pantalla.AddItem IngresaItem
    '                        IngresaItem = !Codigo
    '                        WIndice.AddItem IngresaItem
    '                        .MoveNext
    '                            Else
    '                        Exit Do
    '                    End If
    '                Loop
    '            End With
    '            rstLista.Close
    '        End If
    '
    '    Case Else
    'End Select

End Sub

Private Sub Codigo_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Listas"

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
    ZSql = ZSql + " FROM Lista"
    ZSql = ZSql + " Where Lista.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Lista.Codigo"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstLista.RecordCount > 0 Then
        With rstLista
            .MoveLast
            Codigo.Text = rstLista!Codigo
        End With
        rstLista.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Listas")
    End If
End Sub

Private Sub PantallaFiltrada_Click()
    WIndice = PantallaFiltrada.ListIndex
    WTexto = PantallaFiltrada.List(PantallaFiltrada.ListIndex)
    For i = o To Pantalla.ListCount
        If UCase(Pantalla.List(i)) = UCase(WTexto) Then
            
            Pantalla.ListIndex = i
            
            Call Pantalla_Click
            
            PantallaFiltrada.Visible = False
            
            Exit Sub
        End If
    Next
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM Lista"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstLista.RecordCount > 0 Then
        rstLista.MoveFirst
        ZUltimo = IIf(IsNull(rstLista!CodigoMenor), "", rstLista!CodigoMenor)
        Codigo.Text = ZUltimo
        rstLista.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Lista"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstLista.RecordCount > 0 Then
        rstLista.MoveLast
        ZUltimo = IIf(IsNull(rstLista!CodigoMayor), "", rstLista!CodigoMayor)
        Codigo.Text = ZUltimo
        rstLista.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Lista"
    ZSql = ZSql + " Where Lista.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Lista.Codigo"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstLista.RecordCount > 0 Then
        With rstLista
            .MoveFirst
            Codigo.Text = rstLista!Codigo
        End With
        rstLista.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Listas")
    End If
End Sub


































