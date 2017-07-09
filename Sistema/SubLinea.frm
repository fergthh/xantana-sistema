VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgSubLinea 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de SubLineas de Ventas"
   ClientHeight    =   5355
   ClientLeft      =   945
   ClientTop       =   705
   ClientWidth     =   9765
   LinkTopic       =   "Form2"
   ScaleHeight     =   5355
   ScaleWidth      =   9765
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
      Left            =   6840
      MouseIcon       =   "SubLinea.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":030A
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
      Left            =   5880
      MouseIcon       =   "SubLinea.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":0A56
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
      Left            =   4920
      MouseIcon       =   "SubLinea.frx":0E98
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":11A2
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
      Left            =   3960
      MouseIcon       =   "SubLinea.frx":15E4
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":18EE
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
      Left            =   8760
      MouseIcon       =   "SubLinea.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":203A
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
      Left            =   7800
      MouseIcon       =   "SubLinea.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":2B86
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
      Left            =   3000
      MouseIcon       =   "SubLinea.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":36D2
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
      Left            =   2040
      MouseIcon       =   "SubLinea.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":421E
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
      Left            =   1080
      MouseIcon       =   "SubLinea.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":4D6A
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
      Left            =   120
      MouseIcon       =   "SubLinea.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "SubLinea.frx":58B6
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
      Left            =   240
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.TextBox SubLinea 
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
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
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
         MouseIcon       =   "SubLinea.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "SubLinea.frx":6402
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
         MouseIcon       =   "SubLinea.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "SubLinea.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Hasta 
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
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   720
         Width           =   855
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
   Begin Crystal.CrystalReport Listado 
      Left            =   6600
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "SubLinea.rpt"
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
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Nombre 
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
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   3375
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
      Left            =   240
      TabIndex        =   13
      Top             =   2640
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
      ItemData        =   "SubLinea.frx":6F90
      Left            =   240
      List            =   "SubLinea.frx":6F97
      TabIndex        =   4
      Top             =   2640
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de SubLinea"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "PrgSubLinea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxi As String

Sub Imprime_Nombre()
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM SubLineas"
    ZSql = ZSql + " Where SubLineas.SubLinea = " + "'" + SubLinea.Text + "'"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstSubLinea.RecordCount > 0 Then
        Nombre.Text = Trim(rstSubLinea!Nombre)
        rstSubLinea.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
End Sub

Private Sub Acepta_Click()

    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE SubLineas SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de SubLineas de Venta"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT SubLineas.SubLinea, SubLineas.Nombre, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.SubLineas SubLineas, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "SubLineas.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "SubLineas.SubLinea >= " + Desde.Text + " AND " _
                + "SubLineas.SubLinea <= " + Hasta.Text
    
    Listado.GroupSelectionFormula = "{SubLineas.SubLinea} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{SubLineas.SubLinea} in " + Desde.Text + " to " + Hasta.Text
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    SubLinea.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
    SubLinea.SetFocus
End Sub

Private Sub cmdAdd_Click()
    If SubLinea.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM SubLineas"
        ZSql = ZSql + " Where SubLineas.SubLinea = " + "'" + SubLinea.Text + "'"
        spSubLinea = ZSql
        Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstSubLinea.RecordCount > 0 Then
            rstSubLinea.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE SubLineas SET "
            ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "'"
            ZSql = ZSql + " Where SubLinea = " + "'" + SubLinea.Text + "'"
            spSubLinea = ZSql
            Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO SubLineas ("
            ZSql = ZSql + "SubLinea ,"
            ZSql = ZSql + "Nombre )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + SubLinea.Text + "',"
            ZSql = ZSql + "'" + Nombre.Text + "')"
            spSubLinea = ZSql
            Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        If Val(WEmpresa) = 1 Then
            txtOdbc = "YenadiII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            txtOdbc = "Yenadi"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM SubLineas"
        ZSql = ZSql + " Where SubLineas.SubLinea = " + "'" + SubLinea.Text + "'"
        spSubLinea = ZSql
        Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstSubLinea.RecordCount > 0 Then
            rstSubLinea.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE SubLineas SET "
            ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "'"
            ZSql = ZSql + " Where SubLinea = " + "'" + SubLinea.Text + "'"
            spSubLinea = ZSql
            Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO SubLineas ("
            ZSql = ZSql + "SubLinea ,"
            ZSql = ZSql + "Nombre )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + SubLinea.Text + "',"
            ZSql = ZSql + "'" + Nombre.Text + "')"
            spSubLinea = ZSql
            Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        If Val(WEmpresa) = 1 Then
            txtOdbc = "Yenadi"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            txtOdbc = "YenadiII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        Call CmdLimpiar_Click
        SubLinea.SetFocus
        
    End If
    
End Sub

Private Sub CmdDelete_Click()
    If SubLinea.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM SubLineas"
        ZSql = ZSql + " Where SubLineas.SubLinea = " + "'" + SubLinea.Text + "'"
        spSubLinea = ZSql
        Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstSubLinea.RecordCount > 0 Then
            rstSubLinea.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE SubLineas"
                ZSql = ZSql + " Where SubLinea = " + "'" + SubLinea.Text + "'"
                spSubLinea = ZSql
                Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
                
                If Val(WEmpresa) = 1 Then
                    txtOdbc = "YenadiII"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Else
                    txtOdbc = "Yenadi"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End If
            
                ZSql = ""
                ZSql = ZSql + "DELETE SubLineas"
                ZSql = ZSql + " Where SubLinea = " + "'" + SubLinea.Text + "'"
                spSubLinea = ZSql
                Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    
                If Val(WEmpresa) = 1 Then
                    txtOdbc = "Yenadi"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Else
                    txtOdbc = "YenadiII"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End If
                
                Call CmdLimpiar_Click
            End If
        End If
    
    End If
    SubLinea.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    On Error GoTo WError
    
    SubLinea.Text = "1"
    Nombre.Text = ""
    SubLinea.SetFocus
    
    ZSql = ""
    ZSql = ZSql + "Select Max(SubLinea) as [SubLineaMayor]"
    ZSql = ZSql + " FROM SubLineas"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstSubLinea.RecordCount > 0 Then
        rstSubLinea.MoveLast
        ZUltimo = IIf(IsNull(rstSubLinea!SubLineaMayor), "0", rstSubLinea!SubLineaMayor)
        SubLinea.Text = ZUltimo + 1
        rstSubLinea.Close
    End If
    
    Exit Sub
    
WError:

    Resume Next
        
    
End Sub

Private Sub cmdClose_Click()
    PrgSubLinea.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    Rem If KeyAscii = 13 Then
    Rem     Cuenta.SetFocus
    Rem End If
    If KeyAscii = 27 Then
        Nombre.Text = ""
    End If
End Sub

Private Sub SubLinea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If SubLinea.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM SubLineas"
            ZSql = ZSql + " Where SubLineas.SubLinea = " + "'" + SubLinea.Text + "'"
            spSubLinea = ZSql
            Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstSubLinea.RecordCount > 0 Then
                rstSubLinea.Close
                Call Imprime_Datos
                    Else
                WSubLinea = SubLinea.Text
                CmdLimpiar_Click
                SubLinea.Text = WSubLinea
            End If
        End If
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        SubLinea.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "SubLineas de Ventas"

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
            ZSql = ZSql + " FROM SubLineas"
            ZSql = ZSql + " Order by SubLineas.SubLinea"
            spSubLinea = ZSql
            Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstSubLinea.RecordCount > 0 Then
                With rstSubLinea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!SubLinea) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !SubLinea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSubLinea.Close
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
            Indice = Pantalla.ListIndex
            SubLinea.Text = WIndice.List(Indice)
            Call SubLinea_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    On Error GoTo WError
    
    SubLinea.Text = "1"
    Nombre.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(SubLinea) as [SubLineaMayor]"
    ZSql = ZSql + " FROM SubLineas"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstSubLinea.RecordCount > 0 Then
        rstSubLinea.MoveLast
        ZUltimo = IIf(IsNull(rstSubLinea!SubLineaMayor), "0", rstSubLinea!SubLineaMayor)
        SubLinea.Text = ZUltimo + 1
        rstSubLinea.Close
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
            ZSql = ZSql + " FROM SubLineas"
            ZSql = ZSql + " Where SubLineas.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by SubLineas.SubLinea"
            spSubLinea = ZSql
            Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstSubLinea.RecordCount > 0 Then
                With rstSubLinea
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!SubLinea) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !SubLinea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSubLinea.Close
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

Private Sub SubLinea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "SubLineas de Ventas"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub



Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub SubLinea_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call CmdDelete_Click
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
            Call cmdClose_Click
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
    ZSql = ZSql + " FROM SubLineas"
    ZSql = ZSql + " Where SubLineas.SubLinea < " + "'" + SubLinea.Text + "'"
    ZSql = ZSql + " Order by SubLineas.SubLinea"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstSubLinea.RecordCount > 0 Then
        With rstSubLinea
            .MoveLast
            SubLinea.Text = rstSubLinea!SubLinea
        End With
        rstSubLinea.Close
        Call Imprime_Datos
        SubLinea.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de SubLineas de Venta")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(SubLinea) as [SubLineaMenor]"
    ZSql = ZSql + " FROM SubLineas"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstSubLinea.RecordCount > 0 Then
        rstSubLinea.MoveFirst
        ZUltimo = IIf(IsNull(rstSubLinea!SubLineaMenor), "0", rstSubLinea!SubLineaMenor)
        SubLinea.Text = ZUltimo
        rstSubLinea.Close
        Call Imprime_Datos
        SubLinea.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(SubLinea) as [SubLineaMayor]"
    ZSql = ZSql + " FROM SubLineas"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstSubLinea.RecordCount > 0 Then
        rstSubLinea.MoveLast
        ZUltimo = IIf(IsNull(rstSubLinea!SubLineaMayor), "0", rstSubLinea!SubLineaMayor)
        SubLinea.Text = ZUltimo
        rstSubLinea.Close
        Call Imprime_Datos
        SubLinea.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM SubLineas"
    ZSql = ZSql + " Where SubLineas.SubLinea > " + "'" + SubLinea.Text + "'"
    ZSql = ZSql + " Order by SubLineas.SubLinea"
    spSubLinea = ZSql
    Set rstSubLinea = db.OpenRecordset(spSubLinea, dbOpenSnapshot, dbSQLPassThrough)
    If rstSubLinea.RecordCount > 0 Then
        With rstSubLinea
            .MoveFirst
            SubLinea.Text = rstSubLinea!SubLinea
        End With
        rstSubLinea.Close
        Call Imprime_Datos
        SubLinea.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de SubLineas de Venta")
    End If
End Sub


































