VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgConcepto 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Conceptos de Compra"
   ClientHeight    =   6195
   ClientLeft      =   1125
   ClientTop       =   975
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   ScaleHeight     =   6195
   ScaleWidth      =   9840
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
      MouseIcon       =   "concepto.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Salida"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Registro Siguiente"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":0E98
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Registro Anterior"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":15E4
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":18EE
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Primer Registro"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":203A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Salida"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":2B86
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Impresion "
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":36D2
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Consulta de Datos"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":421E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":4D6A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Elimina el Registro"
      Top             =   1560
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
      MouseIcon       =   "concepto.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "concepto.frx":58B6
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1560
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1680
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   5415
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
         Left            =   3240
         MouseIcon       =   "concepto.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "concepto.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
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
         Left            =   4320
         MouseIcon       =   "concepto.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "concepto.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   28
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   13
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   12
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
         Left            =   1920
         TabIndex        =   11
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
         TabIndex        =   10
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
         Left            =   240
         TabIndex        =   9
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
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.TextBox Cuenta 
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
      MaxLength       =   20
      TabIndex        =   2
      Text            =   " "
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Concepto 
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
   Begin Crystal.CrystalReport Listado 
      Left            =   6840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Conceptos.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Conceptos de Compra"
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
      Left            =   5400
      TabIndex        =   6
      Top             =   120
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
      Height          =   2940
      ItemData        =   "concepto.frx":6F90
      Left            =   240
      List            =   "concepto.frx":6F97
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   7935
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
      Width           =   5535
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
      Height          =   1740
      Left            =   1080
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label DesCuenta 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   5520
      TabIndex        =   15
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Cuenta Contable"
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
      TabIndex        =   14
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion"
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
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Conceptos"
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
      TabIndex        =   3
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Imprime_Nombre()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cuenta"
    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        DesCuenta.Caption = rstCuenta!Descripcion
        rstCuenta.Close
            Else
        DesCuenta.Caption = ""
    End If
End Sub

Sub Verifica_datos()
    If Val(Concepto.Text) = 0 Then
        Concepto.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        Nombre.Text = Trim(rstConceptos!Nombre)
        Cuenta.Text = Trim(rstConceptos!Cuenta)
        rstConceptos.Close
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
    ZSql = ZSql + "UPDATE Conceptos SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Conceptos de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
            
    Listado.SQLQuery = "SELECT Conceptos.Concepto, Conceptos.Nombre, Conceptos.Cuenta, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Conceptos Conceptos, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Conceptos.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Conceptos.Concepto >= " + Desde.Text + " AND " _
                + "Conceptos.Concepto <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Conceptos.Concepto} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{Conceptos.Concepto} in " + Desde.Text + " to " + Hasta.Text
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Concepto.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()


    If Concepto.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Conceptos"
        ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
        spConceptos = ZSql
        Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        If rstConceptos.RecordCount > 0 Then
            rstConceptos.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Conceptos SET "
            ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "',"
            ZSql = ZSql + " Cuenta = " + "'" + Cuenta.Text + "'"
            ZSql = ZSql + " Where Concepto = " + "'" + Concepto.Text + "'"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Conceptos ("
            ZSql = ZSql + "Concepto ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Cuenta )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Concepto.Text + "',"
            ZSql = ZSql + "'" + Nombre.Text + "',"
            ZSql = ZSql + "'" + Cuenta.Text + "')"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
    
    
        If ZZNivel = 0 Then
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
    
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Conceptos"
        ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
        spConceptos = ZSql
        Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        If rstConceptos.RecordCount > 0 Then
            rstConceptos.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Conceptos SET "
            ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "',"
            ZSql = ZSql + " Cuenta = " + "'" + Cuenta.Text + "'"
            ZSql = ZSql + " Where Concepto = " + "'" + Concepto.Text + "'"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Conceptos ("
            ZSql = ZSql + "Concepto ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Cuenta )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Concepto.Text + "',"
            ZSql = ZSql + "'" + Nombre.Text + "',"
            ZSql = ZSql + "'" + Cuenta.Text + "')"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
    
    
        
        If ZZNivel = 0 Then
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        
        
        
        
    
    
    
    
    
    
    
    
    
        Call CmdLimpiar_Click
        Concepto.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Concepto.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Conceptos"
        ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
        spConceptos = ZSql
        Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        If rstConceptos.RecordCount > 0 Then
            rstConceptos.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE Conceptos"
                ZSql = ZSql + " Where Concepto = " + "'" + Concepto.Text + "'"
                spConceptos = ZSql
                Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
        
    
    
        If ZZNivel = 0 Then
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
    
        
        
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Conceptos"
        ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
        spConceptos = ZSql
        Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        If rstConceptos.RecordCount > 0 Then
            rstConceptos.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE Conceptos"
                ZSql = ZSql + " Where Concepto = " + "'" + Concepto.Text + "'"
                spConceptos = ZSql
                Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
        
    
    
        
        If ZZNivel = 0 Then
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "Fragancias"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Else
            Rem txtUserName = "SA"
            Rem txtPassword = "Sw58125812"
            txtOdbc = "FraganciasII"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End If
        
        
        
        
        
        
        
        
    End If
    Concepto.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Concepto.Text = ""
    Nombre.Text = ""
    Cuenta.Text = ""
    DesCuenta.Caption = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Concepto) as [ConceptoMayor]"
    ZSql = ZSql + " FROM Conceptos"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        rstConceptos.MoveLast
        ZUltimo = IIf(IsNull(rstConceptos!ConceptoMayor), "0", rstConceptos!ConceptoMayor)
        Concepto.Text = ZUltimo + 1
        rstConceptos.Close
    End If
    
    Concepto.SetFocus
    
End Sub

Private Sub CmdClose_Click()
    PrgConcepto.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cuenta.SetFocus
    End If
    If KeyAscii = 27 Then
        Nombre.Text = ""
    End If
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            DesCuenta.Caption = IIf(IsNull(rstCuenta!Descripcion), "", rstCuenta!Descripcion)
            rstCuenta.Close
            Nombre.SetFocus
                Else
            DesCuenta.Caption = ""
            Cuenta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cuenta.Text = ""
        DesCuenta.Caption = ""
    End If
End Sub

Private Sub Concepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Concepto.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                rstConceptos.Close
                Call Imprime_Datos
                    Else
                WConcepto = Concepto.Text
                CmdLimpiar_Click
                Concepto.Text = WConcepto
            End If
        End If
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        Concepto.Text = ""
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

     Pantalla.Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Conceptos de Compras"
     Opcion.AddItem "Cuentas Contables"

     Opcion.Visible = True
     
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
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Order by Conceptos.Concepto"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                With rstConceptos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Concepto) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Concepto
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstConceptos.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Order by Cuenta.Cuenta"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                With rstCuenta
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Cuenta + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Cuenta
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCuenta.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
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
            Concepto.Text = WIndice.List(Indice)
            Call Concepto_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Cuenta.Text = WIndice.List(Indice)
            Call Cuenta_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    Concepto.Text = ""
    Nombre.Text = ""
    Cuenta.Text = ""
    DesCuenta.Caption = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Concepto) as [ConceptoMayor]"
    ZSql = ZSql + " FROM Conceptos"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        rstConceptos.MoveLast
        ZUltimo = IIf(IsNull(rstConceptos!ConceptoMayor), "0", rstConceptos!ConceptoMayor)
        Concepto.Text = ZUltimo + 1
        rstConceptos.Close
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
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Where Conceptos.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Conceptos.Concepto"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                With rstConceptos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Concepto) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Concepto
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstConceptos.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cuenta.Cuenta"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                With rstCuenta
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Cuenta + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Cuenta
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCuenta.Close
            End If
            
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Concepto_DblClick()

    Opcion.Clear
    Opcion.AddItem "Conceptos de Compras"
    Opcion.AddItem "Cuentas Contables"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Cuenta_DblClick()

    Opcion.Clear
    Opcion.AddItem "Concepto de Compras"
    Opcion.AddItem "Cuentas Contables"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Concepto_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Cuenta_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call Cancela_click
        Case Else
    End Select
End Sub





Private Sub Anterior_Click()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Where Conceptos.Concepto < " + "'" + Concepto.Text + "'"
    ZSql = ZSql + " Order by Conceptos.Concepto"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        With rstConceptos
            .MoveLast
            Concepto.Text = rstConceptos!Concepto
        End With
        rstConceptos.Close
        Call Imprime_Datos
        Concepto.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de Concepto de Compras")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Concepto) as [ConceptoMenor]"
    ZSql = ZSql + " FROM Conceptos"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        rstConceptos.MoveFirst
        ZUltimo = IIf(IsNull(rstConceptos!ConceptoMenor), "0", rstConceptos!ConceptoMenor)
        Concepto.Text = ZUltimo
        rstConceptos.Close
        Call Imprime_Datos
        Concepto.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Concepto) as [ConceptoMayor]"
    ZSql = ZSql + " FROM Conceptos"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        rstConceptos.MoveLast
        ZUltimo = IIf(IsNull(rstConceptos!ConceptoMayor), "0", rstConceptos!ConceptoMayor)
        Concepto.Text = ZUltimo
        rstConceptos.Close
        Call Imprime_Datos
        Concepto.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Where Conceptos.Concepto > " + "'" + Concepto.Text + "'"
    ZSql = ZSql + " Order by Conceptos.Concepto"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        With rstConceptos
            .MoveFirst
            Concepto.Text = rstConceptos!Concepto
        End With
        rstConceptos.Close
        Call Imprime_Datos
        Concepto.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Concepto de Compras")
    End If
End Sub








