VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form PrgCondPago 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Condiciones de Pago"
   ClientHeight    =   5835
   ClientLeft      =   1125
   ClientTop       =   750
   ClientWidth     =   10545
   LinkTopic       =   "Form2"
   ScaleHeight     =   5835
   ScaleWidth      =   10545
   Begin VB.TextBox Observaciones 
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
      Left            =   4725
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   840
      Width           =   5535
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
      Height          =   735
      Left            =   285
      MouseIcon       =   "condpago.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":030A
      TabIndex        =   23
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borrar  (F2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1245
      MouseIcon       =   "condpago.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":0E56
      TabIndex        =   22
      ToolTipText     =   "Elimina el Registro"
      Top             =   1680
      Width           =   855
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
      Height          =   735
      Left            =   2205
      MouseIcon       =   "condpago.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":19A2
      TabIndex        =   21
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1680
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
      Height          =   735
      Left            =   3165
      MouseIcon       =   "condpago.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":24EE
      TabIndex        =   20
      ToolTipText     =   "Consulta de Datos"
      Top             =   1680
      Width           =   1095
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
      Height          =   735
      Left            =   8445
      MouseIcon       =   "condpago.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":303A
      TabIndex        =   19
      ToolTipText     =   "Impresion "
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Salir (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9405
      MouseIcon       =   "condpago.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":3B86
      TabIndex        =   18
      ToolTipText     =   "Salida"
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primero (F5)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4365
      MouseIcon       =   "condpago.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":46D2
      TabIndex        =   17
      ToolTipText     =   "Primer Registro"
      Top             =   1680
      Width           =   855
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
      Height          =   735
      Left            =   5325
      MouseIcon       =   "condpago.frx":4B14
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":4E1E
      TabIndex        =   16
      ToolTipText     =   "Registro Anterior"
      Top             =   1680
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
      Height          =   735
      Left            =   6285
      MouseIcon       =   "condpago.frx":5260
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":556A
      TabIndex        =   15
      ToolTipText     =   "Registro Siguiente"
      Top             =   1680
      Width           =   1095
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
      Height          =   735
      Left            =   7485
      MouseIcon       =   "condpago.frx":59AC
      MousePointer    =   99  'Custom
      Picture         =   "condpago.frx":5CB6
      TabIndex        =   14
      ToolTipText     =   "Salida"
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox Dias 
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
      Left            =   2325
      MaxLength       =   10
      TabIndex        =   2
      Text            =   " "
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Codigo 
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
      Left            =   2325
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9840
      Top             =   4560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      Destination     =   1
      WindowTitle     =   "Listado de Bancos"
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
      Left            =   9360
      TabIndex        =   5
      Top             =   4200
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
      Left            =   4725
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   2505
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   5535
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
         Left            =   3360
         MouseIcon       =   "condpago.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "condpago.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   4440
         MouseIcon       =   "condpago.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "condpago.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   2040
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
         Left            =   2040
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
         Left            =   1800
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
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
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame ConsultaFrame 
      Height          =   3015
      Left            =   960
      TabIndex        =   28
      Top             =   2640
      Width           =   8535
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
         ItemData        =   "condpago.frx":6F90
         Left            =   240
         List            =   "condpago.frx":6F97
         TabIndex        =   30
         Top             =   600
         Visible         =   0   'False
         Width           =   8175
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
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   8175
      End
      Begin VB.ListBox Opcion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   2280
         TabIndex        =   31
         Top             =   1080
         Visible         =   0   'False
         Width           =   3975
      End
   End
   Begin VB.Label lblLabels 
      Caption         =   "Observaciones"
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
      Index           =   2
      Left            =   3285
      TabIndex        =   27
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Plazo (Meses)"
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
      Left            =   885
      TabIndex        =   13
      Top             =   840
      Width           =   1215
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
      Left            =   3405
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Condicion"
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
      Left            =   285
      TabIndex        =   3
      Top             =   420
      Width           =   2295
   End
End
Attribute VB_Name = "PrgCondPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const WHeightI = 3150
Private Const WHeightII = 6340

Sub Imprime_Nombre()
End Sub

Sub Verifica_datos()
    If Val(Dias.Text) = 0 Then
         Dias.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Dias.Text = Pusing("###,###.##", Dias.Text)
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CondPago"
    ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Codigo.Text + "'"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstCondPago.RecordCount > 0 Then
        Codigo.Text = Trim(rstCondPago!Codigo)
        Nombre.Text = Trim(rstCondPago!Nombre)
        Dias.Text = Str$(Trim(rstCondPago!Dias))
        Observaciones.Text = Trim(rstCondPago!Observaciones)
        rstCondPago.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
End Sub

Private Sub Acepta_Click()
    
    Call ContraerFormulario

    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    
    If Val(Hasta.Text) = 0 And Val(Hasta.Text) = 0 Then
         Hasta.Text = "9999"
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CondPago SET "
    ZSql = ZSql + " CodigoEmpresa = '" & YEmpresa & "'"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Condiciones de Pago"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.ReportFileName = App.Path & "/CondPago.rpt"
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CondPago.Codigo, CondPago.Nombre, CondPago.Dias, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.CondPago CondPago, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "CondPago.CodigoEmpresa = '" & YEmpresa & "' AND " _
                + "CondPago.Codigo >= " + Desde.Text + " AND " _
                + "CondPago.Codigo <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{CondPago.Codigo} in '" + Desde.Text + "' to '" + Hasta.Text + "'"
    Listado.SelectionFormula = "{CondPago.Codigo} in '" + Desde.Text + "' to '" + Hasta.Text + "'"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_Click()
    Call ContraerFormulario

    Codigo.SetFocus
End Sub

Private Sub cmdAdd_Click()
    Dim WCodigo, WNombre, WObservaciones, WDias As String
    
    WCodigo = Trim(Codigo.Text)
    WObservaciones = Trim(Observaciones.Text)
    WNombre = Trim(Nombre.Text)
    WDias = Trim(Dias.Text)
    
    WDias = IIf(WDias = "", "0", WDias)
    
    ' Validamos de que se hayan cargado por lo menos el codigo, el nombre y que el plazo no sea un numero negativo.
    If Trim(Codigo.Text) <> "" And Trim(Nombre.Text) <> "" And Val(WDias) >= 0 Then
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CondPago"
        ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Codigo.Text + "'"
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstCondPago.RecordCount > 0 Then
            rstCondPago.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE CondPago SET "
            ZSql = ZSql + " Nombre = " + "'" + WNombre + "',"
            ZSql = ZSql + " Observaciones = " + "'" + WObservaciones + "',"
            ZSql = ZSql + " Dias = " + "'" + WDias + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
        Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CondPago ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Dias )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WNombre + "',"
            ZSql = ZSql + "'" + WObservaciones + "',"
            ZSql = ZSql + "'" + WDias + "')"
        End If
        
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    
        Call CmdLimpiar_Click
    
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Condicion de Pago")
        
        Codigo.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CondPago"
        ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Codigo.Text + "'"
        spCondPago = ZSql
        Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstCondPago.RecordCount > 0 Then
            rstCondPago.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE CondPago"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spCondPago = ZSql
                Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    ' Reestablecemos el tamaño por defecto de la ventana, ocultando de esta manera la ayuda y ventana de impresion.
    Call ContraerFormulario
    
    Codigo.Text = ""
    Nombre.Text = ""
    Observaciones.Text = ""
    Dias.Text = ""

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM CondPago"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstCondPago.RecordCount > 0 Then
        rstCondPago.MoveLast
        ZUltimo = IIf(IsNull(rstCondPago!CodigoMayor), "0", rstCondPago!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstCondPago.Close
    End If
    
    If Codigo.Visible Then
        Codigo.SetFocus
    End If
    
End Sub

Private Sub CmdClose_Click()
    PrgCondPago.Hide
    Unload Me
    MenuVen.Show
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

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dias.SetFocus
    End If
    If KeyAscii = 27 Then
        Nombre.Text = ""
    End If
End Sub

Private Sub Dias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Format_datos
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Dias.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Codigo = " + "'" + Codigo.Text + "'"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                rstCondPago.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
        End If
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
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
    
    ConsultaFrame.Visible = True
    Pantalla.Visible = False
    Ayuda.Visible = False
    Opcion.Clear
    
    Opcion.AddItem "Condicion de Pago"
    
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
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Order by CondPago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem Trim(IngresaItem)
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
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
            indice = Pantalla.ListIndex
            Codigo.Text = WIndice.List(indice)
            Call Codigo_KeyPress(13)
            
            Call ContraerFormulario
        Case Else
    End Select
    
End Sub

Private Sub ContraerFormulario()
    Frame2.Visible = False
    ConsultaFrame.Visible = False
    Me.Height = WHeightI
    If Codigo.Visible Then
        Codigo.SetFocus
    End If
    
End Sub

Private Sub ExpandirFormulario()
    Me.Height = WHeightII
End Sub

Sub Form_Load()
    
    Call CmdLimpiar_Click
    
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
            ZSql = ZSql + " FROM CondPago"
            ZSql = ZSql + " Where CondPago.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by CondPago.Codigo"
            spCondPago = ZSql
            Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstCondPago.RecordCount > 0 Then
                With rstCondPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem Trim(IngresaItem)
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCondPago.Close
            End If
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        If Trim(Ayuda.Text) = "" Then
            Call ContraerFormulario
            Codigo.SetFocus
            Exit Sub
        End If
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Codigo_DblClick()
    
    Opcion.Clear
    Opcion.AddItem "Condicion de Pago"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    ConsultaFrame.Visible = True
    Call ExpandirFormulario
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Dias_KeyDown(KeyCode As Integer, Shift As Integer)
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
    ZSql = ZSql + " FROM CondPago"
    ZSql = ZSql + " Where CondPago.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by CondPago.Codigo"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstCondPago.RecordCount > 0 Then
        With rstCondPago
            .MoveLast
            Codigo.Text = rstCondPago!Codigo
        End With
        rstCondPago.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Condiciones de Pago")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM CondPago"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstCondPago.RecordCount > 0 Then
        rstCondPago.MoveFirst
        ZUltimo = IIf(IsNull(rstCondPago!CodigoMenor), "0", rstCondPago!CodigoMenor)
        Codigo.Text = ZUltimo
        rstCondPago.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
        ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM CondPago"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstCondPago.RecordCount > 0 Then
        rstCondPago.MoveLast
        ZUltimo = IIf(IsNull(rstCondPago!CodigoMayor), "0", rstCondPago!CodigoMayor)
        Codigo.Text = ZUltimo
        rstCondPago.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CondPago"
    ZSql = ZSql + " Where CondPago.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by CondPago.Codigo"
    spCondPago = ZSql
    Set rstCondPago = db.OpenRecordset(spCondPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstCondPago.RecordCount > 0 Then
        With rstCondPago
            .MoveFirst
            Codigo.Text = rstCondPago!Codigo
        End With
        rstCondPago.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Condiciones de Pago")
    End If
End Sub

















