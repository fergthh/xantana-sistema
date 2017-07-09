VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgLineaInsumo 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Grupos"
   ClientHeight    =   5355
   ClientLeft      =   1050
   ClientTop       =   690
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
      MouseIcon       =   "lineainsumo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":030A
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
      MouseIcon       =   "lineainsumo.frx":074C
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":0A56
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
      MouseIcon       =   "lineainsumo.frx":0E98
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":11A2
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
      MouseIcon       =   "lineainsumo.frx":15E4
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":18EE
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
      MouseIcon       =   "lineainsumo.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":203A
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
      MouseIcon       =   "lineainsumo.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":2B86
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
      MouseIcon       =   "lineainsumo.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":36D2
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
      MouseIcon       =   "lineainsumo.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":421E
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
      MouseIcon       =   "lineainsumo.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":4D6A
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
      MouseIcon       =   "lineainsumo.frx":55AC
      MousePointer    =   99  'Custom
      Picture         =   "lineainsumo.frx":58B6
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
   Begin VB.TextBox Linea 
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
         MouseIcon       =   "lineainsumo.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "lineainsumo.frx":6402
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
         MouseIcon       =   "lineainsumo.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "lineainsumo.frx":6B4E
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
      ReportFileName  =   "LineaInsumos.rpt"
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
      ItemData        =   "lineainsumo.frx":6F90
      Left            =   240
      List            =   "lineainsumo.frx":6F97
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
      Caption         =   "Codigo de Grupo"
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
Attribute VB_Name = "PrgLineaInsumo"
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
    ZSql = ZSql + " FROM LineaInsumo"
    ZSql = ZSql + " Where LineaInsumo.Linea = " + "'" + Linea.Text + "'"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaInsumo.RecordCount > 0 Then
        Nombre.Text = Trim(rstLineaInsumo!Nombre)
        rstLineaInsumo.Close
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
    ZSql = ZSql + "UPDATE LineaInsumo SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + "1" + "'"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Grupos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT LineaInsumo.Linea, LineaInsumo.Nombre, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.LineaInsumo LineaInsumo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "LineaInsumo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "LineaInsumo.Linea >= " + Desde.Text + " AND " _
                + "LineaInsumo.Linea <= " + Hasta.Text
    
    Listado.GroupSelectionFormula = "{LineaInsumo.Linea} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{LineaInsumo.Linea} in " + Desde.Text + " to " + Hasta.Text
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Linea.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
    Linea.SetFocus
End Sub

Private Sub cmdAdd_Click()
    If Linea.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM LineaInsumo"
        ZSql = ZSql + " Where LineaInsumo.Linea = " + "'" + Linea.Text + "'"
        spLineaInsumo = ZSql
        Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLineaInsumo.RecordCount > 0 Then
            rstLineaInsumo.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE LineaInsumo SET "
            ZSql = ZSql + " Nombre = " + "'" + Nombre.Text + "'"
            ZSql = ZSql + " Where Linea = " + "'" + Linea.Text + "'"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO LineaInsumo ("
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Nombre )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Linea.Text + "',"
            ZSql = ZSql + "'" + Nombre.Text + "')"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call CmdLimpiar_Click
        Linea.SetFocus
        
    End If
    
End Sub

Private Sub CmdDelete_Click()
    If Linea.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM LineaInsumo"
        ZSql = ZSql + " Where LineaInsumo.Linea = " + "'" + Linea.Text + "'"
        spLineaInsumo = ZSql
        Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLineaInsumo.RecordCount > 0 Then
            rstLineaInsumo.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE LineaInsumo"
                ZSql = ZSql + " Where Linea = " + "'" + Linea.Text + "'"
                spLineaInsumo = ZSql
                Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
                    
                Call CmdLimpiar_Click
                
            End If
        End If
    
    End If
    Linea.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    On Error GoTo WError
    
    Linea.Text = "1"
    Nombre.Text = ""
    Linea.SetFocus
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Linea) as [LineaMayor]"
    ZSql = ZSql + " FROM LineaInsumo"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaInsumo.RecordCount > 0 Then
        rstLineaInsumo.MoveLast
        ZUltimo = IIf(IsNull(rstLineaInsumo!LineaMayor), "0", rstLineaInsumo!LineaMayor)
        Linea.Text = ZUltimo + 1
        rstLineaInsumo.Close
    End If
    
    Exit Sub
    
WError:

    Resume Next
        
    
End Sub

Private Sub cmdClose_Click()
    PrgLineaInsumo.Hide
    Unload Me
    Menu2.Show
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

Private Sub Linea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Linea.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM LineaInsumo"
            ZSql = ZSql + " Where LineaInsumo.Linea = " + "'" + Linea.Text + "'"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineaInsumo.RecordCount > 0 Then
                rstLineaInsumo.Close
                Call Imprime_Datos
                    Else
                WLinea = Linea.Text
                CmdLimpiar_Click
                Linea.Text = WLinea
            End If
        End If
        Nombre.SetFocus
    End If
    If KeyAscii = 27 Then
        Linea.Text = ""
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

     Opcion.AddItem "Lineas Grupos"

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
            ZSql = ZSql + " FROM LineaInsumo"
            ZSql = ZSql + " Order by LineaInsumo.Linea"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineaInsumo.RecordCount > 0 Then
                With rstLineaInsumo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Linea) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Linea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLineaInsumo.Close
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
            Linea.Text = WIndice.List(Indice)
            Call Linea_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Sub Form_Load()

    On Error GoTo WError
    
    Linea.Text = "1"
    Nombre.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Linea) as [LineaMayor]"
    ZSql = ZSql + " FROM LineaInsumo"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaInsumo.RecordCount > 0 Then
        rstLineaInsumo.MoveLast
        ZUltimo = IIf(IsNull(rstLineaInsumo!LineaMayor), "0", rstLineaInsumo!LineaMayor)
        Linea.Text = ZUltimo + 1
        rstLineaInsumo.Close
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
            ZSql = ZSql + " FROM LineaInsumo"
            ZSql = ZSql + " Where LineaInsumo.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by LineaInsumo.Linea"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineaInsumo.RecordCount > 0 Then
                With rstLineaInsumo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Linea) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Linea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLineaInsumo.Close
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

Private Sub Linea_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Lineas Grupos"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub



Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Linea_KeyDown(KeyCode As Integer, Shift As Integer)
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
    ZSql = ZSql + " FROM LineaInsumo"
    ZSql = ZSql + " Where LineaInsumo.Linea < " + "'" + Linea.Text + "'"
    ZSql = ZSql + " Order by LineaInsumo.Linea"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaInsumo.RecordCount > 0 Then
        With rstLineaInsumo
            .MoveLast
            Linea.Text = rstLineaInsumo!Linea
        End With
        rstLineaInsumo.Close
        Call Imprime_Datos
        Linea.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de Grupos")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Linea) as [LineaMenor]"
    ZSql = ZSql + " FROM LineaInsumo"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaInsumo.RecordCount > 0 Then
        rstLineaInsumo.MoveFirst
        ZUltimo = IIf(IsNull(rstLineaInsumo!LineaMenor), "0", rstLineaInsumo!LineaMenor)
        Linea.Text = ZUltimo
        rstLineaInsumo.Close
        Call Imprime_Datos
        Linea.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Linea) as [LineaMayor]"
    ZSql = ZSql + " FROM LineaInsumo"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaInsumo.RecordCount > 0 Then
        rstLineaInsumo.MoveLast
        ZUltimo = IIf(IsNull(rstLineaInsumo!LineaMayor), "0", rstLineaInsumo!LineaMayor)
        Linea.Text = ZUltimo
        rstLineaInsumo.Close
        Call Imprime_Datos
        Linea.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM LineaInsumo"
    ZSql = ZSql + " Where LineaInsumo.Linea > " + "'" + Linea.Text + "'"
    ZSql = ZSql + " Order by LineaInsumo.Linea"
    spLineaInsumo = ZSql
    Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineaInsumo.RecordCount > 0 Then
        With rstLineaInsumo
            .MoveFirst
            Linea.Text = rstLineaInsumo!Linea
        End With
        rstLineaInsumo.Close
        Call Imprime_Datos
        Linea.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Grupos")
    End If
End Sub


































