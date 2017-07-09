VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaOrdenImportacion 
   AutoRedraw      =   -1  'True
   Caption         =   "Ordenes de Compra Pendientes"
   ClientHeight    =   3105
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3105
   ScaleWidth      =   8145
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
      Height          =   2775
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   5175
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
         Left            =   3960
         MouseIcon       =   "ListaOrdenImportacion.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaOrdenImportacion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salida"
         Top             =   1320
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
         Left            =   2160
         MouseIcon       =   "ListaOrdenImportacion.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaOrdenImportacion.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1320
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
         Left            =   360
         MouseIcon       =   "ListaOrdenImportacion.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaOrdenImportacion.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1320
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
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   5
         Text            =   " "
         Top             =   720
         Width           =   1215
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
         Left            =   2280
         MaxLength       =   8
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Proveedor"
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
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Proveedor"
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
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6720
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Estacli.rpt"
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
      Left            =   6600
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaOrdenImportacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxi As String
Private Auxi As String

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    Rem On Error GoTo WError
    
    Listado.WindowTitle = "Ordenes de Compra de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    If Val(Desde.Text) <> 0 And Val(Hasta.Text) = 0 Then
        Hasta.Text = Desde.Text
    End If

    Listado.SQLQuery = "SELECT OrdenImportacion.Numero, OrdenImportacion.Proveedor, OrdenImportacion.Fecha, OrdenImportacion.FechaLLegada, OrdenImportacion.ProForma, OrdenImportacion.Articulo, OrdenImportacion.Cantidad, OrdenImportacion.Estado, " _
            + "Proveedor.Nombre, " _
            + "Articulo.Descripcion " _
            + "From " _
            + DSQ + ".dbo.OrdenImportacion OrdenImportacion, " _
            + DSQ + ".dbo.Proveedor Proveedor, " _
            + DSQ + ".dbo.Articulo Articulo " _
            + "Where " _
            + "OrdenImportacion.Proveedor = Proveedor.Proveedor AND " _
            + "OrdenImportacion.Articulo = Articulo.Codigo AND " _
            + "OrdenImportacion.Proveedor >= " + Desde.Text + " AND " _
            + "OrdenImportacion.Proveedor <= " + Hasta.Text + " AND " _
            + "OrdenImportacion.Estado = 0"
            
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
            
    Uno = "{OrdenImportacion.Proveedor} in " + Desde.Text + " to " + Hasta.Text
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.ReportFileName = "ListaOrdenImportacion.rpt"
   
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaOrdenImportacion.Hide
    Unload Me
    Menu231.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Frame2.Visible = True
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Panta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub












