VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form PrgLista 
   AutoRedraw      =   -1  'True
   Caption         =   "Lista de Precio"
   ClientHeight    =   5715
   ClientLeft      =   1050
   ClientTop       =   690
   ClientWidth     =   9960
   LinkTopic       =   "Form2"
   ScaleHeight     =   5715
   ScaleWidth      =   9960
   Visible         =   0   'False
   Begin VB.Frame FrameListaPrecios 
      Caption         =   "Listas de Precios Disponibles"
      Height          =   5175
      Left            =   4920
      TabIndex        =   39
      Top             =   240
      Width           =   4455
      Begin VB.ListBox ListasPrecios 
         Height          =   3960
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   3975
      End
      Begin VB.CommandButton btnCerrarListasPrecios 
         Caption         =   "Cerrar Listas"
         Height          =   495
         Left            =   1680
         TabIndex        =   40
         Top             =   4560
         Width           =   1215
      End
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
      Left            =   2280
      TabIndex        =   6
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
      _Version        =   348160
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
      PrintFileLinesPerPage=   60
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
   Begin VB.Frame Articulos 
      Caption         =   "Precios de Venta"
      Height          =   3135
      Left            =   600
      TabIndex        =   28
      Top             =   2280
      Width           =   8655
      Begin VB.ComboBox WCombo1 
         Height          =   315
         Left            =   7560
         TabIndex        =   37
         Text            =   "Combo2"
         Top             =   480
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.TextBox WTexto1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   6600
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox WTexto2 
         Height          =   375
         Left            =   7080
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   2280
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   34
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   33
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   2
         Left            =   4080
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   32
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   3
         Left            =   4680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox WTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   4
         Left            =   3480
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   30
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSMask.MaskEdBox WTexto3 
         Height          =   375
         Left            =   6120
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid Wvector1 
         Height          =   2535
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4471
         _Version        =   393216
         OLEDropMode     =   1
      End
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

' CONTROLES PARA GRILLA
Private WParametros(4, 5)
Private WFormato(100) As String
Private WControl As String
Private WMsgErrores As String


Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Const VK_TAB = &H9

Private Function GetTabState() As Boolean
    GetTabState = False
    If GetKeyState(VK_TAB) And -256 Then
        GetTabState = True
    End If
End Function

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
        Call Traer_Articulos
    End If
End Sub

Private Sub Traer_Articulos()
    ' Cargamos los datos de las Listas de Precios.
    
    Call Limpia_Vector
    
    Dim WLista, WArticulo, WNeto, WPrecio, WClave, WRenglon, XRenglon, WDescripcion
    
    WLista = Trim(Codigo.Text)
    WPrecio = 0
    WNeto = 0
    WDescripcion = ""
    
    WArticulo = ""
    WRenglon = 1
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ListaArticulos, Lista"
    ZSql = ZSql + " Where ListaArticulos.Lista = " + "'" + WLista + "' and ListaArticulos.Lista = Lista.Codigo"
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    If rstLista.RecordCount > 0 Then
    
        With rstLista
            .MoveFirst
            Do While .EOF = False And .BOF = False
                                            
                WArticulo = Trim(IIf(IsNull(!Articulo), "", Trim(!Articulo)))
                WNeto = IIf(IsNull(!Neto), 0, !Neto)
                WPrecio = IIf(IsNull(!Precio), 0, !Precio)
                WDescripcion = IIf(IsNull(!Descripcion), "", Trim(!Descripcion))
                
                    Wvector1.Row = WRenglon
                    Wvector1.Col = 1
                    If Trim(Wvector1.Text) = "" Then
                    
                        Wvector1.Col = 1
                        Wvector1.Text = Trim(WArticulo)
                        
                        Wvector1.Col = 2
                        Wvector1.Text = Trim(WDescripcion)
                        
                        Wvector1.Col = 3
                        Wvector1.Text = Pusing("######.##", Trim(WNeto))
                        
                        Wvector1.Col = 4
                        Wvector1.Text = Pusing("######.##", Trim(WPrecio))
                        
                        WRenglon = WRenglon + 1
                        
                    Else
                    
                        Exit Do
                    
                    End If
            .MoveNext
            Loop
            
        End With
    
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
    Articulos.Visible = True
    'Me.Height = WMinHeight
    
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

Private Sub btnCerrarListasPrecios_Click()
    FrameListaPrecios.Visible = False
    Call Posicionar_En_Grilla
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
    Codigo.SetFocus
    'Me.Height = WMinHeight
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
        
        Call Actualizar_Lista_Precios
       
        'If ZZNivel = 0 Then
        '    txtUserName = "SA"
        '    txtPassword = "Sw58125812"
        '    txtOdbc = "FraganciasII"
        '    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        '    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        '        Else
        '    txtUserName = "SA"
        '    txtPassword = "Sw58125812"
        '    txtOdbc = "Fragancias"
        '    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        '    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        'End If
   
    
        'ZSql = ""
        'ZSql = ZSql + "Select *"
        'ZSql = ZSql + " FROM Lista"
        'ZSql = ZSql + " Where Lista.Codigo = " + "'" + WCodigo + "'"
        'spLista = ZSql
        'Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        'If rstLista.RecordCount > 0 Then
        '    rstLista.Close
         '   ZSql = ""
        '    ZSql = ZSql + "UPDATE Lista SET "
        '    ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
        '    ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
        '    spLista = ZSql
        '    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        'Else
        '    ZSql = ""
        '    ZSql = ZSql + "INSERT INTO Lista ("
        '    ZSql = ZSql + "Codigo ,"
        '    ZSql = ZSql + "Descripcion )"
        '    ZSql = ZSql + "Values ("
        '    ZSql = ZSql + "'" + WCodigo + "',"
        '    ZSql = ZSql + "'" + WDescripcion + "')"
        '    spLista = ZSql
        '    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        'End If
        
         
        'If ZZNivel = 0 Then
        '    txtUserName = "SA"
        '    txtPassword = "Sw58125812"
        '    txtOdbc = "Fragancias"
        '    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        '    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        '        Else
        '    txtUserName = "SA"
        '     txtPassword = "Sw58125812"
        '    txtOdbc = "FraganciasII"
        '    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        '    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        'End If
        
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Lista de Precios")
        
        Call CmdLimpiar_Click
    
        Codigo.SetFocus
        
    End If
    
End Sub

Private Sub Actualizar_Lista_Precios()
    Dim WLista, WArticulo, WNeto, WPrecio, WClave, WRenglon, XRenglon, XLista
    
    WLista = Trim(Codigo.Text)
    WRenglon = 1
    
    ' Borramos la informacion anterior en caso de que hayan, asi no tendremos problemas en los casos en que se elimine alguna lista.
    
    ZSql = ""
    ZSql = ZSql + "DELETE FROM ListaArticulos WHERE Lista = '" + WLista + "'"
    
    spLista = ZSql
    Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
    
    ' Recorremos la grilla y verificamos que haya datos en las columnas de neto y final.
    For i = 1 To Wvector1.Rows
    
        WArticulo = ""
        WNeto = 0
        WPrecio = 0
        WClave = ""
        
        With Wvector1
            .Row = i
            .Col = 1
            If Trim(.Text) <> "" Then
            
                .Col = 1
                WArticulo = Trim(.Text)
                
                .Col = 3
                WNeto = Val(.Text)
                
                .Col = 4
                WPrecio = Val(.Text)
                
                ' Una vez guardada la informacion, la guardamos.
                Auxi = WLista
                Call Ceros(Auxi, 4) ' Solo en caso en que sean numericos las claves de las listas.
                XLista = Auxi
                
                XRenglon = Str$(WRenglon)
                
                Auxi = XRenglon
                Call Ceros(Auxi, 2)
                XRenglon = Auxi
                
                WClave = WArticulo + XLista + XRenglon
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO ListaArticulos "
                ZSql = ZSql + "(Clave, Articulo, Lista, Renglon, Neto, Precio) "
                ZSql = ZSql + "VALUES "
                ZSql = ZSql + "('" + WClave + "','" + WArticulo + "','" + WLista + "','" + XRenglon + "'," + Str$(WNeto) + "," + Str$(WPrecio) + ")"
                
                spLista = ZSql
                Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
                
                WRenglon = WRenglon + 1
                
            Else
            
                Exit For
            
            End If
        
        End With
    
    Next
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
            If Respuestaa% = 6 Then
            
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
    FrameListaPrecios.Visible = False
    
    Articulos.Visible = True
    
    'Me.Height = WMinHeight
    
    Limpia_Vector
    
    Codigo.SetFocus
    
    Exit Sub
    
WError:

    Resume Next
        
    
End Sub

Private Sub Limpia_Vector()

    Wvector1.Clear

    Rem ponga la wvector1 en negritas
    Wvector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    ' Establesco loa Valores de la wvector1
    
    Wvector1.FixedCols = 1
    Wvector1.Cols = 5
    Wvector1.FixedRows = 1
    Wvector1.Rows = 100
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    Wvector1.ColWidth(0) = 200
    Wvector1.Row = 0
    For Ciclo = 1 To Wvector1.Cols - 1
        Wvector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                Wvector1.Text = "Codigo"
                Wvector1.ColWidth(Ciclo) = 1200
                Wvector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                Wvector1.Text = "Descripcion"
                Wvector1.ColWidth(Ciclo) = 5000
                Wvector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 25
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                Wvector1.Text = "Costo"
                Wvector1.ColWidth(Ciclo) = 700
                Wvector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "#####.##"
            Case 4
                Wvector1.Text = "Precio"
                Wvector1.ColWidth(Ciclo) = 700
                Wvector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "######.##"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    Wvector1.Row = 0
    For Ciclo = 1 To Wvector1.Cols - 1
        Wvector1.Col = Ciclo
        WTitulo(Ciclo).Text = Wvector1.Text
        WTitulo(Ciclo).Left = Wvector1.CellLeft + Wvector1.Left
        WTitulo(Ciclo).Top = Wvector1.CellTop + Wvector1.Top
        WTitulo(Ciclo).Width = Wvector1.CellWidth
        WTitulo(Ciclo).Height = Wvector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To Wvector1.Cols - 1
        WAncho = WAncho + Wvector1.ColWidth(Ciclo)
    Next Ciclo
    Wvector1.Width = WAncho

    ' Size the columns.
    Font.Name = Wvector1.Font.Name
    Font.Size = Wvector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el Tamano de las celdas
    Wvector1.AllowUserResizing = flexResizeBoth
    
    Wvector1.Col = 1
    Wvector1.Row = 1
    
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
    'Me.Height = WMaxHeight
    Impresora.Value = False
    
    Ayuda.Visible = False
    Pantalla.Visible = False
    PantallaFiltrada.Visible = False
    
    Articulos.Visible = False
    
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

     Opcion.AddItem "Articulos"
     Opcion.AddItem "Listas"
     
     Pantalla.Visible = False

     Opcion.Visible = True
     
     'Opcion.ListIndex = 0
     'Call Opcion_Click
     
End Sub

Private Sub ListasPrecios_Click()
    Select Case XIndice
        Case 0
            indice = ListasPrecios.ListIndex
            Dim WArticulo
            WArticulo = WIndice.List(indice)
            'Me.Height = WMinHeight
            Articulos.Visible = True
            
            Call Cargar_Articulo(WArticulo)
        Case Else
    End Select
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
    FrameListaPrecios.Visible = False
    Pantalla.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    ListasPrecios.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            Rem ZSql = ZSql + " Where Sector.Codigo = " + "'" + Sector.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            With rstArticulo
                If .RecordCount > 0 Then
                    .MoveFirst
                    
                    Do While .EOF = False And .BOF = False
                        
                        IngresaItem = !Codigo + " " + !Descripcion
                        ListasPrecios.AddItem IngresaItem
                        IngresaItem = !Codigo
                        WIndice.AddItem IngresaItem
                        
                        .MoveNext
                    
                    Loop
                    
                End If
            End With
            
            FrameListaPrecios.Visible = True
        Case 1
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
            Articulos.Visible = False
            Pantalla.Visible = True
            PantallaFiltrada.Visible = False
            Frame2.Visible = False
            'Me.Height = WMaxHeight
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    
    Select Case XIndice
        Case 1
            Pantalla.Visible = False
            Ayuda.Visible = False
        
            indice = Pantalla.ListIndex
            Codigo.Text = WIndice.List(indice)
            'Me.Height = WMinHeight
            Articulos.Visible = True
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Cargar_Articulo(ByVal Codigo As String)

    Dim WRow As Integer
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Codigo = " + "'" + Trim(Codigo) + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    With rstArticulo
        If .RecordCount > 0 Then
            .MoveFirst
            
            For WRow = 1 To 100
            
                Wvector1.Row = WRow
                
                Wvector1.Col = 1
                
                If Wvector1.Text = Trim(!Codigo) Then
                    
                    .Close
                    Exit Sub
                
                End If
                
                If Wvector1.Text = "" Then
                    
                    Exit For
                
                End If
            
            Next
            
            Wvector1.Text = Trim(!Codigo)
            
            Wvector1.Col = 2
            Wvector1.Text = Trim(!Descripcion)
            
            .Close
        End If
    End With
    
    'Call Posicionar_En_Grilla

End Sub

Private Sub Posicionar_En_Grilla()

    Dim WRenglon, WCol
    
    WRenglon = 1
    
    ' Buscamos el primero que este para editar
    For i = WRenglon To Wvector1.Rows
        
        With Wvector1
            .Row = WRenglon
            .Col = 3
            WCol = .Col
            If Trim(.Text) <> "" Then
            
                .Col = 4
                WCol = .Col
                If Trim(.Text) <> "" Then
                
                    WRenglon = WRenglon + 1
                    
                Else
                    Exit For
                End If
                
            Else
                Exit For
            End If
        
        End With
    
    Next
    
    Wvector1.Row = WRenglon
    Wvector1.Col = WCol
    'Wvector1.SetFocus
    Call StartEdit

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
            Articulos.Visible = True
            
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

Private Sub WVector1_DblClick()

    ' Detectamos si se ha querido borrar.
    With Wvector1
        If .Col = 1 Then
            Dim WRow
            WRow = .Row
            
            If .Text = "1" Then
                .Col = 2
                MsgBox "La lista " + Chr$(34) + .Text + Chr$(34) + ", no puede ser eliminada.", vbInformation
                Exit Sub
            End If
            
            T$ = "Borrar Registro"
            m$ = "¿Desea Borrar la lista indicada?" + Chr$(13) + Chr$(13) + "Esta accion no puede deshacerse."
            
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            
            If Respuestaaaaaa% = 6 Then
                Call Borrar_Lista(WRow)
            End If
            
            Exit Sub
        End If
    End With
    
End Sub

Rem
Rem Controles de la wvector1
Rem

Private Sub Borrar_Lista(ByVal iRow)
    Wvector1.RemoveItem (iRow)
End Sub

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = Wvector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = Wvector1.CellLeft + Wvector1.Left
            WTexto1.Top = Wvector1.CellTop + Wvector1.Top
            WTexto1.Width = Wvector1.CellWidth
            WTexto1.Height = Wvector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = Wvector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = Wvector1.CellLeft + Wvector1.Left
            WTexto2.Top = Wvector1.CellTop + Wvector1.Top
            WTexto2.Width = Wvector1.CellWidth
            WTexto2.Height = Wvector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = Wvector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = Wvector1.CellLeft + Wvector1.Left
            WTexto3.Top = Wvector1.CellTop + Wvector1.Top
            WTexto3.Width = Wvector1.CellWidth
            WTexto3.Height = Wvector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(Wvector1.Text) = 10 Then
                        WTexto3.Text = Wvector1.Text
                            Else
                        WTexto3.Mask = ""
                        WTexto3.Text = ""
                        WTexto3.Mask = "##/##/####"
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        Wvector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            If Trim(WTexto1.Text) = "" Then
                Wvector1.Text = ""
            Else
                Wvector1.Text = WTexto1.Text
            End If
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                If Trim(WTexto2.Text) = "" Then
                    Wvector1.Text = ""
                Else
                    Wvector1.Text = WTexto2.Text
                End If
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    If Trim(Replace(WTexto3.Text, "/", "")) = "" Then
                        Wvector1.Text = ""
                    Else
                        Wvector1.Text = WTexto3.Text
                    End If
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(Wvector1.Col) <> "" And Trim(Wvector1.Text) <> "" Then
            Wvector1.Text = Pusing(WFormato(Wvector1.Col), Wvector1.Text)
            'WVector1.Text = WVector1.Text
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = Wvector1.CellLeft + Wvector1.Left
    WCombo1.Top = Wvector1.CellTop + Wvector1.Top
    WCombo1.Width = Wvector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = Wvector1.Text

        Case vbKeyReturn
            ' Finish editing.
            Wvector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.Row < Wvector1.Rows - 1 Then
               'Call Control_Campo
               'If WControl = "S" Then
                    Wvector1.Row = Wvector1.Row + 1
               'End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.Row > Wvector1.FixedRows Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.Row = Wvector1.Row - 1
                'End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.TopRow < Wvector1.Rows - 12 Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.TopRow = Wvector1.TopRow + 12
                    Wvector1.Row = Wvector1.TopRow
                'End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.TopRow - 12 > Wvector1.FixedRows Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.TopRow = Wvector1.TopRow - 12
                    Wvector1.Row = Wvector1.TopRow
                        Else
                    Wvector1.TopRow = 1
                    Wvector1.Row = Wvector1.TopRow
                'End If
            End If
            Call StartEdit

    End Select
End Sub
 
Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = Wvector1.Text

        Case vbKeyReturn
            ' Finish editing.
            Wvector1.SetFocus
            DoEvents
            Call Control_Campo
            If Wvector1.Row < Wvector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    'Wvector1.Row = Wvector1.Row + 1
                    Call Control_wvector1
                End If
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.Row < Wvector1.Rows - 1 Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.Row = Wvector1.Row + 1
                'End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.Row > Wvector1.FixedRows Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.Row = Wvector1.Row - 1
                'End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.TopRow < Wvector1.Rows - 12 Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.TopRow = Wvector1.TopRow + 12
                    Wvector1.Row = Wvector1.TopRow
                'End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.TopRow - 12 > Wvector1.FixedRows Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.TopRow = Wvector1.TopRow - 12
                    Wvector1.Row = Wvector1.TopRow
                        Else
                    Wvector1.TopRow = 1
                    Wvector1.Row = Wvector1.TopRow
                'End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Mask = ""
            WTexto3.Text = ""
            
        Rem F1
        Case 113
            WTexto3.Text = Wvector1.Text

        Case vbKeyReturn
            ' Finish editing.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.Row < Wvector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    'Wvector1.Row = Wvector1.Row + 1
                    Call Control_wvector1
                End If
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.Row < Wvector1.Rows - 1 Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.Row = Wvector1.Row + 1
                'End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.Row > Wvector1.FixedRows Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.Row = Wvector1.Row - 1
                'End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.TopRow < Wvector1.Rows - 12 Then
                 'Call Control_Campo
                 'If WControl = "S" Then
                    Wvector1.TopRow = Wvector1.TopRow + 12
                    Wvector1.Row = Wvector1.TopRow
                 'End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            Wvector1.SetFocus
            DoEvents
            If Wvector1.TopRow - 12 > Wvector1.FixedRows Then
                'Call Control_Campo
                'If WControl = "S" Then
                    Wvector1.TopRow = Wvector1.TopRow - 12
                    Wvector1.Row = Wvector1.TopRow
                        Else
                    Wvector1.TopRow = 1
                    Wvector1.Row = Wvector1.TopRow
                'End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    Wvector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = Wvector1.Col
    Select Case WParametros(4, Wvector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, Wvector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = Wvector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, Wvector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()

    With Wvector1
            
        Select Case .Col
            Case 3
                .Col = .Col + 1
            Case 4
                If .Row < .Rows - 1 Then
                    
                    .Row = Wvector1.Row + 1
                    
                    .Col = 1 ' Controlo que hayan una lista asignada a esa fila.
                    If Trim(.Text) = "" Then
                        ' En caso de que no, me posiciono en la celda de neto en la fila original.
                        .Col = 3
                        .Row = .Row - 1
                        Exit Sub
                    End If
                    
                    .Col = 3
                    
                End If
                Rem .Col = 1
            Case Else
                If .Col < .Cols - 1 Then
                     .Col = .Col + 1
                End If
        End Select
        .SetFocus
        
    End With
    
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = Wvector1.Col
    XFila = Wvector1.Row
    WControl = "S"
    Select Case XColumna
        Case Else
            Wvector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub


































