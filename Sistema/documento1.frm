VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgDocumento1 
   AutoRedraw      =   -1  'True
   Caption         =   "Emision de Documento 1"
   ClientHeight    =   3840
   ClientLeft      =   2175
   ClientTop       =   735
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3840
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
      Height          =   3135
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.CommandButton Panta 
         Caption         =   "Panta F1"
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
         Left            =   600
         MouseIcon       =   "documento1.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "documento1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1800
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
         Left            =   2040
         MouseIcon       =   "documento1.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "documento1.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1800
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
         Left            =   3360
         MouseIcon       =   "documento1.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "documento1.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Salida"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Despacho 
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
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Despacho Numero"
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
         Left            =   480
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Documento1.rpt"
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
End
Attribute VB_Name = "PrgDocumento1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
   
    With rstEmpresa
        .Index = "Empresa"
        Rem .Seek "=", Val(WEmpresa)
        .Seek "=", 1
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Auxi1 = ""
            !varios = Left$(WTitulo, 50)
            .Update
        End If
    End With
    
    For WRenglon = 1 To 100
    
        With rstDespacho
    
            Auxi = Despacho.Text
            Call Ceros(Auxi, 6)
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
        
            .Index = "Clave"
            .Seek "=", Auxi + Auxi1
            If .NoMatch = False Then
                
                WOrden = !Orden
                WArticulo = !Articulo
                WCantidad = !Cantidad
                WCartons = 0
                WInner = 0
                Wcase = 0
                WLugar = 0
                
                With rstOrdenImpo
                    .Index = "Orden"
                    .Seek "=", WOrden
                    Do
                        If .EOF = False Then
                            If WOrden <> !Orden Then
                                Exit Do
                            End If
                            If !Articulo = WArticulo Then
                                WCartons = !Cartons
                                WInner = !Inner
                                Wcase = !Case
                                WLugar = !Lugar
                                Exit Do
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                
                WPrecio = 0
                WUnidades = 0
                With rstArtiExpo
                    .Index = "Codigo"
                    .Seek "=", WArticulo
                    If .NoMatch = False Then
                        WPrecio = !Precio
                        WUnidades = !Unidades
                    End If
                End With
                
                .Edit
                !cajas = WCartons
                !packs = WCantidad
                !pkmaster = Wcase
                !pkinner = WInner
                If WInner = 1 Or WInner = 0 Then
                    !innermaster = 0
                        Else
                    !innermaster = Wcase / WInner
                End If
                !packunitario = WUnidades
                !totalunitario = WUnidades * WCantidad
                !Lugar = WLugar
                !Precio = WPrecio
                .Update
                    
            End If
        
        End With
    
    Next WRenglon
    
    
    Listado.WindowTitle = "Emision de Documento 1"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Despacho.Despacho} in " + Despacho.Text + " to " + Despacho.Text
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstDespacho
        .Close
    End With
    With rstOrdenImpo
        .Close
    End With
    With rstArtiExpo
        .Close
    End With
    DbsAdminis.Close
    Despacho.SetFocus
    PrgDocumento1.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Despacho_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Despacho.Text = ""
    End If
End Sub

Sub Form_Load()
    Despacho.Text = ""
    Frame2.Visible = True
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Despacho
    OPEN_FILE_OrdenImpo
    OPEN_FILE_ArtiExpo
End Sub

Private Sub Despacho_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
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



