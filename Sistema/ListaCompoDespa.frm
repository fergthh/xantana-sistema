VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaCompoDespa 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Componentes por Despacho"
   ClientHeight    =   3810
   ClientLeft      =   2175
   ClientTop       =   735
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3810
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
      Top             =   240
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
         MouseIcon       =   "ListaCompoDespa.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaCompoDespa.frx":030A
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
         MouseIcon       =   "ListaCompoDespa.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaCompoDespa.frx":0E56
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
         MouseIcon       =   "ListaCompoDespa.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaCompoDespa.frx":19A2
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
      ReportFileName  =   "ListaCompoDespa.rpt"
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
Attribute VB_Name = "PrgListaCompoDespa"
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

    Rem On Error GoTo WError
    
    da = ""
    With rstListaCompo
        .Index = "Componente"
        .Seek ">=", da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
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
        
                WArticulo = !Articulo
                WCantidad = !Cantidad
                
                For WRenglon1 = 1 To 100
    
                    With rstFormula
    
                        Auxi1 = WRenglon1
                        Call Ceros(Auxi1, 2)
        
                        .Index = "Clave"
                        .Seek "=", WArticulo + Auxi1
                        If .NoMatch = False Then
        
                            WComponente = !Articulo
                            WCantidadFormula = !Cantidad
                            
                            With rstListaCompo
                                .AddNew
                                !Componente = WComponente
                                !Despacho = Val(Despacho.Text)
                                !ClaveDespacho = Auxi + "01"
                                !Cantidad = WCantidadFormula * WCantidad
                                .Update
                            End With
                
                        End If
        
                    End With
    
                Next WRenglon1
                
            End If
        
        End With
    
    Next WRenglon
    
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
    
    Listado.WindowTitle = "Listado de Componentes por Despacho"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Despacho.Despacho} in " + Despacho.Text + " to " + Despacho.Text
    
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
    With rstListaCompo
        .Close
    End With
    With rstDespacho
        .Close
    End With
    With rstArtiExpo
        .Close
    End With
    With rstFormula
        .Close
    End With
    With rstComponente
        .Close
    End With
    DbsAdminis.Close
    Despacho.SetFocus
    PrgDocumento2.Hide
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
    OPEN_FILE_ListaCompo
    OPEN_FILE_Despacho
    OPEN_FILE_ArtiExpo
    OPEN_FILE_Formula
    OPEN_FILE_Componente
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



