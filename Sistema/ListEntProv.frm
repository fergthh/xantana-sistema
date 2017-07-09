VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListEntProv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Remitos Entregados por Proveedor"
   ClientHeight    =   7875
   ClientLeft      =   2610
   ClientTop       =   570
   ClientWidth     =   6795
   LinkTopic       =   "Form2"
   ScaleHeight     =   7875
   ScaleWidth      =   6795
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
      Left            =   1560
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   2895
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
      TabIndex        =   7
      Top             =   4800
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   4815
      Begin VB.TextBox HastaArt 
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
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   13
         Text            =   " "
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox DesdeArt 
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
         Left            =   2160
         MaxLength       =   12
         TabIndex        =   12
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox DesdeProv 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox HastaProv 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   4
         Text            =   " "
         Top             =   840
         Width           =   975
      End
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   2160
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label6 
         Caption         =   "Hasta Articulo"
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
         TabIndex        =   15
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Desde Articulo"
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
         TabIndex        =   14
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
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
         TabIndex        =   11
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Fecha"
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
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Image Panta 
         Height          =   480
         Left            =   720
         MouseIcon       =   "ListEntProv.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListEntProv.frx":030A
         ToolTipText     =   "Emision por Pantalla"
         Top             =   3480
         Width           =   480
      End
      Begin VB.Image Consulta 
         Height          =   480
         Left            =   2640
         MouseIcon       =   "ListEntProv.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListEntProv.frx":0E56
         ToolTipText     =   "Consulta de Datos"
         Top             =   3480
         Width           =   480
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1575
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
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Image Impre 
         Height          =   480
         Left            =   1680
         MouseIcon       =   "ListEntProv.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListEntProv.frx":19A2
         ToolTipText     =   "Emision por Impresora"
         Top             =   3480
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   3600
         MouseIcon       =   "ListEntProv.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ListEntProv.frx":24EE
         ToolTipText     =   "Menu Principal"
         Top             =   3480
         Width           =   480
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6000
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListEntProv.rpt"
      Destination     =   1
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
      Left            =   5880
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      ItemData        =   "ListEntProv.frx":2D30
      Left            =   240
      List            =   "ListEntProv.frx":2D37
      TabIndex        =   1
      Top             =   5160
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "PrgListEntProv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Proceso_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
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
            .Update
        End If
    End With

    Rem Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    Rem Listado.DataFiles(1) = WEmpresa + "vent.mdb"
    
    Listado.WindowTitle = "Listado de Remitos Entregados por Proveedor"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia

    Listado.GroupSelectionFormula = "{Envio.Articulo} in " + Chr$(34) + DesdeArt.Text + Chr$(34) + " to " + Chr$(34) + HastaArt.Text + Chr$(34) _
        + " and {Envio.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34) _
        + " and {Envio.Proveedor} in " + DesdeProv.Text + " to " + HastaProv.Text
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstEnvio
        .Close
    End With
    With rstProveedor
        .Close
    End With
    With rstArticulo
        .Close
    End With
    DbsAdminis.Close
    DesdeProv.SetFocus
    PrgListEntProv.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub XConsulta_Click()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Call Opcion_Click
End Sub


Private Sub Consulta_Click()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"

    Opcion.Visible = True
    
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstProveedor
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 6)
                        IngresaItem = Auxi + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 1
            With rstArticulo
                .Index = "Codigo"
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstProveedor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Proveedor"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DesdeProv.Text = !Proveedor
                    HastaProv.Text = !Proveedor
                        Else
                    DesdeProv.Text = Claveven$
                    HastaProv.Text = Claveven$
                End If
            End With
            DesdeProv.SetFocus
            
        Case 1
            With rstArticulo

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Codigo"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    DesdeArt.Text = !Codigo
                    HastaArt.Text = !Codigo
                        Else
                    DesdeArt.Text = Claveven$
                    HastaArt.Text = Claveven$
                End If
            End With
            DesdeArt.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Proveedor
    OPEN_FILE_Envio
    OPEN_FILE_Articulo
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub DesdeProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaProv.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            A% = MsgBox(m$, 0, "Listado de Remitos Entregados por Proveedor")
            DesdeFecha.SetFocus
        End If
    End If
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            DesdeArt.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            A% = MsgBox(m$, 0, "Listado de Remitos Entregados por Proveedor")
            HastaFecha.SetFocus
        End If
    End If
End Sub

Private Sub DesdeArt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaArt.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaArt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeProv.SetFocus
    End If
    Rem  Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    DesdeProv.Text = ""
    HastaProv.Text = ""
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    DesdeArt.Text = ""
    HastaArt.Text = ""
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstProveedor
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(!Nombre) - WEspacios
                        For aa = 1 To da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                Auxi = Str$(!Proveedor)
                                Call Ceros(Auxi, 6)
                                IngresaItem = Auxi + " " + !Nombre
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Proveedor
                                WIndice.AddItem IngresaItem
                                Exit For
                            End If
                        Next aa
                        .MoveNext
                                Else
                        Exit Do
                    End If
                Loop
            End With
    
        Case 1
            With rstArticulo
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(!Descripcion) - WEspacios
                        For aa = 1 To da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                IngresaItem = !Codigo + " " + !Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Codigo
                                WIndice.AddItem IngresaItem
                                Exit For
                            End If
                        Next aa
                        .MoveNext
                                Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case Else
    End Select
    
    End If

End Sub

Private Sub DesdeProv_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub HastaProv_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub DesdeArt_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub HastaArt_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub



