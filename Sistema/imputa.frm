VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImputa 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Imputaciones Contables"
   ClientHeight    =   7350
   ClientLeft      =   2790
   ClientTop       =   855
   ClientWidth     =   6525
   LinkTopic       =   "Form2"
   ScaleHeight     =   7350
   ScaleWidth      =   6525
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
      TabIndex        =   13
      Top             =   4200
      Visible         =   0   'False
      Width           =   6015
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
      Left            =   240
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   360
      TabIndex        =   4
      Top             =   240
      Width           =   5655
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Text            =   " "
         Top             =   2520
         Width           =   2175
      End
      Begin VB.TextBox HastaCuenta 
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   3
         Text            =   " "
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox DesdeCuenta 
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
         Left            =   1920
         MaxLength       =   10
         TabIndex        =   2
         Text            =   " "
         Top             =   1440
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4560
         MouseIcon       =   "imputa.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "imputa.frx":030A
         ToolTipText     =   "Menu Principal"
         Top             =   3000
         Width           =   480
      End
      Begin VB.Image Impre 
         Height          =   480
         Left            =   4560
         MouseIcon       =   "imputa.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "imputa.frx":0E56
         ToolTipText     =   "Emision por Impresora"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Consulta 
         Height          =   480
         Left            =   4560
         MouseIcon       =   "imputa.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "imputa.frx":19A2
         ToolTipText     =   "Consulta de Datos"
         Top             =   2160
         Width           =   480
      End
      Begin VB.Image Panta 
         Height          =   480
         Left            =   4560
         MouseIcon       =   "imputa.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "imputa.frx":24EE
         ToolTipText     =   "Emision por Pantalla"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Listado"
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
         Left            =   360
         TabIndex        =   12
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cuenta"
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
         Left            =   360
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cuenta"
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
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6000
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "imputa.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Imputaciones Contables"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgImputa"
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
            !Actividad = "Desde el " + Desde.Text + " hasta el " + Hasta.Text
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Imputaciones Contables de Compras"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Da = ""
    With rstImputac
        .Index = "Clave"
        .Seek ">=", ""
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
    
    With rstImpcyb
        .Index = "Clave"
        .MoveFirst
        Do
            If !OrdFecha >= WDesde And !OrdFecha <= WHasta Then
                If !Cuenta >= DesdeCuenta.Text And !Cuenta <= HastaCuenta.Text Then
                
                    WClave = !Clave
                    WProveedor = !Proveedor
                    WTipo = !Tipo
                    WLetra = !Letra
                    WPunto = !Punto
                    WNumero = !Numero
                    WRenglon = !Renglon
                    WCuenta = !Cuenta
                    WDebito = !Debito
                    WCredito = !Credito
                    WFecha = !Fecha
                    WOrdFecha = !OrdFecha
                    WObservaciones = !Observaciones
                
                    With rstImputac
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Proveedor = WProveedor
                            !TipoComp = WTipo
                            !LetraComp = WLetra
                            !PuntoComp = WPunto
                            !NroComp = WNumero
                            !Renglon = WRenglon
                            !Fecha = WFecha
                            !Observaciones = Left$(WObservaciones, 50)
                            !Cuenta = WCuenta
                            !Debito = WDebito
                            !Credito = WCredito
                            !FechaOrd = WFechaOrd
                            !Titulo = ""
                            !Clave = WClave
                            !Titulolist = ""
                            .Update
                        End If
                    End With
                End If
            End If
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "Imputa.rpt"
            Else
        Listado.ReportFileName = "Imputa2.rpt"
    End If
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstCuenta
        .Close
    End With
    With rstImputac
        .Close
    End With
    With rstImpcyb
        .Close
    End With
    DbsAdminis.Close
    Desde.SetFocus
    PrgImputa.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            DesdeCuenta.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub DesdeCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCuenta.SetFocus
    End If
End Sub

Private Sub HastaCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    With rstCuenta
        .Index = "Cuenta"
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
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    With rstCuenta
        Indice = Pantalla.ListIndex
        .Index = "Cuenta"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            DesdeCuenta.Text = !Cuenta
            HastaCuenta.Text = !Cuenta
                Else
            DesdeCuenta.Text = WIndice.List(Indice)
            HastaCuenta.Text = WIndice.List(Indice)
        End If
    End With
    DesdeCuenta.SetFocus
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeCuenta.Text = ""
    HastaCuenta.Text = ""
    Frame2.Visible = True
    
    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    With rstCuenta
        .Index = "Cuenta"
        .MoveFirst
        Do
            If .EOF = False Then
                Da = Len(!Descripcion) - WEspacios
                For aa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                        IngresaItem = !Cuenta + " " + !Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cuenta
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
    
    End If

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Cuenta
    OPEN_FILE_Impcyb
    OPEN_FILE_Imputac
End Sub

