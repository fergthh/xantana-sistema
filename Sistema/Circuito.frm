VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCircuito 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Circuito de Pago"
   ClientHeight    =   3375
   ClientLeft      =   3210
   ClientTop       =   1845
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3375
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   4215
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Image Panta 
         Height          =   480
         Left            =   720
         MouseIcon       =   "Circuito.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Circuito.frx":030A
         ToolTipText     =   "Emision por Pantalla"
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Impre 
         Height          =   480
         Left            =   1920
         MouseIcon       =   "Circuito.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "Circuito.frx":0E56
         ToolTipText     =   "Emision por Impresora"
         Top             =   2040
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   3120
         MouseIcon       =   "Circuito.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "Circuito.frx":19A2
         ToolTipText     =   "Menu Principal"
         Top             =   2040
         Width           =   480
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
         Left            =   480
         TabIndex        =   4
         Top             =   720
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
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5040
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Circuito.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva Compras"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgCircuito"
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

    Rem Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    
    Listado.WindowTitle = "Listado de Iva Compras"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
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
    
    da = 0
    With rstIva
        .Index = "Clave"
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
            
    With rstIvacomp
            .Index = "Iva"
            .MoveFirst
            Do
                WFecha = Right$(!Periodo, 4) + Mid$(!Periodo, 4, 2) + Left$(!Periodo, 2)
                
                If WDesde <= WFecha And WFecha <= WHasta Then
                
                If Tipo.ListIndex = 1 Or !Letra <> "X" Then
                
                    WClave = !Clave
                    WProveedor = !Proveedor
                    WTipo = !Tipo
                    WLetra = !Letra
                    WPunto = !Punto
                    WNumero = !Numero
                    WFecha = !Fecha
                    WVencimiento = !Vencimiento
                    WPeriodo = !Periodo
                    WNeto = !Neto
                    WIva21 = !Iva21
                    WIva5 = !Iva5
                    WIva27 = !Iva27
                    WIva105 = !Iva105
                    WIb = !Ib
                    WExento = !Exento
                    WImpre = !Impre
                    WOrdfecha = !ordfecha
                    WContado = !Contado
                
                    With rstIva
                        .AddNew
                        !Clave = WClave
                        !Proveedor = WProveedor
                        !Tipo = WTipo
                        !Letra = WLetra
                        !Punto = WPunto
                        !Numero = WNumero
                        !Fecha = WFecha
                        !Vencimiento = WVencimiento
                        !Periodo = WPeriodo
                        !Concepto = WConcepto
                        !Neto = WNeto
                        !Iva21 = WIva21
                        !Iva5 = WIva5
                        !Iva27 = WIva27
                        !Iva105 = WIva105
                        !Ib = WIb
                        !Exento = WExento
                        !Impre = WImpre
                        !ordfecha = WOrdfecha
                        !Contado = WContado
                        !Empresa = 1
                        .Update
                    End With
                End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Rem Listado.GroupSelectionFormula = "{Ivacomp.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstIva
        .Close
    End With
    With rstIvacomp
        .Close
    End With
    DbsAdminis.Close
    Desde.SetFocus
    PrgCircuito.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Iva
    OPEN_FILE_Ivacomp
    OPEN_FILE_Auxiliar
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
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Iva Compras"
    Tipo.AddItem "Completo"
    
    Tipo.ListIndex = 0

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Frame2.Visible = True
End Sub

