VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPosdatfecha 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Pagos Posdatados a Fecha"
   ClientHeight    =   7410
   ClientLeft      =   2655
   ClientTop       =   855
   ClientWidth     =   6900
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
   ScaleWidth      =   6900
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
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   960
      TabIndex        =   4
      Top             =   120
      Width           =   4935
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
         Left            =   3720
         MouseIcon       =   "posdatfecha.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "posdatfecha.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salida"
         Top             =   2760
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
         Left            =   1440
         MouseIcon       =   "posdatfecha.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "posdatfecha.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2760
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
         Left            =   2520
         MouseIcon       =   "posdatfecha.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "posdatfecha.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Consulta de Datos"
         Top             =   2760
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
         MouseIcon       =   "posdatfecha.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "posdatfecha.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox HastaBanco 
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
         MaxLength       =   4
         TabIndex        =   11
         Text            =   " "
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox Desdebanco 
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
         MaxLength       =   4
         TabIndex        =   10
         Text            =   " "
         Top             =   1800
         Width           =   975
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2160
         TabIndex        =   7
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Height          =   300
         Left            =   2160
         TabIndex        =   1
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   300
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
      Begin VB.Label Label5 
         Caption         =   "Fecha Cierre"
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
         Left            =   720
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Banco"
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
         Left            =   720
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Banco"
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
         Left            =   720
         TabIndex        =   8
         Top             =   1800
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
         Left            =   720
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
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
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "posdat.rpt"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1080
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
      ItemData        =   "posdatfecha.frx":2D30
      Left            =   240
      List            =   "posdatfecha.frx":2D37
      TabIndex        =   2
      Top             =   4680
      Visible         =   0   'False
      Width           =   6375
   End
End
Attribute VB_Name = "PrgPosdatfecha"
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

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)

    da = 0
    With rstPosdat
        .Index = "Impre"
        .Seek ">=", 0
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

    With rstPagos
    
            .Index = "CLAVE"
            .MoveFirst
            
            Do
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If !Orden = 19 Then Stop
            
                If !FechaOrd <= WFechaOrd Then
                
                If !Banco2 >= Val(Desdebanco.Text) And !Banco2 <= Val(HastaBanco.Text) Then
                    
                    WFechaCheque = Right$(!Fecha2, 4) + Mid$(!Fecha2, 4, 2) + Left$(!Fecha2, 2)
                    
                    If !Fecha2 <> !Fecha Then
                    
                        If WFechaCheque >= WDesde And WFechaCheque <= WHasta Then
                    
                            WBanco = !Banco2
                            WFecha = !Fecha
                            WImporte = !Importe2
                            WCheque = !Numero2
                            WVencimiento = !Fecha2
                            WProveedor = !Proveedor
                            WObservaciones = !Observaciones
                        
                            With RstProveedor
                                .Index = "Proveedor"
                                .Seek "=", WProveedor
                                If .NoMatch = False Then
                                    WObservaciones = !Nombre
                                End If
                            End With
                
                            With rstPosdat
                                .AddNew
                                !Banco = WBanco
                                !Fecha = WFecha
                                !Cheque = WCheque
                                !Proveedor = WProveedor
                                !Importe = WImporte
                                !Vencimiento = WVencimiento
                                !FechaOrd = Right$(WVencimiento, 4) + Mid$(WVencimiento, 4, 2) + Left$(WVencimiento, 2)
                                !Observaciones = Left$(WObservaciones, 20)
                                .Update
                            End With
                        End If
                    End If
                    
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With

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
            !Actividad = "Del " + Desde.Text + " al " + Hasta.Text + " (" + Fecha.Text + ")"
            .Update
        End If
    End With

    Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    
    Listado.WindowTitle = "Listado Cheque Posdatados a Fecha"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Posdat.Banco} in " + Desde.Text + " to " + Hasta.Text
    
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
    With rstPagos
        .Close
    End With
    With RstProveedor
        .Close
    End With
    With rstPosdat
        .Close
    End With
    With rstBanco
        .Close
    End With
    
    DbsAdminis.Close
    
    Desde.SetFocus
    PrgPosdatfecha.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Pagos
    OPEN_FILE_Proveedor
    OPEN_FILE_Posdat
    OPEN_FILE_Banco
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdebanco.SetFocus
    End If
End Sub

Private Sub DesdeBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaBanco.SetFocus
    End If
End Sub

Private Sub HastaBanco_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Desdebanco.Text = ""
    HastaBanco.Text = ""
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    With rstBanco
        .Index = "Banco"
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = Str$(!Banco) + " " + !Nombre
                Pantalla.AddItem IngresaItem
                IngresaItem = !Banco
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
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    With rstBanco
        Indice = Pantalla.ListIndex
        Claveven$ = WIndice.List(Indice)
        Desdebanco.Text = Claveven$
        .Index = "Banco"
        Claveven$ = Desdebanco.Text
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Desdebanco.Text = !Banco
            HastaBanco.Text = !Banco
                Else
            Desdebanco.Text = Claveven$
            HastaBanco.Text = Claveven$
        End If
    End With
    Desdebanco.SetFocus
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    With rstBanco
        .Index = "Banco"
        .MoveFirst
        Do
            If .EOF = False Then
                da = Len(!Nombre) - WEspacios
                For aa = 1 To da + 1
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        IngresaItem = Str$(!Banco) + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Banco
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
    
    Exit Sub
    
WError:
    Resume Next

End Sub


