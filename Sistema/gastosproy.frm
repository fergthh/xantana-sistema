VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGastosProy 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Control de Gastos por Centro de Costo"
   ClientHeight    =   6765
   ClientLeft      =   3165
   ClientTop       =   1200
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   6765
   ScaleWidth      =   5655
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
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   5055
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
         MouseIcon       =   "gastosproy.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "gastosproy.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2400
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
         Left            =   2760
         MouseIcon       =   "gastosproy.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "gastosproy.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Consulta de Datos"
         Top             =   2400
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
         Left            =   1560
         MouseIcon       =   "gastosproy.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "gastosproy.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2400
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
         Left            =   3840
         MouseIcon       =   "gastosproy.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "gastosproy.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salida"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox HastaProy 
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
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox DesdeProy 
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
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   2640
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
      Begin VB.Label Label4 
         Caption         =   "Hasta Centro de Costos"
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
         TabIndex        =   6
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Centro de Costo"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emision"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5280
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "GastosProy.rpt"
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   1920
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
      Height          =   2400
      ItemData        =   "gastosproy.frx":2D30
      Left            =   120
      List            =   "gastosproy.frx":2D37
      TabIndex        =   1
      Top             =   4320
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "PrgGastosProy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WProyecto As String
Dim WConcepto As String

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
            !Actividad = "al " + Fecha.Text
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Control de Gastos por Proyecto"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFecha = WAno + WMes + "31"
    
    da = ""
    With rstGastosProy
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
    
    With rstImpproy
        .Index = "Clave"
        .MoveFirst
        Do
            If WFecha >= !ordfecha Then
                If DesdeProy.Text <= !Proyecto And !Proyecto <= HastaProy.Text Then
                    WOrdfecha = !ordfecha
                    WProyecto = !Proyecto
                    WConcepto = !Concepto
                    WImporte = !Importe
                    Call Ceros(WConcepto, 4)
                    WClave = WProyecto + WConcepto
                    With rstGastosProy
                        .Index = "Clave"
                        .Seek "=", WClave
                        If .NoMatch Then
                            .AddNew
                            !Clave = WClave
                            !Proyecto = WProyecto
                            !Concepto = Val(WConcepto)
                            If Left$(WFecha, 6) = Left$(WOrdfecha, 6) Then
                                !Importe1 = WImporte
                                    Else
                                !Importe1 = 0
                            End If
                            !Importe2 = WImporte
                            !Importe3 = 0
                            !Importe4 = 0
                            !Porce = 0
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            If Left$(WFecha, 6) = Left$(WOrdfecha, 6) Then
                                !Importe1 = !Importe1 + WImporte
                            End If
                            !Importe2 = !Importe2 + WImporte
                            .Update
                            .Bookmark = .LastModified
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
    
    With rstPto
        .Index = "Clave"
        .MoveFirst
        Do
            If DesdeProy.Text <= !Proyecto And !Proyecto <= HastaProy.Text Then
                WProyecto = !Proyecto
                WConcepto = !Concepto
                WImporte = !Importe
                Call Ceros(WConcepto, 4)
                WClave = WProyecto + WConcepto
                With rstGastosProy
                    .Index = "Clave"
                    .Seek "=", WClave
                    If .NoMatch Then
                        .AddNew
                        !Clave = WClave
                        !Proyecto = WProyecto
                        !Concepto = Val(WConcepto)
                        !Importe1 = 0
                        !Importe2 = 0
                        !Importe3 = WImporte
                        !Importe4 = 0
                        !Porce = 0
                        .Update
                        .Bookmark = .LastModified
                            Else
                        .Edit
                        !Importe3 = !Importe3 + WImporte
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
            End If
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    With rstGastosProy
        .Index = "Clave"
        .MoveFirst
        Do
            .Edit
            WProyecto = !Proyecto
            WPorce = 0
            With rstPto
                .Index = "Proyecto"
                .Seek "=", WProyecto
                If .NoMatch = False Then
                    WPorce = !Avance
                End If
            End With
            !Porce = WPorce
            .Update
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    Rem Listado.GroupSelectionFormula = "{Impproy.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34) + " and {Impproy.proyecto} in " + Chr$(34) + DesdeProy.Text + Chr$(34) + " to " + Chr$(34) + HastaProy.Text + Chr$(34)
    Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstImpproy
        .Close
    End With
    With rstProyecto
        .Close
    End With
    With rstGastosProy
        .Close
    End With
    With rstPto
        .Close
    End With
    DbsAdminis.Close
    Fecha.SetFocus
    PrgGastosProy.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Impproy
    OPEN_FILE_Pto
    OPEN_FILE_Proyecto
    OPEN_FILE_GastosProy
    OPEN_FILE_Auxiliar
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            DesdeProy.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub desdeProy_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaProy.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeProy.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaProy_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaProy.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    DesdeProy.Text = ""
    HastaProy.Text = ""
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    With rstProyecto
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
    With rstProyecto
        Indice = Pantalla.ListIndex
        .Index = "Codigo"
        .Seek "=", WIndice.List(Indice)
        If .NoMatch = False Then
            DesdeProy.Text = !Codigo
            HastaProy.Text = !Codigo
                Else
            DesdeProy.Text = Claveven$
            HastaProy.Text = Claveven$
        End If
    End With
    DesdeProy.SetFocus
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    With rstProyecto
        .Index = "Codigo"
        .MoveFirst
        Do
            If .EOF = False Then
                da = Len(!Descripcion) - WEspacios
                For aa = 1 To da + 1
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
    
    End If
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdeProy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaProy_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Panta_Click
        Case 115
            Call Consulta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub













