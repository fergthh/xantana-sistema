VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgComparaGas 
   Caption         =   "Listado Comparativo de Gastos por Cuenta Contable"
   ClientHeight    =   8115
   ClientLeft      =   2925
   ClientTop       =   480
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   6210
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
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   5895
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
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5520
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5295
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
         Left            =   480
         MouseIcon       =   "ComparaGas.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ComparaGas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2280
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
         Left            =   2880
         MouseIcon       =   "ComparaGas.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ComparaGas.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Consulta de Datos"
         Top             =   2280
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
         Left            =   1680
         MouseIcon       =   "ComparaGas.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ComparaGas.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2280
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
         Left            =   4080
         MouseIcon       =   "ComparaGas.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ComparaGas.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salida"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Mes 
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
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Ano 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   11
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox DesdeCuenta 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   8
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox HastaCuenta 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   7
         Text            =   " "
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Left            =   840
         TabIndex        =   10
         Top             =   1320
         Width           =   1335
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
         Left            =   840
         TabIndex        =   9
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Año"
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
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Mes"
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
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5760
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ComparaGas.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Movimietos de Bancos"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgComparaGas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WMes As String
Private WAno As String

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    On Error GoTo Error_Programa
    
    WMes = Mes.Text
    WAno = Ano.Text
    Call Ceros(WMes, 2)
    Call Ceros(WAno, 4)

    WDesde = WAno + WMes + "01"
    WHasta = WAno + WMes + "31"
    
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
            !Actividad = "Periodo : " + Mes.Text + "/" + Ano.Text
            .Update
        End If
    End With

    With rstCompara
        .Index = "Codigo"
        .MoveFirst
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
    
    For WRenglon = 1 To 1000
    
        With rstPtoCue
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 3)
        
            .Index = "Clave"
            .Seek "=", WMes + WAno + Auxi1
            If .NoMatch = False Then
        
                WCuenta = !Cuenta
                WImporte = !Importe
                With rstCuenta
                    .Index = "Cuenta"
                    .Seek "=", WCuenta
                    If .NoMatch = False Then
                        WDescripcion = !Descripcion
                    End If
                End With
            
                With rstCompara
                    .Index = "Codigo"
                    .AddNew
                    !Codigo = WCuenta
                    !ImpreCodigo = 0
                    !Descripcion = WDescripcion
                    !Impo1 = 0
                    !Importe1 = WImporte
                    !Impo2 = 0
                    !Importe2 = 0
                    !Impo3 = 0
                    !Importe3 = 0
                    !Impo4 = 0
                    !Importe4 = 0
                    !Impo5 = 0
                    !Importe5 = 0
                    !Impo6 = 0
                    !Importe6 = 0
                    !Impo7 = 0
                    !Importe7 = 0
                    .Update
                End With
            
            End If
        
        End With
    
    Next WRenglon
    
    With rstImpcyb
        .Index = "Clave"
        .MoveFirst
        Do
            If WDesde <= !ordfecha And !ordfecha <= WHasta Then
                
                If !Letra <> "X" Then
                
                    WCuenta = !Cuenta
                    WDebito = !Debito
                    WCredito = !Credito
                    
                    With rstCompara
                        .Index = "Codigo"
                        .Seek "=", WCuenta
                        If .NoMatch = False Then
                            .Edit
                            !Impo1 = !Impo1 + WDebito - WCredito
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
    
    Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    Listado.Action = 1
    
    Exit Sub
    
Error_Programa:
     Rem coderr = Err
     Rem Call Errores(coderr, "Error en el sistema", "Se produjo el error " + Str$(coderr))
     Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstCuenta
        .Close
    End With
    With rstPtoCue
        .Close
    End With
    With rstImpcyb
        .Close
    End With
    With rstCompara
        .Close
    End With
    DbsAdminis.Close
    Mes.SetFocus
    PrgComparaGas.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Mes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Mes.Text) > 0 And Val(Mes.Text) < 13 Then
            Ano.SetFocus
                Else
            Mes.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Mes.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCuenta.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesdeCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCuenta.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeCuenta.Text = ""
    End If
End Sub

Private Sub HastaCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mes.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaCuenta.Text = ""
    End If
End Sub

Sub Form_Load()
    
    Mes.Text = ""
    Ano.Text = ""
    DesdeCuenta.Text = ""
    HastaCuenta.Text = ""
    Frame2.Visible = True
    
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
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
    
    Exit Sub
    
WError:
    Resume Next

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

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    With rstCuenta
        .Index = "Cuenta"
        .MoveFirst
        Do
            If .EOF = False Then
                da = Len(!Descripcion) - WEspacios
                For aa = 1 To da + 1
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
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Cuenta
    OPEN_FILE_Compara
    OPEN_FILE_Impcyb
    OPEN_FILE_PtoCue
End Sub

Private Sub Mes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ano_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdeCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
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





