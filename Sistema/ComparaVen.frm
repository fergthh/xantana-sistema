VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgComparaVen 
   Caption         =   "Listado Comparativo de Ventas por Vendedor"
   ClientHeight    =   8115
   ClientLeft      =   2760
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
      Top             =   3720
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
      Top             =   4080
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
      Height          =   3375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   5295
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
         MouseIcon       =   "ComparaVen.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ComparaVen.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Salida"
         Top             =   2160
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
         MouseIcon       =   "ComparaVen.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ComparaVen.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2160
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
         MouseIcon       =   "ComparaVen.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ComparaVen.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Consulta de Datos"
         Top             =   2160
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
         Left            =   480
         MouseIcon       =   "ComparaVen.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ComparaVen.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2160
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
         Left            =   2640
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   11
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox DesdeVendedor 
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   8
         Text            =   " "
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox HastaVendedor 
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
         Left            =   2640
         MaxLength       =   4
         TabIndex        =   7
         Text            =   " "
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Desde Vendedor"
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
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Vendedor"
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
         Width           =   1575
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
      ReportFileName  =   "ComparaVen.rpt"
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
Attribute VB_Name = "PrgComparaVen"
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
    
    For WRenglon = 1 To 100
    
        With rstPtoVend
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
        
            .Index = "Clave"
            .Seek "=", WMes + WAno + Auxi1
            If .NoMatch = False Then
        
                WVendedor = !Vendedor
                WImporte = !Importe
                With rstVendedor
                    .Index = "Codigo"
                    .Seek "=", WVendedor
                    If .NoMatch = False Then
                        WDescripcion = !Nombre
                    End If
                End With
            
                With rstCompara
                    .Index = "Codigo"
                    .AddNew
                    !Codigo = WVendedor
                    !ImpreCodigo = WVendedor
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
    
    With rstCtaCte
        .Index = "Clave"
        .MoveFirst
        Do
            If WDesde <= !OrdFecha And !OrdFecha <= WHasta Then
                
                If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Or Val(!Tipo) = 3 Or Val(!Tipo) = 4 Or Val(!Tipo) = 5 Then
                
                    Wcliente = !Cliente
                
                    With rstClientes
                        .Index = "Cliente"
                        .Seek "=", Wcliente
                        If .NoMatch = False Then
                            WVendedor = !Vendedor
                        End If
                    End With
                    WTotal = !Total
                    
                    With rstCompara
                        .Index = "Codigo"
                        .Seek "=", WVendedor
                        If .NoMatch = False Then
                            .Edit
                            !Impo1 = !Impo1 + WTotal
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
    With rstVendedor
        .Close
    End With
    With rstPtoVend
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    With rstCompara
        .Close
    End With
    DbsAdminis.Close
    Mes.SetFocus
    PrgComparaVen.Hide
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
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeVendedor.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
End Sub

Private Sub DesdeVendedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaVendedor.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeVendedor.Text = ""
    End If
End Sub

Private Sub HastaVendedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mes.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaVendedor.Text = ""
    End If
End Sub

Sub Form_Load()
    
    Mes.Text = ""
    Ano.Text = ""
    DesdeVendedor.Text = ""
    HastaVendedor.Text = ""
    Frame2.Visible = True
    
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    With rstVendedor
        .Index = "Codigo"
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = Str$(!Codigo) + " " + !Nombre
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
    With rstVendedor
        Indice = Pantalla.ListIndex
        .Index = "Codigo"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            DesdeVendedor.Text = !Codigo
            HastaVendedor.Text = !Codigo
                Else
            DesdeVendedor.Text = WIndice.List(Indice)
            HastaVendedor.Text = WIndice.List(Indice)
        End If
    End With
    DesdeVendedor.SetFocus
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    With rstVendedor
        .Index = "Codigo"
        .MoveFirst
        Do
            If .EOF = False Then
                da = Len(!Nombre) - WEspacios
                For aa = 1 To da + 1
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        IngresaItem = Str$(!Codigo) + " " + !Nombre
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

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Vendedor
    OPEN_FILE_Compara
    OPEN_FILE_Ctacte
    OPEN_FILE_PtoVend
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

Private Sub DesdeVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaVendedor_KeyDown(KeyCode As Integer, Shift As Integer)
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




