VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgColor 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Colores"
   ClientHeight    =   6120
   ClientLeft      =   1170
   ClientTop       =   825
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   ScaleHeight     =   6120
   ScaleWidth      =   9465
   Begin VB.TextBox Tipo 
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
      MaxLength       =   6
      TabIndex        =   19
      Text            =   " "
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5160
      TabIndex        =   18
      Top             =   1800
      Width           =   3015
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "color.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "color.frx":030A
         ToolTipText     =   "Primer Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "color.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "color.frx":0A56
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "color.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "color.frx":11A2
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "color.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "color.frx":18EE
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
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
      TabIndex        =   17
      Top             =   3000
      Visible         =   0   'False
      Width           =   6855
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
      Left            =   960
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.ComboBox Pintura 
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
      Left            =   2400
      TabIndex        =   2
      Text            =   " "
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Codigo 
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
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   3240
      TabIndex        =   7
      Top             =   3720
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox Hasta 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   13
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Desde 
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
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   240
         Width           =   975
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
         Left            =   2160
         TabIndex        =   11
         Top             =   1080
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
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   3960
         MouseIcon       =   "color.frx":1D30
         MousePointer    =   99  'Custom
         Picture         =   "color.frx":203A
         ToolTipText     =   "Cancela la Impresion"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Acepta 
         Height          =   480
         Left            =   3960
         MouseIcon       =   "color.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "color.frx":2786
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1080
         Width           =   480
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
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8400
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Color.rpt"
      Destination     =   1
      WindowTitle     =   "Listados de Proveedores"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
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
      ItemData        =   "color.frx":2BC8
      Left            =   240
      List            =   "color.frx":2BCF
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   6855
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
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   1
      Top             =   480
      Width           =   5175
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2760
      MouseIcon       =   "color.frx":2BDD
      MousePointer    =   99  'Custom
      Picture         =   "color.frx":2EE7
      ToolTipText     =   "Consulta de Datos"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3600
      MouseIcon       =   "color.frx":3729
      MousePointer    =   99  'Custom
      Picture         =   "color.frx":3A33
      ToolTipText     =   "Impresion "
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4440
      MouseIcon       =   "color.frx":4275
      MousePointer    =   99  'Custom
      Picture         =   "color.frx":457F
      ToolTipText     =   "Salida"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1080
      MouseIcon       =   "color.frx":4DC1
      MousePointer    =   99  'Custom
      Picture         =   "color.frx":50CB
      ToolTipText     =   "Elimina el Registro"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   240
      MouseIcon       =   "color.frx":590D
      MousePointer    =   99  'Custom
      Picture         =   "color.frx":5C17
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   1920
      MouseIcon       =   "color.frx":6459
      MousePointer    =   99  'Custom
      Picture         =   "color.frx":6763
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label Label10 
      Caption         =   "Pintura"
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
      Left            =   240
      TabIndex        =   15
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label8 
      Caption         =   "Tipo de Color"
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
      Left            =   240
      TabIndex        =   14
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion"
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
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Color"
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
      TabIndex        =   3
      Top             =   180
      Width           =   2055
   End
End
Attribute VB_Name = "PrgColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Auxi As String

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "0"
    End If
    If Val(Tipo.Text) = 0 Then
        Tipo.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()
    With rstColor
        .Index = "Codigo"
        .Seek "=", Codigo.Text
        If .NoMatch = False Then
            Codigo.Text = !Codigo
            Descripcion.Text = !Descripcion
            Tipo.Text = !Tipo
            Pintura.ListIndex = !Pintura
            Call Format_datos
        End If
    End With
    
End Sub

Sub Imprime_Descripcion()
End Sub

Private Sub Acepta_Click()

    With rstEmpresa
        .Index = "Empresa"
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
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Colores"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Listado.GroupSelectionFormula = "{Color.Codigo} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Codigo.SetFocus
    Rem Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
    
        Call Verifica_datos
        WPasa = "S"
    
        If WPasa = "S" Then
    
        With rstColor
            .Index = "Codigo"
            .Seek "=", Codigo.Text
            If .NoMatch Then
                .AddNew
                Call Verifica_datos
                !Codigo = Codigo.Text
                !Descripcion = Descripcion.Text
                !Tipo = Tipo.Text
                !Pintura = Pintura.ListIndex
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                Call Verifica_datos
                !Codigo = Codigo.Text
                !Descripcion = Descripcion.Text
                !Tipo = Tipo.Text
                !Pintura = Pintura.ListIndex
                .Update
                .Bookmark = .LastModified
            End If
        End With
        Call CmdLimpiar_Click
        
        End If
        
        Codigo.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
        With rstColor
            .Index = "Codigo"
            .Seek "=", Codigo.Text
            If .NoMatch = False Then
                T$ = "Borrar Registro"
                m$ = "Desea Borrar el Registro "
                Respuesta% = MsgBox(m$, 32 + 4, T$)
                If Respuesta% = 6 Then
                    .Delete
                    Call CmdLimpiar_Click
                End If
            End If
        End With
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Codigo.Text = ""
    Descripcion.Text = ""
    Tipo.Text = ""
    Pintura.ListIndex = 0
    Codigo.SetFocus
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    
    With rstColor
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    
    DbsAdminis.Close
    
    Codigo.SetFocus
    PrgColor.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Anterior_Click()
    With rstColor
        .Index = "Codigo"
        .Seek "=", Codigo.Text
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                m$ = "No exsite registro Anterior"
                A% = MsgBox(m$, 0, "Archivo de Colores")
                .MoveFirst
            End If
            Codigo.Text = !Codigo
            Call Imprime_Datos
            Codigo.SetFocus
        End If
    End With
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Color
    OPEN_FILE_Auxiliar
End Sub


Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Tipo.SetFocus
    End If
End Sub

Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Codigo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            With rstColor
                .Index = "Codigo"
                Claveven$ = Codigo.Text
                .Seek "=", Codigo.Text
                If .NoMatch Then
                    CmdLimpiar_Click
                    Codigo.Text = Claveven$
                        Else
                    Codigo.Text = !Codigo
                    Call Imprime_Datos
                End If
            End With
        End If
        Descripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Color_DblClick()
    Call Consulta_Click
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

    Opcion.Clear
    
    Opcion.AddItem "Colores"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Call Opcion_Click
End Sub


Private Sub WConsulta_Click()

    Opcion.Clear
    
    Opcion.AddItem "Colores"

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
            With rstColor
                .Index = "Descripcion"
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = Str$(!Codigo)
                        Call Ceros(Auxi, 4)
                        IngresaItem = Auxi + " " + !Descripcion
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
            With rstColor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Codigo.Text = Claveven$
                .Index = "Codigo"
                Claveven$ = Codigo.Text
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Codigo.Text = !Codigo
                    Call Imprime_Datos
                        Else
                    CmdLimpiar_Click
                    Codigo.Text = Claveven$
                End If
            End With
            Codigo.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    Rem On Error GoTo Error_primer
    With rstColor
        .Index = "Codigo"
        .MoveFirst
        Codigo.Text = !Codigo
        Call Imprime_Datos
        Codigo.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Colores", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Ultimo_Click()
    Rem On Error GoTo Error_ultimo
    With rstColor
        .Index = "Codigo"
        .MoveLast
        Codigo.Text = !Codigo
        Call Imprime_Datos
        Codigo.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Color", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstColor
        .Index = "Codigo"
        Claveven$ = Codigo.Text
        .Seek "=", Claveven$
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                m$ = "No exsite registro Posterior"
                A% = MsgBox(m$, 0, "Archivo de Colores")
                Call Ultimo_Click
            End If
            Codigo.Text = !Codigo
            Call Imprime_Datos
            Codigo.SetFocus
        End If
    End With
End Sub

Sub Form_Load()

    Codigo.Text = ""
    Descripcion.Text = ""
    Tipo.Text = ""
    
    Pintura.Clear
    
    Pintura.AddItem ""
    Pintura.AddItem "Lleva Pintura"
    Pintura.AddItem "No Lleva Pintura"
    
    Pintura.ListIndex = 0

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            With rstColor
                .Index = "Descripcion"
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(!Descripcion) - WEspacios
                        For aa = 1 To da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                Auxi = Str$(!Codigo)
                                Call Ceros(Auxi, 4)
                                IngresaItem = Auxi + " " + !Descripcion
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


Private Sub Codigo_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Color"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub


