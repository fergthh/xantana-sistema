VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgComponente 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Componentes"
   ClientHeight    =   6165
   ClientLeft      =   1320
   ClientTop       =   870
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   ScaleHeight     =   6165
   ScaleWidth      =   9750
   Begin VB.TextBox Costo 
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
      MaxLength       =   12
      TabIndex        =   27
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1680
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela F12"
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
         Left            =   4440
         MouseIcon       =   "Componente.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Componente.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Confirma F11"
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
         MouseIcon       =   "Componente.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "Componente.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   600
         Width           =   855
      End
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   720
         Width           =   855
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
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   11
         Text            =   " "
         Top             =   360
         Width           =   855
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
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   1455
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo F8"
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
      Left            =   6840
      MouseIcon       =   "Componente.frx":0E98
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":11A2
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Salida"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Siguiente 
      Caption         =   "Siguien. F7"
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
      Left            =   5880
      MouseIcon       =   "Componente.frx":15E4
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":18EE
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Registro Siguiente"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "Anterior F6"
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
      Left            =   4920
      MouseIcon       =   "Componente.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":203A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Registro Anterior"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primer F5"
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
      Left            =   3960
      MouseIcon       =   "Componente.frx":247C
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":2786
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Primer Registro"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
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
      Left            =   8760
      MouseIcon       =   "Componente.frx":2BC8
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":2ED2
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Salida"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Lista 
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
      Left            =   7800
      MouseIcon       =   "Componente.frx":3714
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":3A1E
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Impresion "
      Top             =   1560
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
      Left            =   3000
      MouseIcon       =   "Componente.frx":4260
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":456A
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Consulta de Datos"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpia F3"
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
      MouseIcon       =   "Componente.frx":4DAC
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":50B6
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borra  F2"
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
      Left            =   1080
      MouseIcon       =   "Componente.frx":58F8
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":5C02
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Elimina el Registro"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
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
      Left            =   120
      MouseIcon       =   "Componente.frx":6444
      MousePointer    =   99  'Custom
      Picture         =   "Componente.frx":674E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1560
      Width           =   855
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
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   8175
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
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Componente.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Bancos"
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
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   975
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
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   4935
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   1320
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   3975
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
      Height          =   2940
      ItemData        =   "Componente.frx":6F90
      Left            =   120
      List            =   "Componente.frx":6F97
      TabIndex        =   4
      Top             =   3120
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Costo"
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
      Index           =   2
      Left            =   120
      TabIndex        =   28
      Top             =   1080
      Width           =   2175
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
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
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
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgComponente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Imprime_Descripcion()
End Sub

Sub Verifica_datos()
    If Val(Costo.Text) = 0 Then
        Costo.Text = "0"
    End If
End Sub

Sub Format_datos()
    If Val(Costo.Text) <> 0 Then
        Costo.Text = Pusing("###,###.##", Costo.Text)
    End If
End Sub

Sub Imprime_Datos()
    With rstComponente
        .Index = "Codigo"
        .Seek "=", Val(Codigo.Text)
        If .NoMatch = False Then
            Codigo.Text = !Codigo
            Descripcion.Text = !Descripcion
            Costo.Text = Str$(!Costo)
            Call Format_datos
            Call Imprime_Descripcion
        End If
    End With
End Sub

Private Sub Acepta_Click()
    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    Listado.GroupSelectionFormula = "{Componente.Codigo} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Val(Codigo.Text) <> 0 Then
        With rstComponente
            .Index = "Codigo"
            .Seek "=", Val(Codigo.Text)
            If .NoMatch Then
                .AddNew
                Call Verifica_datos
                !Codigo = Val(Codigo.Text)
                !Descripcion = Descripcion.Text
                !Costo = Val(Costo.Text)
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                Call Verifica_datos
                !Codigo = Val(Codigo.Text)
                !Descripcion = Descripcion.Text
                !Costo = Val(Costo.Text)
                .Update
                .Bookmark = .LastModified
            End If
        End With
        Call CmdLimpiar_Click
        Codigo.SetFocus
    End If
End Sub

Private Sub CmdDelete_Click()
    If Val(Codigo.Text) <> 0 Then
        With rstComponente
            .Index = "Codigo"
            .Seek "=", Val(Codigo.Text)
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
    Costo.Text = ""

    With rstComponente
        .Index = "Codigo"
        .Seek "<", 9999
        If .NoMatch = False Then
            Codigo.Text = Mid$(Str$(!Codigo + 1), 2, 4)
                Else
            Codigo.Text = "1"
        End If
    End With
    
    Codigo.SetFocus
    
End Sub

Private Sub CmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstComponente
        .Close
    End With
    DbsAdminis.Close
    Codigo.SetFocus
    PrgComponente.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()
    With rstComponente
        .Index = "Codigo"
        .Seek "=", Val(Codigo.Text)
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                m$ = "No exsite registro Anterior"
                a% = MsgBox(m$, 0, "Archivo de Componentes")
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
    OPEN_FILE_Auxiliar
    OPEN_FILE_Componente
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
        Costo.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Costo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo.Text = Pusing("###,###.##", Costo.Text)
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Costo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            With rstComponente
                .Index = "Codigo"
                Claveven$ = Codigo.Text
                .Seek "=", Val(Codigo.Text)
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
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Componentes"

     Rem Opcion.Visible = True
     Opcion.ListIndex = 0
     Call Opcion_Click
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            With rstComponente
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Codigo) + " " + !Descripcion
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
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            With rstComponente

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Codigo.Text = Val(Claveven$)
                .Index = "Codigo"
                Claveven$ = Codigo.Text
                .Seek "=", Val(Claveven$)
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

    On Error GoTo Error_primer
    
    With rstComponente
        .Index = "Codigo"
        .MoveFirst
        Codigo.Text = !Codigo
        Call Imprime_Datos
        Codigo.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Componentes", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Ultimo_Click()

    On Error GoTo Error_ultimo
    
    With rstComponente
        .Index = "Codigo"
        .MoveLast
        Codigo.Text = !Codigo
        Call Imprime_Datos
        Codigo.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Componentes", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstComponente
        .Index = "Codigo"
        Claveven$ = Val(Codigo.Text)
        .Seek "=", Claveven$
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                m$ = "No exsite registro Posterior"
                a% = MsgBox(m$, 0, "Archivo de Componentes")
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
    
    With rstComponente
        .Index = "Codigo"
        .Seek "<", 9999
        If .NoMatch = False Then
            Codigo.Text = Mid$(Str$(!Codigo + 1), 2, 4)
                Else
            Codigo.Text = "1"
        End If
    End With
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            With rstComponente
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(!Descripcion) - WEspacios
                        For aa = 1 To da + 1
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                IngresaItem = Str$(!Codigo) + " " + !Descripcion
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
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Codigo_DblClick()

    Opcion.Clear
    Opcion.AddItem "Componentes"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Costo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Panta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impresora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 113
            Call CmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call Primer_Click
        Case 117
            Call Anterior_Click
        Case 118
            Call Siguiente_Click
        Case 119
            Call Ultimo_Click
        Case 120
            Call Lista_Click
        Case 121
            Call CmdClose_Click
        Case 122
            Call Acepta_Click
        Case 123
            Call Cancela_click
        Case Else
    End Select
End Sub


