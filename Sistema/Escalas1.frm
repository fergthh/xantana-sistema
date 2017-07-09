VERSION 5.00
Begin VB.Form PrgEscalas1 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Proyectos"
   ClientHeight    =   7875
   ClientLeft      =   2220
   ClientTop       =   450
   ClientWidth     =   7740
   LinkTopic       =   "Form2"
   ScaleHeight     =   7875
   ScaleWidth      =   7740
   Begin VB.Frame Cuadro1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Parametros Ganancias (2784)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   3495
      Left            =   480
      TabIndex        =   18
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Minimo1 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   26
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Minimo2 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   25
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Minimo3 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   24
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Escala1 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   23
         Text            =   " "
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Escala2 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   22
         Text            =   " "
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Escala3 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   21
         Text            =   " "
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Escala4 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   20
         Text            =   " "
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox RetMinima 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   19
         Text            =   " "
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo Pago Otros"
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
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo Pago Honorarios"
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
         TabIndex        =   33
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Minimo Pago Alquileres"
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
         TabIndex        =   32
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 1"
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
         Index           =   3
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 2"
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
         Index           =   4
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 3"
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
         Index           =   5
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 4"
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
         Index           =   6
         Left            =   240
         TabIndex        =   28
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Retencion Minima"
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
         Index           =   7
         Left            =   240
         TabIndex        =   27
         Top             =   3120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parametros Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1815
      Left            =   480
      TabIndex        =   9
      Top             =   5640
      Width           =   5175
      Begin VB.TextBox Text15 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   13
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text14 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text13 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   11
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text12 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   10
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 1"
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
         Index           =   15
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 2"
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
         Index           =   14
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 3"
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
         Index           =   13
         Left            =   360
         TabIndex        =   15
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 4"
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
         Index           =   12
         Left            =   360
         TabIndex        =   14
         Top             =   1320
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parametros Retencion de Iva (3125)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1815
      Left            =   480
      TabIndex        =   0
      Top             =   3720
      Width           =   5175
      Begin VB.TextBox Text11 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   7
         Text            =   " "
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text10 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   5
         Text            =   " "
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox Text9 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   3
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Text8 
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   1
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 4"
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
         Index           =   11
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Importe Escala 3"
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
         Index           =   10
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Porcenta"
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
         Index           =   9
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "POrcentaje Compra Bienes"
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
         Index           =   8
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6480
      MouseIcon       =   "Escalas1.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Escalas1.frx":030A
      ToolTipText     =   "Salida"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   6480
      MouseIcon       =   "Escalas1.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "Escalas1.frx":0E56
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1200
      Width           =   480
   End
End
Attribute VB_Name = "PrgEscalas1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Imprime_Nombre()
    With rstCuenta
        .Index = "Cuenta"
        .Seek "=", Cuenta.Text
        If .NoMatch = False Then
            DesCuenta.Caption = !Descripcion
                Else
            DesCuenta.Caption = ""
        End If
    End With
    With rstClientes
        .Index = "Cliente"
        .Seek "=", Cliente.Text
        If .NoMatch = False Then
            DesCliente.Caption = !Razon
                Else
            DesCliente.Caption = ""
        End If
    End With
End Sub

Sub Verifica_datos()
    Rem If Val(Cuenta.text) = 0 Then
    Rem     Cuenta.text = "0"
    Rem End If
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()
    With rstProyecto
        .Index = "Codigo"
        .Seek "=", Codigo.Text
        If .NoMatch = False Then
            Codigo.Text = !Codigo
            Descripcion.Text = !Descripcion
            Cliente.Text = !Cliente
            Cuenta.Text = !Cuenta
            Tipo.ListIndex = !Tipo
            Call Format_datos
            Call Imprime_Nombre
        End If
    End With
End Sub

Private Sub Acepta_Click()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            .Update
        End If
    End With

    Listado.GroupSelectionFormula = "{Proyecto.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
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
    If Codigo.Text <> "" Then
        With rstProyecto
            .Index = "Codigo"
            .Seek "=", Codigo.Text
            If .NoMatch Then
                .AddNew
                Call Verifica_datos
                !Codigo = Codigo.Text
                !Descripcion = Descripcion.Text
                !Cliente = Val(Cliente.Text)
                !Cuenta = Cuenta.Text
                !Tipo = Tipo.ListIndex
                .Update
                .Bookmark = .LastModified
                    Else
                .Edit
                Call Verifica_datos
                !Codigo = Codigo.Text
                !Descripcion = Descripcion.Text
                !Cliente = Val(Cliente.Text)
                !Cuenta = Cuenta.Text
                !Tipo = Tipo.ListIndex
                .Update
                .Bookmark = .LastModified
            End If
        End With
        Call CmdLimpiar_Click
        Codigo.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
        With rstProyecto
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
    Cliente.Text = ""
    Cuenta.Text = ""
    DesCuenta.Caption = ""
    DesCliente.Caption = ""
    Tipo.ListIndex = 0
    Codigo.SetFocus
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstProyecto
        .Close
    End With
    With rstClientes
        .Close
    End With
    With rstCuenta
        .Close
    End With
    DbsAdminis.Close
    Codigo.SetFocus
    PrgProyecto.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Anterior_Click()
    With rstProyecto
        .Index = "Codigo"
        .Seek "=", Codigo.Text
        If .NoMatch = False Then
            .MovePrevious
            If .BOF = True Then
                .MoveFirst
                m$ = "No exsite registro Anterior"
                A% = MsgBox(m$, 0, "Archivo de Proyectos de Ventas")
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
    OPEN_FILE_Proyecto
    OPEN_FILE_Cuenta
    OPEN_FILE_Clientes
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
        Cliente.SetFocus
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstClientes
            .Index = "Cliente"
            .Seek "=", Cliente.Text
            If .NoMatch = False Then
                DesCliente.Caption = !Razon
                Cuenta.SetFocus
                    Else
                Cliente.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstCuenta
            .Index = "Cuenta"
            .Seek "=", Cuenta.Text
            If .NoMatch = False Then
                DesCuenta.Caption = !Descripcion
                Codigo.SetFocus
                    Else
                Cuenta.SetFocus
            End If
        End With
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            With rstProyecto
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
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proyectos"
     Opcion.AddItem "Clientes"
     Opcion.AddItem "Cuentas Contables"

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
            
        Case 1
            With rstClientes
                .Index = "Cliente"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Cliente) + " " + !Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 2
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
            With rstProyecto
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Codigo"
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
            
        Case 1
            With rstClientes
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Cliente"
                .Seek "=", Val(Claveven$)
                If .NoMatch = False Then
                    Cliente.Text = !Cliente
                    DesCliente.Caption = !Razon
                End If
            End With
            Cliente.SetFocus
            
        Case 2
            With rstCuenta
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Cuenta"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Cuenta.Text = !Cuenta
                    DesCuenta.Caption = !Descripcion
                End If
            End With
            Cuenta.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()
    Rem On Error GoTo Error_primer
    With rstProyecto
        .Index = "Codigo"
        .MoveFirst
        Codigo.Text = !Codigo
        Call Imprime_Datos
        Codigo.SetFocus
    End With
    Exit Sub

Error_primer:
     coderr = Err
     Call Errores(coderr, "Proyecto de Ventas", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Ultimo_Click()
    Rem On Error GoTo Error_ultimo
    With rstProyecto
        .Index = "Codigo"
        .MoveLast
        Codigo.Text = !Codigo
        Call Imprime_Datos
        Codigo.SetFocus
    End With
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Proyecto de Ventas", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Siguiente_Click()
    With rstProyecto
        .Index = "Codigo"
        Claveven$ = Codigo.Text
        .Seek "=", Claveven$
        If .NoMatch = False Then
            .MoveNext
            If .EOF = True Then
                m$ = "No exsite registro Posterior"
                A% = MsgBox(m$, 0, "Archivo de Proyecto de Ventas")
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
    Cliente.Text = ""
    Cuenta.Text = ""
    DesCuenta.Caption = ""
    DesCliente.Caption = ""
    
    Tipo.Clear
    
    Tipo.AddItem ""
    Tipo.AddItem "Activa"
    Tipo.AddItem "Inactiva"
    
    Tipo.ListIndex = 0
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            With rstProyecto
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
    
        Case 1
            With rstClientes
                .Index = "Cliente"
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(!Razon) - WEspacios
                        For aa = 1 To da
                            If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                                IngresaItem = Str$(!Cliente) + " " + !Razon
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Cliente
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
            
        Case 2
            With rstCuenta
                .Index = "Cuenta"
                .MoveFirst
                Do
                    If .EOF = False Then
                        da = Len(!Descripcion) - WEspacios
                        For aa = 1 To da
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
        Case Else
    End Select
    
    End If

End Sub

Private Sub Codigo_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proyectos"
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Cliente_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proyectos"
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub Cuenta_DblClick()

    Opcion.Clear
    Opcion.AddItem "Proyectos"
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Cuentas Contables"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub Text1_Change()

End Sub
