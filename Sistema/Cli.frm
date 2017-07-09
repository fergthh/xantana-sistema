VERSION 5.00
Begin VB.Form prgcli 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Datos de Clientes"
   ClientHeight    =   5190
   ClientLeft      =   1035
   ClientTop       =   1290
   ClientWidth     =   10020
   LinkTopic       =   "Form2"
   ScaleHeight     =   5190
   ScaleWidth      =   10020
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar F10"
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
      MouseIcon       =   "Cli.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Cli.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Cerrar Consulta"
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox Provincia 
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
      Left            =   2280
      TabIndex        =   29
      Text            =   " "
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox fax 
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
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   27
      Text            =   " "
      Top             =   3240
      Width           =   2055
   End
   Begin VB.TextBox email 
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
      Left            =   6360
      MaxLength       =   20
      TabIndex        =   26
      Text            =   " "
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Observaciones 
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   16
      Text            =   " "
      Top             =   3240
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Condicion de Iva"
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
      Height          =   1575
      Left            =   5520
      TabIndex        =   14
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton Iva6 
         Caption         =   "Inscripto Esp."
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
         Left            =   2040
         TabIndex        =   23
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Iva5 
         Caption         =   "Monotributo"
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
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Iva4 
         Caption         =   "Exento"
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
         Left            =   2040
         TabIndex        =   20
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Iva3 
         Caption         =   "Cons. Final"
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
         Left            =   2040
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Iva2 
         Caption         =   "No Inscripto"
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Iva1 
         Caption         =   "Inscripto"
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
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Cuit 
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
      Left            =   6360
      MaxLength       =   13
      TabIndex        =   13
      Text            =   " "
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Telefono 
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
      Left            =   2280
      MaxLength       =   20
      TabIndex        =   12
      Text            =   " "
      Top             =   2520
      Width           =   3135
   End
   Begin VB.TextBox Postal 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   11
      Text            =   " "
      Top             =   2160
      Width           =   735
   End
   Begin VB.TextBox Localidad 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   10
      Text            =   " "
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox Direccion 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   5
      Text            =   " "
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Cliente 
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
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Razon 
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
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label9 
      Caption         =   "Provincia"
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
      Left            =   120
      TabIndex        =   28
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label20 
      Caption         =   "Fax"
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
      Left            =   5520
      TabIndex        =   25
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "E-Mail"
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
      Left            =   5520
      TabIndex        =   24
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Contacto"
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
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Cuit"
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
      Left            =   5520
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Telefono"
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
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo Postal"
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
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Poblaci 
      Caption         =   "Localidad"
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
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Razon Social"
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
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Cliente"
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
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "prgcli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub Imprime_Descripcion()
End Sub

Sub Format_datos()
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
    
        Cliente.Text = Trim(rstCliente!Cliente)
        Razon.Text = Trim(rstCliente!Razon)
        Direccion.Text = Trim(rstCliente!Direccion)
        Localidad.Text = Trim(rstCliente!Localidad)
        Postal.Text = Trim(rstCliente!Postal)
        Telefono.Text = Trim(rstCliente!Telefono)
        Observaciones.Text = Trim(rstCliente!Observaciones)
        Cuit.Text = Trim(rstCliente!Cuit)
        EMail.Text = Trim(rstCliente!EMail)
        fax.Text = Trim(rstCliente!fax)
        Iva1.Value = False
        Iva2.Value = False
        Iva3.Value = False
        Iva4.Value = False
        Iva5.Value = False
        Iva6.Value = False
        Provincia.ListIndex = Val(rstCliente!Provincia)
        Select Case Val(rstCliente!Iva)
            Case 1
                Iva1.Value = True
            Case 2
                Iva2.Value = True
            Case 3
                Iva3.Value = True
            Case 4
                Iva4.Value = True
            Case 5
                Iva5.Value = True
            Case 6
                Iva6.Value = True
            Case Else
        End Select
        
        
        rstCliente.Close
        Call Format_datos
        Call Imprime_Descripcion
    End If

End Sub

Private Sub CmdLimpiar_Click()
    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    Rem Vendedor.Text = ""
    Rem DesVendedor.Caption = ""
    EMail.Text = ""
    fax.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    Provincia.ListIndex = 25
    Cliente.SetFocus
End Sub

Private Sub Cerrar_Click()
    prgcli.Hide
    Unload Me
    PrgCtaCte1.Show
End Sub

Private Sub cmdClose_Click()
    prgcli.Hide
    Unload Me
    PrgCtaCte1.Show
End Sub

Sub Form_Load()
    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    Rem Vendedor.Text = ""
    Rem DesVendedor.Caption = ""
    EMail.Text = ""
    fax.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    
    Provincia.Clear
    
    Provincia.AddItem "Capital Federal"
    Provincia.AddItem "Buenos Aires"
    Provincia.AddItem "Catamarca"
    Provincia.AddItem "Cordoba"
    Provincia.AddItem "Corrientes"
    Provincia.AddItem "Chaco"
    Provincia.AddItem "Chubut"
    Provincia.AddItem "Entre Rios"
    Provincia.AddItem "Formosa"
    Provincia.AddItem "Jujuy"
    Provincia.AddItem "La Pampa"
    Provincia.AddItem "La Rioja"
    Provincia.AddItem "Mendoza"
    Provincia.AddItem "Misiones"
    Provincia.AddItem "Neuquen"
    Provincia.AddItem "Rio Negro"
    Provincia.AddItem "Salta"
    Provincia.AddItem "San Juan"
    Provincia.AddItem "San Luis"
    Provincia.AddItem "Santa Cruz"
    Provincia.AddItem "Santa Fe"
    Provincia.AddItem "Santiago del Estero"
    Provincia.AddItem "Tucuman"
    Provincia.AddItem "Tierra del Fuego"
    Provincia.AddItem "Exterior"
    Provincia.AddItem ""
    
    Cliente.Text = PCliente
    Call Imprime_Datos
    
End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 121
            Call Cerrar_Click
        Case Else
    End Select
End Sub















