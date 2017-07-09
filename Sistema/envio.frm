VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEnvio 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Remitos a Proveedores"
   ClientHeight    =   8175
   ClientLeft      =   330
   ClientTop       =   555
   ClientWidth     =   11160
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11160
   Visible         =   0   'False
   Begin MSFlexGridLib.MSFlexGrid Grilla 
      Height          =   3975
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   7011
      _Version        =   65541
      Rows            =   101
      Cols            =   7
   End
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
      Left            =   4560
      TabIndex        =   18
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   4560
      MaxLength       =   8
      TabIndex        =   14
      Text            =   " "
      Top             =   480
      Width           =   1095
   End
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
      TabIndex        =   10
      Top             =   5880
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   9735
      Begin VB.TextBox WDesColor 
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
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   22
         Text            =   " "
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox WColor 
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
         Left            =   6000
         TabIndex        =   21
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox WRecibida 
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox WCantidad 
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
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox WArticulo 
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
         MaxLength       =   13
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox WDescripcion 
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
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   9
         Text            =   " "
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox WLinea 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10560
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
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
      Height          =   1260
      Left            =   2280
      TabIndex        =   6
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   8400
      TabIndex        =   5
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.TextBox Numero 
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
      Left            =   1560
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9360
      TabIndex        =   2
      Top             =   7920
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
      Height          =   1740
      ItemData        =   "envio.frx":0000
      Left            =   120
      List            =   "envio.frx":0007
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox FechaEntrega 
      Height          =   285
      Left            =   1560
      TabIndex        =   16
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Image GrabaII 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "envio.frx":0015
      MousePointer    =   99  'Custom
      Picture         =   "envio.frx":031F
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Impresion 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "envio.frx":0761
      MousePointer    =   99  'Custom
      Picture         =   "envio.frx":0A6B
      ToolTipText     =   "Impresion "
      Top             =   5160
      Width           =   480
   End
   Begin VB.Label DesProveedor 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   480
      Width           =   4095
   End
   Begin VB.Image Borra 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "envio.frx":12AD
      MousePointer    =   99  'Custom
      Picture         =   "envio.frx":15B7
      ToolTipText     =   "Borra Renglon"
      Top             =   2640
      Width           =   480
   End
   Begin VB.Image Graba 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "envio.frx":1DF9
      MousePointer    =   99  'Custom
      Picture         =   "envio.frx":2103
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "envio.frx":2945
      MousePointer    =   99  'Custom
      Picture         =   "envio.frx":2C4F
      ToolTipText     =   "Salida"
      Top             =   6000
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "envio.frx":3491
      MousePointer    =   99  'Custom
      Picture         =   "envio.frx":379B
      ToolTipText     =   "Consulta de Datos"
      Top             =   3480
      Width           =   480
   End
   Begin VB.Image Limpia 
      Height          =   480
      Left            =   10440
      MouseIcon       =   "envio.frx":3FDD
      MousePointer    =   99  'Custom
      Picture         =   "envio.frx":42E7
      ToolTipText     =   "Limpia la pantalla"
      Top             =   4320
      Width           =   480
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha"
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
      Left            =   7560
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Proveedor"
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
      Left            =   3240
      TabIndex        =   15
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo Remito"
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
      Left            =   3240
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Entrega"
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
      TabIndex        =   4
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nro.Remito"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEnvio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private XIndice As Single
Private Clave As String
Private Auxi As String
Dim Ciclo As Integer
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Cantidad As Double
Private WFecha As String
Private WAnterior As Integer
Private WDescri As String
Dim WGrilla(1000, 10) As String
Dim Renglon As Integer

Private Sub Borra_Click()

    Grilla.Col = 1
    Grilla.Text = ""
    
    Grilla.Col = 2
    Grilla.Text = ""
    
    Grilla.Col = 3
    Grilla.Text = ""
    
    Grilla.Col = 4
    Grilla.Text = ""
    
    Grilla.Col = 5
    Grilla.Text = ""
    
    Grilla.Col = 6
    Grilla.Text = ""
    
    WCantidad.Text = ""
    WRecibida.Text = ""
    WArticulo.Text = ""
    WDescripcion.Text = ""
    WColor.Text = ""
    WDesColor.Text = ""
    WLinea.Text = ""
    
    WCantidad.SetFocus
    
    Erase WGrilla
    EntraGrilla = 0
    
    For Ciclo = 1 To 100
    
        Grilla.Row = Ciclo
    
        Grilla.Col = 1
        WAuxi1 = Grilla.Text
    
        Grilla.Col = 3
        WAuxi2 = Grilla.Text
    
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
        
            EntraGrilla = EntraGrilla + 1
    
            For Ciclo1 = 1 To 6
                Grilla.Col = Ciclo1
                WGrilla(EntraGrilla, Ciclo1) = Grilla.Text
            Next Ciclo1
            
        End If
        
    Next Ciclo
    
    Grilla.Clear
    
    Grilla.Row = 0
    
    Grilla.Col = 1
    Grilla.Text = "Cantidad"
    
    Grilla.Col = 2
    Grilla.Text = "Recibida"
    
    Grilla.Col = 3
    Grilla.Text = "Articulo"
    
    Grilla.Col = 4
    Grilla.Text = "Descripcion"
    
    Grilla.Col = 5
    Grilla.Text = "Color"
    
    Grilla.Col = 6
    Grilla.Text = "Descripcion"
    
    Renglon = EntraGrilla
    
    For Ciclo = 1 To EntraGrilla
    
        Grilla.Row = Ciclo
        
        For da = 1 To 6
            Grilla.Col = da
            Grilla.Text = WGrilla(Ciclo, da)
        Next da
    
    Next Ciclo
    
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Proveedores"
     Opcion.AddItem "Articulos"
     Opcion.AddItem "Color"

     Opcion.Visible = True
     
End Sub


Private Sub GrabaII_Click()

    With rstEnvio
        .Index = "Numero"
        Claveven$ = "999999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Numero.Text = !Numero + 1
                Else
            Numero.Text = ""
        End If
    End With
    
    If Val(Numero.Text) < 100000 Then
        Numero.Text = "100000"
    End If
    
    Call Graba_Click

End Sub


Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Ayuda.Visible = True
    Ayuda.Text = ""
    
    Select Case XIndice
        Case 0
            With rstProveedor
                .Index = "Nombre"
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 6)
                        IngresaItem = Auxi + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
    
        Case 1
            With rstArticulo
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
        Case 2
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
            
    Ayuda.SetFocus
    Pantalla.Visible = True

End Sub

Private Sub Grilla_GotFocus()
    
    WCol = Grilla.Col
    WRow = Grilla.Row
    
    Grilla.Col = 1
    WAuxi1 = Grilla.Text
    
    Grilla.Col = 3
    WAuxi2 = Grilla.Text
    
    If WAuxi1 = "" And WAuxi2 = "" Then
        WCantidad.Text = ""
        WRecibida.Text = ""
        WArticulo.Text = ""
        WDescripcion.Text = ""
        WColor.Text = ""
        WDesColor.Text = ""
        WLinea.Text = ""
             Else
        WLinea.Text = Grilla.Row
    End If
    
    Grilla.Col = 1
    If Val(Grilla.Text) <> 0 Then
        WCantidad.Text = Grilla.Text
            Else
        WCantidad.Text = ""
    End If
    
    Grilla.Col = 2
    If Val(Grilla.Text) <> 0 Then
        WRecibida.Text = Grilla.Text
            Else
        WRecibida.Text = ""
    End If
    
    Grilla.Col = 3
    WArticulo.Text = Grilla.Text
    
    Grilla.Col = 4
    WDescripcion.Text = Grilla.Text
    
    Grilla.Col = 5
    WColor.Text = Grilla.Text
    
    Grilla.Col = 6
    WDesColor.Text = Grilla.Text

    WCantidad.SetFocus
    
End Sub


Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstEnvio
        .Close
    End With
    With rstArticulo
        .Close
    End With
    With rstColor
        .Close
    End With
    
    DbsAdminis.Close
    PrgEnvio.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    For WRenglon = 1 To 100
    
        With rstEnvio
    
            Auxi = Numero.Text
            Call Ceros(Auxi, 6)
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
        
            .Index = "Clave"
            .Seek "=", Auxi + Auxi1
            If .NoMatch = False Then
                .Delete
            End If
        
        End With
    
    Next WRenglon

    WRenglon = 0
    With rstEnvio
        
        .Index = "Clave"
        For iRow = 1 To 100
        
            Grilla.Row = iRow
            
            Grilla.Col = 1
            Cantidad = Val(Grilla.Text)
            
            Grilla.Col = 2
            Recibida = Val(Grilla.Text)
            
            Grilla.Col = 3
            Articulo = Grilla.Text
            
            Grilla.Col = 4
            XDescri = Grilla.Text
            
            Grilla.Col = 5
            XColor = Val(Grilla.Text)
                    
            If Cantidad <> 0 Or Articulo <> "" Then
                    
                WRenglon = WRenglon + 1
                Auxi = Str$(WRenglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 6)
                    
                .AddNew
                !Numero = Numero.Text
                !Renglon = WRenglon
                !Fecha = Fecha.Text
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !Tipo = Tipo.ListIndex
                !Proveedor = Proveedor.Text
                !FechaEntrega = FechaEntrega.Text
                !OrdFechaEntrega = Right$(FechaEntrega.Text, 4) + Mid$(FechaEntrega.Text, 4, 2) + Left$(FechaEntrega.Text, 2)
                !Cantidad = Cantidad
                !Recibida = Recibida
                !Articulo = Articulo
                !Color = XColor
                !Clave = Auxi1 + Auxi
                !Descripcion = XDescri
                .Update
                        
            End If
            
        Next iRow
            
    End With
    
    Call Impresion_Click
        
    With rstEmpresa
        .Index = "Empresa"
        Rem .Seek "=", Val(WEmpresa)
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

    Call Limpia_Click

    Grilla.Col = 1
    Grilla.Row = 1
        
    Numero.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WCantidad.Text = ""
    WRecibida.Text = ""
    WArticulo.Text = ""
    WDescripcion.Text = ""
    WColor.Text = ""
    WDesColor.Text = ""
    WLinea.Text = ""
    
    WCantidad.SetFocus
    
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    FechaEntrega.Text = "  /  /    "
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Tipo.ListIndex = 0
    
    WCantidad.Text = ""
    WRecibida.Text = ""
    WArticulo.Text = ""
    WDescripcion.Text = ""
    WColor.Text = ""
    WDesColor.Text = ""
    WLinea.Text = ""
    
    Renglon = 0
    
    Grilla.Clear
    
    Grilla.Row = 0
    
    Grilla.Col = 1
    Grilla.Text = "Cantidad"
    
    Grilla.Col = 2
    Grilla.Text = "Recibida"
    
    Grilla.Col = 3
    Grilla.Text = "Articulo"
    
    Grilla.Col = 4
    Grilla.Text = "Descripcion"
    
    Grilla.Col = 5
    Grilla.Text = "Color"
    
    Grilla.Col = 6
    Grilla.Text = "Descripcion"
    
    With rstEnvio
        .Index = "Numero"
        Claveven$ = "99999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Numero.Text = !Numero + 1
                Else
            Numero.Text = ""
        End If
    End With
    
    Graba.Enabled = True
    Borra.Enabled = True
    Rem Ingresa.Enabled = True
    
    Grilla.Col = 1
    Grilla.Row = 1
    
    Numero.SetFocus

End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WArticulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WArticulo.Text = UCase(WArticulo.Text)
        If WArticulo.Text <> "999999" Then
            With rstArticulo
                .Index = "Codigo"
                Claveven$ = WArticulo.Text
                .Seek "=", WArticulo.Text
                If .NoMatch = False Then
                    WDescripcion.Text = !Descripcion
                    WColor.SetFocus
                        Else
                    WArticulo.SetFocus
                End If
            End With
                Else
            WDescripcion.SetFocus
        End If
    End If
End Sub

Private Sub WDescripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WColor.SetFocus
    End If
End Sub

Private Sub WColor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstColor
            .Index = "Codigo"
            Claveven$ = WColor.Text
            .Seek "=", WColor.Text
            If .NoMatch = False Then
                WDesColor.Text = !Descripcion
                Call Alta_Vector
                Call Ingresa_Click
                WCantidad.Text = ""
                WRecibida.Text = ""
                WArticulo.Text = ""
                WDescripcion.Text = ""
                WColor.Text = ""
                WDesColor.Text = ""
                WLinea.Text = ""
                WCantidad.SetFocus
                    Else
                WColor.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    Rem Opcion.Visible = False
    Select Case XIndice
        Case 0
            With rstProveedor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Proveedor"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Proveedor.Text = !Proveedor
                    DesProveedor.Caption = !Nombre
                End If
            End With
            Proveedor.SetFocus
            Rem Ayuda.Visible = False
        Case 1
            With rstArticulo
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Codigo"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WArticulo.Text = !Codigo
                    WDescripcion.Text = !Descripcion
                End If
            End With
            WArticulo.SetFocus
            Rem Ayuda.Visible = False
            
        Case 2
            With rstColor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Codigo"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    WColor.Text = !Codigo
                    WDesColor.Text = !Descripcion
                End If
            End With
            WColor.SetFocus
            Rem Ayuda.Visible = False
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Grilla.Clear
     
    Grilla.ColWidth(0) = 150
    Grilla.ColWidth(1) = 1000
    Grilla.ColWidth(2) = 1000
    Grilla.ColWidth(3) = 1300
    Grilla.ColWidth(4) = 2400
    Grilla.ColWidth(5) = 1000
    Grilla.ColWidth(6) = 2400
    
    Grilla.ColAlignment(3) = 0
    Grilla.ColAlignment(4) = 0
    Grilla.ColAlignment(6) = 0
    Grilla.Font.Bold = True

    Grilla.Row = 0
    
    Grilla.Col = 1
    Grilla.Text = "Cantidad"
    
    Grilla.Col = 2
    Grilla.Text = "Recibida"
    
    Grilla.Col = 3
    Grilla.Text = "Articulo"
    
    Grilla.Col = 4
    Grilla.Text = "Descripcion"
    
    Grilla.Col = 5
    Grilla.Text = "Color"
    
    Grilla.Col = 6
    Grilla.Text = "Descripcion"
    
    Grilla.Col = 1
    Grilla.Row = 1

    Numero.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    FechaEntrega.Text = "  /  /    "
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    
    WCantidad.Text = ""
    WRecibida.Text = ""
    WArticulo.Text = ""
    WDescripcion.Text = ""
    WColor.Text = ""
    WDesColor.Text = ""
    WLinea.Text = ""
    
    Renglon = 0
    
    Tipo.Clear
    
    Tipo.AddItem ""
    Tipo.AddItem "Entrega Normal"
    Tipo.AddItem "Devolucion de Mercaderia"
    
    Tipo.ListIndex = 0
    
    With rstEnvio
        .Index = "Numero"
        Claveven$ = "99999"
        .Seek "<=", Claveven$
        If .NoMatch = False Then
            Numero.Text = !Numero + 1
                Else
            Numero.Text = "1"
        End If
    End With
    
End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Grilla.Row = Renglon
            WAnterior = Renglon
                
            Grilla.Col = 1
            Grilla.Text = WCantidad.Text
            
            Grilla.Col = 2
            Grilla.Text = WRecibida.Text
            
            Grilla.Col = 3
            Grilla.Text = WArticulo.Text
            
            Grilla.Col = 4
            Grilla.Text = WDescripcion.Text
            
            Grilla.Col = 5
            Grilla.Text = WColor.Text
            
            Grilla.Col = 6
            Grilla.Text = WDesColor.Text
            
            Grilla.Row = Renglon
            Grilla.Col = 0
            
            Primer = (Int(Renglon / 15) * 15) + 1
            Grilla.TopRow = Primer
            
                Else
                
            Grilla.Row = Val(WLinea.Text)
            
            WAnterior = Grilla.Row
            
            Grilla.Col = 1
            Grilla.Text = WCantidad.Text
            
            Grilla.Col = 2
            Grilla.Text = WRecibida.Text
            
            Grilla.Col = 3
            Grilla.Text = WArticulo.Text
            
            Grilla.Col = 4
            Grilla.Text = WDescripcion.Text
            
            Grilla.Col = 5
            Grilla.Text = WColor.Text
            
            Grilla.Col = 6
            Grilla.Text = WDesColor.Text
            
            Grilla.Row = Renglon
            Grilla.Col = 0
            
    End If

End Sub

Private Sub Proceso_Click()

    Grilla.Clear
    Grilla.Row = 0
    
    Grilla.Col = 1
    Grilla.Text = "Cantidad"
    
    Grilla.Col = 2
    Grilla.Text = "Recibida"
    
    Grilla.Col = 3
    Grilla.Text = "Articulo"
    
    Grilla.Col = 4
    Grilla.Text = "Descripcion"
    
    Grilla.Col = 5
    Grilla.Text = "Color"
    
    Grilla.Col = 6
    Grilla.Text = "Descripcion"
    
    Grilla.Col = 1
    Grilla.Row = 1
    
    
    For WRenglon = 1 To 100
    
    With rstEnvio
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 6)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        .Index = "Clave"
        .Seek "=", Auxi + Auxi1
        If .NoMatch = False Then
        
            Grilla.Row = WRenglon
            Renglon = WRenglon
                
            Grilla.Col = 1
            Grilla.Text = !Cantidad
            
            Grilla.Col = 2
            Grilla.Text = !Recibida
            
            Grilla.Col = 3
            Grilla.Text = !Articulo
            WWArticulo = !Articulo
            
            Grilla.Col = 4
            Grilla.Text = !Descripcion
            
            Grilla.Col = 5
            Grilla.Text = !Color
            WWColor = !Color
            
            With rstArticulo
                .Index = "Codigo"
                .Seek "=", WWArticulo
                If .NoMatch = False Then
                    Grilla.Col = 4
                    Grilla.Text = !Descripcion
                End If
            End With
            
            With rstColor
                .Index = "Codigo"
                .Seek "=", WWColor
                If .NoMatch = False Then
                    Grilla.Col = 6
                    Grilla.Text = !Descripcion
                End If
            End With
                
        End If
        
    End With
    
    Next WRenglon
    
    With rstProveedor
        .Index = "Proveedor"
        Claveven$ = Proveedor.Text
        .Seek "=", Proveedor.Text
        If .NoMatch = False Then
            DesProveedor.Caption = !Nombre
        End If
    End With
    
    Graba.Enabled = True
    Borra.Enabled = True

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstEnvio
            Auxi = Numero.Text
            Call Ceros(Auxi, 6)
            .Index = "Clave"
            
            Claveven$ = Auxi + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Fecha.Text = !Fecha
                FechaEntrega.Text = !FechaEntrega
                Proveedor.Text = !Proveedor
                Tipo.ListIndex = !Tipo
                Call Proceso_Click
                    Else
                Graba.Enabled = True
                Borra.Enabled = True
                Rem Ingresa.Enabled = True
                WNumero = Numero.Text
                Numero.Text = WNumero
                Fecha.SetFocus
            End If
        End With
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            FechaEntrega.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            A% = MsgBox(m$, 0, "Remitos a Proveedores")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub FechaEntrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaEntrega.Text, Auxi)
        If Auxi = "S" Then
            Proveedor.SetFocus
                Else
            m$ = "Formato de fecha de Entrega invalido"
            A% = MsgBox(m$, 0, "Remitos a Proveedores")
            FechaEntrega.SetFocus
        End If
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstProveedor
            .Index = "Proveedor"
            Claveven$ = Proveedor.Text
            .Seek "=", Proveedor.Text
            If .NoMatch = False Then
                DesProveedor.Caption = !Nombre
                Grilla.Col = 1
                Grilla.Row = 1
                Grilla.SetFocus
                    Else
                Proveedor.SetFocus
            End If
        End With
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    Busqueda = Left$(Ayuda.Text, WEspacios)
    
    Select Case XIndice
        Case 0
            With rstProveedor
                .Index = "Nombre"
                .Seek ">=", Busqueda
                Do
                    If .EOF = False Then
                        If Left$(Ayuda.Text, WEspacios) <> Left(!Nombre, WEspacios) Then
                            Exit Do
                        End If
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 6)
                        IngresaItem = Auxi + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
        Case 1
            With rstArticulo
                .Index = "Descripcion"
                .Seek ">=", Busqueda
                Do
                    If .EOF = False Then
                        If Left$(Ayuda.Text, WEspacios) <> Left(!Descripcion, WEspacios) Then
                            Exit Do
                        End If
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
        Case 2
            With rstColor
                .Index = "Descripcion"
                .Seek ">=", Busqueda
                Do
                    If .EOF = False Then
                        If Left$(Ayuda.Text, WEspacios) <> Left(!Descripcion, WEspacios) Then
                            Exit Do
                        End If
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
            
    End If

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Articulo
    OPEN_FILE_Proveedor
    OPEN_FILE_Color
    OPEN_FILE_Envio
End Sub

Private Sub Impresion_Click()

    With rstProveedor
        .Index = "Proveedor"
        Claveven$ = Proveedor.Text
        .Seek "=", Proveedor.Text
        If .NoMatch = False Then
            WDireccion = !Direccion
            WPostal = !Postal
            WLocalidad = !Localidad
            WCuit = !Cuit
        End If
    End With

    Rem Open "lpt1" For Output As #1
    Open "dada.txt" For Output As #1
    
    If Val(Numero.Text) < 100000 Then

        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, Tab(60); Left$(Fecha.Text$, 2);
        Print #1, Tab(65); Mid$(Fecha.Text, 4, 2);
        Print #1, Tab(70); Right$(Fecha.Text, 4)
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1,
        Print #1, Tab(15); DesProveedor.Caption; "  "; Proveedor.Text
        Print #1,
        Print #1, Tab(15); WDireccion;
        Print #1, Tab(55); WPostal
        Print #1,
        Print #1, Tab(15); WLocalidad
        Print #1,
        Print #1, Tab(55); WCuit
        Print #1,
        Print #1,
        Print #1,
        If Tipo.ListIndex = 2 Then
            Print #1, Tab(30); "DEVOLUCION"
                Else
            Print #1,
        End If
        Print #1,
        Print #1,

        Lin = 0

        For Counter = 1 To 100
                
            Grilla.Row = Counter
                        
            Grilla.Col = 1
            Cantidad = Val(Grilla.Text)
                        
            Grilla.Col = 2
            Recibida = Val(Grilla.Text)
            
            Grilla.Col = 3
            Articulo = Grilla.Text
                        
            Grilla.Col = 4
            WDescri = Grilla.Text
                        
            Grilla.Col = 5
            XColor = Val(Grilla.Text)
                        
            Grilla.Col = 6
            WDescriColor = Grilla.Text
                        
            If Cantidad <> 0 Then

                Print #1, Tab(1); Alinea("###,###.##", Str$(Cantidad));
                Print #1, Tab(15); Articulo;
                Print #1, Tab(30); Left$(WDescri, 30);
                Print #1, Tab(62); WDescriColor

                Lin = Lin + 1

            End If

        Next Counter
        
        For da = Lin To 24
            Print #1, ""
        Next da

            Else

        Print #1, Chr$(27) + Chr$(64);
        Print #1, Chr$(27) + Chr$(67) + Chr$(72);
        Print #1, Chr$(18)

        For WDa = 1 To 2

            If WDa = 1 Then
                Print #1,
                Print #1, "Numero : "; Val(Numero.Text)
                Print #1,
                    Else
                Print #1,
                Print #1, Proveedor.Text
                Print #1, DesProveedor.Caption
            End If
            Print #1, Tab(60); Left$(Fecha.Text, 2);
            Print #1, Tab(65); Mid$(Fecha.Text, 4, 2);
            Print #1, Tab(70); Right$(Fecha.Text, 4)
                        
            If WDa = 1 Then
                Print #1, ""
                Print #1, Tab(15); DesProveedor.Caption; "  "; Proveedor.Text
                Print #1, ""
                Print #1, Tab(15); WDireccion;
                Print #1, Tab(55); WPostal
                Print #1, Tab(15); WLocalidad
                Print #1, Tab(55); WCuit
                    Else
                Print #1,
                Print #1,
            End If
            Print #1,

            Lin = 0

            For Counter = 1 To 100
                        
                Grilla.Row = Counter
                        
                Grilla.Col = 1
                Cantidad = Val(Grilla.Text)
                        
                Grilla.Col = 2
                Recibida = Val(Grilla.Text)
            
                Grilla.Col = 3
                Articulo = Grilla.Text
                        
                Grilla.Col = 4
                WDescri = Grilla.Text
            
                Grilla.Col = 5
                XColor = Val(Grilla.Text)
                        
                Grilla.Col = 6
                WDescriColor = Grilla.Text
                        
                If Cantidad <> 0 Then

                    Print #1, Tab(1); Alinea("###,###.##", Str$(Cantidad));
                    Print #1, Tab(15); Articulo;
                    Print #1, Tab(30); WDescri;
                    Print #1, Tab(62); WDescriColor

                    Lin = Lin + 1
                                
                End If

            Next Counter
            For da = Lin To 22
                Print #1, ""
            Next da
        Next WDa

        Print #1, Chr$(12)
        
    End If
    
    Close #1

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Color"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub WArticulo_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Color"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub WColor_DblClick()

    Opcion.Clear
    
    Opcion.AddItem "Proveedores"
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Color"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 2
    
    Call Opcion_Click

End Sub

Private Sub Grilla_DblClick()

    Call Borra_Click

End Sub

