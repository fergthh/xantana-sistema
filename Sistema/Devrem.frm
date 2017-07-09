VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDevrem 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emision de Devolucion de Remito"
   ClientHeight    =   8175
   ClientLeft      =   390
   ClientTop       =   405
   ClientWidth     =   11160
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11160
   Visible         =   0   'False
   Begin VB.TextBox Vendedor 
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
      MaxLength       =   6
      TabIndex        =   35
      Text            =   " "
      Top             =   720
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Condicion de venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6120
      TabIndex        =   29
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton ctacte 
         Caption         =   "Cuenta Corriente"
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
         Left            =   2160
         TabIndex        =   31
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Contado 
         Caption         =   "Contado"
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
         Left            =   360
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton WImpresion 
      Caption         =   "Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9120
      TabIndex        =   28
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta de &Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6840
      TabIndex        =   23
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Ingresa 
      Caption         =   "Ingresa Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5760
      TabIndex        =   22
      Top             =   7440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   19
      Top             =   5280
      Width           =   10335
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
         Left            =   5400
         MaxLength       =   50
         TabIndex        =   33
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox WTalle 
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
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   32
         Text            =   " "
         Top             =   240
         Width           =   1215
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
         Left            =   6600
         TabIndex        =   27
         Top             =   240
         Width           =   1695
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
         Left            =   240
         MaxLength       =   10
         TabIndex        =   26
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   24
         Text            =   " "
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox WLinea 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   " "
         Top             =   240
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.TextBox WPrecio 
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
         Left            =   8280
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   20
         Text            =   " "
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton Calcula 
      Caption         =   "Calcula Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6840
      TabIndex        =   18
      Top             =   6480
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   15
      Top             =   6480
      Width           =   2775
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   960
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000A&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10440
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "factura.rpt"
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8040
      TabIndex        =   14
      Top             =   7440
      Width           =   975
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
      Height          =   1560
      Left            =   1680
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox Cliente 
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
      MaxLength       =   6
      TabIndex        =   11
      Text            =   " "
      Top             =   360
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4440
      TabIndex        =   9
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
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
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "&Limpia Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5760
      TabIndex        =   6
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton Borra 
      Caption         =   "Borra Renglon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6840
      TabIndex        =   5
      Top             =   7440
      Width           =   975
   End
   Begin VB.CommandButton Graba 
      Caption         =   "&Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5760
      TabIndex        =   4
      Top             =   6960
      Width           =   975
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Height          =   4095
      Left            =   360
      OleObjectBlob   =   "Devrem.frx":0000
      TabIndex        =   3
      Top             =   1080
      Width           =   10695
   End
   Begin VB.ListBox WIndice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10080
      TabIndex        =   2
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1740
      ItemData        =   "Devrem.frx":09DE
      Left            =   120
      List            =   "Devrem.frx":09E5
      TabIndex        =   1
      Top             =   6360
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label DesVendedor 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   3240
      TabIndex        =   36
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Vendedor"
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
      TabIndex        =   34
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   3240
      TabIndex        =   12
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Cliente"
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
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "DEVOL REMITO"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "PrgDevrem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mTotalRows& ' Contiene las filas totales del conjunto de registros
Private UserData() As Variant ' Matriz de 2 dimensiones que contiene registros
Private Const MAXCOLS = 6 ' Número máximo de campos del conjunto de registros.
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WDescuento As Double
Private WCodIva As String
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private WTipo As String
Private WProvincia As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private WPago As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private Mes(0 To 30) As String
Private XIndice As Single
Private WNume As String
Private WLista As Integer
Private WTipopro As Integer
Private XTalle As String
Private XColor As String
Private XArticulo As String

Private Sub Calcula_FechaVto()

    Rem With rstPago
    Rem    .Index = "Pago"
    Rem    .Seek "=", WPago1
    Rem    If .NoMatch = False Then
    Rem        WPlazo1 = !Plazo
    Rem        WTasa = !Tasa
    Rem        WDescuento = !Descuento
    Rem        WPago = !Nombre
    Rem    End If
    Rem End With
    
    Rem WFecha = Fecha.Text
    Rem Call Calcula_vencimiento(WFecha, WPlazo1, Wvencimiento)
    
    Rem With rstPago
    Rem     .Index = "Pago"
    Rem     .Seek "=", WPago2
    Rem     If .NoMatch = False Then
    Rem         WPlazo2 = !Plazo
    Rem     End If
    Rem End With
    
    Rem Call Calcula_vencimiento(WFecha, WPlazo2, WVencimiento1)

End Sub

Private Sub Borra_Click()

    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""
    
    DBGrid1.Col = 2
    DBGrid1.Text = ""
    
    DBGrid1.Col = 3
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""
    
    DBGrid1.Col = 4
    DBGrid1.Text = ""

    WArticulo.Text = ""
    WDescripcion.Text = ""
    WTalle.Text = ""
    WColor.Text = ""
    WCantidad.Text = ""
    WPrecio.Text = ""
    WLinea.Text = ""
    
    WArticulo.SetFocus

End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Clientes"
     Opcion.AddItem "Articulos"
     Opcion.AddItem "Color"
     Opcion.AddItem "Vendedor"

     Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Estadistica
    OPEN_FILE_Articulo
    OPEN_FILE_Numero
    OPEN_FILE_Color
    OPEN_FILE_Stock
    OPEN_FILE_Vendedor
End Sub

Private Sub WImpresion_Click()

    Rem Call Impresion
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
        
    Numero.SetFocus
    
End Sub

Private Sub Opcion_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Rem XIndice = 0
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
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
            
        Case 1
            Ayuda.Visible = True
            Ayuda.Text = ""
            With rstArticulo
                .Index = "Descripcion"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Descripcion
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
            Ayuda.Visible = True
            Ayuda.Text = ""
            With rstColor
                .Index = "Codigo"
                .MoveFirst
                Do
                    If .EOF = False Then
                            IngresaItem = !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
        Case 3
            Rem Ayuda.Visible = True
            Rem Ayuda.Text = ""
            With rstVendedor
                .Index = "Vendedor"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(!Vendedor) + " " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Vendedor
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

End Sub

Private Sub DBGrid1_GotFocus()
    
    WCol = DBGrid1.Col
    WRow = DBGrid1.Row
    
    DBGrid1.Col = WCol
    DBGrid1.Row = WRow
    
    DBGrid1.Col = 0
    WAuxi1 = DBGrid1.Text
    
    Rem DBGrid1.Col = 1
    Rem WAuxi1 = DBGrid1.Text
    
    DBGrid1.Col = 2
    WAuxi2 = DBGrid1.Text
    
    DBGrid1.Col = 3
    WAuxi3 = DBGrid1.Text
    
    DBGrid1.Col = 4
    WAuxi4 = DBGrid1.Text
    
    DBGrid1.Col = 5
    WAuxi5 = DBGrid1.Text
    
    If WAuxi1 = "" Then
        WArticulo.Text = ""
        WDescripcion.Text = ""
        WTalle.Text = ""
        WColor.Text = ""
        WCantidad.Text = ""
        WPrecio.Text = ""
        WLinea.Text = ""
            Else
        WLinea.Text = DBGrid1.Row + 1
        WArticulo.Text = DBGrid1.Text
    End If
    
    DBGrid1.Col = 0
    WArticulo.Text = DBGrid1.Text
    
    DBGrid1.Col = 1
    WDescripcion.Text = DBGrid1.Text
    
    DBGrid1.Col = 2
    WTalle.Text = DBGrid1.Text
    
    DBGrid1.Col = 3
    WColor.Text = DBGrid1.Text
    
    DBGrid1.Col = 4
    If Val(DBGrid1.Text) <> 0 Then
        WCantidad.Text = DBGrid1.Text
            Else
        WCantidad.Text = ""
    End If

    DBGrid1.Col = 5
    If Val(DBGrid1.Text) <> 0 Then
        WPrecio.Text = DBGrid1.Text
            Else
        WPrecio.Text = ""
    End If
    
    WArticulo.SetFocus
    
    If Fecha.Text = "  /  /    " Or Cliente.Text = "" Then
         Numero.SetFocus
    End If

End Sub

Private Sub Calcula_Click()

    WNeto = 0

    For a = 0 To 3
        
        Suma = a * 10
        DBGrid1.FirstRow = Suma
            
        For iRow = 0 To 9
                
            WRow = iRow
            DBGrid1.Row = WRow
                    
            DBGrid1.Col = 4
            WCantidad = Val(DBGrid1.Text)
            
            DBGrid1.Col = 5
            WPrecio = Val(DBGrid1.Text)
                    
            WNeto = WNeto + (WPrecio * WCantidad)
                    
        Next iRow
            
    Next a
    
    Call Calcula_Importe
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    
End Sub

Private Sub Calcula_Importe()

    XNeto = WNeto
    WImpoDto = 0
    
    If WDescuento <> 0 Then
        WImpoDto = WNeto * WDescuento / 100
        Call Redondeo(WImpoDto)
        WNeto = WNeto - WImpoDto
    End If
    
    WIva1 = 0
    WIva2 = 0
    
    Rem If Normal.Value = True Then
    Rem
    Rem Select Case Val(WCodIva)
    Rem     Case 2
    Rem         WIva1 = WNeto * 0.21
    Rem         WIva2 = WNeto * 0.105
    Rem         Call Redondeo(WIva1)
    Rem         Call Redondeo(WIva2)
    Rem     Case Else
    Rem         WIva1 = WNeto * 0.21
    Rem         Call Redondeo(WIva1)
    Rem End Select
            
    Rem End If
    
    If WNeto <> 0 Then
        Call Convierte1_datos(Str$(WNeto), Auxi)
        Rem Neto.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Neto.Caption = "0.00"
    End If
    
    If WImpoDto <> 0 Then
        Call Convierte1_datos(Str$(WImpoDto), Auxi)
        Rem Dto.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Dto.Caption = "0.00"
    End If
    
    If WIva1 <> 0 Then
        Call Convierte1_datos(Str$(WIva1), Auxi)
        Rem Iva1.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Iva1.Caption = "0.00"
    End If
    
    If WIva2 <> 0 Then
        Call Convierte1_datos(Str$(WIva2), Auxi)
        Rem Iva2.Caption = Pusing("###,###.##", Auxi)
            Else
        Rem Iva2.Caption = "0.00"
    End If
    
    WTotal = WNeto + WIva1 + WIva2
    Call Convierte1_datos(Str$(WTotal), Auxi)
    Total.Caption = Pusing("###,###.##", Auxi)

End Sub

Private Sub cmdClose_Click()

    Call Limpia_Click

    With rstAuxiliar
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstClientes
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    With rstArticulo
        .Close
    End With
    With rstEstadistica
        .Close
    End With
    With rstStock
        .Close
    End With
    With rstColor
        .Close
    End With
    With rstVendedor
        .Close
    End With
    
    DbsAdminis.Close
    PrgDevrem.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

        For WRenglon = 1 To 40
           
            With rstEstadistica
    
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
        
                Auxi1 = WRenglon
                Call Ceros(Auxi1, 2)
        
                .Index = "Clave"
                .Seek "=", "19" + Auxi + Auxi1
                If .NoMatch = False Then
                
                    Articulo = !Articulo
                    Cantidad = Abs(!Cantidad)
                    Talle = IIf(IsNull(!Talle), "0", !Talle)
                    XXColor = IIf(IsNull(!Color), "0", !Color)
                
                    With rstArticulo
                        .Index = "Codigo"
                        Claveven$ = Articulo
                        .Seek "=", Articulo
                        If .NoMatch = False Then
                            .Edit
                            !Stock = !Stock - Cantidad
                            .Update
                        End If
                    End With
                    
                    If XXColor <> 0 And Talle <> 0 Then
                        With rstStock
                            .Index = "Clave"
                            XArticulo = Left$(Articulo + Space$(10), 10)
                            XColor = XXColor
                            XTalle = Talle
                            Call Ceros(XColor, 4)
                            Call Ceros(XTalle, 4)
                            WClave = XArticulo + XTalle + XColor
                            .Seek "=", WClave
                            If .NoMatch Then
                                .AddNew
                                !Clave = WClave
                                !Articulo = XArticulo
                                !Talle = Talle
                                !Color = XXColor
                                !Stock = Cantidad * -1
                                .Update
                                .Bookmark = .LastModified
                                    Else
                                .Edit
                                !Stock = !Stock - Cantidad
                                .Update
                                .Bookmark = .LastModified
                            End If
                        End With
                    End If
                
                    .Delete
                    
                End If
                
            End With
    
        Next WRenglon


        Renglon = Renglon + 1
        Lugar1 = Int((Renglon - 1) / 15) * 15
        Lugar2 = Renglon - Lugar1
                
        DBGrid1.FirstRow = Lugar1
        DBGrid1.Row = Lugar2 - 1
    
        DBGrid1.Col = 0
        DBGrid1.Text = ""

        Call Calcula_Click
        Call Calcula_Click
        
        Rem If Val(WCodIva) <> 1 And Val(WCodIva) <> 2 Then
        Rem    WPrecio = WNeto
        Rem    WNeto = WNeto / 1.21
        Rem    Call Redondeo(WNeto)
        Rem    WIva1 = WPrecio - WNeto
        Rem    WIva2 = 0
        Rem End If
      

        With rstCtaCte
        
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            WTipo = "19"
            
            Claveven$ = WTipo + Auxi + "01"
            .Index = "Clave"
            .Seek "=", Claveven$
            If .NoMatch = True Then
        
                .AddNew
            
                !Tipo = "19"
                !Impre = "DR"
                !Numero = Numero.Text
                !Renglon = "01"
                !Cliente = Cliente.Text
                !Fecha = Fecha.Text
                !Estado = "0"
                !Vencimiento = "  /  /    "
                !Vencimiento1 = "  /  /    "
                !Total = WTotal * -1
                !TotalUs = WTotal * -1
                If Contado.Value = True Then
                    !Tipofac = 0
                    !Saldo = 0
                    !SaldoUs = 0
                        Else
                    !Tipofac = 1
                    !Saldo = WTotal * -1
                    !SaldoUs = WTotal * -1
                End If
                !Neto = WNeto * -1
                !Iva1 = WIva1 * -1
                !Iva2 = WIva2 * -1
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdVencimiento = "00000000"
                !OrdVencimiento1 = "00000000"
                !Pedido = ""
                !Remito = ""
                !Orden = ""
                !Paridad = 0
                !Provincia = WProvincia
                !Vendedor = Val(Vendedor.Text)
                !Rubro = 0
                !Comprobante = ""
                !Aceptada = ""
                !Costo = 0
                !Importe1 = 0
                !Importe2 = 0
                !Importe3 = 0
                !Importe4 = 0
                !Importe5 = 0
                !Importe6 = 0
                !Importe7 = 0
                !Empresa = 1
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                !Condventa = ""
                !OCompra = ""
                !Remito = ""
            
                !Clave = !Tipo + Auxi + "01"
                !WDate = Date$
                .Update
                
                    Else
                    
                .Edit
            
                !Tipo = "19"
                !Impre = "DR"
                !Numero = Numero.Text
                !Renglon = "01"
                !Cliente = Cliente.Text
                !Fecha = Fecha.Text
                !Estado = "0"
                !Vencimiento = "  /  /    "
                !Vencimiento1 = "  /  /    "
                !Total = WTotal * -1
                !TotalUs = WTotal * -1
                If Contado.Value = True Then
                    !Saldo = 0
                    !SaldoUs = 0
                        Else
                    !Saldo = WTotal * -1
                    !SaldoUs = WTotal * -1
                End If
                !Neto = WNeto * -1
                !Iva1 = WIva1 * -1
                !Iva2 = WIva2 * -1
                !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                !OrdVencimiento = "00000000"
                !OrdVencimiento1 = "00000000"
                !Pedido = ""
                !Remito = ""
                !Orden = ""
                !Paridad = 0
                !Provincia = WProvincia
                !Vendedor = Val(Vendedor.Text)
                !Rubro = 0
                !Comprobante = ""
                !Aceptada = ""
                !Costo = 0
                !Importe1 = 0
                !Importe2 = 0
                !Importe3 = 0
                !Importe4 = 0
                !Importe5 = 0
                !Importe6 = 0
                !Importe7 = 0
                !Empresa = 1
                Auxi = Numero.Text
                Call Ceros(Auxi, 8)
                !Condventa = ""
                !OCompra = ""
                !Remito = ""
            
                !Clave = !Tipo + Auxi + "01"
                !WDate = Date$
                .Update
                
            End If
            
        End With
                        
        Renglon = 0
        WRenglon = 0
        DBGrid1.Refresh
        
        With rstEstadistica
        
            Renglon = 0
            .Index = "Clave"
                                        
            For a = 0 To 3
        
                Suma = a * 10
                DBGrid1.FirstRow = Suma
            
                For iRow = 0 To 9
                
                    WRenglon = WRenglon + 1
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Articulo = DBGrid1.Text
                    
                    DBGrid1.Col = 2
                    Talle = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 3
                    XXColor = Val(DBGrid1.Text)
    
                    DBGrid1.Col = 4
                    Cantidad = Val(DBGrid1.Text) * -1
                    
                    DBGrid1.Col = 5
                    Precio = Val(DBGrid1.Text)
                    
                    If Cantidad <> 0 Then
                    
                        Renglon = Renglon + 1
                        Auxi = Str$(Renglon)
                        Call Ceros(Auxi, 2)
                        
                        Auxi1 = Str$(Numero.Text)
                        Call Ceros(Auxi1, 8)
                    
                        .AddNew
                        !Tipo = "19"
                        !Numero = Numero.Text
                        !Renglon = Renglon
                        !Articulo = Articulo
                        !Cantidad = Cantidad
                        !Precio = Precio
                        !PrecioUs = Precio
                        !Importe = Precio * Cantidad
                        !ImporteUs = Precio * Cantidad
                        !Cliente = Cliente.Text
                        !Paridad = 0
                        !Vendedor = Val(Vendedor.Text)
                        !Rubro = 0
                        !Linea = 0
                        !Costo1 = 0
                        !Costo2 = 0
                        !Coeficiente = 0
                        !Pedido = 0
                        !Fecha = Fecha.Text
                        !Importe1 = 0
                        !Importe2 = 0
                        !Importe3 = 0
                        !Importe4 = 0
                        !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
                        !WArticulo = ""
                        !Remito = ""
                        !Clave = "19" + Auxi1 + Auxi
                        !WDate = Date$
                        !Clavectacte = Left$(!Clave, 10) + "01"
                        !Imprefactura = "DEVOLUCION"
                        !NroFactura = Auxi1
                        !Talle = Talle
                        !Color = XXColor
                        .Update
                        
                        With rstArticulo
                            .Index = "Codigo"
                            Claveven$ = Articulo
                            .Seek "=", Articulo
                            If .NoMatch = False Then
                                .Edit
                                !Stock = !Stock - Cantidad
                                .Update
                            End If
                        End With
                        
                        If Talle <> 0 And XXColor <> 0 Then
                            With rstStock
                                .Index = "Clave"
                                XArticulo = Left$(Articulo + Space$(10), 10)
                                XColor = XXColor
                                XTalle = Talle
                                Call Ceros(XColor, 4)
                                Call Ceros(XTalle, 4)
                                WClave = XArticulo + XTalle + XColor
                                .Seek "=", WClave
                                If .NoMatch Then
                                    .AddNew
                                    !Clave = WClave
                                    !Articulo = XArticulo
                                    !Talle = Talle
                                    !Color = XXColor
                                    !Stock = Cantidad * -1
                                    .Update
                                    .Bookmark = .LastModified
                                        Else
                                    .Edit
                                    !Stock = !Stock - Cantidad
                                    .Update
                                    .Bookmark = .LastModified
                                End If
                            End With
                        End If
                        
                    End If
                                        
                Next iRow
            
            Next a
            
        End With
        
        With rstNumero
            .Index = "Codigo"
            .Seek "=", "19"
            If .NoMatch = False Then
                .Edit
                If Val(Numero.Text) > !Numero Then
                    !Numero = Val(Numero.Text)
                End If
                .Update
                    Else
                .AddNew
                !Codigo = "19"
                !Numero = Val(Numero.Text)
                .Update
            End If
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
                .Update
            End If
        End With

        Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
        Rem Listado.GroupSelectionFormula = "{Pedido.Pedido} in " + Pedido.Text + " to " + Pedido.Text
        Rem Listado.Destination = 1
        Rem Listado.Action = 1
        
        Rem Call Impresion
        
        T$ = "DEVOLUCION DE REMITOS"
        m$ = "Desea Imprimir el Comprobante"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            WUno = "{Estadistica.Numero} in " + Numero.Text + " to " + Numero.Text + ""
            WDos = " and {Estadistica.Tipo} in 9 to 9"
            Listado.GroupSelectionFormula = WUno + WDos
            Listado.Destination = 1
            Listado.Action = 1
        End If
        
        Call Limpia_Click

        DBGrid1.FirstRow = 0
        DBGrid1.Col = 0
        DBGrid1.Row = 0
        
        Numero.SetFocus
        
End Sub

Private Sub Ingresa_Click()

    WLinea.Text = ""
    WArticulo.Text = ""
    WDescripcion.Text = ""
    WTalle.Text = ""
    WColor.Text = ""
    WCantidad.Text = ""
    WPrecio.Text = ""
    
    WDescripcion.SetFocus
    
End Sub

Private Sub Limpia_Click()

    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Vendedor.Text = "1"
    DesVendedor.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    WLista = 0
    
    Contado.Value = True
    ctacte.Value = False
    
    WLinea.Text = ""
    WArticulo.Text = ""
    WDescripcion.Text = ""
    WTalle.Text = ""
    WColor.Text = ""
    WCantidad.Text = ""
    WPrecio.Text = ""
  
    For a = 0 To 3
        Suma = a * 10
        DBGrid1.FirstRow = Suma
        For iRow = 0 To 9
            For iCol = 0 To 5
                DBGrid1.Col = iCol
                DBGrid1.Row = iRow
                DBGrid1.Text = ""
            Next iCol
        Next iRow
    Next a
    
    DBGrid1.FirstRow = 0
    Renglon = 0
    
    Rem Neto.Caption = ""
    Rem Iva1.Caption = ""
    Rem Iva2.Caption = ""
    Total.Caption = ""
    Rem Dto.Caption = ""
    
    With rstNumero
        .Index = "Codigo"
        Claveven$ = "19"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Numero.Text = !Numero + 1
                Else
            Numero.Text = ""
        End If
    End With
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = "1"
    
    With rstClientes
        .Index = "Cliente"
        Claveven$ = Cliente.Text
        .Seek "=", Cliente.Text
        If .NoMatch = False Then
            Cliente.Text = !Cliente
            DesCliente.Caption = !Razon
            Rem WVendedor = !Vendedor
            WProvincia = !Provincia
            WCodIva = !Iva
            WRazon = !Razon
            WDireccion = !Direccion
            WLocalidad = !Localidad
            WPostal = !Postal
            WCuit = !Cuit
            WLista = !Precio
        End If
    End With
    
    With rstVendedor
        .Index = "Vendedor"
        Claveven$ = Vendedor.Text
        .Seek "=", Vendedor.Text
        If .NoMatch = False Then
            Vendedor.Text = !Vendedor
            DesVendedor.Caption = !Nombre
        End If
    End With

    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    DBGrid1.SetFocus
    
    Rem Factura.Value = True
    Rem Debito.Value = False
    Rem Credito.Value = False
    Rem Normal.Value = True
    Rem Exenta.Value = False
    
    Graba.Enabled = True
    Borra.Enabled = True
    Ingresa.Enabled = True
    
    Numero.SetFocus

End Sub

Private Sub WArticulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstArticulo
            .Index = "Codigo"
            Claveven$ = WArticulo.Text
            .Seek "=", WArticulo.Text
            If .NoMatch = False Then
                WDescripcion.Text = !Descripcion
                If WLista = 0 Then
                    WPrecio.Text = !Precio
                        Else
                    WPrecio.Text = !Precio1
                End If
                        
                WPrecio.Text = Pusing("###,###.##", WPrecio.Text)
                WTipopro = IIf(IsNull(!Tipo), "0", !Tipo)
                If WTipopro = 0 Then
                    WTalle.Text = "0"
                    WColor.Text = "0"
                    WCantidad.SetFocus
                        Else
                    WTalle.SetFocus
                End If
                
                    Else
                WArticulo.SetFocus
            End If
        End With
    End If
End Sub

Private Sub WTalle_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(WTalle.Text) <> 0 Then
            WColor.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WColor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstColor
            .Index = "Codigo"
            Claveven$ = WColor.Text
            .Seek "=", WColor.Text
            If .NoMatch = False Then
                WCantidad.SetFocus
                    Else
                WColor.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WCantidad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WCantidad.Text = Pusing("###,###.##", WCantidad.Text)
        Call Alta_Vector
        Call Ingresa_Click
        Call Calcula_Click
        Call Calcula_Click
        WArticulo.Text = ""
        WDescripcion.Text = ""
        WTalle.Text = ""
        WColor.Text = ""
        WCantidad.Text = ""
        WPrecio.Text = ""
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WPrecio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPrecio.Text = Pusing("###,###.##", WPrecio.Text)
        Call Alta_Vector
        Call Ingresa_Click
        Call Calcula_Click
        Call Calcula_Click
        WArticulo.Text = ""
        WDescripcion.Text = ""
        WTalle.Text = ""
        WColor.Text = ""
        WCantidad.Text = ""
        WPrecio.Text = ""
        WArticulo.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    Rem Opcion.Visible = False
    Select Case XIndice
        Case 0
            With rstClientes
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Cliente"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Cliente.Text = !Cliente
                    DesCliente.Caption = !Razon
                    Rem WVendedor = !Vendedor
                    WProvincia = !Provincia
                    WCodIva = !Iva
                    WRazon = !Razon
                    WDireccion = !Direccion
                    WLocalidad = !Localidad
                    WPostal = !Postal
                    WCuit = !Cuit
                    WLista = !Precio
                End If
            End With
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
                    WTipopro = IIf(IsNull(!Tipo), "0", !Tipo)
                    If WLista = 0 Then
                        WPrecio.Text = !Precio
                            Else
                        WPrecio.Text = !Precio1
                    End If
                    
                    Rem DBGrid1.Col = 0
                    Rem DBGrid1.Text = !Codigo
                    Rem DBGrid1.Col = 1
                    Rem DBGrid1.Text = !Descripcion1
                    Rem DBGrid1.Col = 3
                    Rem DBGrid1.Text = Pusing("###,###.##", !Precio)
                    
                    Call Alta_Vector
                    WLinea.Text = WAnterior + 1
                    If Val(WLinea.Text) > 0 Then
                        DBGrid1.Row = Val(WLinea.Text) - 1
                    End If
                    
                    Call DBGrid1.SetFocus
                    If WTipopro = 0 Then
                        WTalle.Text = "0"
                        WColor.Text = "0"
                        WCantidad.SetFocus
                            Else
                        WTalle.SetFocus
                    End If
                
                End If
            End With
            Rem Ayuda.Visible = False
            
        Case 3
            With rstVendedor
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Vendedor"
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Vendedor.Text = !Vendedor
                    DesVendedor.Caption = !Nombre
                End If
            End With
            Rem Ayuda.Visible = False
            Rem Pantalla.Visible = False
            Vendedor.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub DbGrid1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case DBGrid1.Col
            Case 0, 1, 2, 3, 4, 5, 6
                Select Case KeyCode
                    Case 13
                        If DBGrid1.Row < 40 Then
                            DBGrid1.Row = DBGrid1.Row + 1
                            DBGrid1.Col = 0
                            KeyCode = 0
                        End If
                        Call Calcula_Click
                        Call Calcula_Click
                        DBGrid1.Row = WRow

                    Case Else
                        Rem If KeyCode <> 0 Then Stop
                
            End Select
            
    End Select

    
End Sub


' Cuando el usuario hace clic en el icono Agregar, esta subrutina agrega una
' nueva fila a la variable RowBuf y un marcador a la variable NewRowBookmark
Private Sub DBGrid1_UnboundAddData(ByVal RowBuf As RowBuffer, NewRowBookmark As Variant)
Dim iCol As Integer

mTotalRows = mTotalRows + 1
ReDim Preserve UserData(MAXCOLS - 1, mTotalRows - 1)
NewRowBookmark = mTotalRows - 1 'Establece el marcador a la última fila.

' El bucle siguiente agrega un nuevo registro a la base de datos.
For iCol = 0 To UBound(UserData, 1)
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, mTotalRows - 1) = RowBuf.Value(0, iCol)
    Else
        ' Si no se establece ningún valor para la columna, usa DefaultValue
        UserData(iCol, mTotalRows - 1) = DBGrid1.Columns(iCol).DefaultValue
    End If
Next iCol

End Sub

' Esta subrutina elimina una fila basándose en su marcador.
Private Sub DBGrid1_UnboundDeleteRow(Bookmark As Variant)
Dim iCol As Integer, iRow As Integer

' Mueve todas las filas encima de la fila eliminada de
' la matriz.

For iRow = Bookmark + 1 To mTotalRows - 1
    For iCol = 0 To MAXCOLS - 1
        UserData(iCol, iRow - 1) = UserData(iCol, iRow)
    Next iCol
Next iRow
mTotalRows = mTotalRows - 1

End Sub

' Se llama a esta subrutina cada vez que DBGrid quiere mostrar
' datos nuevos.
Private Sub DBGrid1_UnboundReadData(ByVal RowBuf As RowBuffer, StartLocation As Variant, ByVal ReadPriorRows As Boolean)
Dim CurRow&, iRow As Integer, iCol As Integer, iRowsFetched As Integer, iIncr As Integer
' DBGrid está solicitando filas, así que se las damos

If ReadPriorRows Then
    iIncr = -1
Else
    iIncr = 1
End If

' Si StartLocation es Null, empieza a leer por el final
' o por el principio del conjunto de datos.
If IsNull(StartLocation) Then
    If ReadPriorRows Then
        CurRow& = RowBuf.RowCount - 1
    Else
        CurRow& = 0
    End If
Else
    ' Busca la posición para empezar a leer, basándose en el marcador
    ' StartLocation y en la variable iIncr
    CurRow& = CLng(StartLocation) + iIncr
End If

' Transfiere datos de nuestra matriz de conjunto de datos al objeto RowBuf
' que DBGrid utiliza para presentar los datos
For iRow = 0 To RowBuf.RowCount - 1
    If CurRow& < 0 Or CurRow& >= mTotalRows& Then Exit For
    For iCol = 0 To UBound(UserData, 1)
        RowBuf.Value(iRow, iCol) = UserData(iCol, CurRow&)
    Next iCol
    ' Establece el marcador mediante CurRow&, que es también
    ' nuestro índice de matriz
    RowBuf.Bookmark(iRow) = CStr(CurRow&)
    CurRow& = CurRow& + iIncr
    iRowsFetched = iRowsFetched + 1
Next iRow
RowBuf.RowCount = iRowsFetched
End Sub

' Esta subrutina actualiza los datos de la matriz después de
' haberse modificado.

Private Sub DBGrid1_UnboundWriteData(ByVal RowBuf As RowBuffer, WriteLocation As Variant)
Dim iCol As Integer
' Se están actualizando los datos

' Actualiza cada columna de la matriz de conjuntos de datos
For iCol = 0 To MAXCOLS - 1
    If Not IsNull(RowBuf.Value(0, iCol)) Then
        UserData(iCol, WriteLocation) = RowBuf.Value(0, iCol)
    End If
Next iCol

End Sub

Private Sub Form_Load()

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
    
    Mes(1) = "Enero"
    Mes(2) = "Febrero"
    Mes(3) = "Marzo"
    Mes(4) = "Abril"
    Mes(5) = "Mayo"
    Mes(6) = "Junio"
    Mes(7) = "Julio"
    Mes(8) = "Agosto"
    Mes(9) = "Septiembre"
    Mes(10) = "Octubre"
    Mes(11) = "Noviembre"
    Mes(12) = "Diciembre"
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"

' 3 columnas, 15 filas de datos
ReDim UserData(0 To 5, 0 To 40)

mTotalRows& = 40

Dim oldcnt As Integer, newcnt As Integer

Me.Show
oldcnt = DBGrid1.Columns.Count
newcnt = 0
Dim i As Integer

' Quita las columnas antiguas
For i = DBGrid1.Columns.Count - 1 To 0 Step -1
      DBGrid1.Columns.Remove i
Next i

' Agrega nuevas columnas
For i = 0 To 5
    DBGrid1.Columns.Add newcnt
     Select Case i
         Case 0
             DBGrid1.Columns(newcnt).Caption = "Articulo"
             DBGrid1.Columns(newcnt).Width = 1300
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 1
             DBGrid1.Columns(newcnt).Caption = "Descripcion"
             DBGrid1.Columns(newcnt).Width = 2600
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Locked = True
         Case 2
             DBGrid1.Columns(newcnt).Caption = "Talle"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 3
             DBGrid1.Columns(newcnt).Caption = "Color"
             DBGrid1.Columns(newcnt).Width = 1200
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 4
             DBGrid1.Columns(newcnt).Caption = "Cantidad"
             DBGrid1.Columns(newcnt).Width = 1700
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
         Case 5
             DBGrid1.Columns(newcnt).Caption = "Precio"
             DBGrid1.Columns(newcnt).Width = 2000
             DBGrid1.Columns(newcnt).AllowSizing = False
             DBGrid1.Columns(newcnt).Alignment = 1
             DBGrid1.Columns(newcnt).Locked = True
             
         Case Else

     End Select
     DBGrid1.Columns(newcnt).Visible = True
     newcnt = newcnt + 1
         
Next i
 
    Rem DBGrid1.FirstRow = 0
    Rem DBGrid1.Col = 0
    Rem DBGrid1.Row = 0
    
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Vendedor.Text = "1"
    DesVendedor.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    WLinea.Text = ""
    WDescripcion.Text = ""
    WTalle.Text = ""
    WColor.Text = ""
    WPrecio.Text = ""
    
    Rem Factura.Value = True
    Rem Debito.Value = False
    Rem Credito.Value = False
    
    Rem Normal.Value = True
    Rem Exenta.Value = False
    
    Contado.Value = True
    ctacte.Value = False
     
     With rstNumero
        .Index = "Codigo"
        Claveven$ = "19"
        .Seek "=", Claveven$
        If .NoMatch = False Then
            Numero.Text = !Numero + 1
                Else
            Numero.Text = "1"
        End If
    End With
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = "1"
    WLista = 0
    
    With rstClientes
        .Index = "Cliente"
        Claveven$ = Cliente.Text
        .Seek "=", Cliente.Text
        If .NoMatch = False Then
            Cliente.Text = !Cliente
            DesCliente.Caption = !Razon
            Rem WVendedor = !Vendedor
            WProvincia = !Provincia
            WCodIva = !Iva
            WRazon = !Razon
            WDireccion = !Direccion
            WLocalidad = !Localidad
            WPostal = !Postal
            WCuit = !Cuit
            WLista = !Precio
        End If
    End With
    
    With rstVendedor
        .Index = "Vendedor"
        Claveven$ = Vendedor.Text
        .Seek "=", Vendedor.Text
        If .NoMatch = False Then
            Vendedor.Text = !Vendedor
            DesVendedor.Caption = !Nombre
        End If
    End With
    
    DBGrid1.FirstRow = 0
    DBGrid1.Col = 0
    DBGrid1.Row = 0
    DBGrid1.SetFocus

    Numero.SetFocus
    
End Sub

Private Sub Alta_Vector()

    If Val(WLinea.Text) = 0 Then

            Renglon = Renglon + 1
            
            Lugar1 = Int((Renglon - 1) / 15) * 15
            Lugar2 = Renglon - Lugar1
                
            DBGrid1.FirstRow = Lugar1
            DBGrid1.Row = Lugar2 - 1
            
            WAnterior = DBGrid1.Row
                
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WTalle.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WColor.Text
            
            If Val(WCantidad.Text) <> 0 Then
                DBGrid1.Col = 4
                DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                    Else
                DBGrid1.Col = 4
                DBGrid1.Text = ""
            End If
            
            If Val(WPrecio.Text) <> 0 Then
                DBGrid1.Col = 5
                DBGrid1.Text = Pusing("###,###.##", WPrecio.Text)
                    Else
                DBGrid1.Col = 5
                DBGrid1.Text = ""
            End If
            
            Rem DbGrid1.Row = Renglon
            Rem DbGrid1.Col = 0
            
            If Renglon < 15 Then
                Lugar1 = Int((Renglon - 1) / 15) * 15
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                DBGrid1.Col = 0
                    Else
                DBGrid1.FirstRow = 0
                DBGrid1.Row = 0
                DBGrid1.Col = 0
                DBGrid1.SetFocus
            End If
            
            
                Else
                
            DBGrid1.Row = Val(WLinea.Text) - 1
            
            WAnterior = DBGrid1.Row
            
            DBGrid1.Col = 0
            DBGrid1.Text = WArticulo.Text
            
            DBGrid1.Col = 1
            DBGrid1.Text = WDescripcion.Text
            
            DBGrid1.Col = 2
            DBGrid1.Text = WTalle.Text
            
            DBGrid1.Col = 3
            DBGrid1.Text = WColor.Text
            
            If Val(WCantidad.Text) <> 0 Then
                DBGrid1.Col = 4
                DBGrid1.Text = Pusing("###,###.##", WCantidad.Text)
                    Else
                DBGrid1.Col = 4
                DBGrid1.Text = ""
            End If
            
            If Val(WPrecio.Text) <> 0 Then
                DBGrid1.Col = 5
                DBGrid1.Text = Pusing("###,###.##", WPrecio.Text)
                    Else
                DBGrid1.Col = 5
                DBGrid1.Text = ""
            End If
            
            Rem DbGrid1.Row = Renglon
            Rem DbGrid1.Col = 0
            
            If Renglon < 15 Then
                Lugar1 = Int((Renglon - 1) / 15) * 15
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                DBGrid1.Col = 0
                    Else
                DBGrid1.FirstRow = 0
                DBGrid1.Row = 0
                DBGrid1.Col = 0
                DBGrid1.SetFocus
            End If
            
            
    End If

End Sub

Private Sub Proceso_Click()

    For a = 0 To 3
    Suma = a * 10
    DBGrid1.FirstRow = Suma
    For iRow = 0 To 9
        For iCol = 0 To 5
            DBGrid1.Col = iCol
            DBGrid1.Row = iRow
            DBGrid1.Text = ""
        Next iCol
    Next iRow
    Next a
    
    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 40
    
    With rstEstadistica
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
        
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        .Index = "Clave"
        .Seek "=", "19" + Auxi + Auxi1
        If .NoMatch = False Then
        
            Canti = !Cantidad
            
            If Abs(Canti) > 0 Then
        
                Renglon = Renglon + 1
            
                Lugar1 = Int((Renglon - 1) / 15) * 15
                Lugar2 = Renglon - Lugar1
                
                DBGrid1.FirstRow = Lugar1
                DBGrid1.Row = Lugar2 - 1
                
                DBGrid1.Col = 0
                DBGrid1.Text = !Articulo
                Auxi1 = !Articulo
                
                DBGrid1.Col = 2
                DBGrid1.Text = IIf(IsNull(!Talle), "0", !Talle)
                
                DBGrid1.Col = 3
                DBGrid1.Text = IIf(IsNull(!Color), "0", !Color)
                
                DBGrid1.Col = 4
                DBGrid1.Text = Pusing("###,###.##", Str$(Abs(!Cantidad)))
                
                DBGrid1.Col = 5
                DBGrid1.Text = Pusing("###,###.##", Str$(!Precio))
            
                With rstArticulo
                    .Index = "Codigo"
                    .Seek "=", Auxi1
                    If .NoMatch = False Then
                        DBGrid1.Col = 1
                        DBGrid1.Text = !Descripcion
                    End If
                End With
                
                If Val(Canti) <> 0 Then
                    WNeto = WNeto + (Val(Cantidad) * Precio)
                End If
                
            End If
                
        End If
        
    End With
    
    Next WRenglon

    DBGrid1.FirstRow = 0
    
    Renglon = Renglon + 1
    Lugar1 = Int((Renglon - 1) / 15) * 15
    Lugar2 = Renglon - Lugar1
                
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    DBGrid1.Col = 0
    DBGrid1.Text = ""
    
    DBGrid1.Col = 1
    DBGrid1.Text = ""
    
    DBGrid1.Col = 2
    DBGrid1.Text = "0"
    
    DBGrid1.Col = 3
    DBGrid1.Text = "0"
    
    DBGrid1.Col = 4
    DBGrid1.Text = "0"
    
    DBGrid1.Col = 5
    DBGrid1.Text = "0"
    
    Call Calcula_Click
    Call Calcula_Click
    
    Renglon = Renglon - 1
    Lugar1 = Int((Renglon - 1) / 15) * 15
    Lugar2 = Renglon - Lugar1
    DBGrid1.FirstRow = Lugar1
    DBGrid1.Row = Lugar2 - 1
    
    Graba.Enabled = True
    Borra.Enabled = True

End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstCtaCte
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            .Index = "Clave"
            
            WTipo = "19"
            
            Claveven$ = WTipo + Auxi + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
                Fecha.Text = !Fecha
                Cliente.Text = !Cliente
                Vendedor.Text = !Vendedor
                
                With rstClientes
                    .Index = "Cliente"
                    Claveven$ = Cliente.Text
                    .Seek "=", Cliente.Text
                    If .NoMatch = False Then
                        Cliente.Text = !Cliente
                        DesCliente.Caption = !Razon
                        Rem WVendedor = !Vendedor
                        WProvincia = !Provincia
                        WCodIva = !Iva
                        WRazon = !Razon
                        WDireccion = !Direccion
                        WLocalidad = !Localidad
                        WPostal = !Postal
                        WCuit = !Cuit
                        WLista = !Precio
                    End If
                End With
                With rstVendedor
                    .Index = "Vendedor"
                    Claveven$ = Vendedor.Text
                    .Seek "=", Vendedor.Text
                    If .NoMatch = False Then
                        Vendedor.Text = !Vendedor
                        DesVendedor.Caption = !Nombre
                    End If
                End With
                Call Proceso_Click
                    Else
                Rem .Index = "Numero"
                Rem .Seek "=", Val(Numero.Text)
                Rem If .NoMatch = False Then
                Rem     m$ = "Comprobante ya existente"
                Rem   A% = MsgBox(m$, 0, "Ingreso de comprobantes varias")
                Rem     Numero.SetFocus
                Rem        Else
                Rem    Graba.Enabled = True
                Rem    Borra.Enabled = True
                Rem    Ingresa.Enabled = True
                Rem    WNumero = Numero.Text
                Rem    Numero.Text = WNumero
                Rem    Fecha.SetFocus
                Rem End If
                Graba.Enabled = True
                Borra.Enabled = True
                Ingresa.Enabled = True
                WNumero = Numero.Text
                Numero.Text = WNumero
                Fecha.SetFocus
                
            End If
        End With
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstClientes
            .Index = "Cliente"
            Claveven$ = Cliente.Text
            .Seek "=", Cliente.Text
            If .NoMatch = False Then
                Cliente.Text = !Cliente
                DesCliente.Caption = !Razon
                Rem WVendedor = !Vendedor
                WProvincia = !Provincia
                WCodIva = !Iva
                WRazon = !Razon
                WDireccion = !Direccion
                WLocalidad = !Localidad
                WPostal = !Postal
                WCuit = !Cuit
                WLista = !Precio
                DBGrid1.FirstRow = 0
                DBGrid1.Col = 0
                DBGrid1.Row = 0
                DBGrid1.SetFocus
                    Else
                Cliente.SetFocus
            End If
        End With
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Rem With rstCambios
            Rem     .Index = "Fecha"
            Rem     .Seek "=", Fecha.Text
            Rem     If .NoMatch = False Then
            Rem         Paridad.Text = Pusing("###,###.##", Str$(!Cambio))
            Rem             Else
            Rem         Paridad.Text = "1.00"
            Rem     End If
            Rem End With
            Cliente.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            DBGrid1.FirstRow = 0
            DBGrid1.Col = 0
            DBGrid1.Row = 0
            DBGrid1.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstVendedor
            .Index = "Vendedor"
            Claveven$ = Vendedor.Text
            .Seek "=", Vendedor.Text
            If .NoMatch = False Then
                Vendedor.Text = !Vendedor
                DesVendedor.Caption = !Nombre
                DBGrid1.FirstRow = 0
                DBGrid1.Col = 0
                DBGrid1.Row = 0
                DBGrid1.SetFocus
                    Else
                Vendedor.SetFocus
            End If
        End With
    End If
End Sub

Sub Impresion()

    Rem Open "lpt1" For Output As #1
    
    Printer.EndDoc

    Printer.Font = "Times New Roman"
    Printer.FontSize = "12"
    Printer.Print ""
    Printer.FontSize = "12"
    
    Rem Width #1, 255

    For XX% = 1 To 1
    
        WNume = Numero.Text
        Call Ceros(WNume, 8)
    
    
        Rem Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
        Printer.Print ""
        Printer.Print "|--------------------------------------------------------------------------------|"
        Printer.Print "|                                                                                |"
        Printer.Print "| YENADI S.A.                                  | DEVOLUCION DE MERCADERIA        |"
        Printer.Print "|                                              |Factura N§         :0000-" + WNume;
        Printer.Print Tab(82); "|"
        Printer.Print "|                                              |Fecha              :" + Fecha.Text;
        Printer.Print Tab(82); "|"
        Printer.Print "|                                              |Cuit Nro           :11-11111111-1|"
        Printer.Print "|--------------------------------------------------------------------------------|"
        Printer.Print "|                                                                                |"
        Printer.Print "|Senor    : "; Left$(WRazon, 30);
        Printer.Print Tab(82); "|"
        Printer.Print "|Domicilio:"; ; Left$(WDomicilio, 30);
        Printer.Print Tab(82); "|"
        Printer.Print "|Localidad:"; ; Left$(WLocalidad, 30);
        Printer.Print Tab(82); "|"
        Printer.Print "|C.U.I.T. :";
        Printer.Print Tab(82); "|"
        Printer.Print "|                                                                                |"
        Printer.Print "|--------------------------------------------------------------------------------|"
        Printer.Print "|   Cantidad   |           Descripcion             |   Precio     |    Total     |"
        Printer.Print "|--------------------------------------------------------------------------------|"
        Printer.Print "|              |                                   |              |              |"
        
        Rem Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "16" + Chr$(72)
        
        Impre = 0

        For a = 0 To 1
        
        
            Suma = a * 10
            DBGrid1.FirstRow = Suma
            
            For iRow = 0 To 9
            
                If Impre < 11 Then
                
                    WRow = iRow
                    DBGrid1.Row = WRow
                    
                    DBGrid1.Col = 0
                    Arti = DBGrid1.Text
                    
                    DBGrid1.Col = 1
                    Descri = DBGrid1.Text
                
                    DBGrid1.Col = 2
                    Cantidad = Val(DBGrid1.Text)
                    
                    DBGrid1.Col = 3
                    Precio = Val(DBGrid1.Text)
                    
                    WUnidad = ""
                    WDescri1 = ""
                    
                    With rstArticulo
                        .Index = "Codigo"
                        Claveven$ = Arti
                        .Seek "=", Arti
                        If .NoMatch = False Then
                            Descri = !Descripcion
                        End If
                    End With
                    
                    parcial = Cantidad * Precio
                    
                    If Cantidad <> 0 Then
                    
                        Printer.Print "|"; Alinea("###,###.##", Str$(Cantidad));
                        Printer.Print Tab(16); "|"; Descri;
                        Printer.Print Tab(52); "|"; Alinea("###,###.##", Str$(Precio));
                        Printer.Print Tab(67); "|"; Alinea("###,###.##", Str$(parcial));
                        Printer.Print Tab(82); "|"
                    
                        Impre = Impre + 1
                    
                    End If
                    
                End If
                    
            Next iRow
            
        Next a


        For Imprelinea% = Impre To 20
                Printer.Print "|              |                                   |              |              |"
        Next Imprelinea%

        Printer.Print "|--------------------------------------------------------------------------------|"

        Printer.Print ""
        Printer.Print ""
        Printer.Print ""

        Printer.Print "|--------------------------------------------------------------------------------|"
        Printer.Print "|                                                     Sub-Total : "; Alinea("###,###.##", Total.Caption);
        Printer.Print Tab(82); "|"
        Printer.Print "|--------------------------------------------------------------------------------|"


        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""
        Printer.Print ""

        

        Printer.Print ""
        Printer.Print ""
        
        Rem Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        Rem Print #1, Chr$(12)
        
        Rem Printer.NewPage
        
    Next XX%
    
    Printer.EndDoc

    Rem Close #1

End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    If XIndice = 0 Then
    
        With rstClientes
            .Index = "Razon"
            .MoveFirst
            Do
                If .EOF = False Then
            
                    da = Len(!Razon) - WEspacios
                
                    For aa = 1 To da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                            Auxi = !Cliente
                            IngresaItem = Auxi + "    " + !Razon
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
        
            Else
            
        With rstArticulo
            .Index = "Descripcion"
            .MoveFirst
            Do
                If .EOF = False Then
                
                    da = Len(!Codigo) - WEspacios
                
                    For aa = 1 To da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Codigo, aa, WEspacios) Then
                            IngresaItem = !Descripcion
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
    End If

End Sub


