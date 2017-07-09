VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFactuExpo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Facturas de Exportacion"
   ClientHeight    =   8385
   ClientLeft      =   630
   ClientTop       =   405
   ClientWidth     =   10800
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8385
   ScaleWidth      =   10800
   Visible         =   0   'False
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
      Left            =   9840
      MouseIcon       =   "FactuExpo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "FactuExpo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Menu Principal"
      Top             =   2640
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
      Left            =   9840
      MouseIcon       =   "FactuExpo.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "FactuExpo.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Consulta de Datos"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Limpia 
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
      Left            =   8880
      MouseIcon       =   "FactuExpo.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "FactuExpo.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Borra 
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
      Left            =   8880
      MouseIcon       =   "FactuExpo.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "FactuExpo.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Elimina el Registro"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Graba 
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
      Left            =   8880
      MouseIcon       =   "FactuExpo.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "FactuExpo.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox Paridad 
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
      Left            =   6240
      MaxLength       =   6
      TabIndex        =   29
      Text            =   " "
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   3360
      TabIndex        =   24
      Top             =   2400
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   3360
      TabIndex        =   22
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox Letra 
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
      MaxLength       =   1
      TabIndex        =   21
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Punto 
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
      Left            =   2880
      MaxLength       =   4
      TabIndex        =   1
      Top             =   480
      Width           =   735
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
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Comprobante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8880
      TabIndex        =   16
      Top             =   120
      Width           =   1815
      Begin VB.OptionButton Credito 
         Caption         =   "Nota Credito"
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
         Left            =   240
         TabIndex        =   19
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Debito 
         Caption         =   "Nota Debito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   480
         Width           =   1455
      End
      Begin VB.OptionButton Factura 
         Caption         =   "Factura "
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
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   5280
      TabIndex        =   13
      Top             =   5640
      Width           =   2895
      Begin VB.Label Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Total"
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
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9840
      Top             =   7560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "FactuFor.rpt"
      CopiesToPrinter =   2
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
      Left            =   2160
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSMask.MaskEdBox Vencimiento 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   2400
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   6240
      TabIndex        =   7
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   3720
      MaxLength       =   8
      TabIndex        =   2
      Text            =   " "
      Top             =   480
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   9600
      TabIndex        =   4
      Top             =   7440
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
      Height          =   1980
      ItemData        =   "FactuExpo.frx":387C
      Left            =   120
      List            =   "FactuExpo.frx":3883
      TabIndex        =   3
      Top             =   6000
      Visible         =   0   'False
      Width           =   5055
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3360
      TabIndex        =   27
      Top             =   2880
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4335
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   7646
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Label11 
      Caption         =   "Paridad"
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
      Height          =   285
      Left            =   5280
      TabIndex        =   30
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Vencimiento"
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
      TabIndex        =   10
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00C0C000&
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
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   4575
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
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   120
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5280
      TabIndex        =   6
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Numero de Comprobante"
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
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "PrgFactuExpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private WVencimiento As String
Private WNeto As Double
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
Private WDias As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private Mes(0 To 20) As String
Private WNumero As String
Private XTexto1 As String
Private XTexto2 As String
Private WPlazo1 As Integer
Dim CantiFac As Integer
Dim CantiRem As Integer
Dim CantiArti As Integer

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

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

    With rstCtaCte
        
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
            
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
            
        .Index = "Clave"
            
        If Factura.Value = True Then
            WTipo = "03"
        End If
        If Debito.Value = True Then
            WTipo = "04"
        End If
        If Credito.Value = True Then
            WTipo = "05"
        End If
            
        .Seek "=", Letra.Text + WTipo + WPunto + Auxi + "01"
        If .NoMatch = False Then
        
            If !Saldo <> !Total Then
            
                m$ = "El comprobante se encuentra total o parcialmente cancelado"
                a% = MsgBox(m$, 0, "Eliminacion de Comprobantes")
                
                    Else
                    
                .Delete
                
                Renglon = 0
                                
                With rstDesccomp
                    For a = 1 To 30
                        Renglon = Renglon + 1
                        Auxi1 = Str$(Renglon)
                        Call Ceros(Auxi1, 2)
                        .Index = "Clave"
                        .Seek "=", Letra.Text + WTipo + WPunto + Auxi + Auxi1
                        If .NoMatch = False Then
                            .Delete
                        End If
                    Next a
                End With
                
            End If
            
        End If
            
    End With
    
    Call Limpia_Click

End Sub

Private Sub Consulta_Click()

     Opcion.Clear
     Opcion.AddItem "Clientes"
     Opcion.Visible = True
     
 End Sub

Private Sub Impresion_Click()

End Sub

 Private Sub Opcion_Click()

    On Error GoTo WError
    
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
                .Index = "Razon"
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Calcula_Click()

    WNeto = 0
    For a = 1 To 30
        WImporte = Val(WVector1.TextMatrix(a, 2))
        WNeto = WNeto + WImporte
    Next a
    
    Call Calcula_Importe
    
End Sub

Private Sub Calcula_Importe()

    WImpoDto = 0
    
    WIva1 = 0
    WIva2 = 0
    
    WTotal = WNeto + WIva1 + WIva2
    Total.Caption = Str$(WTotal)
    Total.Caption = Pusing("###,###.##", Total.Caption)

End Sub

Private Sub CmdClose_Click()

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
    With rstDesccomp
        .Close
    End With
    With rstConfiguracion
        .Close
    End With
    With rstFactura
        .Close
    End With
    With rstPlantilla
        .Close
    End With
    
    DbsAdminis.Close
    PrgFactuExpo.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Graba_Click()

    Rem If WLicencia <> "1234-5678-ABCD-EFGH" And Val(Numero.Text) > 10 Then
    Rem     m$ = "La version del sistema es para un uso limitado de movimientos." + Chr$(13) + _
    rem          "El objetivo es el de verificar las opciones y el funcionamiento del mismo." + Chr$(13) + _
    rem          "Para poder utilizar el sistema sin limite de movimientos se debe adquirir la version definitiva."
    Rem     a% = MsgBox(m$, 0, "Sistema de Control de Gestion")
    Rem     Exit Sub
    Rem End If

    If Val(Paridad.Text) = 0 Then Exit Sub

    Call Calcula_Click

    With rstCtaCte
        
        .Index = "Clave"
        .AddNew
        
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
            
        WPunto = Punto.Text
        Call Ceros(WPunto, 4)
        
        If Factura.Value = True Then
            !Tipo = "03"
            !Impre = "FV"
        End If
        If Debito.Value = True Then
            !Tipo = "04"
            !Impre = "ND"
        End If
        If Credito.Value = True Then
            !Tipo = "05"
            !Impre = "NC"
        End If
            
        !Punto = WPunto
        !Letra = Letra.Text
        !Numero = Auxi
        !Renglon = "01"
        !Cliente = Val(Cliente.Text)
        !Fecha = Fecha.Text
        !Estado = "0"
        !Vencimiento = Vencimiento.Text
        If Credito.Value = False Then
            !Total = WTotal
            !Saldo = WTotal
            !Totalus = WTotal
            !Saldous = WTotal
            !Neto = WNeto * Val(Paridad.Text)
            !Iva1 = WIva1 * Val(Paridad.Text)
            !Iva2 = WIva2 * Val(Paridad.Text)
                Else
            !Total = WTotal * -1
            !Saldo = WTotal * -1
            !Totalus = WTotal * -1
            !Saldous = WTotal * -1
            !Neto = WNeto * Val(Paridad.Text) * -1
            !Iva1 = WIva1 * Val(Paridad.Text) * -1
            !Iva2 = WIva2 * Val(Paridad.Text) * -1
        End If
        !OrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        !OrdVencimiento = Right$(Vencimiento.Text, 4) + Mid$(Vencimiento.Text, 4, 2) + Left$(Vencimiento.Text, 2)
        !Pedido = ""
        !Remito = ""
        !Orden = ""
        !Provincia = WProvincia
        !Vendedor = WVendedor
        !Costo = 0
        !Importe1 = 0
        !Importe2 = 0
        !Importe3 = 0
        !Importe4 = 0
        !Importe5 = 0
        !Importe6 = 0
        !Importe7 = 0
        !Tipoventa = 0
        !Proyecto = 0
        !Paridad = Val(Paridad.Text)
            
        !Clave = !Letra + !Tipo + WPunto + Auxi + "01"
        !Busqueda = !Letra + WPunto + Auxi
        
        .Update
            
    End With
                        
    Renglon = 0
    WRenglon = 0
    
    With rstDesccomp
        
        Renglon = 0
        .Index = "Clave"
                                        
        For a = 1 To 30
        
            WVector1.Row = a
            
            WVector1.Col = 1
            WDescripcion = WVector1.Text
                    
            WVector1.Col = 2
            WImporte = Val(WVector1.Text)
                    
            If WDescripcion <> "" Or WImporte <> 0 Then
                    
                Renglon = Renglon + 1
                Auxi = Str$(Renglon)
                Call Ceros(Auxi, 2)
                        
                Auxi1 = Str$(Numero.Text)
                Call Ceros(Auxi1, 8)
                    
                .AddNew
                        
                If Factura.Value = True Then
                    !Tipo = "03"
                End If
                If Debito.Value = True Then
                    !Tipo = "04"
                End If
                If Credito.Value = True Then
                    !Tipo = "05"
                End If
                        
                !Punto = Punto.Text
                !Letra = Letra.Text
                !Numero = Numero.Text
                !Renglon = Renglon
                !Descripcion = WDescripcion
                !Importe = WImporte
                !Clave = !Letra + !Tipo + WPunto + Auxi1 + Auxi
                
                .Update
                        
            End If
                                        
        Next a
            
    End With

    Call Limpia_Click
    Cliente.SetFocus
        
End Sub

Private Sub Numtolet()

    'Convertir en letras el número en Text1
    
    Dim Numero As String
    Dim Letras As String
    Dim sCentimos As String
    Dim sMoneda As String
            
    sMoneda = ""
    sCentimos = "centavos"
    
    Numero = CStr(Val(Total.Caption))
    
    XTexto1 = Numero2Letra(Numero, , sMoneda & " ", sCentimos & " ")
    XTexto1 = XTexto1 + Space$(100)
    
    Pasa = 0
    
    For da = 60 To 1 Step -1
        If Mid$(XTexto1, da, 1) = Space$(1) Then
            Pasa = 1
        End If
        If Pasa = 1 Then
            If Mid$(XTexto1, da, 1) <> Space$(1) Then
                Exit For
            End If
        End If
    Next da
    
    XTexto2 = Mid$(XTexto1, da + 2, 100)
    XTexto1 = Left$(XTexto1, da)
    
End Sub

Private Sub Limpia_Click()

    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Total.Caption = ""
    
    Rem With rstNumero
    Rem     .Index = "Codigo"
    Rem     Claveven$ = "01"
    Rem   .Seek "=", Claveven$
    Rem     If .NoMatch = False Then
    Rem         Numero.Text = !Numero + 1
    Rem             Else
    Rem         Numero.Text = ""
    Rem     End If
    Rem End With
    
    Select Case WVarios
        Case 1
            Factura.Value = True
            Debito.Value = False
            Credito.Value = False
        Case 2
            Factura.Value = False
            Debito.Value = True
            Credito.Value = False
        Case 3
            Factura.Value = False
            Debito.Value = False
            Credito.Value = True
        Case Else
    End Select
    
    Graba.Enabled = True
    
    Call Limpia_Vector
    
    Cliente.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
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
                    WDias = !Dias
                End If
            End With
            Ayuda.Visible = False
            Cliente.SetFocus
            
        Case Else
    End Select
    
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
    
    Letra.Text = ""
    Punto.Text = ""
    Numero.Text = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Select Case WVarios
        Case 1
            Factura.Value = True
            Debito.Value = False
            Credito.Value = False
        Case 2
            Factura.Value = False
            Debito.Value = True
            Credito.Value = False
        Case 3
            Factura.Value = False
            Debito.Value = False
            Credito.Value = True
        Case Else
    End Select
    
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Vencimiento.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Call Limpia_Vector
    
    With rstConfiguracion
       .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            ConfigIva1 = !Iva1
            ConfigIva2 = !Iva2
            ConfigPercepcion = !Percepcion
            ConfigPunto = !Punto
            CantiFac = !CantiFac
            CantiRem = !CantiRem
            CantiArti = !CantiArti
        End If
    End With
     
    Rem With rstNumero
    Rem    .Index = "Codigo"
    Rem    Claveven$ = "01"
    Rem    .Seek "=", Claveven$
    Rem    If .NoMatch = False Then
    Rem        Numero.Text = !Numero + 1
    Rem            Else
    Rem        Numero.Text = ""
    Rem    End If
    Rem  End With
    
End Sub

Private Sub Proceso_Click()
    
    Call Limpia_Vector

    Renglon = 0
    For WRenglon = 1 To 30
    
        With rstDesccomp
    
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
        
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
        
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
        
            If Factura.Value = True Then
                WTipo = "03"
            End If
            If Debito.Value = True Then
                WTipo = "04"
            End If
            If Credito.Value = True Then
                WTipo = "05"
            End If
            
            .Index = "Clave"
            .Seek "=", Letra.Text + WTipo + WPunto + Auxi + Auxi1
            If .NoMatch = False Then
        
                Renglon = Renglon + 1
            
                WVector1.Row = Renglon
                
                WVector1.Col = 1
                WVector1.Text = !Descripcion
                
                WVector1.Col = 2
                If !Importe <> 0 Then
                    WVector1.Text = Pusing("###,###.##", Str$(!Importe))
                        Else
                    WVector1.Text = ""
                End If
                
            End If
        
        End With
    
    Next WRenglon

    Call Calcula_Click
    
    Graba.Enabled = False

End Sub

Private Sub Punto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WPunto = Numero.Text
        Call Ceros(WPunto, 4)
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Punto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstCtaCte
        
            WPunto = Punto.Text
            Call Ceros(WPunto, 4)
            
            Auxi = Numero.Text
            Call Ceros(Auxi, 8)
            
            .Index = "Clave"
            
            If Factura.Value = True Then
                WTipo = "03"
            End If
            If Debito.Value = True Then
                WTipo = "04"
            End If
            If Credito.Value = True Then
                WTipo = "05"
            End If
            
            Claveven$ = Letra.Text + WTipo + WPunto + Auxi + "01"
            .Seek "=", Claveven$
            If .NoMatch = False Then
            
                Fecha.Text = !Fecha
                Cliente.Text = !Cliente
                Vencimiento.Text = !Vencimiento
                
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
                        WDias = !Dias
                    End If
                End With
                
                Call Proceso_Click
                
                    Else
                    
                Graba.Enabled = True
                WNumero = Numero.Text
                Numero.Text = WNumero
                Fecha.SetFocus
                
            End If
        End With
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
                WDias = !Dias
                Select Case Val(WCodIva)
                    Case 1, 2
                        Letra.Text = "A"
                    Case 6
                        Letra.Text = "E"
                    Case Else
                        Letra.Text = "B"
                End Select
                
                WPunto = Str(ConfigPunto)
                Call Ceros(WPunto, 4)
                Punto.Text = WPunto
                
                Numero.Text = "1"
                
                If Factura.Value = True Then
                    WTipo = "03"
                End If
                If Debito.Value = True Then
                    WTipo = "04"
                End If
                If Credito.Value = True Then
                    WTipo = "05"
                End If
                
                With rstCtaCte
                    .Index = "Numero"
                    WClave = Letra.Text + Punto.Text + "99999999"
                    .Seek "<=", Letra.Text + Punto.Text + "99999999"
                    If .NoMatch = False Then
                        If Letra.Text = !Letra And Punto.Text = !Punto Then
                            Numero.Text = Str$(Val(!Numero) + 1)
                        End If
                    End If
                End With
                
                Punto.SetFocus
                    Else
                Cliente.SetFocus
            End If
        End With
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            WFecha = Fecha.Text
            WPlazo1 = WDias + 1
            Call Calcula_vencimiento(WFecha, WPlazo1, WVencimiento)
            Vencimiento.Text = WVencimiento
            Vencimiento.SetFocus
                Else
            m$ = "Formato de fecha invalido"
            a% = MsgBox(m$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Vencimiento_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vencimiento.Text, Auxi)
        If Auxi = "S" Then
            Paridad.SetFocus
                Else
            Vencimiento.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vencimiento.Text = "  /  /    "
    End If
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Paridad.Text = Pusing("###,###.##", Paridad.Text)
        WVector1.Col = 1
        WVector1.Row = 1
        Call StartEdit
    End If
    If KeyAscii = 27 Then
        Paridad.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        Dim IngresaItem As String
        Pantalla.Clear
        WIndice.Clear

        Opcion.Visible = False
        XIndice = Opcion.ListIndex
    
        Rem XIndice = 0
    
        Select Case XIndice
            Case 0
                WEspacios = Len(Ayuda.Text)
                With rstClientes
                    .Index = "Razon"
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
                
            Case Else
        End Select
    
    End If
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Auxiliar
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_DescComp
    OPEN_FILE_Configuracion
    OPEN_FILE_Factura
    OPEN_FILE_Plantilla
End Sub


Rem
Rem Controles de la grilla
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.Visible = True
            WTexto1.SetFocus
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.Visible = True
            WTexto2.SetFocus
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            WTexto3.Visible = True
            WTexto3.SetFocus
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Call Calcula_Click
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
            
        Case 34
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 10 Then
                WVector1.TopRow = WVector1.TopRow + 10
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit
            
        Case 33
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 10 > WVector1.FixedRows Then
                WVector1.TopRow = WVector1.TopRow - 10
                WVector1.Row = WVector1.TopRow
                    Else
                WVector1.TopRow = 1
                WVector1.Row = WVector1.TopRow
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_Grilla
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_Grilla()
    Select Case WVector1.Col
        Case 1
            WVector1.Col = WVector1.Col + 1
        Case 2
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            Rem If Val(WVector1.Text) <> 0 Then
            Rem     With rstConceptos
            Rem         .Index = "Concepto"
            Rem         .Seek "=", Val(WVector1.Text)
            Rem         If .NoMatch = False Then
            Rem             WVector1.Col = 2
            Rem             WVector1.Text = !Nombre
            Rem             WVector1.Col = XColumna
            Rem                 Else
            Rem            WControl = "N"
            Rem         End If
            Rem     End With
            Rem End If
            WVector1.Col = XColumna
        Case 2
            WVector1.Col = XColumna
        Case Else
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Then
        Exit Sub
    End If

    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        WVector1.Col = 2
        WAuxi2 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la grilla en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector1.FixedCols = 1
    WVector1.Cols = 3
    WVector1.FixedRows = 1
    WVector1.Rows = 31
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 6000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Importe"
                WVector1.ColWidth(Ciclo) = 1800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Cliente_DblClick()
    Opcion.Clear
    Opcion.AddItem "Clientes"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    Call Opcion_Click
End Sub

Private Sub Lee_Plantilla()

    Renglon = 0
    For WRenglon = 1 To 30
        With rstPlantilla
            Auxi = Plantilla.Text
            Call Ceros(Auxi, 6)
            Auxi1 = WRenglon
            Call Ceros(Auxi1, 2)
            .Index = "Clave"
            .Seek "=", Auxi + Auxi1
            If .NoMatch = False Then
                Renglon = Renglon + 1
                WVector1.Row = Renglon
                WVector1.Col = 1
                WVector1.Text = !Descripcion
                WVector1.Col = 2
                If !Importe <> 0 Then
                    WVector1.Text = Pusing("###,###.##", Str$(!Importe))
                        Else
                    WVector1.Text = ""
                End If
            End If
        End With
    Next WRenglon
    
    Call Calcula_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Letra_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Punto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vencimiento_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Paridad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Graba_Click
        Case 113
            Call Borra_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call CmdClose_Click
        Case Else
    End Select
End Sub




















