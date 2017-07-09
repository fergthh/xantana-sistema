VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgClienteAdi 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Discriminacion de Vendedore y Descuento por Linea"
   ClientHeight    =   8175
   ClientLeft      =   405
   ClientTop       =   435
   ClientWidth     =   11520
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8175
   ScaleWidth      =   11520
   Visible         =   0   'False
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
      Index           =   7
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3360
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
      Index           =   6
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3720
      Width           =   375
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
      Left            =   10440
      MouseIcon       =   "ClienteAdi.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "ClienteAdi.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Consulta de Datos"
      Top             =   1800
      Width           =   855
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
      Index           =   5
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2760
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
      Index           =   4
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2880
      Width           =   375
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   1455
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
      Left            =   10440
      MouseIcon       =   "ClienteAdi.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "ClienteAdi.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   600
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
      Left            =   10440
      MouseIcon       =   "ClienteAdi.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "ClienteAdi.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Menu Principal"
      Top             =   2880
      Width           =   855
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
      TabIndex        =   11
      Top             =   2400
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3360
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   390
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
      Left            =   4200
      TabIndex        =   9
      Top             =   2400
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2880
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
      Index           =   2
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2880
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
      Index           =   3
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2880
      Width           =   375
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
      TabIndex        =   5
      Top             =   5880
      Visible         =   0   'False
      Width           =   6855
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
      TabIndex        =   4
      Top             =   6240
      Visible         =   0   'False
      Width           =   4215
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
      ItemData        =   "ClienteAdi.frx":21E4
      Left            =   120
      List            =   "ClienteAdi.frx":21EB
      TabIndex        =   1
      Top             =   6240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   4800
      TabIndex        =   12
      Top             =   2400
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
      Height          =   5175
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9128
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label DesCliente 
      BackColor       =   &H00C0C000&
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
      Left            =   2880
      TabIndex        =   16
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "PrgClienteAdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFecha As String
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private XIndice As Single
Dim Vector(100, 10) As String
Private Auxi As String
Private XColor As String
Private XArticulo As String
Private WTipopro As Integer
Private RenglonConjunto As String
Private ZDetalle As String

Rem para el vector

Dim WBorra(1000, 10) As String
Dim WParametros(10, 10) As Double
Dim WFormato(10) As String
Dim WControl As String

Private Sub Consulta_Click()

    Opcion.Clear
    Opcion.AddItem "Linea"
    Opcion.AddItem "Vendedor"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 0
    Call Opcion_Click
     
End Sub


Private Sub Form_Activate()
    Call Proceso_Click
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Ayuda.Visible = True
            Ayuda.Text = ""
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Order by Lineas.Linea"
            spLineas = ZSql
            Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineas.RecordCount > 0 Then
                With rstLineas
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstLineas!Linea + " " + rstLineas!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstLineas!Linea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLineas.Close
            End If
            Ayuda.SetFocus
            
        Case 1
            Ayuda.Visible = True
            Ayuda.Text = ""
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Order by Vendedor.Codigo"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstVendedor!Codigo + " " + rstVendedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstVendedor!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
            Ayuda.SetFocus
            
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub cmdClose_Click()
    PrgClienteAdi.Hide
    Unload Me
    prgcliente.Show
End Sub

Private Sub Graba_Click()
    
    ZSql = ""
    ZSql = ZSql + "DELETE ClienteAdicional"
    ZSql = ZSql + " Where ClienteAdicional.Cliente = " + "'" + Cliente.Text + "'"
    spClienteAdicional = ZSql
    Set rstClienteAdicional = db.OpenRecordset(spClienteAdicional, dbOpenSnapshot, dbSQLPassThrough)

    Renglon = 0
    RenglonII = 0
    WRenglon = 0
        
    For IRow = 1 To 100
            
        WVector1.Row = IRow
            
        WVector1.Col = 1
        Linea = WVector1.Text
        
        WVector1.Col = 3
        Descuento1 = WVector1.Text
        
        WVector1.Col = 4
        Descuento2 = WVector1.Text
        
        WVector1.Col = 5
        Descuento3 = WVector1.Text
        
        WVector1.Col = 6
        Vendedor = WVector1.Text
        
        If Linea <> "" Then
                    
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
            
            Auxi1 = Cliente.Text
            Call Ceros(Auxi1, 6)
            
                    
            ZZClave = Auxi1 + Auxi
            ZZRenglon = Str$(Renglon)
            ZZCliente = Cliente.Text
            ZZLinea = Linea
            ZZDescuento1 = Descuento1
            ZZDescuento2 = Descuento2
            ZZDescuento3 = Descuento3
            ZZVendedor = Vendedor
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ClienteAdicional ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Linea  ,"
            ZSql = ZSql + "Descuento1 ,"
            ZSql = ZSql + "Descuento2 ,"
            ZSql = ZSql + "Descuento3 ,"
            ZSql = ZSql + "Vendedor )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + Cliente.Text + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZLinea + "',"
            ZSql = ZSql + "'" + ZZDescuento1 + "',"
            ZSql = ZSql + "'" + ZZDescuento2 + "',"
            ZSql = ZSql + "'" + ZZDescuento3 + "',"
            ZSql = ZSql + "'" + ZZVendedor + "')"
            spClienteAdicional = ZSql
            Set rstClienteAdicional = db.OpenRecordset(spClienteAdicional, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
                                        
    Next IRow

    Call cmdClose_Click
        
End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            WTexto1.Visible = False
            WTexto2.Visible = False
            Indice = Pantalla.ListIndex
            WVector1.TextMatrix(WVector1.Row, 1) = WIndice.List(Indice)
            WTexto1.Text = WIndice.List(Indice)
            Call WTexto1_KeyDown(13, 0)
            Ayuda.Visible = False
            
        Case 0
            WTexto1.Visible = False
            WTexto2.Visible = False
            Indice = Pantalla.ListIndex
            WVector1.TextMatrix(WVector1.Row, 6) = WIndice.List(Indice)
            WTexto1.Text = WIndice.List(Indice)
            Call WTexto1_KeyDown(13, 0)
            Ayuda.Visible = False
            
        Case Else
    End Select
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    
    Cliente.Text = ZZPasaCliente
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = Trim(rstCliente!Fantasia)
        rstCliente.Close
    End If

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 100
    
        Auxi = Cliente.Text
        Call Ceros(Auxi, 6)
            
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ClienteAdicional"
        ZSql = ZSql + " Where ClienteAdicional.Clave = " + "'" + WClave + "'"
        spClienteAdicional = ZSql
        Set rstClienteAdicional = db.OpenRecordset(spClienteAdicional, dbOpenSnapshot, dbSQLPassThrough)
        If rstClienteAdicional.RecordCount > 0 Then
        
            Renglon = Renglon + 1
            WVector1.Row = Renglon
                
            WVector1.Col = 1
            WVector1.Text = rstClienteAdicional!Linea
            Auxi1 = rstClienteAdicional!Linea
                
            WVector1.Col = 3
            WVector1.Text = Str$(rstClienteAdicional!Descuento1)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            
            WVector1.Col = 4
            WVector1.Text = Str$(rstClienteAdicional!Descuento2)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
            
            WVector1.Col = 5
            WVector1.Text = Str$(rstClienteAdicional!Descuento3)
            WVector1.Text = Pusing("###,###.##", WVector1.Text)
                
            WVector1.Col = 6
            WVector1.Text = rstClienteAdicional!Vendedor
            Auxi2 = rstClienteAdicional!Vendedor
            
            rstClienteAdicional.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Where Lineas.Linea = " + "'" + Auxi1 + "'"
            spLineas = ZSql
            Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineas.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstLineas!Nombre
                rstLineas.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Codigo = " + "'" + Auxi2 + "'"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                WVector1.Col = 7
                WVector1.Text = rstVendedor!Nombre
                rstVendedor.Close
            End If
        
        End If
    
    Next WRenglon
    
    WVector1.Col = 1
    WVector1.Row = 1
    Call StartEdit

End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    Pantalla.Clear
    WIndice.Clear
    
    If KeyAscii > 31 Then
        ZAyuda = Ayuda.Text + Chr$(KeyAscii)
            Else
        Select Case KeyAscii
            Case 27
                Ayuda.Text = ""
                ZAyuda = ""
            Case 8
                WEspacios = Len(Ayuda.Text)
                If WEspacios > 0 Then
                    WEspacios = WEspacios - 1
                    ZAyuda = Left$(Ayuda.Text, WEspacios)
                End If
            Case Else
                ZAyuda = Ayuda.Text
        End Select
    End If
    WEspacios = Len(ZAyuda)
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Where Lineas.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Lineas.Linea"
            spLineas = ZSql
            Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineas.RecordCount > 0 Then
                With rstLineas
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Linea) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Linea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLineas.Close
            End If
            
        Case 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Vendedor.Codigo"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
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
                rstVendedor.Close
            End If
            
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Rem
Rem Controles de la wvector1
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
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
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
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
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
            WTexto3.Visible = True
            WTexto3.SetFocus
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
        Rem Call Suma_Datos
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
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
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
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
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
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
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

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 6, 7
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
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Where Lineas.Linea = " + "'" + WVector1.Text + "'"
            spLineas = ZSql
            Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineas.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstLineas!Nombre
                rstLineas.Close
                    Else
                WControl = "N"
            End If
            
        Case 6
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Codigo = " + "'" + WVector1.Text + "'"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                WVector1.Col = 7
                WVector1.Text = rstVendedor!Nombre
                rstVendedor.Close
                    Else
                WControl = "N"
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
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
        If WAuxi1 <> "" Then
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
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
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

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 8
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
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
                WVector1.Text = "Linea"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Desc. 1"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 4
                WVector1.Text = "Desc. 2"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 5
                WVector1.Text = "Desc. 3"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = "###,###.##"
            Case 6
                WVector1.Text = "Vendedor"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Nombre"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
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
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tama�o de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

        Opcion.Clear
    
        Opcion.AddItem "Linea"
        Opcion.AddItem "Vendedor"

        Rem Opcion.Visible = False
    
        Opcion.ListIndex = 0
    
        Call Opcion_Click
    
    End If
    
    If WVector1.Col = 6 Then

        Opcion.Clear
    
        Opcion.AddItem "Linea"
        Opcion.AddItem "Vendedor"

        Rem Opcion.Visible = False
    
        Opcion.ListIndex = 1
    
        Call Opcion_Click
    
    End If
    
End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

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
        Case 115
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub






















