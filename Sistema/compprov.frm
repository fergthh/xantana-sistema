VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompprov 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Compras por Proveedor"
   ClientHeight    =   7125
   ClientLeft      =   3165
   ClientTop       =   975
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7125
   ScaleWidth      =   5655
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
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   4935
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
         Left            =   240
         MouseIcon       =   "compprov.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "compprov.frx":030A
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
         Left            =   2640
         MouseIcon       =   "compprov.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "compprov.frx":0E56
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
         Left            =   1440
         MouseIcon       =   "compprov.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "compprov.frx":19A2
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
         Left            =   3840
         MouseIcon       =   "compprov.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "compprov.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salida"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Hastaprv 
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
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox Desdeprv 
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
         Left            =   2520
         MaxLength       =   6
         TabIndex        =   9
         Top             =   1200
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Top             =   720
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   2280
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label Label4 
         Caption         =   "Hasta Proveedor"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Proveedor"
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
         Left            =   600
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
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
         Left            =   600
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
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
         Left            =   600
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Compprv.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva Compras"
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
      Left            =   4680
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
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
      Height          =   2790
      ItemData        =   "compprov.frx":2D30
      Left            =   120
      List            =   "compprov.frx":2D37
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "PrgCompprov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    On Error GoTo WError
    
    If Val(Desdeprv.Text) = 0 Then
        Desdeprv.Text = "0"
    End If
    If Val(Hastaprv.Text) = 0 Then
        Hastaprv.Text = "0"
    End If
    
    WTitulo = "Del " + Desde.Text + " al " + Hasta.Text
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Actividad = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE IvaComp SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Compras por Proveedor"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT IvaComp.Proveedor, IvaComp.Letra, IvaComp.Punto, IvaComp.Fecha, IvaComp.Neto, IvaComp.Exento, IvaComp.Impre, IvaComp.Ordfecha, IvaComp.Concepto, IvaComp.ImpreNumero, " _
                + "Proveedor.Nombre, " _
                + "Auxiliar.Nombre, Auxiliar.Actividad, " _
                + "Conceptos.Nombre " _
                + "From " _
                + DSQ + ".dbo.IvaComp IvaComp, " _
                + DSQ + ".dbo.Proveedor Proveedor, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Conceptos Conceptos " _
                + "Where " _
                + "IvaComp.Proveedor = Proveedor.Proveedor AND " _
                + "IvaComp.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "IvaComp.Concepto = Conceptos.Concepto AND " _
                + "IvaComp.Proveedor >= " + Desdeprv.Text + " AND " _
                + "IvaComp.Proveedor <= " + Hastaprv.Text + " AND " _
                + "IvaComp.Ordfecha >= '" + WDesde + "' AND " _
                + "IvaComp.Ordfecha <= '" + WHasta + "'"
    
    Listado.Connect = Connect()
    
    Uno = "{Ivacomp.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {Ivacomp.Proveedor} in " + Desdeprv.Text + " to " + Hastaprv.Text
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgCompprov.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Desde.Text = "  /  /    "
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desdeprv.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Private Sub Desdeprv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastaprv.SetFocus
    End If
    If KeyAscii = 27 Then
        Desdeprv.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastaprv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hastaprv.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Desdeprv.Text = ""
    Hastaprv.Text = ""
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Order by Proveedor.Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        With rstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(!Proveedor) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProveedor.Close
    End If
            
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
    Indice = Pantalla.ListIndex
    WCodigo = WIndice.List(Indice)
    Desdeprv.Text = WCodigo
    Hastaprv.Text = WCodigo
    Desdeprv.SetFocus
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
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Proveedor"
    ZSql = ZSql + " Where Proveedor.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
    ZSql = ZSql + " Order by Proveedor.Proveedor"
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstProveedor.RecordCount > 0 Then
        With rstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(!Proveedor) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstProveedor.Close
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdePrv_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaPrv_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call Cancela_Click
        Case Else
    End Select
End Sub









