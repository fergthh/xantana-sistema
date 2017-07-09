VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCtaCtefecha 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Clientes a Fecha"
   ClientHeight    =   5955
   ClientLeft      =   2490
   ClientTop       =   585
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   ScaleHeight     =   5955
   ScaleWidth      =   7215
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
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Frame Frame2 
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
      Height          =   3015
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   5895
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
         Left            =   720
         MouseIcon       =   "ctactefecha.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ctactefecha.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1800
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
         Left            =   3120
         MouseIcon       =   "ctactefecha.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ctactefecha.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Consulta de Datos"
         Top             =   1800
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
         Left            =   1920
         MouseIcon       =   "ctactefecha.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ctactefecha.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1800
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
         Left            =   4320
         MouseIcon       =   "ctactefecha.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ctactefecha.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salida"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Hasta 
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
         MaxLength       =   10
         TabIndex        =   7
         Text            =   " "
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox Desde 
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
         MaxLength       =   10
         TabIndex        =   1
         Text            =   " "
         Top             =   840
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   360
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
      Begin VB.Label Label3 
         Caption         =   "Fecha Emision"
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
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1335
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
         Left            =   360
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
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
         Left            =   360
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6480
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impctacte.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Clientes"
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
      Left            =   4560
      TabIndex        =   3
      Top             =   4320
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
      Height          =   2205
      ItemData        =   "ctactefecha.frx":2D30
      Left            =   120
      List            =   "ctactefecha.frx":2D37
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   6735
   End
End
Attribute VB_Name = "PrgCtaCtefecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WPasa As String
Private WTitulo As String
Private Importe3 As Double
Dim BajaRecibo(10000) As String
Dim BajaVarios(5000, 4) As String
Dim WSaldo As Double

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
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFechaOrd = WAno + WMes + WDia
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Importe7 = Saldo" + ","
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase BajaRecibo
    LugarRecibo = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.FechaOrd > " + "'" + WFechaOrd + "'"
    ZSql = ZSql + " and Recibos.Cliente >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Recibos.Cliente <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and Recibos.Renglon = " + "'" + "01" + "'"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
    
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Val(rstRecibos!Recibo) < 900000 Then
                        LugarRecibo = LugarRecibo + 1
                        BajaRecibo(LugarRecibo) = rstRecibos!Recibo
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstRecibos.Close
    End If
    
    For Cicla = 1 To LugarRecibo
    
        WRecibo = BajaRecibo(Cicla)
        
        For da = 1 To 99
        
            Auxi1 = WRecibo
            Call Ceros(Auxi1, 6)
            
            Auxi2 = Str$(da)
            Call Ceros(Auxi2, 2)
            
            ZClave = Auxi1 + Auxi2
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.Clave = " + "'" + ZClave + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
            
                WTipoRec = rstRecibos!Tiporec
                WLetra = rstRecibos!Letra1
                WTipo = rstRecibos!Tipo1
                WPunto = rstRecibos!Punto1
                WNumero = rstRecibos!Numero1
                WImporte = rstRecibos!Importe1
                WCliente = rstRecibos!Cliente
                WTipoReg = rstRecibos!Tiporeg
                
                rstRecibos.Close
                
                If Val(WTipoReg) = 1 Then
                
                    WClave = WLetra + WTipo + WPunto + WNumero + "01"
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CtaCte SET "
                    ZSql = ZSql + " Importe7 = Importe7 + " + "'" + Str$(WImporte) + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                    spCtaCte = ZSql
                    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
            End If
            
        Next da
        
    Next Cicla
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Erase BajaVarios
    LugarVarios = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.OrdFecha > " + "'" + WFechaOrd + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
    
        With rstCtaCte
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Val(rstCtaCte!Remito) <> 0 Then
                        Select Case Val(rstCtaCte!Tipo)
                            Case 2, 3, 4, 5
                                LugarVarios = LugarVarios + 1
                                BajaVarios(LugarVarios, 1) = rstCtaCte!Remito
                                BajaVarios(LugarVarios, 2) = rstCtaCte!Letra
                                BajaVarios(LugarVarios, 3) = rstCtaCte!Punto
                                BajaVarios(LugarVarios, 4) = rstCtaCte!Total * -1
                            Case Else
                        End Select
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCte.Close
    End If
    
    For Cicla = 1 To LugarVarios
    
        WNumero = BajaVarios(Cicla, 1)
        WLetra = BajaVarios(Cicla, 2)
        WPunto = BajaVarios(Cicla, 3)
        WImporte = BajaVarios(Cicla, 4)
        WTipo = "01"
        
        Auxi$ = WNumero
        Call Ceros(Auxi$, 8)
        WClave = WLetra + WTipo + WPunto + Auxi$ + "01"
        
        
        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCte SET "
        ZSql = ZSql + " Importe7 = Importe7 + " + "'" + Str$(WImporte) + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Cicla
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Listado.WindowTitle = "Listado de Cuenta Corriente a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = "SELECT CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.Total, CtaCte.Saldo, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Remito, CtaCte.Linea, CtaCte.Importe7, " _
            + "Cliente.Razon, Cliente.Telefono, Cliente.Vendedor, " _
            + "Auxiliar.Nombre, Auxiliar.Varios " _
            + "From " _
            + DSQ + ".dbo.CtaCte CtaCte, " _
            + DSQ + ".dbo.Cliente Cliente, " _
            + DSQ + ".dbo.Auxiliar Auxiliar " _
            + "Where " _
            + "CtaCte.Cliente = Cliente.Cliente AND " _
            + "CtaCte.CodigoEmpresa = Auxiliar.Empresa AND " _
            + "CtaCte.Cliente >= '" + Desde.Text + "' AND " _
            + "CtaCte.Cliente <= '" + Hasta.Text + "' AND " _
            + "CtaCte.Importe7 <> 0 AND " _
            + "CtaCte.OrdFecha >= '" + "00000000" + "' AND " _
            + "CtaCte.OrdFecha <= '" + WFechaOrd + "'"
    
    Uno = "{CtaCte.Importe7} <> 0 "
    Dos = " and {CtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Tres = " and {CtaCte.OrdFecha} in " + Chr$(34) + "00000000" + Chr$(34) + " to " + Chr$(34) + WFechaOrd + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
    
    Listado.ReportFileName = "CtacteFecha.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgCtaCtefecha.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    Ayuda.Visible = True
    Ayuda.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Order by Cliente.Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = !Cliente + " " + !Razon
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Cliente
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Indice = Pantalla.ListIndex
    Desde.Text = WIndice.List(Indice)
    Hasta.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    Desde.SetFocus
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Frame2.Visible = True
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
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
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

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
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




