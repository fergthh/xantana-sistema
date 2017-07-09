VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCompconConsol 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Compras por Concepto"
   ClientHeight    =   7410
   ClientLeft      =   3165
   ClientTop       =   1200
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7410
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
      Top             =   4200
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.Frame Frame2 
      Height          =   3975
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   4815
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
         MouseIcon       =   "CompConConsol.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "CompConConsol.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2880
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
         MouseIcon       =   "CompConConsol.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "CompConConsol.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Consulta de Datos"
         Top             =   2880
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
         MouseIcon       =   "CompConConsol.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "CompConConsol.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2880
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
         Left            =   3720
         MouseIcon       =   "CompConConsol.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "CompConConsol.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salida"
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Hastacon 
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
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Desdecon 
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
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   2520
         TabIndex        =   1
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
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   2520
         TabIndex        =   0
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
      Begin VB.Label Label4 
         Caption         =   "Hasta Concepto"
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
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Concepto"
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
         Top             =   1320
         Width           =   1455
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
         Top             =   840
         Width           =   1335
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
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5280
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Compcon.rpt"
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
      Left            =   5040
      TabIndex        =   3
      Top             =   1920
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
      Height          =   2400
      ItemData        =   "CompConConsol.frx":2D30
      Left            =   120
      List            =   "CompConConsol.frx":2D37
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   5415
   End
End
Attribute VB_Name = "PrgCompconConsol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZSuma(10000) As Double

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    Rem On Error GoTo WError
    
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
    
    Listado.WindowTitle = "Listado de Compras por Concepto"
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
    
    If Val(Desdecon.Text) = 0 Then
        Desdecon.Text = "0"
    End If
    If Val(Hastacon.Text) = 0 Then
        Hastacon.Text = "0"
    End If
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Conceptos SET "
    ZSql = ZSql + " Importe = 0"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)


    Erase ZZSuma


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ivacomp"
    ZSql = ZSql + " Where Ivacomp.OrdFecha >= '" + WDesde + "'"
    ZSql = ZSql + " and Ivacomp.OrdFecha <= '" + WHasta + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveFirst
            Do
            
                ZZConcepto = rstIvaComp!Concepto
                ZZSuma(ZZConcepto) = ZZSuma(ZZConcepto) + rstIvaComp!Exento + rstIvaComp!Neto
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstIvaComp.Close
    End If
    
    
    XEmpresa = WEmpresa
    
    If XEmpresa = 1 Then
            
        WEmpresa = "2"
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        
        txtOdbc = "FraganciasII"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                Else
            
        WEmpresa = "1"
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    End If


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Ivacomp"
    ZSql = ZSql + " Where Ivacomp.OrdFecha >= '" + WDesde + "'"
    ZSql = ZSql + " and Ivacomp.OrdFecha <= '" + WHasta + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveFirst
            Do
            
                ZZConcepto = rstIvaComp!Concepto
                ZZSuma(ZZConcepto) = ZZSuma(ZZConcepto) + rstIvaComp!Exento + rstIvaComp!Neto
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstIvaComp.Close
    End If
    
    
    
    
    If XEmpresa = 1 Then
    
        WEmpresa = "1"
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            Else
            
        WEmpresa = "2"
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        
        txtOdbc = "FraganciasII"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    End If
    
    
    
    
    
    
    For Ciclo = 1 To 10000
        If ZZSuma(Ciclo) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Conceptos SET "
            ZSql = ZSql + " Importe = " + "'" + Str$(ZZSuma(Ciclo)) + "'"
            ZSql = ZSql + " Where Concepto = " + "'" + Str$(Ciclo) + "'"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        End If
    Next Ciclo
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Conceptos SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Conceptos.Concepto, Conceptos.Nombre, Conceptos.Importe, " _
                + "Auxiliar.Nombre " _
                + "From " _
                + DSQ + ".dbo.Conceptos Conceptos, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Conceptos.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "conceptos.Importe <> 0"
                
    Listado.Connect = Connect()
    
    Uno = "{Conceptos.Importe} <> 0 "
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.ReportFileName = "CompConConsol.rpt"
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgCompconConsol.Hide
    Unload Me
    MenuAdminis.Show
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
            Desdecon.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Private Sub Desdecon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastacon.SetFocus
    End If
    If KeyAscii = 27 Then
        Desdecon.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastacon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hastacon.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Desdecon.Text = ""
    Hastacon.Text = ""
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Order by Conceptos.Concepto"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        With rstConceptos
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(!Concepto) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Concepto
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstConceptos.Close
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
    Desdecon.Text = WCodigo
    Hastacon.Text = WCodigo
    Desdecon.SetFocus
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
    ZSql = ZSql + " FROM Conceptos"
    ZSql = ZSql + " Where Conceptos.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
    ZSql = ZSql + " Order by Conceptos.Concepto"
    spConceptos = ZSql
    Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
    If rstConceptos.RecordCount > 0 Then
        With rstConceptos
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(!Concepto) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Concepto
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstConceptos.Close
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

Private Sub DesdeCon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaCon_KeyDown(KeyCode As Integer, Shift As Integer)
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







