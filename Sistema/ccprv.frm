VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCcprv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Proveedores"
   ClientHeight    =   7770
   ClientLeft      =   1440
   ClientTop       =   525
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   9120
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
      Height          =   300
      Left            =   480
      TabIndex        =   1
      Text            =   " "
      Top             =   3360
      Width           =   7695
   End
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   1320
      TabIndex        =   5
      Top             =   240
      Width           =   6255
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
         Left            =   4080
         MouseIcon       =   "ccprv.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ccprv.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   480
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
         Left            =   5160
         MouseIcon       =   "ccprv.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ccprv.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salida"
         Top             =   1680
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
         Left            =   5160
         MouseIcon       =   "ccprv.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ccprv.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Impresion x Impresora"
         Top             =   480
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
         Left            =   4080
         MouseIcon       =   "ccprv.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ccprv.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consulta de Datos"
         Top             =   1680
         Width           =   855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Listado"
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
         Height          =   855
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   3375
         Begin VB.OptionButton Tipo2 
            Caption         =   "Completo"
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
            Left            =   1560
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Tipo1 
            Caption         =   "Pendiente"
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
            TabIndex        =   9
            Top             =   360
            Width           =   1335
         End
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
         MaxLength       =   8
         TabIndex        =   2
         Text            =   " "
         Top             =   1080
         Width           =   1215
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
         MaxLength       =   8
         TabIndex        =   0
         Text            =   " "
         Top             =   600
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
         Left            =   480
         TabIndex        =   7
         Top             =   1080
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
         Left            =   480
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   -120
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ccprv.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Proveedores"
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
      Left            =   720
      TabIndex        =   4
      Top             =   360
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
      Height          =   3570
      ItemData        =   "ccprv.frx":2D30
      Left            =   480
      List            =   "ccprv.frx":2D37
      TabIndex        =   3
      Top             =   3720
      Width           =   7695
   End
End
Attribute VB_Name = "PrgCcprv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Acumula As Double
Private Pasa As Single
Private WSaldo As Double

Dim ZVector(10000, 3) As String

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

    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCtePrv SET "
    ZSql = ZSql + " Empresa = " + "'" + WNombreEmpresa + "'"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)



    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCteprv SET "
    ZSql = ZSql + " Saldo = 0"
    ZSql = ZSql + " Where Saldo < 0.01 and Saldo > -0.01"
    ZSql = ZSql + " and CtaCtePrv.Proveedor >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.Proveedor <= " + "'" + Hasta.Text + "'"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)


    
    Erase ZVector
    ZLugar = 0
    ZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCtePrv"
    ZSql = ZSql + " Where CtaCtePrv.Proveedor >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.Proveedor <= " + "'" + Hasta.Text + "'"
    If Tipo1.Value = True Then
        ZSql = ZSql + " and CtaCtePRV.Saldo <> 0"
    End If
    ZSql = ZSql + " Order by CtaCtePrv.Proveedor,CtaCtePrv.OrdFecha,CtaCtePrv.Impre,CtaCtePrv.Numero"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCtePrv.RecordCount > 0 Then
        With rstCtaCtePrv
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZPasa = 0 Then
                        ZPasa = 1
                        ZCorte = rstCtaCtePrv!Proveedor
                        ZSuma = 0
                    End If
                    
                    If ZCorte <> rstCtaCtePrv!Proveedor Then
                        ZCorte = rstCtaCtePrv!Proveedor
                        ZSuma = 0
                    End If
                    
                    ZSuma = ZSuma + rstCtaCtePrv!Saldo
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar, 1) = rstCtaCtePrv!Clave
                    ZVector(ZLugar, 2) = Str$(ZSuma)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCtePrv.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCtePrv SET "
        ZSql = ZSql + " Importe7 = " + "'" + ZVector(Ciclo, 2) + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZVector(Ciclo, 1) + "'"
        spCtaCtePrv = ZSql
        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo














    Listado.WindowTitle = "Listado de Cuenta Corriente de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Tipo1.Value = True Then
            
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT CtaCtePrv.Proveedor, CtaCtePrv.Letra, CtaCtePrv.Punto, CtaCtePrv.Numero, CtaCtePrv.fecha, CtaCtePrv.Vencimiento, CtaCtePrv.Total, CtaCtePrv.Saldo, CtaCtePrv.OrdFecha, CtaCtePrv.Impre, CtaCtePrv.Observaciones, CtaCtePrv.Empresa, " _
                    + "Proveedor.Nombre, Proveedor.NombreCheque " _
                    + "From " _
                    + DSQ + ".dbo.CtaCtePrv CtaCtePrv, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "CtaCtePrv.Proveedor = Proveedor.Proveedor AND " _
                    + "CtaCtePrv.Proveedor >= '" + Desde.Text + "' AND " _
                    + "CtaCtePrv.Proveedor <= '" + Hasta.Text + "' AND " _
                    + "CtaCtePrv.Saldo <> 0"
                
        Listado.Connect = Connect()
    
        Uno = "{CtaCtePrv.Saldo} <> 0 "
        Dos = " and {CtaCtePrv.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
        
        Listado.ReportFileName = "CcprvPend.rpt"
    
            Else
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
    
        Listado.SQLQuery = "SELECT CtaCtePrv.Proveedor, CtaCtePrv.Letra, CtaCtePrv.Punto, CtaCtePrv.Numero, CtaCtePrv.fecha, CtaCtePrv.Vencimiento, CtaCtePrv.Total, CtaCtePrv.Saldo, CtaCtePrv.OrdFecha, CtaCtePrv.Impre, CtaCtePrv.Observaciones, CtaCtePrv.Empresa, " _
                    + "Proveedor.Nombre, Proveedor.NombreCheque " _
                    + "From " _
                    + DSQ + ".dbo.CtaCtePrv CtaCtePrv, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "CtaCtePrv.Proveedor = Proveedor.Proveedor AND " _
                    + "CtaCtePrv.Proveedor >= '" + Desde.Text + "' AND " _
                    + "CtaCtePrv.Proveedor <= '" + Hasta.Text + "'"
                
        Listado.Connect = Connect()
        
        Uno = "{CtaCtePrv.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
        
        Listado.GroupSelectionFormula = Uno
        Listado.SelectionFormula = Uno
        
        Listado.ReportFileName = "Ccprv.rpt"
            
    End If
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgCcprv.Hide
    Unload Me
    MenuAdminis.Show
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
                    IngresaItem = !Proveedor + " " + !Nombre
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
            
    Rem Pantalla.Visible = True
    Ayuda.Text = ""
    Rem Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Indice = Pantalla.ListIndex
    Desde.Text = WIndice.List(Indice)
    Hasta.Text = WIndice.List(Indice)
    Desde.SetFocus
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Tipo1.Value = True
    Tipo2.Value = False
    Frame2.Visible = True
    Call Consulta_Click
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
                    IngresaItem = !Proveedor + " " + !Nombre
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


