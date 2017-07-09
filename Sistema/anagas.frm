VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAnagas 
   Caption         =   "Analisis de Presupuesto de Gastos"
   ClientHeight    =   6060
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6060
   ScaleWidth      =   8145
   Begin VB.ListBox Opcion 
      Height          =   840
      Left            =   1440
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Ano 
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
         MaxLength       =   4
         TabIndex        =   17
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Mes 
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
         MaxLength       =   2
         TabIndex        =   16
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Hasta 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Desde 
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Left            =   3120
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Año"
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
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Mes"
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
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Producto"
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
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Producto"
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Anagas.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
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
      Left            =   6120
      TabIndex        =   4
      Top             =   720
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
      ItemData        =   "anagas.frx":0000
      Left            =   240
      List            =   "anagas.frx":0007
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   7455
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5760
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6120
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgAnagas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstUnidad As Recordset
Dim spUnidad As String
Dim rstPtoGas As Recordset
Dim spPtoGas As String
Dim rstAnaGas As Recordset
Dim spAnaGas As String
Dim rstImpu As Recordset
Dim spImpu As String
Dim rstIvacomp As Recordset
Dim spIvacomp As String
Dim WProducto As String
Dim WProveedor As String
Dim WLetra As String
Dim WTipo As String
Dim WPunto As String
Dim WNumero As String
Dim WRenglon As String
Dim WTipomovi As String
Dim Vector(10000, 6) As String
Dim Vector1(10000, 6) As String

Private Sub Acepta_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WEmpe = !Nombre
        End If
    End With

    With rstAuxiliar
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            .Edit
            !Nombre = WEmpe
            !varios = "Periodo Año : " + Mes.Text + "/" + Ano.Text
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Presupuesto de Gastos"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Anagas.Producto} in " + Desde.Text + " to " + Hasta.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Set rstAnaGas = db.OpenRecordset("BorrarAnagas ", dbOpenSnapshot, dbSQLPassThrough)
    
    Rem -------------------------------------------------------------------
    Rem TOMA LOS PRESUPUESTADO
    Rem -------------------------------------------------------------------
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                + Hasta.Text + "'"
    
    spPtoGas = "ListaPtogasProducto " + XParam
    Set rstPtoGas = db.OpenRecordset(spPtoGas, dbOpenSnapshot, dbSQLPassThrough)
    If rstPtoGas.RecordCount > 0 Then
        With rstPtoGas
            .MoveFirst
            Do
                If .EOF = False Then
            
                    If !Ano = Val(Ano.Text) Then
                    
                        WImpo1 = "0"
                        WImpo2 = "0"
                        WImpo3 = "0"
                        WImpo4 = "0"
                        
                        Select Case Val(Mes.Text)
                             Case 1
                                WImpo1 = Str$(rstPtoGas!Horas1)
                                WImpo2 = Str$(rstPtoGas!Importe1)
                             Case 2
                                WImpo1 = Str$(rstPtoGas!Horas2)
                                WImpo2 = Str$(rstPtoGas!Importe2)
                             Case 3
                                WImpo1 = Str$(rstPtoGas!Horas3)
                                WImpo2 = Str$(rstPtoGas!Importe3)
                             Case 4
                                WImpo1 = Str$(rstPtoGas!Horas4)
                                WImpo2 = Str$(rstPtoGas!Importe4)
                             Case 5
                                WImpo1 = Str$(rstPtoGas!Horas5)
                                WImpo2 = Str$(rstPtoGas!Importe5)
                             Case 6
                                WImpo1 = Str$(rstPtoGas!Horas6)
                                WImpo2 = Str$(rstPtoGas!Importe6)
                             Case 7
                                WImpo1 = Str$(rstPtoGas!Horas7)
                                WImpo2 = Str$(rstPtoGas!Importe7)
                             Case 8
                                WImpo1 = Str$(rstPtoGas!Horas8)
                                WImpo2 = Str$(rstPtoGas!Importe8)
                             Case 9
                                WImpo1 = Str$(rstPtoGas!Horas9)
                                WImpo2 = Str$(rstPtoGas!Importe9)
                             Case 10
                                WImpo1 = Str$(rstPtoGas!Horas10)
                                WImpo2 = Str$(rstPtoGas!Importe10)
                             Case 11
                                WImpo1 = Str$(rstPtoGas!Horas11)
                                WImpo2 = Str$(rstPtoGas!Importe11)
                             Case 12
                                WImpo1 = Str$(rstPtoGas!Horas12)
                                WImpo2 = Str$(rstPtoGas!Importe12)
                            Case Else
                        End Select
                        
                        If Val(WImpo1) <> 0 Or Val(WImpo2) <> 0 Then
                        
                            WProducto = rstPtoGas!Producto
                            WGasto = rstPtoGas!gasto
                            WTarea = rstPtoGas!Tarea
                            WGerencia = rstPtoGas!Gerencia
                            
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = WProducto
                            Vector(Renglon, 2) = WGasto
                            Vector(Renglon, 3) = WImpo1
                            Vector(Renglon, 4) = WImpo2
                            Vector(Renglon, 5) = WTarea
                            Vector(Renglon, 6) = WGerencia
                            
                        End If
                    
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPtoGas.Close
    End If
    
    For Cicla1 = 1 To Renglon
    
        WProducto = Vector(Cicla1, 1)
        WGasto = Vector(Cicla1, 2)
        WTarea = Vector(Cicla1, 5)
        WGerencia = Vector(Cicla1, 6)
        WImpo1 = Vector(Cicla1, 3)
        WImpo2 = Vector(Cicla1, 4)
        WImpo3 = ""
        WImpo4 = ""
        
        XParam = "'" + WProducto + "','" _
                + WGasto + "','" _
                + WTarea + "','" _
                + WGerencia + "'"
        spAnaGas = "ConsultaAnagas " + XParam
        Set rstAnaGas = db.OpenRecordset(spAnaGas, dbOpenSnapshot, dbSQLPassThrough)
        If rstAnaGas.RecordCount > 0 Then
        
            WImpo1 = Str$(Val(WImpo1) + rstAnaGas!Impo1)
            WImpo2 = Str$(Val(WImpo2) + rstAnaGas!Impo2)
            WImpo3 = Str$(rstAnaGas!Impo3)
            WImpo4 = Str$(rstAnaGas!Impo4)
            rstAnaGas.Close
                    
            XParam = "'" + WProducto + "','" _
                    + WGasto + "','" _
                    + WImpo1 + "','" _
                    + WImpo2 + "','" _
                    + WImpo3 + "','" _
                    + WImpo4 + "','" _
                    + WTarea + "','" _
                    + WGerencia + "'"
            
            Set rstAnaGas = db.OpenRecordset("ModificaAnagas " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            XParam = "'" + WProducto + "','" _
                    + WGasto + "','" _
                    + WImpo1 + "','" _
                    + WImpo2 + "','" _
                    + WImpo3 + "','" _
                    + WImpo4 + "','" _
                    + WTarea + "','" _
                    + WGerencia + "'"
            
            Set rstAnaGas = db.OpenRecordset("AltaAnaGas " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
        End If
    
    Next Cicla1

    Rem -------------------------------------------------------------------
    Rem TOMA LOS REALIZADO
    Rem -------------------------------------------------------------------
    
    Erase Vector
    Erase Vector1
    Renglon = 0
    Renglon1 = 0
    
    spIvacomp = "ListaIvacomp"
    Set rstIvacomp = db.OpenRecordset(spIvacomp, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstIvacomp.RecordCount > 0 Then
        With rstIvacomp
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Val(Desde.Text) <= rstIvacomp!Unidad And Val(Hasta.Text) >= rstIvacomp!Unidad Then
            
                        If rstIvacomp!Ano = Val(Ano.Text) Then
                        
                            If rstIvacomp!Mes = Val(Mes.Text) Then
                        
                                WProducto = rstIvacomp!Unidad
                                WProveedor = rstIvacomp!Proveedor
                                WLetra = rstIvacomp!Letra
                                WTipo = rstIvacomp!Tipo
                                WPunto = rstIvacomp!Punto
                                WNumero = rstIvacomp!Numero
                        
                                Renglon1 = Renglon1 + 1
                                Vector1(Renglon1, 1) = WProducto
                                Vector1(Renglon1, 2) = WProveedor
                                Vector1(Renglon1, 3) = WLetra
                                Vector1(Renglon1, 4) = WTipo
                                Vector1(Renglon1, 5) = WPunto
                                Vector1(Renglon1, 6) = WNumero
                                
                            End If
                        
                        End If
                    
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstIvacomp.Close
    End If
    
    For Cicla2 = 1 To Renglon1
    
        WProducto = Vector1(Cicla2, 1)
        WProveedor = Vector1(Cicla2, 2)
        WLetra = Left$(Vector1(Cicla2, 3), 1)
        WTipo = Vector1(Cicla2, 4)
        WPunto = Vector1(Cicla2, 5)
        WNumero = Vector1(Cicla2, 6)
        
        Call Ceros(WProveedor, 11)
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 4)
        Call Ceros(WNumero, 8)
        
        For A = 1 To 20
        
            WTipomovi = "2"
            
            Auxi1 = Str$(A)
            Call Ceros(Auxi1, 2)
            WRenglon = Auxi1
            
            ClaveImputac = WTipomovi + WProveedor + WLetra + WTipo + WPunto + WNumero + WRenglon
            aa = Len(ClaveImputac)
            spImputac = "Consultaimputac " + "'" + ClaveImputac + "'"
            Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
            If rstImputac.RecordCount > 0 Then
            
                If Left$(rstImputac!Cuenta, 1) = "5" Then
                
                    WDebito = rstImputac!Debito
                    WCredito = rstImputac!Credito
                    WGasto = Mid$(rstImputac!Cuenta, 6, 1) + "00"
                    WTarea = Mid$(rstImputac!Cuenta, 6, 3)
                    WGerencia = Mid$(rstImputac!Cuenta, 2, 1)
                    
                    If WDebito <> 0 Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = WProducto
                        Vector(Renglon, 2) = WGasto
                        Vector(Renglon, 3) = Abs(WDebito)
                        Vector(Renglon, 5) = WTarea
                        Vector(Renglon, 6) = WGerencia
                    End If
                    
                    If WCredito <> 0 Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = WProducto
                        Vector(Renglon, 2) = WGasto
                        Vector(Renglon, 3) = Abs(WCredito) * -1
                        Vector(Renglon, 5) = WTarea
                        Vector(Renglon, 6) = WGerencia
                    End If
                    
                End If
                
            End If
        Next A
            
    Next Cicla2
        
    For Cicla3 = 1 To Renglon
    
        WProducto = Vector(Cicla3, 1)
        WGasto = Vector(Cicla3, 2)
        WTarea = Vector(Cicla3, 5)
        WGerencia = Vector(Cicla3, 6)
        WImpo1 = ""
        WImpo2 = ""
        WImpo3 = ""
        WImpo4 = Vector(Cicla3, 3)
        
        XParam = "'" + WProducto + "','" _
                + WGasto + "','" _
                + WTarea + "','" _
                + WGerencia + "'"
        spAnaGas = "ConsultaAnagas " + XParam
        Set rstAnaGas = db.OpenRecordset(spAnaGas, dbOpenSnapshot, dbSQLPassThrough)
        If rstAnaGas.RecordCount > 0 Then
        
            WImpo1 = Str$(rstAnaGas!Impo1)
            WImpo2 = Str$(rstAnaGas!Impo2)
            WImpo3 = Str$(rstAnaGas!Impo3)
            WImpo4 = Str$(Val(WImpo4) + rstAnaGas!Impo4)
            rstAnaGas.Close
                    
            XParam = "'" + WProducto + "','" _
                    + WGasto + "','" _
                    + WImpo1 + "','" _
                    + WImpo2 + "','" _
                    + WImpo3 + "','" _
                    + WImpo4 + "','" _
                    + WTarea + "','" _
                    + WGerencia + "'"
            
            Set rstAnaGas = db.OpenRecordset("ModificaAnagas " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            XParam = "'" + WProducto + "','" _
                    + WGasto + "','" _
                    + WImpo1 + "','" _
                    + WImpo2 + "','" _
                    + WImpo3 + "','" _
                    + WImpo4 + "','" _
                    + WTarea + "','" _
                    + WGerencia + "'"
            
            Set rstAnaGas = db.OpenRecordset("AltaAnaGas " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
        End If
    
    Next Cicla3
        
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Anagas.Producto, Anagas.Gasto, Anagas.Impo2, Anagas.Impo4, Anagas.Tarea, Anagas.Gerencia, " _
                        + "Unidad.Nombre, Gasto.Nombre, Tarea.Nombre, Gerencia.Nombre " _
                        + "From " _
                        + DSQ + ".dbo.Anagas Anagas, " _
                        + DSQ + ".dbo.Unidad Unidad, " _
                        + DSQ + ".dbo.Gasto Gasto, " _
                        + DSQ + ".dbo.Tarea Tarea, " _
                        + DSQ + ".dbo.Gerencia Gerencia " _
                        + "Where " _
                        + "Anagas.Producto = Unidad.Codigo AND " _
                        + "Anagas.Gasto = Gasto.Codigo AND " _
                        + "Anagas.Tarea = Tarea.Codigo AND " _
                        + "Anagas.Gerencia = Gerencia.Codigo AND " _
                        + "Anagas.Producto >= 0 AND Anagas.Producto <= 9999"
    
    Rem Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    
    With rstEmpresa
        .Close
    End With
    
    Desde.SetFocus
    PrgAnagas.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Mes.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Mes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ano.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Mes.Text = ""
    Ano.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

     Opcion.Clear

     Opcion.AddItem "Productos"
     Rem Opcion.AddItem "Cuentas Contables"

     Rem Opcion.Visible = True
     
     Opcion.ListIndex = 0
     
     Call Opcion_Click
     
End Sub


Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Ayuda.Text = ""
    Ayuda.Visible = True
    
    Select Case XIndice
        Case 0
            spUnidad = "ListaUnidades"
            Set rstUnidad = db.OpenRecordset(spUnidad, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstUnidad
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstUnidad!Codigo) + " " + rstUnidad!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstUnidad!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstUnidad.Close
            
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WUnidad = WIndice.List(Indice)
            spUnidad = "ConsultaUnidad " + "'" + Str$(WUnidad) + "'"
            Set rstUnidad = db.OpenRecordset(spUnidad, dbOpenSnapshot, dbSQLPassThrough)
            If rstUnidad.RecordCount > 0 Then
                Desde.Text = rstUnidad!Codigo
                Hasta.Text = rstUnidad!Codigo
                rstUnidad.Close
                        Else
                Desde.Text = WUnidad
                Hasta.Text = WUnidad
            End If
            Desde.SetFocus
            
        Case Else
    End Select
    
    Ayuda.Text = ""
    Ayuda.Visible = False
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spUnidad = "ListaUnidades"
    Set rstUnidad = db.OpenRecordset(spUnidad, dbOpenSnapshot, dbSQLPassThrough)
    If rstUnidad.RecordCount > 0 Then
        With rstUnidad
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Da = Len(rstUnidad!Nombre) - WEspacios
                
                    For aa = 1 To Da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                            Auxi = rstUnidad!Codigo
                            IngresaItem = Auxi + "    " + rstUnidad!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstUnidad!Codigo
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
        rstUnidad.Close
    End If
    End If

End Sub




