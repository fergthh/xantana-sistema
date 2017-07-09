VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgComparaIV 
   Caption         =   "Listado Comparativo de Egresos"
   ClientHeight    =   7365
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   8145
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   1560
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   3015
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
      Left            =   360
      TabIndex        =   10
      Top             =   3840
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
      Height          =   3255
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
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
         Left            =   1920
         TabIndex        =   19
         Top             =   2520
         Width           =   1935
      End
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
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1920
         Width           =   855
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
         Left            =   2280
         MaxLength       =   2
         TabIndex        =   17
         Top             =   1440
         Width           =   495
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
         TabIndex        =   14
         Top             =   960
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
         Top             =   600
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
         Left            =   3240
         TabIndex        =   9
         Top             =   1200
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
         Left            =   3240
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   480
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
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         TabIndex        =   16
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label7 
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
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Unidad"
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
         TabIndex        =   13
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Unidad"
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
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7560
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ComparaIV.rpt"
      Destination     =   1
      WindowTitle     =   "Listado Comparativo de Ingresos"
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
      ItemData        =   "ComparaIV.frx":0000
      Left            =   360
      List            =   "ComparaIV.frx":0007
      TabIndex        =   3
      Top             =   4200
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
Attribute VB_Name = "PrgComparaIV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WImporte1 As String
Private WImporte2 As String
Private WImporte3 As String
Private WImporte4 As String
Private WImporte5 As String
Private WImporte6 As String
Private WImporte7 As String
Private WImporte8 As String
Private WImporte9 As String
Private WImporte10 As String
Private WImporte11 As String
Private WImporte12 As String
Private WRubro As String
Private WSubRubro As String
Private WConcepto As String
Private WUnidad As String
Private WCliente As String
Dim rstPtoAdmiII As Recordset
Dim spPtoAdmII As String
Dim rstLista As Recordset
Dim spLista As String
Dim rstIvacomp As Recordset
Dim spIvacomp As String
Dim rstImputac As Recordset
Dim spImputac As String
Dim rstUnidad As Recordset
Dim spUnidad As String
Dim Vector(5000, 30) As String
Dim Vector1(10000, 7) As String
Dim Apertura(100, 2) As String
Dim Semestre As Integer

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
            !varios = "Periodo Mes/Año : " + Mes.Text + "/" + Ano.Text
            If Val(Mes.Text) > 6 Then
                !Impre1 = "Julio"
                !Impre2 = "Agosto"
                !Impre3 = "Septiembre"
                !Impre4 = "Octubre"
                !Impre5 = "Noviembre"
                !Impre6 = "Diciembre"
                Semestre = 2
                    Else
                !Impre1 = "Enero"
                !Impre2 = "Febrero"
                !Impre3 = "Marzo"
                !Impre4 = "Abril"
                !Impre5 = "Mayo"
                !Impre6 = "Junio"
                Semestre = 1
            End If
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado Comparativo de Egresos"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{Lista.Rubro} in " + Desde.Text + " to " + Hasta.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Renglon = 0
    Set rstLista = db.OpenRecordset("BorrarLista ", dbOpenSnapshot, dbSQLPassThrough)
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                + Hasta.Text + "'"
    
    spPtoAdmII = "ListaPtoUnidadAdmII " + XParam
    Set rstPtoAdmiII = db.OpenRecordset(spPtoAdmII, dbOpenSnapshot, dbSQLPassThrough)
            
    If rstPtoAdmiII.RecordCount > 0 Then
    With rstPtoAdmiII
        .MoveFirst
        Do
            If .EOF = False Then
            
                If !Ano = Val(Ano.Text) Then
                
                    Renglon = Renglon + 1
                
                    WUnidad = Str$(rstPtoAdmiII!Unidad)
                    WConcepto = Str$(rstPtoAdmiII!Concepto)
                    WImporte1 = "0"
                    WImporte2 = "0"
                    WImporte3 = "0"
                    WImporte4 = "0"
                    WImporte5 = "0"
                    WImporte6 = "0"
                    WImporte7 = "0"
                    
                    Select Case Val(Mes.Text)
                            Case 1
                                WImporte1 = Str$(rstPtoAdmiII!Importe1)
                            Case 2
                                WImporte1 = Str$(rstPtoAdmiII!Importe1)
                                WImporte2 = Str$(rstPtoAdmiII!Importe2)
                            Case 3
                                WImporte1 = Str$(rstPtoAdmiII!Importe1)
                                WImporte2 = Str$(rstPtoAdmiII!Importe2)
                                WImporte3 = Str$(rstPtoAdmiII!Importe3)
                            Case 4
                                WImporte1 = Str$(rstPtoAdmiII!Importe1)
                                WImporte2 = Str$(rstPtoAdmiII!Importe2)
                                WImporte3 = Str$(rstPtoAdmiII!Importe3)
                                WImporte4 = Str$(rstPtoAdmiII!Importe4)
                            Case 5
                                WImporte1 = Str$(rstPtoAdmiII!Importe1)
                                WImporte2 = Str$(rstPtoAdmiII!Importe2)
                                WImporte3 = Str$(rstPtoAdmiII!Importe3)
                                WImporte4 = Str$(rstPtoAdmiII!Importe4)
                                WImporte5 = Str$(rstPtoAdmiII!Importe5)
                            Case 6
                                WImporte1 = Str$(rstPtoAdmiII!Importe1)
                                WImporte2 = Str$(rstPtoAdmiII!Importe2)
                                WImporte3 = Str$(rstPtoAdmiII!Importe3)
                                WImporte4 = Str$(rstPtoAdmiII!Importe4)
                                WImporte5 = Str$(rstPtoAdmiII!Importe5)
                                WImporte6 = Str$(rstPtoAdmiII!Importe6)
                            Case 7
                                WImporte1 = Str$(rstPtoAdmiII!Importe7)
                                WImporte7 = Str$(rstPtoAdmiII!Importe1 + rstPtoAdmiII!Importe2 + rstPtoAdmiII!Importe3 + rstPtoAdmiII!Importe4 + rstPtoAdmiII!Importe5 + rstPtoAdmiII!Importe6)
                            Case 8
                                WImporte1 = Str$(rstPtoAdmiII!Importe7)
                                WImporte2 = Str$(rstPtoAdmiII!Importe8)
                                WImporte7 = Str$(rstPtoAdmiII!Importe1 + rstPtoAdmiII!Importe2 + rstPtoAdmiII!Importe3 + rstPtoAdmiII!Importe4 + rstPtoAdmiII!Importe5 + rstPtoAdmiII!Importe6)
                            Case 9
                                WImporte1 = Str$(rstPtoAdmiII!Importe7)
                                WImporte2 = Str$(rstPtoAdmiII!Importe8)
                                WImporte3 = Str$(rstPtoAdmiII!Importe9)
                                WImporte7 = Str$(rstPtoAdmiII!Importe1 + rstPtoAdmiII!Importe2 + rstPtoAdmiII!Importe3 + rstPtoAdmiII!Importe4 + rstPtoAdmiII!Importe5 + rstPtoAdmiII!Importe6)
                            Case 10
                                WImporte1 = Str$(rstPtoAdmiII!Importe7)
                                WImporte2 = Str$(rstPtoAdmiII!Importe8)
                                WImporte3 = Str$(rstPtoAdmiII!Importe9)
                                WImporte4 = Str$(rstPtoAdmiII!Importe10)
                                WImporte7 = Str$(rstPtoAdmiII!Importe1 + rstPtoAdmiII!Importe2 + rstPtoAdmiII!Importe3 + rstPtoAdmiII!Importe4 + rstPtoAdmiII!Importe5 + rstPtoAdmiII!Importe6)
                            Case 11
                                WImporte1 = Str$(rstPtoAdmiII!Importe7)
                                WImporte2 = Str$(rstPtoAdmiII!Importe8)
                                WImporte3 = Str$(rstPtoAdmiII!Importe9)
                                WImporte4 = Str$(rstPtoAdmiII!Importe10)
                                WImporte5 = Str$(rstPtoAdmiII!Importe11)
                                WImporte7 = Str$(rstPtoAdmiII!Importe1 + rstPtoAdmiII!Importe2 + rstPtoAdmiII!Importe3 + rstPtoAdmiII!Importe4 + rstPtoAdmiII!Importe5 + rstPtoAdmiII!Importe6)
                            Case 12
                                WImporte1 = Str$(rstPtoAdmiII!Importe7)
                                WImporte2 = Str$(rstPtoAdmiII!Importe8)
                                WImporte3 = Str$(rstPtoAdmiII!Importe9)
                                WImporte4 = Str$(rstPtoAdmiII!Importe10)
                                WImporte5 = Str$(rstPtoAdmiII!Importe11)
                                WImporte6 = Str$(rstPtoAdmiII!Importe12)
                                WImporte7 = Str$(rstPtoAdmiII!Importe1 + rstPtoAdmiII!Importe2 + rstPtoAdmiII!Importe3 + rstPtoAdmiII!Importe4 + rstPtoAdmiII!Importe5 + rstPtoAdmiII!Importe6)
                            Case Else
                    End Select
                    
                    Vector(Renglon, 1) = WUnidad
                    Vector(Renglon, 2) = WConcepto
                    Vector(Renglon, 4) = WImporte1
                    Vector(Renglon, 5) = WImporte2
                    Vector(Renglon, 6) = WImporte3
                    Vector(Renglon, 7) = WImporte4
                    Vector(Renglon, 8) = WImporte5
                    Vector(Renglon, 9) = WImporte6
                    Vector(Renglon, 10) = WImporte7
                    
                End If
                
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstPtoAdmiII.Close
    End If
    
    For Da = 1 To Renglon
    
        WUnidad = Vector(Da, 1)
        WConcepto = Vector(Da, 2)
        WImporte1 = Vector(Da, 4)
        WImporte2 = Vector(Da, 5)
        WImporte3 = Vector(Da, 6)
        WImporte4 = Vector(Da, 7)
        WImporte5 = Vector(Da, 8)
        WImporte6 = Vector(Da, 9)
        WImporte7 = Vector(Da, 10)
        
        WRubro = "0"
        WSubRubro = "0"
        WCliente = "0"
        
        Rem spUnidad = "ConsultaUnidad " + "'" + WUnidad + "'"
        Rem Set rstUnidad = db.OpenRecordset(spUnidad, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstUnidad.RecordCount > 0 Then
        Rem     WSubRubro = Str$(rstUnidad!SubRubro)
        Rem     rstUnidad.Close
        Rem End If
        Rem
        Rem If Val(WSubRubro) <> 0 Then
        Rem     spSubRubro = "ConsultaSubRubros " + "'" + WSubRubro + "'"
        Rem     Set rstSubRubro = db.OpenRecordset(spSubRubro, dbOpenSnapshot, dbSQLPassThrough)
        Rem     If rstSubRubro.RecordCount > 0 Then
        Rem         WRubro = Str$(rstSubRubro!Rubro)
        Rem         rstSubRubro.Close
        Rem     End If
        Rem End If
        Rem
        Rem Call Ceros(WConcepto, 4)
        Rem Call Ceros(WUnidad, 4)
        Rem
        Rem WClave = WConcepto + WUnidad
        Rem
        Rem WMarca = 9
        Rem spConcepto = "ConsultaConcepto " + "'" + WClave + "'"
        Rem Set rstConcepto = db.OpenRecordset(spConcepto, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstConcepto.RecordCount > 0 Then
        Rem     WMarca = rstConcepto!Marca
        Rem End If
            
        Call Ceros(WRubro, 4)
        Call Ceros(WSubRubro, 4)
        Call Ceros(WUnidad, 4)
        Call Ceros(WConcepto, 4)
        Call Ceros(WCliente, 6)
            
        WClave = WRubro + WSubRubro + WUnidad + WConcepto + WCliente
        WImporte8 = ""
        WImporte9 = ""
        WImporte10 = ""
        WImporte11 = ""
        WImporte12 = ""
        WNota1 = ""
        WNota2 = ""
        WNota3 = ""
        WNota4 = ""
        WNota5 = ""
        WNota6 = ""
        WNota7 = ""
        WNota8 = ""
        WNota9 = ""
        WNota10 = ""
        WNota11 = ""
        WNota12 = ""
        
        XParam = "'" + WClave + "','" _
                    + WRubro + "','" _
                    + WSubRubro + "','" _
                    + WUnidad + "','" _
                    + WConcepto + "','" _
                    + WCliente + "','" _
                    + WImporte1 + "','" + WImporte2 + "','" + WImporte3 + "','" _
                    + WImporte4 + "','" + WImporte5 + "','" + WImporte6 + "','" _
                    + WImporte7 + "','" + WImporte8 + "','" + WImporte9 + "','" _
                    + WImporte10 + "','" + WImporte11 + "','" + WImporte12 + "','" _
                    + WImpo1 + "','" + WImpo2 + "','" + WImpo3 + "','" _
                    + WImpo4 + "','" + WImpo5 + "','" + WImpo6 + "','" _
                    + WImpo7 + "','" + WImpo8 + "','" + WImpo9 + "','" _
                    + WImpo10 + "','" + WImpo11 + "','" + WImpo12 + "','" _
                    + WNota1 + "','" + WNota2 + "','" + WNota3 + "','" _
                    + WNota4 + "','" + WNota5 + "','" + WNota6 + "','" _
                    + WNota7 + "','" + WNota8 + "','" + WNota9 + "','" _
                    + WNota10 + "','" + WNota11 + "','" + WNota12 + "'"
                         
        Set rstLista = db.OpenRecordset("AltaLista " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
    Next Da
    
    
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
            
                        WMes = Str$(rstIvacomp!Mes)
                        
                        If rstIvacomp!Ano = Val(Ano.Text) And Val(WMes) <= Val(Mes.Text) Then
                        
                            WProducto = Str$(rstIvacomp!Unidad)
                            WProveedor = Str$(rstIvacomp!Proveedor)
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
                            Vector1(Renglon1, 7) = WMes
                        
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
        WMes = Vector1(Cicla2, 7)
        
        Call Ceros(WProveedor, 11)
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 4)
        Call Ceros(WNumero, 8)
        
        For a = 1 To 20
        
            WTipomovi = "2"
            
            Auxi1 = Str$(a)
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
                    WGasto = Mid$(rstImputac!Cuenta, 6, 3)
                    WTarea = Mid$(rstImputac!Cuenta, 6, 3)
                    WGerencia = Mid$(rstImputac!Cuenta, 2, 1)
                    
                    WImporte1 = "0"
                    WImporte2 = "0"
                    WImporte3 = "0"
                    WImporte4 = "0"
                    WImporte5 = "0"
                    WImporte6 = "0"
                    WImporte7 = "0"
                    
                    If WDebito <> 0 Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = WProducto
                        Vector(Renglon, 2) = WGasto
                        Rem Vector(Renglon, 3) = Abs(WDebito)
                        Vector(Renglon, 5) = WTarea
                        Vector(Renglon, 6) = WGerencia
                        
                        WNeto = Abs(WDebito)
                        Select Case Val(WMes)
                            Case 1
                                If Semestre = 1 Then
                                    WImporte1 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 2
                                If Semestre = 1 Then
                                    WImporte2 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 3
                                If Semestre = 1 Then
                                    WImporte3 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 4
                                If Semestre = 1 Then
                                    WImporte4 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 5
                                If Semestre = 1 Then
                                    WImporte5 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 6
                                If Semestre = 1 Then
                                    WImporte6 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 7
                                WImporte1 = Str$(WNeto)
                            Case 8
                                WImporte2 = Str$(WNeto)
                            Case 9
                                WImporte3 = Str$(WNeto)
                            Case 10
                                WImporte4 = Str$(WNeto)
                            Case 11
                                WImporte5 = Str$(WNeto)
                            Case 12
                                WImporte6 = Str$(WNeto)
                            Case Else
                        End Select
                        Vector(Renglon, 7) = WImporte1
                        Vector(Renglon, 8) = WImporte2
                        Vector(Renglon, 9) = WImporte3
                        Vector(Renglon, 10) = WImporte4
                        Vector(Renglon, 11) = WImporte5
                        Vector(Renglon, 12) = WImporte6
                        Vector(Renglon, 13) = WImporte7
                    End If
                    
                    WImporte1 = "0"
                    WImporte2 = "0"
                    WImporte3 = "0"
                    WImporte4 = "0"
                    WImporte5 = "0"
                    WImporte6 = "0"
                    WImporte7 = "0"
                    
                    If WCredito <> 0 Then
                        Renglon = Renglon + 1
                        Vector(Renglon, 1) = WProducto
                        Vector(Renglon, 2) = WGasto
                        Rem Vector(Renglon, 3) = Abs(WCredito) * -1
                        Vector(Renglon, 5) = WTarea
                        Vector(Renglon, 6) = WGerencia
                        
                        WNeto = Abs(WCredito) * -1
                        Select Case Val(WMes)
                            Case 1
                                If Semestre = 1 Then
                                    WImporte1 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 2
                                If Semestre = 1 Then
                                    WImporte2 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 3
                                If Semestre = 1 Then
                                    WImporte3 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 4
                                If Semestre = 1 Then
                                    WImporte4 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 5
                                If Semestre = 1 Then
                                    WImporte5 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 6
                                If Semestre = 1 Then
                                    WImporte6 = Str$(WNeto)
                                        Else
                                    WImporte7 = Str$(WNeto)
                                End If
                            Case 7
                                WImporte1 = Str$(WNeto)
                            Case 8
                                WImporte2 = Str$(WNeto)
                            Case 9
                                WImporte3 = Str$(WNeto)
                            Case 10
                                WImporte4 = Str$(WNeto)
                            Case 11
                                WImporte5 = Str$(WNeto)
                            Case 12
                                WImporte6 = Str$(WNeto)
                            Case Else
                        End Select
                        Vector(Renglon, 7) = WImporte1
                        Vector(Renglon, 8) = WImporte2
                        Vector(Renglon, 9) = WImporte3
                        Vector(Renglon, 10) = WImporte4
                        Vector(Renglon, 11) = WImporte5
                        Vector(Renglon, 12) = WImporte6
                        Vector(Renglon, 13) = WImporte7
                    End If
                    
                End If
                
            End If
        Next a
            
    Next Cicla2
        
    For Cicla3 = 1 To Renglon
    
        WUnidad = Vector(Cicla3, 1)
        WConcepto = Vector(Cicla3, 2)
        WTarea = Vector(Cicla3, 5)
        WGerencia = Vector(Cicla3, 6)
        Rem WImpo = Vector(Cicla3, 3)
        WRubro = "0"
        WSubRubro = "0"
        WCliente = "0"
        WImpo1 = Vector(Cicla3, 7)
        WImpo2 = Vector(Cicla3, 8)
        WImpo3 = Vector(Cicla3, 9)
        WImpo4 = Vector(Cicla3, 10)
        WImpo5 = Vector(Cicla3, 11)
        WImpo6 = Vector(Cicla3, 12)
        WImpo7 = Vector(Cicla3, 13)
        
        Call Ceros(WRubro, 4)
        Call Ceros(WSubRubro, 4)
        Call Ceros(WUnidad, 4)
        Call Ceros(WConcepto, 4)
        Call Ceros(WCliente, 6)
            
        WClave = WRubro + WSubRubro + WUnidad + WConcepto + WCliente
        
        spLista = "ConsultaLista " + "'" + WClave + "'"
        Set rstLista = db.OpenRecordset(spLista, dbOpenSnapshot, dbSQLPassThrough)
        If rstLista.RecordCount > 0 Then
        
            WImpo1 = Str$(rstLista!Impo1 + WImpo1)
            WImpo2 = Str$(rstLista!Impo2 + WImpo2)
            WImpo3 = Str$(rstLista!Impo3 + WImpo3)
            WImpo4 = Str$(rstLista!Impo4 + WImpo4)
            WImpo5 = Str$(rstLista!Impo5 + WImpo5)
            WImpo6 = Str$(rstLista!Impo6 + WImpo6)
            WImpo7 = Str$(rstLista!Impo7 + WImpo7)
            XParam = "'" + WClave + "','" _
                        + WImpo1 + "','" _
                        + WImpo2 + "','" _
                        + WImpo3 + "','" _
                        + WImpo4 + "','" _
                        + WImpo5 + "','" _
                        + WImpo6 + "','" _
                        + WImpo7 + "'"
                    
            Set rstLista = db.OpenRecordset("ModificaListaII " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Else
                                               
            WImporte1 = "0"
            WImporte2 = "0"
            WImporte3 = "0"
            WImporte4 = "0"
            WImporte5 = "0"
            WImporte6 = "0"
            WImporte7 = "0"
            WImporte8 = "0"
            WImporte9 = "0"
            WImporte10 = "0"
            WImporte11 = "0"
            WImporte12 = "0"
            WNota1 = "0"
            WNota2 = "0"
            WNota3 = "0"
            WNota4 = "0"
            WNota5 = "0"
            WNota6 = "0"
            WNota7 = "0"
            WNota8 = "0"
            WNota9 = "0"
            WNota10 = "0"
            WNota11 = "0"
            WNota12 = "0"
            WImpo1 = Str$(WImpo1)
            WImpo2 = Str$(WImpo2)
            WImpo3 = Str$(WImpo3)
            WImpo4 = Str$(WImpo4)
            WImpo5 = Str$(WImpo5)
            WImpo6 = Str$(WImpo6)
            WImpo7 = Str$(WImpo7)
            WImpo8 = "0"
            WImpo9 = "0"
            WImpo10 = "0"
            WImpo11 = "0"
            WImpo12 = "0"

            XParam = "'" + WClave + "','" _
                            + WRubro + "','" _
                            + WSubRubro + "','" _
                            + WUnidad + "','" _
                            + WConcepto + "','" _
                            + WCliente + "','" _
                            + WImporte1 + "','" + WImporte2 + "','" + WImporte3 + "','" _
                            + WImporte4 + "','" + WImporte5 + "','" + WImporte6 + "','" _
                            + WImporte7 + "','" + WImporte8 + "','" + WImporte9 + "','" _
                            + WImporte10 + "','" + WImporte11 + "','" + WImporte12 + "','" _
                            + WImpo1 + "','" + WImpo2 + "','" + WImpo3 + "','" _
                            + WImpo4 + "','" + WImpo5 + "','" + WImpo6 + "','" _
                            + WImpo7 + "','" + WImpo8 + "','" + WImpo9 + "','" _
                            + WImpo10 + "','" + WImpo11 + "','" + WImpo12 + "','" _
                            + WNota1 + "','" + WNota2 + "','" + WNota3 + "','" _
                            + WNota4 + "','" + WNota5 + "','" + WNota6 + "','" _
                            + WNota7 + "','" + WNota8 + "','" + WNota9 + "','" _
                            + WNota10 + "','" + WNota11 + "','" + WNota12 + "'"
                                
            Set rstLista = db.OpenRecordset("AltaLista " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
    Next Cicla3
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Lista.Unidad, Lista.Confac, Lista.Importe1, Lista.Importe2, Lista.Importe3, Lista.Importe4, Lista.Importe5, Lista.Importe6, Lista.Importe7, Lista.Impo1, Lista.Impo2, Lista.Impo3, Lista.Impo4, Lista.Impo5, Lista.Impo6, Lista.Impo7, " _
                        + "Unidad.Nombre, " _
                        + "Conpto.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.Lista Lista, " _
                        + DSQ + ".dbo.Unidad Unidad, " _
                        + DSQ + ".dbo.Conpto Conpto " _
                        + "Where " _
                        + "Lista.Unidad = Unidad.Codigo AND " _
                        + "Lista.Confac = Conpto.Codigo AND " _
                        + "Lista.Unidad >= 0 AND Lista.Unidad <= 9999"
    
    Rem Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "comparaiv.rpt"
            Else
        Listado.ReportFileName = "Sumaiv.rpt"
    End If
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    
    Desde.SetFocus
    PrgComparaIV.Hide
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

    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0

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

     Opcion.AddItem "Unidades de Negocios"
     Rem Opcion.AddItem "Rubros"
     Rem Opcion.AddItem "Clientes"

     Opcion.Visible = True
     
     Rem Opcion.ListIndex = 0
     
     Rem Call Opcion_Click
     
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
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
            
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
            spUnidad = "ConsultaUnidad " + "'" + WUnidad + "'"
            Set rstUnidad = db.OpenRecordset(spUnidad, dbOpenSnapshot, dbSQLPassThrough)
            If rstUnidad.RecordCount > 0 Then
                Desde.Text = rstUnidad!Codigo
                Hasta.Text = rstUnidad!Codigo
                rstRubro.Close
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
    
    Select Case XIndice
            
        Case 0
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
            
        Case Else
    End Select
    End If

End Sub




