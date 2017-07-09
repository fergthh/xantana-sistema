VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgListaFormulaCosto 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Formulas (Costo)"
   ClientHeight    =   3585
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3585
   ScaleWidth      =   8145
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
      Height          =   3135
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   6135
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
         Left            =   2520
         TabIndex        =   10
         Top             =   1320
         Width           =   2175
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
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   1215
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
         Left            =   2520
         MaxLength       =   8
         TabIndex        =   1
         Text            =   " "
         Top             =   720
         Width           =   1215
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
         Left            =   4440
         MouseIcon       =   "ListaFormulaCosto.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ListaFormulaCosto.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salida"
         Top             =   1920
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
         Left            =   2640
         MouseIcon       =   "ListaFormulaCosto.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ListaFormulaCosto.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1920
         Width           =   855
      End
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
         MouseIcon       =   "ListaFormulaCosto.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ListaFormulaCosto.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
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
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   360
         Width           =   1575
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   720
         Width           =   1935
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7920
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PrecioFam.rpt"
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
      Left            =   7440
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaFormulaCosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Producto As String
Private Costo As Double
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

    ZSql = ""
    ZSql = ZSql + "UPDATE Formula SET "
    ZSql = ZSql + " Lista = " + "'" + "S" + "'"
    ZSql = ZSql + " Where Formula.Articulo >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Formula.Articulo <= " + "'" + Hasta.Text + "'"
    spFormula = ZSql
    Set rstFormula = db.OpenRecordset(spFormula, dbOpenSnapshot, dbSQLPassThrough)


    Erase ZVector
    ZLugar = 0
    ZPasa = 0
    ZSuma = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Formula"
    ZSql = ZSql + " Where Formula.Articulo >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Formula.Articulo <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and Formula.Proveedor <> 0"
    ZSql = ZSql + " Order by Formula.Articulo, Formula.Color, Formula.Insumo"
    spFormula = ZSql
    Set rstFormula = db.OpenRecordset(spFormula, dbOpenSnapshot, dbSQLPassThrough)
    If rstFormula.RecordCount > 0 Then
        With rstFormula
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If rstFormula!Proveedor <> 12 Then
                
                        If ZPasa = 0 Then
                            ZCorteI = rstFormula!Articulo
                            ZCorteII = rstFormula!Color
                            ZCorteIII = rstFormula!Insumo
                            ZSuma = 0
                            ZPasa = 1
                        End If
                        
                        If ZCorteI <> rstFormula!Articulo Or ZCorteII <> rstFormula!Color Then
                            ZCorteI = rstFormula!Articulo
                            ZCorteII = rstFormula!Color
                            ZCorteIII = rstFormula!Insumo
                            ZSuma = 0
                        End If
                        
                        ZSuma = ZSuma + 1
                        If ZSuma > 1 Then
                            ZLugar = ZLugar + 1
                            ZVector(ZLugar, 1) = rstFormula!Clave
                        End If
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFormula.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZClave = ZVector(Ciclo, 1)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Formula SET "
        ZSql = ZSql + " Lista = " + "'" + "N" + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
        spFormula = ZSql
        Set rstFormula = db.OpenRecordset(spFormula, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo

    Listado.WindowTitle = "Listado de Costos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.SQLQuery = "SELECT Formula.Articulo, Formula.Color, Formula.Renglon, Formula.Insumo, Formula.Proveedor, Formula.Cantidad, Formula.CantidadII, Formula.Base, Formula.Corte, Formula.Lista,  " _
            + "Articulo.Descripcion, " _
            + "Insumo.Descripcion, Insumo.Costo " _
            + "From " _
            + DSQ + ".dbo.Formula Formula, " _
            + DSQ + ".dbo.Articulo Articulo, " _
            + DSQ + ".dbo.Insumo Insumo " _
            + "Where " _
            + "Formula.Articulo = Articulo.Codigo AND " _
            + "Formula.Insumo = Insumo.Codigo AND " _
            + "Formula.Articulo >= '" + Desde.Text + "' AND " _
            + "Formula.Articulo <= '" + Hasta.Text + "' AND " _
            + "Formula.Lista = '" + "S" + "'"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
            
    Uno = "{Formula.Articulo} in " + Chr$(34) + Desde.Text + Chr(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Dos = " and {Formula.Lista} = " + Chr$(34) + "S" + Chr(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
        
    Select Case Tipo.ListIndex
        Case 0
            Listado.ReportFileName = "ListaFormulaCosto.rpt"
        Case 1
            Listado.ReportFileName = "ListaFormulaCostoResu.rpt"
        Case Else
            Listado.ReportFileName = "ListaFormulaCostoPromedi.rpt"
    End Select
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaStockInsumos.Hide
    Unload Me
    Menu4.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Desde.Text) <> "" Then
            ZZVeri = UCase(Left$(Desde.Text, 1))
            If ZZVeri < "A" Or ZZVeri > "Z" Then
                ZZVeri = Left$(Desde.Text, 1)
                Desde.Text = ZZVeri + Desde.Text
            End If
            Auxi = UCase(Left$(Desde.Text, 1))
            Auxi1 = Mid$(Desde.Text, 2, 5)
            Call Ceros(Auxi1, 5)
            Desde.Text = Auxi + Auxi1
            Hasta.Text = Auxi + Auxi1
        End If
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(Hasta.Text) <> "" Then
            ZZVeri = UCase(Left$(Hasta.Text, 1))
            If ZZVeri < "A" Or ZZVeri > "Z" Then
                ZZVeri = Left$(Hasta.Text, 1)
                Hasta.Text = ZZVeri + Hasta.Text
            End If
            Auxi = UCase(Left$(Hasta.Text, 1))
            Auxi1 = Mid$(Hasta.Text, 2, 5)
            Call Ceros(Auxi1, 5)
            Rem Desde.Text = Auxi + Auxi1
            Hasta.Text = Auxi + Auxi1
        End If
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Abierto"
    Tipo.AddItem "Resumido"
    Tipo.AddItem "Promedio"
    
    Tipo.ListIndex = 0

    Desde.Text = ""
    Hasta.Text = ""
    Frame2.Visible = True
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
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub
















