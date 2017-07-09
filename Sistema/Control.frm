VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgControl 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Control de Chequeras"
   ClientHeight    =   3045
   ClientLeft      =   3435
   ClientTop       =   1350
   ClientWidth     =   5130
   LinkTopic       =   "Form2"
   ScaleHeight     =   3045
   ScaleWidth      =   5130
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
      Height          =   2655
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   3735
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
         Left            =   360
         MouseIcon       =   "Control.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Control.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1440
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
         Left            =   1560
         MouseIcon       =   "Control.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "Control.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1440
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
         Left            =   2760
         MouseIcon       =   "Control.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "Control.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salida"
         Top             =   1440
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
      Begin MSMask.MaskEdBox Desdefec 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
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
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4440
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Control.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Valores en Cartera"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Clave1 As String
Private Clave2 As String
Private Clave As String
Private WAuxiliar As String
Private WLinea As Single
Private WRecibo As String
Private WCheque As String
Private WBanco As String
Private WFecha As String
Private WImporte As Double
Private WFechaImpre As String
Private WVencimiento As String
Private Baja As Integer
Dim XParam As String
Dim XCliente As String
Dim WVenci As String

Dim ZVector(1000, 10) As String

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
    ZSql = ZSql + "DELETE Control"
    spControl = ZSql
    Set rstControl = db.OpenRecordset(spControl, dbOpenSnapshot, dbSQLPassThrough)

    WDesdefecha = Right$(Desdefec.Text, 4) + Mid$(Desdefec.Text, 4, 2) + Left$(Desdefec.Text, 2)
    WHastafecha = Right$(Hastafec.Text, 4) + Mid$(Hastafec.Text, 4, 2) + Left$(Hastafec.Text, 2)
    
    WAno = Right$(Desdefec.Text, 4)
    WMes = Mid$(Desdefec.Text, 4, 2)
    WDia = Left$(Desdefec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hastafec.Text, 4)
    WMes = Mid$(Hastafec.Text, 4, 2)
    WDia = Left$(Hastafec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WTitulo = "Del " + Desdefec.Text + " al " + Hastafec.Text
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Actividad = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Chequera"
    ZSql = ZSql + " Where Chequera.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Chequera.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " Order by Banco, OrdFecha"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
    
        With rstChequera
            .MoveFirst
            Do
                ZLugar = ZLugar + 1
                ZVector(ZLugar, 1) = !Banco
                ZVector(ZLugar, 2) = !Desde
                ZVector(ZLugar, 3) = !Hasta
                ZVector(ZLugar, 4) = !Fecha
                ZVector(ZLugar, 5) = !ordfecha
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstChequera.Close
    End If
    
    Corte = 0
    
    For ZCiclo = 1 To ZLugar
    
        WBanco = ZVector(ZCiclo, 1)
        WDesde = ZVector(ZCiclo, 2)
        WHasta = ZVector(ZCiclo, 3)
        WFecha = ZVector(ZCiclo, 4)
        WOrdFecha = ZVector(ZCiclo, 5)
        WCodigoEmpresa = "1"
        
        Corte = Corte + 1
        
        For Ciclo = Val(WDesde) To Val(WHasta)
        
            Clave1 = WBanco
            Clave2 = Ciclo
            Call Ceros(Clave1, 4)
            Call Ceros(Clave2, 8)
            Clave = Clave1 + Clave2
            
            Graba = "S"
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pagos"
            ZSql = ZSql + " Where Pagos.TipoReg = '2'"
            ZSql = ZSql + " and Pagos.Tipo2 = '02'"
            ZSql = ZSql + " and Pagos.Numero2 = '" + Clave2 + "'"
            ZSql = ZSql + " and Pagos.Banco2 = '" + WBanco + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            If rstPagos.RecordCount > 0 Then
                Graba = "N"
                rstPagos.Close
            End If
            
            If Graba = "S" Then
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Control ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Banco ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Corte ,"
                ZSql = ZSql + "Desde ,"
                ZSql = ZSql + "Hasta ,"
                ZSql = ZSql + "CodigoEmpresa )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + Clave + "',"
                ZSql = ZSql + "'" + WBanco + "',"
                ZSql = ZSql + "'" + WFecha + "',"
                ZSql = ZSql + "'" + WOrdFecha + "',"
                ZSql = ZSql + "'" + Clave2 + "',"
                ZSql = ZSql + "'" + Str$(Corte) + "',"
                ZSql = ZSql + "'" + WDesde + "',"
                ZSql = ZSql + "'" + WHasta + "',"
                ZSql = ZSql + "'" + WCodigoEmpresa + "')"
                spControl = ZSql
                Set rstControl = db.OpenRecordset(spControl, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
        Next Ciclo
        
    Next ZCiclo
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Control.Banco, Control.Fecha, Control.Numero, Control.Corte, Control.Desde, Control.Hasta, " _
            + "Banco.Nombre, " _
            + "Auxiliar.Nombre " _
            + "From " _
            + DSQ + ".dbo.Control Control, " _
            + DSQ + ".dbo.Banco Banco, " _
            + DSQ + ".dbo.Auxiliar Auxiliar " _
            + "Where " _
            + "Control.Banco = Banco.Banco AND " _
            + "Control.CodigoEmpresa = Auxiliar.Empresa AND " _
            + "Control.Banco >= 0 AND " _
            + "Control.Banco <= 9999"
    
    Listado.Connect = Connect()
    
    Rem Uno = "{IvaComp.Letra} <> " + Chr$(34) + "X" + Chr$(34)
    Rem Dos = " and {IvaComp.OrdPeriodo} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    
    Rem Listado.GroupSelectionFormula = Uno + Dos
    Rem Listado.SelectionFormula = Uno + Dos
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgControl.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desdefec.Text, Auxi)
        If Auxi = "S" Then
            Hastafec.SetFocus
                Else
            Desdefec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Desdefec.Text = "  /  /    "
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hastafec.Text, Auxi)
        If Auxi = "S" Then
            Desdefec.SetFocus
                Else
            Hastafec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hastafec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Desdefec.Text = "  /  /    "
    Hastafec.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub DesdeFec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 120 Or KeyCode = 121 Then
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
            Call Cancela_Click
        Case Else
    End Select
End Sub







