VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPosicion 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Posicion de Iva"
   ClientHeight    =   3000
   ClientLeft      =   3165
   ClientTop       =   1425
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3000
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   4455
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
         Left            =   480
         MouseIcon       =   "Posicion.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Posicion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1320
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
         Left            =   1800
         MouseIcon       =   "Posicion.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "Posicion.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1320
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
         Left            =   3120
         MouseIcon       =   "Posicion.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "Posicion.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salida"
         Top             =   1320
         Width           =   855
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   2400
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
         Left            =   2400
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
         Left            =   600
         TabIndex        =   5
         Top             =   720
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
         Left            =   600
         TabIndex        =   4
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
      ReportFileName  =   "Posicion.rpt"
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
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgPosicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZSuma1 As Double
Dim ZSuma2 As Double
Dim ZSuma3 As Double
Dim ZSuma4 As Double
Dim ZSuma5 As Double
Dim ZSuma6 As Double
Dim ZSuma7 As Double
Dim ZSuma8 As Double

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


    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    WTitulo = "Del " + Desde.Text + " al " + Hasta.Text
    
    ZSql = ""
    ZSql = ZSql + "DELETE Posicion"
    spPosicion = ZSql
    Set rstPosicion = db.OpenRecordset(spPosicion, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSuma1 = 0
    ZSuma2 = 0
    ZSuma3 = 0
    ZSuma4 = 0
    ZSuma5 = 0
    ZSuma6 = 0
    ZSuma7 = 0
    ZSuma8 = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and CtaCte.OrdFecha <= " + "'" + WHasta + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        With rstCtaCte
            .MoveFirst
            Do
                If Val(rstCtaCte!Tipo) <= 5 Then
                
                    WNeto = rstCtaCte!Neto
                    WIva1 = rstCtaCte!Iva1
                    WIva2 = rstCtaCte!Iva2
                
                    If WIva1 <> 0 Then
                        WSuma1 = WSuma1 + WNeto + WIva1
                        WSuma2 = WSuma2 + WNeto
                        WSuma3 = WSuma3 + WIva1
                            Else
                        WSuma1 = WSuma1 + WNeto + WIva1
                        WSuma3 = WSuma3 + WIva1
                        WSuma8 = WSuma8 + WNeto
                    End If
                    
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstCtaCte.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "INSERT INTO Posicion ("
    ZSql = ZSql + "Codigo ,"
    ZSql = ZSql + "Descripcion ,"
    ZSql = ZSql + "Importe1 ,"
    ZSql = ZSql + "Importe2 ,"
    ZSql = ZSql + "Importe3 ,"
    ZSql = ZSql + "Importe4 ,"
    ZSql = ZSql + "Importe5 ,"
    ZSql = ZSql + "Importe6 ,"
    ZSql = ZSql + "Importe7 ,"
    ZSql = ZSql + "Importe8 ,"
    ZSql = ZSql + "DesEmpresa ,"
    ZSql = ZSql + "Periodo )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + "1" + "',"
    ZSql = ZSql + "'" + "Ventas" + "',"
    ZSql = ZSql + "'" + Str$(WSuma1) + "',"
    ZSql = ZSql + "'" + Str$(WSuma2) + "',"
    ZSql = ZSql + "'" + Str$(WSuma3) + "',"
    ZSql = ZSql + "'" + Str$(WSuma4) + "',"
    ZSql = ZSql + "'" + Str$(WSuma5) + "',"
    ZSql = ZSql + "'" + Str$(WSuma6) + "',"
    ZSql = ZSql + "'" + Str$(WSuma7) + "',"
    ZSql = ZSql + "'" + Str$(WSuma8) + "',"
    ZSql = ZSql + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + "'" + WTitulo + "')"
    spPosicion = ZSql
    Set rstPosicion = db.OpenRecordset(spPosicion, dbOpenSnapshot, dbSQLPassThrough)






    ZSuma1 = 0
    ZSuma2 = 0
    ZSuma3 = 0
    ZSuma4 = 0
    ZSuma5 = 0
    ZSuma6 = 0
    ZSuma7 = 0
    ZSuma8 = 0


    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    ZSql = ZSql + " Where IvaComp.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and IvaComp.OrdFecha <= " + "'" + WHasta + "'"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
        With rstIvaComp
            .MoveFirst
            Do
                If Val(rstIvaComp!Tipo) <= 5 Or Val(rstIvaComp!Tipo) = 7 Then
                
                    If rstIvaComp!Letra <> "X" Then
                
                        WNeto = rstIvaComp!Neto
                        WIva1 = rstIvaComp!Iva21
                        WIva2 = rstIvaComp!Iva5
                        WIva3 = rstIvaComp!Iva27
                        WIva4 = rstIvaComp!Ib + rstIvaComp!ImpInterno + rstIvaComp!ImpCombustible
                        WIva5 = rstIvaComp!Iva105
                        WExento = rstIvaComp!Exento
                        
                        WSuma1 = WSuma1 + ((WNeto + WIva1 + WIva2 + WIva3 + WIva4 + WIva5 + WExento) * -1)
                        WSuma2 = WSuma2 + (WNeto * -1)
                        WSuma3 = WSuma3 + (WIva1 * -1)
                        WSuma4 = WSuma4 + (WIva2 * -1)
                        WSuma5 = WSuma5 + (WIva3 * -1)
                        WSuma6 = WSuma6 + (WIva4 * -1)
                        WSuma7 = WSuma7 + (WIva5 * -1)
                        WSuma8 = WSuma8 + (WExento * -1)
                        
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstIvaComp.Close
    End If
        
        
    
    
    ZSql = ""
    ZSql = ZSql + "INSERT INTO Posicion ("
    ZSql = ZSql + "Codigo ,"
    ZSql = ZSql + "Descripcion ,"
    ZSql = ZSql + "Importe1 ,"
    ZSql = ZSql + "Importe2 ,"
    ZSql = ZSql + "Importe3 ,"
    ZSql = ZSql + "Importe4 ,"
    ZSql = ZSql + "Importe5 ,"
    ZSql = ZSql + "Importe6 ,"
    ZSql = ZSql + "Importe7 ,"
    ZSql = ZSql + "Importe8 ,"
    ZSql = ZSql + "DesEmpresa ,"
    ZSql = ZSql + "Periodo )"
    ZSql = ZSql + "Values ("
    ZSql = ZSql + "'" + "2" + "',"
    ZSql = ZSql + "'" + "Compras" + "',"
    ZSql = ZSql + "'" + Str$(WSuma1) + "',"
    ZSql = ZSql + "'" + Str$(WSuma2) + "',"
    ZSql = ZSql + "'" + Str$(WSuma3) + "',"
    ZSql = ZSql + "'" + Str$(WSuma4) + "',"
    ZSql = ZSql + "'" + Str$(WSuma5) + "',"
    ZSql = ZSql + "'" + Str$(WSuma6) + "',"
    ZSql = ZSql + "'" + Str$(WSuma7) + "',"
    ZSql = ZSql + "'" + Str$(WSuma8) + "',"
    ZSql = ZSql + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + "'" + WTitulo + "')"
    spPosicion = ZSql
    Set rstPosicion = db.OpenRecordset(spPosicion, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    

    Listado.WindowTitle = "Posicion de Iva"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Posicion.Codigo, Posicion.Descripcion, Posicion.Importe1, Posicion.Importe2, Posicion.Importe3, Posicion.Importe4, Posicion.Importe5, Posicion.Importe6, Posicion.Importe7, Posicion.Importe8, Posicion.DesEmpresa, Posicion.Periodo " _
                + "From " _
                + DSQ + ".dbo.Posicion Posicion " _
                + "Where " _
                + "Posicion.Codigo >= 0 AND " _
                + "Posicion.Codigo <= 999999"
    
    Listado.Connect = Connect()
    
    Uno = "{Posicion.Codigo} in 0 to 999999"
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgPosicion.Hide
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
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
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




