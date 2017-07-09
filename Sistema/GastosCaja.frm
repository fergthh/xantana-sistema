VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGastosCaja 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Gastos"
   ClientHeight    =   6810
   ClientLeft      =   1005
   ClientTop       =   420
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   ScaleHeight     =   6810
   ScaleWidth      =   9900
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   1440
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9495
      Begin VB.TextBox Comprobante 
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
         MaxLength       =   20
         TabIndex        =   19
         Text            =   " "
         Top             =   1440
         Width           =   1695
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpia F3"
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
         Left            =   3000
         MouseIcon       =   "GastosCaja.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "GastosCaja.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Limpia la pantalla"
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Proceso 
         Caption         =   "Graba"
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
         MouseIcon       =   "GastosCaja.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "GastosCaja.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Concepto 
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
         TabIndex        =   15
         Text            =   " "
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Codigo 
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
         MaxLength       =   10
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Importe 
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
         MaxLength       =   20
         TabIndex        =   9
         Text            =   " "
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Observaciones 
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
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   1800
         Width           =   5535
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
         Left            =   4680
         MouseIcon       =   "GastosCaja.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "GastosCaja.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Consulta de Datos"
         Top             =   2640
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
         Left            =   6360
         MouseIcon       =   "GastosCaja.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "GastosCaja.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salida"
         Top             =   2640
         Width           =   855
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   5040
         TabIndex        =   13
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
      Begin VB.Label Label2 
         Caption         =   "Comprobante"
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
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label DesConcepto 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   3000
         TabIndex        =   18
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label12 
         Caption         =   "Fecha"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Numero Movimiento"
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
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Concepto"
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Importe"
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
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Observaciones"
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
         TabIndex        =   8
         Top             =   1800
         Width           =   1575
      End
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
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   7815
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ventclie.rpt"
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
      Left            =   6840
      TabIndex        =   2
      Top             =   5160
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
      ItemData        =   "GastosCaja.frx":2D30
      Left            =   120
      List            =   "GastosCaja.frx":2D37
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgGastosCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZProvincia(100) As String

Private Sub CmdLimpiar_Click()

    Codigo.Text = "1"
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Concepto.Text = ""
    DesConcepto.Caption = ""
    Importe.Text = ""
    Comprobante.Text = ""
    Observaciones.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM GastosCaja"
    spGastosCaja = ZSql
    Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
    If rstGastosCaja.RecordCount > 0 Then
        rstGastosCaja.MoveLast
        ZUltimo = IIf(IsNull(rstGastosCaja!CodigoMayor), "0", rstGastosCaja!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstGastosCaja.Close
    End If
    
    Codigo.SetFocus

End Sub

Private Sub Proceso_Click()

    ZZOrdFecha = ""
    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM GastosCaja"
    ZSql = ZSql + " Where GastosCaja.Codigo = " + "'" + Codigo.Text + "'"
    spGastosCaja = ZSql
    Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
    If rstGastosCaja.RecordCount > 0 Then
    
        rstGastosCaja.Close
    
        ZSql = ""
        ZSql = ZSql + "UPDATE GastosCaja SET "
        ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
        ZSql = ZSql + " OrdFecha = " + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + " Concepto = " + "'" + Concepto.Text + "',"
        ZSql = ZSql + " Importe = " + "'" + Importe.Text + "',"
        ZSql = ZSql + " Comprobante = " + "'" + Comprobante.Text + "',"
        ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        
        spGastosCaja = ZSql
        Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
    
            Else

        ZSql = ""
        ZSql = ZSql + "INSERT INTO GastosCaja ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Concepto ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Comprobante ,"
        ZSql = ZSql + "Observaciones )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + "'" + Concepto.Text + "',"
        ZSql = ZSql + "'" + Importe.Text + "',"
        ZSql = ZSql + "'" + Comprobante.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "')"
                                
        spGastosCaja = ZSql
        Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    Call CmdLimpiar_Click
    
End Sub

Private Sub Cancela_Click()
    PrgGastosCaja.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM GastosCaja"
            ZSql = ZSql + " Where GastosCaja.Codigo = " + "'" + Codigo.Text + "'"
            spGastosCaja = ZSql
            Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
            If rstGastosCaja.RecordCount > 0 Then
                Fecha.Text = rstGastosCaja!Fecha
                Concepto.Text = rstGastosCaja!Concepto
                Importe.Text = Str$(rstGastosCaja!Importe)
                Comprobante.Text = rstGastosCaja!Comprobante
                Observaciones.Text = rstGastosCaja!Observaciones
                rstGastosCaja.Close
            End If
            Fecha.SetFocus
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                DesConcepto.Caption = rstConceptos!Nombre
                rstConceptos.Close
                    Else
                DesConcepto.Caption = ""
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha1(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Concepto.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Concepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Conceptos"
        ZSql = ZSql + " Where Conceptos.Concepto = " + "'" + Concepto.Text + "'"
        spConceptos = ZSql
        Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
        If rstConceptos.RecordCount > 0 Then
            DesConcepto.Caption = rstConceptos!Nombre
            rstConceptos.Close
            Importe.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Concepto.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Importe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comprobante.SetFocus
    End If
    If KeyAscii = 27 Then
        Importe.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Comprobante_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Comprobante.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub




Sub Form_Load()

    Codigo.Text = "1"
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Concepto.Text = ""
    DesConcepto.Caption = ""
    Importe.Text = ""
    Comprobante.Text = ""
    Observaciones.Text = ""
 
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM GastosCaja"
    spGastosCaja = ZSql
    Set rstGastosCaja = db.OpenRecordset(spGastosCaja, dbOpenSnapshot, dbSQLPassThrough)
    If rstGastosCaja.RecordCount > 0 Then
        rstGastosCaja.MoveLast
        ZUltimo = IIf(IsNull(rstGastosCaja!CodigoMayor), "0", rstGastosCaja!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstGastosCaja.Close
    End If
    
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Conceptos"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
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
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Concepto.Text = WIndice.List(Indice)
            Call Concepto_KeyPress(13)
            
        Case Else
    End Select
    
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
            
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Concepto_DblClick()

    Opcion.Clear
    Opcion.AddItem "Concepto"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub


Private Sub Concepto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub


Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Importe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Comprobante_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
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
        Case 115
            Call Consulta_Click
        Case 121
            Call Cancela_Click
        Case Else
    End Select
End Sub










