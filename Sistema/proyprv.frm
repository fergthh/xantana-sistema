VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProyPrv 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Proyeccion de Pagos"
   ClientHeight    =   8145
   ClientLeft      =   2910
   ClientTop       =   525
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   8145
   ScaleWidth      =   6135
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
      TabIndex        =   12
      Top             =   5040
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   4695
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
         Left            =   3600
         MouseIcon       =   "proyprv.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "proyprv.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Consulta de Datos"
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Panta 
         Caption         =   "Panta F1"
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
         Left            =   3600
         MouseIcon       =   "proyprv.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "proyprv.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   360
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
         Left            =   3600
         MouseIcon       =   "proyprv.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "proyprv.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Left            =   3600
         MouseIcon       =   "proyprv.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "proyprv.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salida"
         Top             =   3600
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
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   1
         Text            =   " "
         Top             =   840
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
         Left            =   1920
         MaxLength       =   8
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Vence4 
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence3 
         Height          =   375
         Left            =   1080
         TabIndex        =   6
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence2 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence1 
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Parametros de Fechas"
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
         Left            =   720
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5640
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ProyPrv.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Saldos de Cuenta Corriente de Proveedores"
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
      Top             =   2640
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
      Height          =   2595
      ItemData        =   "proyprv.frx":2D30
      Left            =   120
      List            =   "proyprv.frx":2D37
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   5895
   End
End
Attribute VB_Name = "PrgProyPrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WTrabajo(10000, 10) As String
Dim Pasa As Integer
Dim Lugar As Integer
Dim Impo1 As Double
Dim Impo2 As Double
Dim Impo3 As Double
Dim Impo4 As Double
Dim Impo5 As Double
Dim Impo6 As Double

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
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Auxi1 = " + "'" + Vence1.Text + "',"
    ZSql = ZSql + " Auxi2 = " + "'" + Vence2.Text + "',"
    ZSql = ZSql + " Auxi3 = " + "'" + Vence3.Text + "',"
    ZSql = ZSql + " Auxi4 = " + "'" + Vence4.Text + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Proveedor SET "
    ZSql = ZSql + " CodigoEmpresa = 1 ,"
    ZSql = ZSql + " Importe1 = 0 ,"
    ZSql = ZSql + " Importe2 = 0 ,"
    ZSql = ZSql + " Importe3 = 0 ,"
    ZSql = ZSql + " Importe4 = 0 ,"
    ZSql = ZSql + " Importe5 = 0 ,"
    ZSql = ZSql + " Importe6 = 0 "
    spProveedor = ZSql
    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Proyeccion de Pagos"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Fecha1 = Right$(Vence1.Text, 4) + Mid$(Vence1.Text, 4, 2) + Left$(Vence1.Text, 2)
    Fecha2 = Right$(Vence2.Text, 4) + Mid$(Vence2.Text, 4, 2) + Left$(Vence2.Text, 2)
    Fecha3 = Right$(Vence3.Text, 4) + Mid$(Vence3.Text, 4, 2) + Left$(Vence3.Text, 2)
    Fecha4 = Right$(Vence4.Text, 4) + Mid$(Vence4.Text, 4, 2) + Left$(Vence4.Text, 2)
    
    Pasa = 0
    Lugar = 0
    Impo1 = 0
    Impo2 = 0
    Impo3 = 0
    Impo4 = 0
    Impo5 = 0
    Impo6 = 0
    Erase WTrabajo

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCtePrv"
    ZSql = ZSql + " Where CtaCtePrv.Proveedor >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.Proveedor <= " + "'" + Hasta.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.Saldo <> 0 "
    ZSql = ZSql + " Order by CtaCtePrv.Proveedor,CtaCtePrv.OrdFecha,CtaCtePrv.Impre,CtaCtePrv.Numero"
        
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCtePrv.RecordCount > 0 Then
    
        With rstCtaCtePrv
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Pasa = 0 Then
                        Corte = rstCtaCtePrv!Proveedor
                        Pasa = 1
                        Impo1 = 0
                        Impo2 = 0
                        Impo3 = 0
                        Impo4 = 0
                        Impo5 = 0
                        Impo6 = 0
                    End If
                    
                    If Corte <> rstCtaCtePrv!Proveedor Then
                        Lugar = Lugar + 1
                        WTrabajo(Lugar, 1) = Corte
                        WTrabajo(Lugar, 2) = Str$(Impo1)
                        WTrabajo(Lugar, 3) = Str$(Impo2)
                        WTrabajo(Lugar, 4) = Str$(Impo3)
                        WTrabajo(Lugar, 5) = Str$(Impo4)
                        WTrabajo(Lugar, 6) = Str$(Impo5)
                        WTrabajo(Lugar, 7) = Str$(Impo6)
                        Impo1 = 0
                        Impo2 = 0
                        Impo3 = 0
                        Impo4 = 0
                        Impo5 = 0
                        Impo6 = 0
                        Corte = rstCtaCtePrv!Proveedor
                    End If
                        
                    WSaldo = !Saldo
                    WVencimiento = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)
                    
                    Impo6 = Impo6 + WSaldo
                    If WVencimiento <= Fecha1 Then
                        Impo1 = Impo1 + WSaldo
                            Else
                        If WVencimiento <= Fecha2 Then
                            Impo2 = Impo2 + WSaldo
                                Else
                            If WVencimiento <= Fecha3 Then
                                Impo3 = Impo3 + WSaldo
                                    Else
                                If WVencimiento <= Fecha4 Then
                                    Impo4 = Impo4 + WSaldo
                                        Else
                                    Impo5 = Impo5 + WSaldo
                                End If
                            End If
                        End If
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCtePrv.Close
    End If
    
    If Pasa <> 0 Then
        Lugar = Lugar + 1
        WTrabajo(Lugar, 1) = Corte
        WTrabajo(Lugar, 2) = Str$(Impo1)
        WTrabajo(Lugar, 3) = Str$(Impo2)
        WTrabajo(Lugar, 4) = Str$(Impo3)
        WTrabajo(Lugar, 5) = Str$(Impo4)
        WTrabajo(Lugar, 6) = Str$(Impo5)
        WTrabajo(Lugar, 7) = Str$(Impo6)
    End If
    
    For Ciclo = 1 To Lugar
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Proveedor SET "
        ZSql = ZSql + " Importe1 = " + "'" + WTrabajo(Ciclo, 2) + "',"
        ZSql = ZSql + " Importe2 = " + "'" + WTrabajo(Ciclo, 3) + "',"
        ZSql = ZSql + " Importe3 = " + "'" + WTrabajo(Ciclo, 4) + "',"
        ZSql = ZSql + " Importe4 = " + "'" + WTrabajo(Ciclo, 5) + "',"
        ZSql = ZSql + " Importe5 = " + "'" + WTrabajo(Ciclo, 6) + "',"
        ZSql = ZSql + " Importe6 = " + "'" + WTrabajo(Ciclo, 7) + "'"
        ZSql = ZSql + " Where Proveedor = " + "'" + WTrabajo(Ciclo, 1) + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    Rem If Val(Desde.Text) = 0 Then
    Rem     Desde.Text = "0"
    Rem End If
    Rem If Val(Hasta.Text) = 0 Then
    Rem     Hasta.Text = "0"
    Rem End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Proveedor.Proveedor, Proveedor.Nombre, Proveedor.Importe1, Proveedor.Importe2, Proveedor.Importe3, Proveedor.Importe4, Proveedor.Importe5, Proveedor.Importe6, " _
                + "Auxiliar.Nombre, Auxiliar.Auxi1, Auxiliar.Auxi2, Auxiliar.Auxi3, Auxiliar.Auxi4 " _
                + "From " _
                + DSQ + ".dbo.Proveedor Proveedor, " _
                + DSQ + ".dbo.Auxiliar Auxiliar " _
                + "Where " _
                + "Proveedor.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Proveedor.Proveedor >= '" + Desde.Text + "' AND " _
                + "Proveedor.Proveedor <= '" + Hasta.Text + "' AND " _
                + "Proveedor.Importe6 <> 0"
    
    Listado.Connect = Connect()
    
    Uno = "{Proveedor.Importe6} <> 0.00 "
    Dos = " and {Proveedor.Proveedor} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    
    Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgProyPrv.Hide
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
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Vence1.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Vence1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Then
            Vence2.SetFocus
                Else
            Vence1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence1.Text = "  /  /    "
    End If
End Sub

Private Sub Vence2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence2.Text, Auxi)
        If Auxi = "S" Then
            Vence3.SetFocus
                Else
            Vence2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence2.Text = "  /  /    "
    End If
End Sub

Private Sub Vence3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence3.Text, Auxi)
        If Auxi = "S" Then
            Vence4.SetFocus
                Else
            Vence3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence3.Text = "  /  /    "
    End If
End Sub

Private Sub Vence4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence4.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Vence4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vence4.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Vence1.Text = "  /  /    "
    Vence2.Text = "  /  /    "
    Vence3.Text = "  /  /    "
    Vence4.Text = "  /  /    "
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

Private Sub Vence1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vence4_KeyDown(KeyCode As Integer, Shift As Integer)
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



