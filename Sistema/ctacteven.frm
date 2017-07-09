VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCtaCteVen 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Clientes por Vendedor"
   ClientHeight    =   6840
   ClientLeft      =   2490
   ClientTop       =   585
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   ScaleHeight     =   6840
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
      TabIndex        =   10
      Top             =   3960
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
      Height          =   3735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Dias 
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
         MaxLength       =   8
         TabIndex        =   18
         Text            =   " "
         Top             =   1440
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
         Left            =   4920
         MouseIcon       =   "ctacteven.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ctacteven.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salida"
         Top             =   2400
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
         Left            =   2520
         MouseIcon       =   "ctacteven.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ctacteven.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Impresion x Impresora"
         Top             =   2400
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
         Left            =   3720
         MouseIcon       =   "ctacteven.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ctacteven.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Consulta de Datos"
         Top             =   2400
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
         Left            =   1320
         MouseIcon       =   "ctacteven.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "ctacteven.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox HastaVend 
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
         MaxLength       =   8
         TabIndex        =   2
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox DesdeVend 
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
         MaxLength       =   8
         TabIndex        =   0
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   5040
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   285
         Left            =   5040
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
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
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Visible         =   0   'False
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
      Begin VB.Label Label7 
         Caption         =   "Dias"
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
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Left            =   3480
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   3480
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Vendedor"
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Vendedor"
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
         TabIndex        =   8
         Top             =   720
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
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
      TabIndex        =   6
      Top             =   5280
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
      ItemData        =   "ctacteven.frx":2D30
      Left            =   120
      List            =   "ctacteven.frx":2D37
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   6735
   End
End
Attribute VB_Name = "PrgCtaCteVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WPasa As String
Private WTitulo As String
Private Importe3 As Double

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

    Rem On Error GoTo WError
    
    If DesdeFecha.Text = "  /  /    " Then
        DesdeFecha.Text = "00/00/0000"
    End If
    If HastaFecha.Text = "  /  /    " Then
        HastaFecha.Text = "99/99/9999"
    End If
    
    
    WAno = Right$(DesdeFecha.Text, 4)
    WMes = Mid$(DesdeFecha.Text, 4, 2)
    WDia = Left$(DesdeFecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia
    
    ZDesdeDias = "0"
    ZHastaDias = "999999"
    If Val(Dias.Text) <> 0 Then
        ZDesdeDias = Dias.Text
    End If
    

    ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Varios = " + "'" + ZFecha + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + " UPDATE Ctacte SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "',"
    ZSql = ZSql + " CtaCte.Vendedor = Cliente.Vendedor"
    ZSql = ZSql + " From CtaCte, Cliente"
    ZSql = ZSql + " Where Ctacte.Cliente = Cliente.Cliente"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)




    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Lista = " + "'" + "0" + "',"
    ZSql = ZSql + " Importe7 = " + "'" + "0" + "',"
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    ZSql = ZSql + " Where CtaCte.Vendedor >= " + "'" + DesdeVend.Text + "'"
    ZSql = ZSql + " and CtaCte.Vendedor <= " + "'" + HastaVend.Text + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Saldo = 0"
    ZSql = ZSql + " Where Saldo < 0.01 and Saldo > -0.01"
    ZSql = ZSql + " and CtaCte.Vendedor >= " + "'" + DesdeVend.Text + "'"
    ZSql = ZSql + " and CtaCte.Vendedor <= " + "'" + HastaVend.Text + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " SaldoUs = 0"
    ZSql = ZSql + " Where SaldoUs < 0.01 and SaldoUs > -0.01"
    ZSql = ZSql + " and CtaCte.Vendedor >= " + "'" + DesdeVend.Text + "'"
    ZSql = ZSql + " and CtaCte.Vendedor <= " + "'" + HastaVend.Text + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZFechaActual = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Erase ZVector
    ZLugar = 0
    ZPasa = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.Vendedor >= " + "'" + DesdeVend.Text + "'"
    ZSql = ZSql + " and CtaCte.Vendedor <= " + "'" + HastaVend.Text + "'"
    ZSql = ZSql + " and CtaCte.Saldo <> 0"
    ZSql = ZSql + " Order by CtaCte.Cliente,CtaCte.OrdFecha,CtaCte.Numero"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        With rstCtaCte
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZPasa = 0 Then
                        ZPasa = 1
                        ZCorte = rstCtaCte!Cliente
                        ZSuma = 0
                    End If
                    
                    If ZCorte <> rstCtaCte!Cliente Then
                        ZCorte = rstCtaCte!Cliente
                        ZSuma = 0
                    End If
                    
                    Select Case ZZNivel
                        Case 0
                            ZSuma = ZSuma + rstCtaCte!Saldo
                        Case 1
                            ZSuma = ZSuma + rstCtaCte!Saldous
                        Case Else
                            ZSuma = ZSuma + rstCtaCte!Saldous - rstCtaCte!Saldo
                    End Select
                    
                    ZZfecha = rstCtaCte!Fecha
                    ZDias = 0
                    ZDias = DateDiff("d", ZZfecha, ZZFechaActual)
                    If ZDias < 0 Then
                        ZDias = 0
                    End If
                    
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar, 1) = rstCtaCte!Clave
                    ZVector(ZLugar, 2) = Str$(ZSuma)
                    ZVector(ZLugar, 3) = Str$(ZDias)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCte.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZSql = ""
        ZSql = ZSql + "UPDATE CtaCte SET "
        ZSql = ZSql + " Lista = " + "'" + ZVector(Ciclo, 3) + "',"
        ZSql = ZSql + " Importe7 = " + "'" + ZVector(Ciclo, 2) + "'"
        ZSql = ZSql + " Where Clave = " + "'" + ZVector(Ciclo, 1) + "'"
        spCtaCte = ZSql
        Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    
    Listado.WindowTitle = "Listado de Cuenta Corriente de Clientes por Vendedor"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Val(DesdeVend.Text) = 0 Then
        DesdeVend.Text = "0"
    End If
    If Val(HastaVend.Text) = 0 Then
        HastaVend.Text = "0"
    End If

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    If ZZNivel = 0 Then
    
        Listado.SQLQuery = "SELECT CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.Total, CtaCte.Saldo, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.Remito, CtaCte.Importe7, CtaCte.Lista, " _
                + "Cliente.Razon, Cliente.Vendedor, " _
                + "Auxiliar.Nombre, Auxiliar.Varios, " _
                + "Vendedor.Nombre " _
                + "From " _
                + DSQ + ".dbo.CtaCte CtaCte, " _
                + DSQ + ".dbo.Cliente Cliente, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Vendedor Vendedor " _
                + "Where " _
                + "CtaCte.Cliente = Cliente.Cliente AND " _
                + "CtaCte.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Ctacte.Vendedor = Vendedor.Codigo AND " _
                + "CtaCte.Saldo <> 0 AND " _
                + "CtaCte.OrdFecha >= '" + WDesde + "' AND " _
                + "CtaCte.OrdFecha <= '" + WHasta + "' AND " _
                + "Cliente.Vendedor >= " + DesdeVend.Text + " AND " _
                + "Cliente.Vendedor <= " + HastaVend.Text + " AND " _
                + "CtaCte.Lista >= " + ZDesdeDias + " AND " _
                + "CtaCte.Lista <= " + ZHastaDias
                    
        Uno = "{CtaCte.Saldo} <> 0 "
        Dos = " and {CtaCte.Vendedor} in " + DesdeVend.Text + " to " + HastaVend.Text
        Tres = " and {CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        Cuatro = " and {CtaCte.Lista} in " + ZDesdeDias + " to " + ZHastaDias
            
        Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
        Listado.SelectionFormula = Uno + Dos + Tres + Cuatro
            
        Listado.ReportFileName = "CtacteVen.rpt"
        
            Else
            
        Listado.SQLQuery = "SELECT CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte.Total, CtaCte.Saldo, CtaCte.OrdFecha, CtaCte.Impre, CtaCte.TotalUs, CtaCte.SaldoUs, CtaCte.Lista, " _
                + "Cliente.Razon, Cliente.Vendedor, " _
                + "Auxiliar.Nombre, Auxiliar.Varios, " _
                + "Vendedor.Nombre " _
                + "From " _
                + DSQ + ".dbo.CtaCte CtaCte, " _
                + DSQ + ".dbo.Cliente Cliente, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Vendedor Vendedor " _
                + "Where " _
                + "CtaCte.Cliente = Cliente.Cliente AND " _
                + "CtaCte.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Ctacte.Vendedor = Vendedor.Codigo AND " _
                + "CtaCte.SaldoUs <> 0 AND " _
                + "CtaCte.OrdFecha >= '" + WDesde + "' AND " _
                + "CtaCte.OrdFecha <= '" + WHasta + "' AND " _
                + "Cliente.Vendedor >= " + DesdeVend.Text + " AND " _
                + "Cliente.Vendedor <= " + HastaVend.Text + " AND " _
                + "CtaCte.Lista >= " + ZDesdeDias + " AND " _
                + "CtaCte.Lista <= " + ZHastaDias
                    
        Uno = "{CtaCte.SaldoUs} <> 0 "
        Dos = " and {CtaCte.Vendedor} in " + DesdeVend.Text + " to " + HastaVend.Text
        Tres = " and {CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
        Cuatro = " and {CtaCte.Lista} in " + ZDesdeDias + " to " + ZHastaDias
            
        Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
        Listado.SelectionFormula = Uno + Dos + Tres + Cuatro
            
        Listado.ReportFileName = "CtacteVenTotal.rpt"
            
    End If
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgCtaCteVen.Hide
    Unload Me
    Menu41.Show
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
    ZSql = ZSql + " FROM Vendedor"
    ZSql = ZSql + " Order by Vendedor.Codigo"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = Str$(!Codigo) + " " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Codigo
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
    DesdeVend.Text = WIndice.List(Indice)
    HastaVend.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    DesdeVend.SetFocus
    
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            DesdeVend.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub DesdeVend_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaVend.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeVend.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub HastaVend_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dias.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaVend.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Dias_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeVend.SetFocus
    End If
    If KeyAscii = 27 Then
        Dias.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFecha.Text, Auxi)
        If Auxi = "S" Then
            HastaFecha.SetFocus
                Else
            DesdeFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFecha.Text = "  /  /    "
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFecha.Text, Auxi)
        If Auxi = "S" Then
            Fecha.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    DesdeVend.Text = ""
    HastaVend.Text = ""
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Dias.Text = ""
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
    
    XIndice = 0
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Where Vendedor.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Vendedor.Codigo"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstVendedor!Codigo) + " " + rstVendedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstVendedor!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
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

Private Sub DesdeVend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaVend_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdeFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFecha_KeyDown(KeyCode As Integer, Shift As Integer)
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



