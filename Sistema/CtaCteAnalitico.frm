VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCtaCteAnalitico 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Clientes Analitico"
   ClientHeight    =   6105
   ClientLeft      =   2490
   ClientTop       =   585
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   ScaleHeight     =   6105
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
      TabIndex        =   7
      Top             =   3360
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
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6735
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
         Left            =   4800
         MouseIcon       =   "CtaCteAnalitico.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteAnalitico.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   2400
         MouseIcon       =   "CtaCteAnalitico.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteAnalitico.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1920
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
         Left            =   3600
         MouseIcon       =   "CtaCteAnalitico.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteAnalitico.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Consulta de Datos"
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
         Left            =   1200
         MouseIcon       =   "CtaCteAnalitico.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "CtaCteAnalitico.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox DesdeCli 
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
         Top             =   360
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
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
         Left            =   1800
         TabIndex        =   1
         Top             =   960
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
      Begin VB.Label DesCliente 
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
         Left            =   3240
         TabIndex        =   14
         Top             =   360
         Width           =   3375
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
         Left            =   240
         TabIndex        =   13
         Top             =   960
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
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1215
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
      TabIndex        =   4
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
      ItemData        =   "CtaCteAnalitico.frx":2D30
      Left            =   120
      List            =   "CtaCteAnalitico.frx":2D37
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   6735
   End
End
Attribute VB_Name = "PrgCtaCteAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WPasa As String
Private WTitulo As String
Private Importe3 As Double

Dim ZVector(10000, 15) As String
Dim ZZRecibo(100, 5) As String

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
    
    ZZfecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    ZZTitulo = "Desde el " + DesdeFecha.Text + " al " + HastaFecha.Text
    
    ZSql = ""
    ZSql = ZSql + "DELETE ImpCtaCte"
    spImpCtaCte = ZSql
    Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Saldo = 0"
    ZSql = ZSql + " Where Saldo < 0.01 and Saldo > -0.01"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " SaldoUs = 0"
    ZSql = ZSql + " Where SaldoUs < 0.01 and SaldoUs > -0.01"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)

    Listado.WindowTitle = "Listado de Cuenta Corriente de Clientes Analitico"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Erase ZVector
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.Cliente >= " + "'" + DesdeCli.Text + "'"
    ZSql = ZSql + " and CtaCte.Cliente <= " + "'" + DesdeCli.Text + "'"
    ZSql = ZSql + " and CtaCte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and CtaCte.OrdFecha <= " + "'" + WHasta + "'"
    ZSql = ZSql + " Order by CtaCte.Cliente,CtaCte.OrdFecha,CtaCte.Numero"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
        With rstCtaCte
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugar = ZLugar + 1
                    
                    ZVector(ZLugar, 1) = rstCtaCte!Clave
                    ZVector(ZLugar, 2) = rstCtaCte!Letra
                    ZVector(ZLugar, 3) = rstCtaCte!Tipo
                    ZVector(ZLugar, 4) = rstCtaCte!Punto
                    ZVector(ZLugar, 5) = rstCtaCte!Numero
                    ZVector(ZLugar, 6) = rstCtaCte!Renglon
                    ZVector(ZLugar, 7) = rstCtaCte!Cliente
                    ZVector(ZLugar, 8) = rstCtaCte!Fecha
                    ZVector(ZLugar, 9) = rstCtaCte!Remito
                    ZVector(ZLugar, 10) = rstCtaCte!vencimiento
                    ZVector(ZLugar, 11) = Str$(rstCtaCte!Total)
                    ZVector(ZLugar, 12) = Str$(rstCtaCte!Saldo)
                    ZVector(ZLugar, 13) = rstCtaCte!ordfecha
                    ZVector(ZLugar, 14) = rstCtaCte!OrdVencimiento
                    ZVector(ZLugar, 15) = rstCtaCte!Impre
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCte.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        If ZVector(Ciclo, 3) <> 6 Then
    
            ZZClave = ZVector(Ciclo, 1)
            ZZLetra = ZVector(Ciclo, 2)
            ZZTipo = ZVector(Ciclo, 3)
            ZZPunto = ZVector(Ciclo, 4)
            ZZNumero = ZVector(Ciclo, 5)
            ZZRenglon = ZVector(Ciclo, 6)
            ZZCliente = ZVector(Ciclo, 7)
            ZZfecha = ZVector(Ciclo, 8)
            ZZEstado = ""
            ZZVencimiento = ZVector(Ciclo, 10)
            ZZTotal = ZVector(Ciclo, 11)
            ZZSaldo = ZVector(Ciclo, 12)
            ZZOrdFecha = ZVector(Ciclo, 13)
            ZZOrdVencimiento = ZVector(Ciclo, 14)
            ZZImpre = ZVector(Ciclo, 15)
            ZZPeriodo = ZZTitulo
            ZZDesEmpresa = WNombreEmpresa
            ZZRemito = ZVector(Ciclo, 9)
            ZZAgrupa = ZZClave
            
            If Val(ZZTipo) = 2 Or Val(ZZTipo) = 5 Then
                If Val(ZZRemito) <> 0 Then
                    Auxi = ZZRemito
                    Call Ceros(Auxi, 8)
                    ZZAgrupa = Left$(ZZAgrupa, 1) + "010001" + Auxi + "01"
                End If
            End If
                            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO ImpCtaCte ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Letra ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Punto ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "Vencimiento ,"
            ZSql = ZSql + "Total ,"
            ZSql = ZSql + "Saldo ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "OrdVencimiento ,"
            ZSql = ZSql + "Impre ,"
            ZSql = ZSql + "Periodo ,"
            ZSql = ZSql + "DesEmpresa ,"
            ZSql = ZSql + "Agrupa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZLetra + "',"
            ZSql = ZSql + "'" + ZZTipo + "',"
            ZSql = ZSql + "'" + ZZPunto + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZEstado + "',"
            ZSql = ZSql + "'" + ZZVencimiento + "',"
            ZSql = ZSql + "'" + ZZTotal + "',"
            ZSql = ZSql + "'" + ZZSaldo + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
            ZSql = ZSql + "'" + ZZImpre + "',"
            ZSql = ZSql + "'" + ZZPeriodo + "',"
            ZSql = ZSql + "'" + ZZDesEmpresa + "',"
            ZSql = ZSql + "'" + ZZAgrupa + "')"
            spImpCtaCte = ZSql
            Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    
            Erase ZZRecibo
            ZZLugar = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Recibos"
            ZSql = ZSql + " Where Recibos.Letra1 = " + "'" + ZZLetra + "'"
            ZSql = ZSql + " and Recibos.Tipo1 = " + "'" + ZZTipo + "'"
            ZSql = ZSql + " and Recibos.Punto1 = " + "'" + ZZPunto + "'"
            ZSql = ZSql + " and Recibos.Numero1 = " + "'" + ZZNumero + "'"
            spRecibos = ZSql
            Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
            If rstRecibos.RecordCount > 0 Then
            
                With rstRecibos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            ZZLugar = ZZLugar + 1
                            ZZRecibo(ZZLugar, 1) = rstRecibos!Recibo
                            ZZRecibo(ZZLugar, 2) = Str$(rstRecibos!Importe1)
                            ZZRecibo(ZZLugar, 3) = rstRecibos!Fecha
                            ZZRecibo(ZZLugar, 4) = rstRecibos!fechaord
                                
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
            End If
            
            For ZZCiclo = 1 To ZZLugar
            
                ZZNroRecibo = ZZRecibo(ZZCiclo, 1)
                ZZImpoRecibo = ZZRecibo(ZZCiclo, 2)
                ZZFechaRecibo = ZZRecibo(ZZCiclo, 3)
                ZZOrdFechaRecibo = ZZRecibo(ZZCiclo, 4)
            
                ZZClave = ZVector(Ciclo, 1)
                ZZLetra = "X"
                ZZTipo = "06"
                ZZPunto = "0000"
                ZZNumero = ZZNroRecibo
                ZZRenglon = "01"
                ZZCliente = ZVector(Ciclo, 7)
                ZZfecha = ZZFechaRecibo
                ZZEstado = "0"
                ZZVencimiento = ZZFechaRecibo
                ZZTotal = Str$(Val(ZZImpoRecibo) * -1)
                ZZSaldo = "0"
                ZZOrdFecha = ZZOrdFechaRecibo
                ZZOrdVencimiento = ZZOrdFechaRecibo
                ZZImpre = "RC"
                ZZPeriodo = ZZTitulo
                ZZDesEmpresa = WNombreEmpresa
                ZZAgrupa = ZVector(Ciclo, 1)
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO ImpCtaCte ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Letra ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Punto ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Cliente ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Estado ,"
                ZSql = ZSql + "Vencimiento ,"
                ZSql = ZSql + "Total ,"
                ZSql = ZSql + "Saldo ,"
                ZSql = ZSql + "OrdFecha ,"
                ZSql = ZSql + "OrdVencimiento ,"
                ZSql = ZSql + "Impre ,"
                ZSql = ZSql + "Periodo ,"
                ZSql = ZSql + "DesEmpresa ,"
                ZSql = ZSql + "Agrupa )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZZClave + "',"
                ZSql = ZSql + "'" + ZZLetra + "',"
                ZSql = ZSql + "'" + ZZTipo + "',"
                ZSql = ZSql + "'" + ZZPunto + "',"
                ZSql = ZSql + "'" + ZZNumero + "',"
                ZSql = ZSql + "'" + ZZRenglon + "',"
                ZSql = ZSql + "'" + ZZCliente + "',"
                ZSql = ZSql + "'" + ZZfecha + "',"
                ZSql = ZSql + "'" + ZZEstado + "',"
                ZSql = ZSql + "'" + ZZVencimiento + "',"
                ZSql = ZSql + "'" + ZZTotal + "',"
                ZSql = ZSql + "'" + ZZSaldo + "',"
                ZSql = ZSql + "'" + ZZOrdFecha + "',"
                ZSql = ZSql + "'" + ZZOrdVencimiento + "',"
                ZSql = ZSql + "'" + ZZImpre + "',"
                ZSql = ZSql + "'" + ZZPeriodo + "',"
                ZSql = ZSql + "'" + ZZDesEmpresa + "',"
                ZSql = ZSql + "'" + ZZAgrupa + "')"
                spImpCtaCte = ZSql
                Set rstImpCtaCte = db.OpenRecordset(spImpCtaCte, dbOpenSnapshot, dbSQLPassThrough)
                
            Next ZZCiclo
    
        End If
    
    Next Ciclo
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    Listado.SQLQuery = "SELECT Impctacte.Letra, Impctacte.Tipo, Impctacte.Punto, Impctacte.Numero, Impctacte.Cliente, Impctacte.fecha, Impctacte.Total, Impctacte.OrdFecha, Impctacte.Impre, Impctacte.Periodo, Impctacte.DesEmpresa, Impctacte.Agrupa, " _
        + "Cliente.Razon " _
        + "From " _
        + DSQ + ".dbo.Impctacte Impctacte," _
        + DSQ + ".dbo.Cliente Cliente " _
        + "Where " _
        + "Impctacte.Cliente = Cliente.Cliente AND " _
        + "Impctacte.Cliente >= '" + DesdeCli.Text + "' AND " _
        + "Impctacte.Cliente <= '" + DesdeCli.Text + "'"
    
    Uno = "{IMPCtaCte.Cliente} in " + Chr$(34) + DesdeCli.Text + Chr$(34) + " to " + Chr$(34) + DesdeCli.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.ReportFileName = "CtacteAnalitico.rpt"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgCtaCteAnalitico.Hide
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
    ZSql = ZSql + " FROM Cliente"
    ZSql = ZSql + " Order by Cliente.Cliente"
    spCliente = ZSql
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = !Cliente + " " + !Razon
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Cliente
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
    Ayuda.Visible = False
    DesdeCli.Text = WIndice.List(Indice)
    DesdeCli_Keypress (13)
    
End Sub

Private Sub DesdeCli_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(DesdeCli.Text) <> "" Then
            Auxi = UCase(Left$(DesdeCli.Text, 1))
            Auxi1 = Mid$(DesdeCli.Text, 2, 5)
            Call Ceros(Auxi1, 3)
            DesdeCli.Text = Auxi + "-" + Auxi1
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + DesdeCli.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            rstCliente.Close
            DesdeFecha.SetFocus
        End If
    
    End If
    If KeyAscii = 27 Then
        DesdeCli.Text = ""
        DesCliente.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
            DesdeCli.SetFocus
                Else
            HastaFecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    DesdeCli.Text = ""
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
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
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Cliente.Cliente"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
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

Private Sub DesdeCli_KeyDown(KeyCode As Integer, Shift As Integer)
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



