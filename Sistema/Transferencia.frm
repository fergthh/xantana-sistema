VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgTransferencia 
   AutoRedraw      =   -1  'True
   Caption         =   "Transferencia entre Cuentas"
   ClientHeight    =   7785
   ClientLeft      =   1005
   ClientTop       =   420
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   ScaleHeight     =   7785
   ScaleWidth      =   9900
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   9495
      Begin VB.TextBox NroCheque 
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
         TabIndex        =   27
         Text            =   " "
         Top             =   1440
         Width           =   1695
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
         TabIndex        =   22
         Text            =   " "
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox BancoII 
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
         TabIndex        =   21
         Text            =   " "
         Top             =   2160
         Width           =   975
      End
      Begin VB.ComboBox TipoII 
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
         TabIndex        =   20
         Top             =   1800
         Width           =   1695
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
         MouseIcon       =   "Transferencia.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Transferencia.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salida"
         Top             =   3480
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
         Left            =   4680
         MouseIcon       =   "Transferencia.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "Transferencia.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Consulta de Datos"
         Top             =   3480
         Width           =   855
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
         TabIndex        =   10
         Text            =   " "
         Top             =   3000
         Width           =   5535
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
      Begin VB.TextBox BancoI 
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
         TabIndex        =   9
         Text            =   " "
         Top             =   1080
         Width           =   975
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
         MouseIcon       =   "Transferencia.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "Transferencia.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   3480
         Width           =   855
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
         MouseIcon       =   "Transferencia.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "Transferencia.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Limpia la pantalla"
         Top             =   3480
         Width           =   855
      End
      Begin VB.ComboBox TipoI 
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
         TabIndex        =   6
         Top             =   720
         Width           =   1695
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
      Begin VB.Label Label3 
         Caption         =   "Nro de Cheque"
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
         TabIndex        =   28
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
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
         TabIndex        =   26
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Banco"
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
         TabIndex        =   25
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label DesBancoII 
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
         TabIndex        =   24
         Top             =   2160
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Entrada"
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
         TabIndex        =   23
         Top             =   1800
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
         TabIndex        =   19
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Banco"
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1575
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
         TabIndex        =   17
         Top             =   360
         Width           =   1815
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
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label DesBancoI 
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
         TabIndex        =   15
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Salida"
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
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
   End
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
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   3975
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
      Top             =   4920
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
      ItemData        =   "Transferencia.frx":2D30
      Left            =   120
      List            =   "Transferencia.frx":2D37
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZProvincia(100) As String

Private Sub CmdLimpiar_Click()

    Codigo.Text = "1"
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    BancoI.Text = ""
    DesBancoI.Caption = ""
    BancoII.Text = ""
    DesBancoII.Caption = ""
    Importe.Text = ""
    Observaciones.Text = ""
    NroCheque.Text = ""
    
    TipoI.ListIndex = 0
    TipoII.ListIndex = 0
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Transferencia"
    spTransferencia = ZSql
    Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    If rstTransferencia.RecordCount > 0 Then
        rstTransferencia.MoveLast
        ZUltimo = IIf(IsNull(rstTransferencia!CodigoMayor), "0", rstTransferencia!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstTransferencia.Close
    End If
    
    Codigo.SetFocus

End Sub

Private Sub Form_Activate()

    TipoI.Clear
    
    TipoI.AddItem ""
    TipoI.AddItem "Efectivo"
    TipoI.AddItem "Banco"
    TipoI.AddItem "Caja"
    
    TipoI.ListIndex = 0

    TipoII.Clear
    
    TipoII.AddItem ""
    TipoII.AddItem "Efectivo"
    TipoII.AddItem "Banco"
    TipoII.AddItem "Caja"
    
    TipoII.ListIndex = 0

End Sub

Private Sub Proceso_Click()

    ZZOrdFecha = ""
    ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Transferencia"
    ZSql = ZSql + " Where Transferencia.Codigo = " + "'" + Codigo.Text + "'"
    spTransferencia = ZSql
    Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    If rstTransferencia.RecordCount > 0 Then
    
        rstTransferencia.Close
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Transferencia SET "
        ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
        ZSql = ZSql + " OrdFecha = " + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + " TipoI = " + "'" + Str$(TipoI.ListIndex) + "',"
        ZSql = ZSql + " BancoI = " + "'" + BancoI.Text + "',"
        ZSql = ZSql + " TipoII = " + "'" + Str$(TipoII.ListIndex) + "',"
        ZSql = ZSql + " BancoII = " + "'" + BancoII.Text + "',"
        ZSql = ZSql + " Importe = " + "'" + Importe.Text + "',"
        ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "',"
        ZSql = ZSql + " NroCheque = " + "'" + NroCheque.Text + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        
        spTransferencia = ZSql
        Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    
            Else

        ZSql = ""
        ZSql = ZSql + "INSERT INTO Transferencia ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "TipoI ,"
        ZSql = ZSql + "BancoI ,"
        ZSql = ZSql + "TipoII ,"
        ZSql = ZSql + "BancoII ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "NroCheque ,"
        ZSql = ZSql + "Observaciones )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + "'" + Str$(TipoI.ListIndex) + "',"
        ZSql = ZSql + "'" + BancoI.Text + "',"
        ZSql = ZSql + "'" + Str$(TipoII.ListIndex) + "',"
        ZSql = ZSql + "'" + BancoII.Text + "',"
        ZSql = ZSql + "'" + Importe.Text + "',"
        ZSql = ZSql + "'" + NroCheque.Text + "',"
        ZSql = ZSql + "'" + Observaciones.Text + "')"
                                
        spTransferencia = ZSql
        Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    Call CmdLimpiar_Click
    
End Sub

Private Sub Cancela_Click()
    PrgTransferencia.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Transferencia"
            ZSql = ZSql + " Where Transferencia.Codigo = " + "'" + Codigo.Text + "'"
            spTransferencia = ZSql
            Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
            If rstTransferencia.RecordCount > 0 Then
                Fecha.Text = rstTransferencia!Fecha
                TipoI.ListIndex = rstTransferencia!TipoI
                BancoI.Text = rstTransferencia!BancoI
                TipoII.ListIndex = rstTransferencia!TipoII
                BancoII.Text = rstTransferencia!BancoII
                Importe.Text = Str$(rstTransferencia!Importe)
                Observaciones.Text = rstTransferencia!Observaciones
                rstTransferencia.Close
            End If
            Fecha.SetFocus
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Banco = " + "'" + BancoI.Text + "'"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                DesBancoI.Caption = rstBanco!Nombre
                rstBanco.Close
                    Else
                DesBancoI.Caption = ""
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Banco = " + "'" + BancoII.Text + "'"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                DesBancoII.Caption = rstBanco!Nombre
                rstBanco.Close
                    Else
                DesBancoII.Caption = ""
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
            TipoI.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub TipoI_Click()

    On Error GoTo WError
    
        If TipoI.ListIndex = 2 Then
            BancoI.SetFocus
                Else
            TipoII.SetFocus
        End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub TipoI_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TipoI.ListIndex = 2 Then
            BancoI.SetFocus
                Else
            TipoII.SetFocus
        End If
    End If
End Sub

Private Sub BancoI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + BancoI.Text + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            DesBancoI.Caption = rstBanco!Nombre
            rstBanco.Close
            NroCheque.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        BancoI.Text = ""
        DesBancoI.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub NroCheque_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TipoII.SetFocus
    End If
    If KeyAscii = 27 Then
        NroCheque.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Private Sub TipoII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TipoII.ListIndex = 2 Then
            BancoII.SetFocus
                Else
            Importe.SetFocus
        End If
    End If
End Sub

Private Sub TipoII_Click()

    On Error GoTo WError
    
        If TipoII.ListIndex = 2 Then
            BancoII.SetFocus
                Else
            Importe.SetFocus
        End If
    
    Exit Sub
    
WError:
    Resume Next

        
End Sub

Private Sub BancoII_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + BancoII.Text + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            DesBancoII.Caption = rstBanco!Nombre
            rstBanco.Close
            Importe.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        BancoII.Text = ""
        DesBancoII.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Importe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Importe.Text = ""
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

    TipoI.Clear
    
    TipoI.AddItem ""
    TipoI.AddItem "Efectivo"
    TipoI.AddItem "Banco"
    TipoI.AddItem "Caja"
    
    TipoI.ListIndex = 0

    TipoII.Clear
    
    TipoII.AddItem ""
    TipoII.AddItem "Efectivo"
    TipoII.AddItem "Banco"
    TipoII.AddItem "Caja"
    
    TipoII.ListIndex = 0


    Codigo.Text = "1"
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    BancoI.Text = ""
    DesBancoI.Caption = ""
    BancoII.Text = ""
    DesBancoII.Caption = ""
    Importe.Text = ""
    Observaciones.Text = ""
    NroCheque.Text = ""
 
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Transferencia"
    spTransferencia = ZSql
    Set rstTransferencia = db.OpenRecordset(spTransferencia, dbOpenSnapshot, dbSQLPassThrough)
    If rstTransferencia.RecordCount > 0 Then
        rstTransferencia.MoveLast
        ZUltimo = IIf(IsNull(rstTransferencia!CodigoMayor), "0", rstTransferencia!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstTransferencia.Close
    End If
    
    Rem Codigo.SetFocus
    
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Bancos"

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
        Case 0, 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Order by Banco.Banco"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                With rstBanco
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Banco) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Banco
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstBanco.Close
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
            BancoI.Text = WIndice.List(Indice)
            Call BancoI_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            BancoII.Text = WIndice.List(Indice)
            Call BancoII_KeyPress(13)
            
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
        Case 0, 1
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Banco.Banco"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                With rstBanco
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Banco) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Banco
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstBanco.Close
            End If
            
        Case Else
    End Select
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub BancoI_DblClick()

    Opcion.Clear
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub BancoII_DblClick()

    Opcion.Clear
    Opcion.AddItem "Bancos"
    Opcion.AddItem "Cuentas Contables"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub

Private Sub BancoI_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub BancoII_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TipoII_KeyDown(KeyCode As Integer, Shift As Integer)
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










