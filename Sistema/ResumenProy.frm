VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgResumenProy 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Resumen de Movimientos "
   ClientHeight    =   2655
   ClientLeft      =   3135
   ClientTop       =   1815
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   2655
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   4695
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
         Left            =   600
         MouseIcon       =   "ResumenProy.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "ResumenProy.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   1080
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
         Left            =   1920
         MouseIcon       =   "ResumenProy.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "ResumenProy.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1080
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
         Left            =   3240
         MouseIcon       =   "ResumenProy.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "ResumenProy.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salida"
         Top             =   1080
         Width           =   855
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Label Label1 
         Caption         =   "Fecha Emision"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5280
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ResumenProy.rpt"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgResumenProy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WProyecto As String
Dim WConcepto As String

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
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Actividad = "al " + Fecha.Text
            .Update
        End If
    End With
    
    Listado.WindowTitle = "Listado de Control de Gastos por Proyecto"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFecha = WAno + WMes + "31"
    
    da = ""
    With rstGastosProy
        .Index = "Clave"
        .Seek ">=", da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    With rstImpproy
        .Index = "Clave"
        .MoveFirst
        Do
            If Left$(WFecha, 6) = Left$(!ordfecha, 6) Then
                WOrdfecha = !ordfecha
                WProyecto = ""
                WConcepto = !Concepto
                WImporte = !Importe
                WTipo = "1"
                Call Ceros(WConcepto, 4)
                WClave = WTipo + WProyecto + WConcepto
                With rstGastosProy
                    .Index = "Clave"
                    .Seek "=", WClave
                    If .NoMatch Then
                        .AddNew
                        !Clave = WClave
                        !Proyecto = WProyecto
                        !Concepto = Val(WConcepto)
                        !Importe1 = WImporte
                        !Importe2 = 0
                        !Importe3 = 0
                        !Importe4 = 0
                        !Porce = 0
                        !Tipo = 1
                        !Descripcion = ""
                        .Update
                        .Bookmark = .LastModified
                            Else
                        .Edit
                        !Importe1 = !Importe1 + WImporte
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
            End If
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    With rstCtaCte
        .Index = "Clave"
        .MoveFirst
        Do
            If Left$(WFecha, 6) = Left$(!ordfecha, 6) And !Neto <> 0 Then
                WOrdfecha = !ordfecha
                WProyecto = !Proyecto
                WConcepto = 0
                WImporte = !Neto
                WTipo = "2"
                Call Ceros(WConcepto, 4)
                WClave = WTipo + WProyecto + WConcepto
                With rstGastosProy
                    .Index = "Clave"
                    .Seek "=", WClave
                    If .NoMatch Then
                        .AddNew
                        !Clave = WClave
                        !Proyecto = WProyecto
                        !Concepto = Val(WConcepto)
                        !Importe1 = 0
                        !Importe2 = WImporte
                        !Importe3 = 0
                        !Importe4 = 0
                        !Porce = 0
                        !Tipo = 2
                        .Update
                        .Bookmark = .LastModified
                            Else
                        .Edit
                        !Importe2 = !Importe2 + WImporte
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
            End If
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    Rem Listado.GroupSelectionFormula = "{Impproy.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34) + " and {Impproy.proyecto} in " + Chr$(34) + DesdeProy.Text + Chr$(34) + " to " + Chr$(34) + HastaProy.Text + Chr$(34)
    Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    Listado.Action = 1
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstImpproy
        .Close
    End With
    With rstProyecto
        .Close
    End With
    With rstGastosProy
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    DbsAdminis.Close
    Fecha.SetFocus
    PrgGastosProy.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Impproy
    OPEN_FILE_Ctacte
    OPEN_FILE_Proyecto
    OPEN_FILE_GastosProy
    OPEN_FILE_Auxiliar
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
    End If
End Sub

Sub Form_Load()
    Fecha.Text = "  /  /    "
    Frame2.Visible = True
End Sub

