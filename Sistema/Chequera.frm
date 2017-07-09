VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgChequera 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Chequeras"
   ClientHeight    =   6255
   ClientLeft      =   1170
   ClientTop       =   930
   ClientWidth     =   9750
   LinkTopic       =   "Form2"
   ScaleHeight     =   6255
   ScaleWidth      =   9750
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
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
      Left            =   120
      MouseIcon       =   "Chequera.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "Borra  F2"
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
      Left            =   1080
      MouseIcon       =   "Chequera.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Elimina el Registro"
      Top             =   1920
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
      Left            =   2040
      MouseIcon       =   "Chequera.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Limpia la pantalla"
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
      Left            =   3000
      MouseIcon       =   "Chequera.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Consulta de Datos"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Lista 
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
      Left            =   7800
      MouseIcon       =   "Chequera.frx":2D30
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":303A
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Impresion "
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton CmdClose 
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
      Left            =   8760
      MouseIcon       =   "Chequera.frx":387C
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":3B86
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Salida"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Primer 
      Caption         =   "Primer F5"
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
      Left            =   3960
      MouseIcon       =   "Chequera.frx":43C8
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":46D2
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Primer Registro"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Anterior 
      Caption         =   "Anterior F6"
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
      MouseIcon       =   "Chequera.frx":4B14
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":4E1E
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Registro Anterior"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Siguiente 
      Caption         =   "Siguien. F7"
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
      Left            =   5880
      MouseIcon       =   "Chequera.frx":5260
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":556A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Registro Siguiente"
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Ultimo 
      Caption         =   "Ultimo F8"
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
      Left            =   6840
      MouseIcon       =   "Chequera.frx":59AC
      MousePointer    =   99  'Custom
      Picture         =   "Chequera.frx":5CB6
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Salida"
      Top             =   1920
      Width           =   855
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
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   18
      Text            =   " "
      Top             =   1560
      Width           =   1335
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
      Left            =   3120
      MaxLength       =   8
      TabIndex        =   16
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   1800
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   5775
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancela F12"
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
         Left            =   4560
         MouseIcon       =   "Chequera.frx":60F8
         MousePointer    =   99  'Custom
         Picture         =   "Chequera.frx":6402
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Confirma F11"
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
         Left            =   3480
         MouseIcon       =   "Chequera.frx":6844
         MousePointer    =   99  'Custom
         Picture         =   "Chequera.frx":6B4E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   480
         Width           =   855
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   840
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
         Left            =   1680
         TabIndex        =   21
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
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1215
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
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Banco 
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
      Left            =   3120
      MaxLength       =   4
      TabIndex        =   1
      Text            =   " "
      Top             =   840
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
      Left            =   3120
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Chequera.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Bancos"
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
      Left            =   7080
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      ItemData        =   "Chequera.frx":6F90
      Left            =   120
      List            =   "Chequera.frx":6F97
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   8175
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
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   3120
      TabIndex        =   14
      Top             =   480
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
   Begin VB.Label Label5 
      Caption         =   "Hasta Numero"
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
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Desde Numero"
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
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fecha de Emision"
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
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label DesBanco 
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
      Left            =   4200
      TabIndex        =   6
      Top             =   840
      Width           =   3735
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nro de Movimiento"
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
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   2535
   End
End
Attribute VB_Name = "PrgChequera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Imprime_Nombre()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Banco"
    ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
    spBanco = ZSql
    Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
    If rstBanco.RecordCount > 0 Then
        DesBanco.Caption = rstBanco!Nombre
        rstBanco.Close
            Else
        DesBanco.Caption = ""
    End If
    
End Sub

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
         Codigo.Text = "0"
    End If
    If Val(Banco.Text) = 0 Then
         Banco.Text = "0"
    End If
    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Comision.text = PUsing("###,###.##", Comision.text)
End Sub

Sub Imprime_Datos()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Chequera"
    ZSql = ZSql + " Where Chequera.Codigo = " + "'" + Codigo.Text + "'"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        Codigo.Text = rstChequera!Codigo
        Banco.Text = rstChequera!Banco
        Fecha.Text = rstChequera!Fecha
        Desde.Text = rstChequera!Desde
        Hasta.Text = rstChequera!Hasta
        rstChequera.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
    
End Sub

Private Sub Acepta_Click()

    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Chequera SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Listado.WindowTitle = "Listado de Chequeras"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Chequera.Codigo, Chequera.Fecha, Chequera.Ordfecha, Chequera.Banco, Chequera.Desde, Chequera.Hasta, " _
                    + "Auxiliar.Nombre, " _
                    + "Banco.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Chequera Chequera, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Banco Banco " _
                    + "Where " _
                    + "Chequera.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Chequera.Banco = Banco.Banco AND " _
                    + "Chequera.Ordfecha >= '" + WDesde + "' AND " _
                    + "Chequera.Ordfecha <= '" + WHasta + "'"
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Chequera.ordFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.SelectionFormula = "{Chequera.ordFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    
    Rem Codigo.SetFocus
    Rem Frame2.Visible = False
End Sub

Private Sub Cancela_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
    
        ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Chequera"
        ZSql = ZSql + " Where Chequera.Codigo = " + "'" + Codigo.Text + "'"
        spChequera = ZSql
        Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
        If rstChequera.RecordCount > 0 Then
            rstChequera.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Chequera SET "
            ZSql = ZSql + " Banco = " + "'" + Banco.Text + "',"
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + " Desde = " + "'" + Desde.Text + "',"
            ZSql = ZSql + " Hasta = " + "'" + Hasta.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spChequera = ZSql
            Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Chequera ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Banco ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Desde ,"
            ZSql = ZSql + "Hasta )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Banco.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + Desde.Text + "',"
            ZSql = ZSql + "'" + Hasta.Text + "')"
            spChequera = ZSql
            Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Chequera"
        ZSql = ZSql + " Where Chequera.Codigo = " + "'" + Codigo.Text + "'"
        spChequera = ZSql
        Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
        If rstChequera.RecordCount > 0 Then
            rstChequera.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE Chequera"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spChequera = ZSql
                Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Banco.Text = ""
    DesBanco.Caption = ""
    Fecha.Text = "  /  /    "
    Desde.Text = ""
    Hasta.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Chequera"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        rstChequera.MoveLast
        ZUltimo = IIf(IsNull(rstChequera!CodigoMayor), "0", rstChequera!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstChequera.Close
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgChequera.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Lista_Click()
    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desdefecha.SetFocus
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Banco.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Banco"
        ZSql = ZSql + " Where Banco.Banco = " + "'" + Banco.Text + "'"
        spBanco = ZSql
        Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
        If rstBanco.RecordCount > 0 Then
            DesBanco.Caption = rstBanco!Nombre
            Desde.SetFocus
            rstBanco.Close
                Else
            DesBanco.Caption = ""
            Banco.SetFocus
        End If
        
    End If
    
    If KeyAscii = 27 Then
        Banco.Text = ""
        DesBanco.Caption = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Chequera"
            ZSql = ZSql + " Where Chequera.Codigo = " + "'" + Codigo.Text + "'"
            spChequera = ZSql
            Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
            If rstChequera.RecordCount > 0 Then
                rstChequera.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
        End If
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Desdefecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefecha.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaFecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
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
        Case 0
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
            Banco.Text = WIndice.List(Indice)
            Call Banco_KeyPress(13)
            
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    Codigo.Text = ""
    Fecha.Text = "  /  /    "
    Banco.Text = ""
    DesBanco.Caption = ""
    Desde.Text = ""
    Hasta.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Chequera"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        rstChequera.MoveLast
        ZUltimo = IIf(IsNull(rstChequera!CodigoMayor), "0", rstChequera!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstChequera.Close
    End If
    
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
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Banco_DblClick()

    Opcion.Clear
    Opcion.AddItem "Banco"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Banco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdeFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaFecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Panta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impresora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 113
            Call cmdDelete_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 116
            Call Primer_Click
        Case 117
            Call Anterior_Click
        Case 118
            Call Siguiente_Click
        Case 119
            Call Ultimo_Click
        Case 120
            Call Lista_Click
        Case 121
            Call cmdClose_Click
        Case 122
            Call Acepta_Click
        Case 123
            Call Cancela_Click
        Case Else
    End Select
End Sub















Private Sub Anterior_Click()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Chequera"
    ZSql = ZSql + " Where Chequera.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Chequera.Codigo"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        With rstChequera
            .MoveLast
            Codigo.Text = rstChequera!Codigo
        End With
        rstChequera.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Tipo de Proveedores")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM Chequera"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        rstChequera.MoveFirst
        ZUltimo = IIf(IsNull(rstChequera!CodigoMenor), "0", rstChequera!CodigoMenor)
        Codigo.Text = ZUltimo
        rstChequera.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM Chequera"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        rstChequera.MoveLast
        ZUltimo = IIf(IsNull(rstChequera!CodigoMayor), "0", rstChequera!CodigoMayor)
        Codigo.Text = ZUltimo
        rstChequera.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Chequera"
    ZSql = ZSql + " Where Chequera.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by Chequera.Codigo"
    spChequera = ZSql
    Set rstChequera = db.OpenRecordset(spChequera, dbOpenSnapshot, dbSQLPassThrough)
    If rstChequera.RecordCount > 0 Then
        With rstChequera
            .MoveFirst
            Codigo.Text = rstChequera!Codigo
        End With
        rstChequera.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Tipo de Proveedores")
    End If
End Sub



