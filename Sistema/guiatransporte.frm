VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgGuiaTransporte 
   AutoRedraw      =   -1  'True
   Caption         =   "Guia de Transporte"
   ClientHeight    =   8385
   ClientLeft      =   195
   ClientTop       =   390
   ClientWidth     =   11550
   LinkTopic       =   "Form2"
   ScaleHeight     =   8385
   ScaleWidth      =   11550
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   11415
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
         Left            =   2760
         MouseIcon       =   "guiatransporte.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "guiatransporte.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4080
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
         Left            =   1200
         MouseIcon       =   "guiatransporte.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "guiatransporte.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Graba los Datos Ingresados"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Expreso 
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
         TabIndex        =   29
         Text            =   " "
         Top             =   3000
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
         TabIndex        =   1
         Text            =   " "
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Facturas 
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
         Left            =   5160
         MaxLength       =   50
         TabIndex        =   24
         Text            =   " "
         Top             =   3360
         Width           =   6015
      End
      Begin VB.TextBox Valor 
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
         TabIndex        =   21
         Text            =   " "
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox Provincia 
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
         TabIndex        =   19
         Text            =   " "
         Top             =   2640
         Width           =   9255
      End
      Begin VB.TextBox Cliente 
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
         Top             =   840
         Width           =   975
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
         Left            =   4440
         MouseIcon       =   "guiatransporte.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "guiatransporte.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Consulta de Datos"
         Top             =   4080
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
         Left            =   6120
         MouseIcon       =   "guiatransporte.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "guiatransporte.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salida"
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox Bultos 
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
         MaxLength       =   6
         TabIndex        =   10
         Text            =   " "
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Razon 
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
         TabIndex        =   9
         Text            =   " "
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox Cuit 
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
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   8
         Text            =   " "
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Direccion 
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
         Top             =   1920
         Width           =   9255
      End
      Begin VB.TextBox Localidad 
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
         TabIndex        =   6
         Text            =   " "
         Top             =   2280
         Width           =   9255
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   5040
         TabIndex        =   27
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
      Begin VB.Label DesExpreso 
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
         TabIndex        =   32
         Top             =   3000
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
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "Numero Despacho"
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
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Facturas"
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
         Left            =   4080
         TabIndex        =   25
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Expreso"
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
         TabIndex        =   23
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Valor Declarado"
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
         TabIndex        =   22
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia"
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
         Top             =   2640
         Width           =   1575
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
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Razon Social"
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
         TabIndex        =   17
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Bultos"
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
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Cuit"
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
         Left            =   4080
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion"
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
         TabIndex        =   14
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Localidad"
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
         TabIndex        =   13
         Top             =   2280
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
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   7815
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9000
      Top             =   5520
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
      TabIndex        =   3
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
      ItemData        =   "guiatransporte.frx":2D30
      Left            =   120
      List            =   "guiatransporte.frx":2D37
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgGuiaTransporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZProvincia(100) As String

Private Sub CmdLimpiar_Click()

    Codigo.Text = "1"
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = ""
    Razon.Text = ""
    Bultos.Text = ""
    Cuit.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Provincia.Text = ""
    Expreso.Text = ""
    Valor.Text = ""
    Facturas.Text = ""
    DesExpreso.Caption = ""
 
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM GuiaTransporte"
    spGuiaTransporte = ZSql
    Set rstGuiaTransporte = db.OpenRecordset(spGuiaTransporte, dbOpenSnapshot, dbSQLPassThrough)
    If rstGuiaTransporte.RecordCount > 0 Then
        rstGuiaTransporte.MoveLast
        ZUltimo = IIf(IsNull(rstGuiaTransporte!CodigoMayor), "0", rstGuiaTransporte!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstGuiaTransporte.Close
    End If
    
    Cliente.SetFocus

End Sub

Private Sub Proceso_Click()

    ZZOrdFecha = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM GuiaTransporte"
    ZSql = ZSql + " Where GuiaTransporte.Codigo = " + "'" + Codigo.Text + "'"
    spGuiaTransporte = ZSql
    Set rstGuiaTransporte = db.OpenRecordset(spGuiaTransporte, dbOpenSnapshot, dbSQLPassThrough)
    If rstGuiaTransporte.RecordCount > 0 Then
    
        rstGuiaTransporte.Close
    
        ZSql = ""
        ZSql = ZSql + "UPDATE GuiaTransporte SET "
        ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
        ZSql = ZSql + " Cliente = " + "'" + Cliente.Text + "',"
        ZSql = ZSql + " Razon = " + "'" + Razon.Text + "',"
        ZSql = ZSql + " Bultos = " + "'" + Bultos.Text + "',"
        ZSql = ZSql + " Cuit = " + "'" + Cuit.Text + "',"
        ZSql = ZSql + " Direccion = " + "'" + Direccion.Text + "',"
        ZSql = ZSql + " Localidad = " + "'" + Localidad.Text + "',"
        ZSql = ZSql + " Provincia = " + "'" + Provincia.Text + "',"
        ZSql = ZSql + " Expreso = " + "'" + Expreso.Text + "',"
        ZSql = ZSql + " Valor = " + "'" + Valor.Text + "',"
        ZSql = ZSql + " Facturas = " + "'" + Facturas.Text + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        
        spGuiaTransporte = ZSql
        Set rstGuiaTransporte = db.OpenRecordset(spGuiaTransporte, dbOpenSnapshot, dbSQLPassThrough)
    
            Else

        ZSql = ""
        ZSql = ZSql + "INSERT INTO GuiaTransporte ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Razon ,"
        ZSql = ZSql + "Bultos ,"
        ZSql = ZSql + "Cuit ,"
        ZSql = ZSql + "Direccion ,"
        ZSql = ZSql + "Localidad ,"
        ZSql = ZSql + "Provincia ,"
        ZSql = ZSql + "Expreso ,"
        ZSql = ZSql + "Valor ,"
        ZSql = ZSql + "Facturas )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Codigo.Text + "',"
        ZSql = ZSql + "'" + Fecha.Text + "',"
        ZSql = ZSql + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + "'" + Cliente.Text + "',"
        ZSql = ZSql + "'" + Razon.Text + "',"
        ZSql = ZSql + "'" + Bultos.Text + "',"
        ZSql = ZSql + "'" + Cuit.Text + "',"
        ZSql = ZSql + "'" + Direccion.Text + "',"
        ZSql = ZSql + "'" + Localidad.Text + "',"
        ZSql = ZSql + "'" + Provincia.Text + "',"
        ZSql = ZSql + "'" + Expreso.Text + "',"
        ZSql = ZSql + "'" + Valor.Text + "',"
        ZSql = ZSql + "'" + Facturas.Text + "')"
                                
        spGuiaTransporte = ZSql
        Set rstGuiaTransporte = db.OpenRecordset(spGuiaTransporte, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Expreso"
    ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
    spExpreso = ZSql
    Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
    If rstExpreso.RecordCount > 0 Then
        ZZDireccionExpreso = rstExpreso!Direccion
        rstExpreso.Close
            Else
        ZZDireccionExpreso = ""
    End If
    

    Rem Open "lpt1" For Output As #1
    Open "dada.txt" For Output As #1

    For Ciclo = 1 To 3

        Print #1, "=============================================================================================="
        
        Print #1, Tab(5); Chr$(14) + "CELUGAMA";
        Print #1, "   "; Chr$(14) + "Guia de Transporte"
        
        Print #1, Tab(7); "Sarmiento 5535"
        
        Print #1, Tab(7); "(1653) Villa Ballester";
        Print #1, Tab(50); "Numero : " + Codigo.Text
        
        Print #1, Tab(7); "Republica Argentina";
        Print #1, Tab(50); "Fecha : " + Fecha.Text
        
        Print #1, Tab(7); "Tel. 4768-1775/4764-4786"
        
        Print #1, Tab(7); "        4767-5364"
        
        Print #1, Tab(7); "    Fax: 4764-2968";
        Print #1, Tab(50); "Cuit : 20-63767162-9"

        Print #1, "=============================================================================================="
        
        Print #1, Tab(7); "Señor/es: "; Razon.Text; " "; Cliente.Text
        
        Print #1, Tab(7); "Direccion: "; Left$(Direccion.Text, 40);
        Print #1, Tab(50); "Cuit: "; Cuit.Text
        
        Print #1, Tab(15); Trim(Localidad.Text) + "   " + Trim(Provincia.Text)
        
        Print #1, "-----------------------------------------------------------------------------------------------"
        
        Print #1, Tab(5); "Remitimos a ustedes lo siguiente: Facturas : "; Facturas.Text
        
        Print #1, ""
        
        Print #1, Tab(5); "Articulo de Perfumeria"
        Print #1, Tab(5); "Despacho Nro.: " + Chr$(14) + Codigo.Text
        
        Print #1, "-----------------------------------------------------------------------------------------------"
        
        Print #1, Tab(7); "Cantidad de Bultos: "; Bultos.Text;
        Print #1, Tab(50); "Valor Declarado: $";
        Print #1, Alinea("###,###.##", Valor.Text)

        Print #1, "=============================================================================================="
        
        Print #1, ""
        
        Print #1, Tab(2); "Transportista:" + Left$(DesExpreso.Caption, 30);
        Print #1, Tab(50); ".....................";
        
        Print #1, Tab(2); "Direccion:" + Left$(ZZDireccionExpreso, 45);
        Print #1, Tab(50); "Recibo Conforme"
        
        Print #1, "=============================================================================================="
        
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        Print #1, ""
        
    Next Ciclo
    
    Close #1
    
    Listado.WindowTitle = "Guia de Transporte"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
        
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT GuiaTransporte.Codigo, GuiaTransporte.Fecha, GuiaTransporte.OrdFecha, GuiaTransporte.Cliente, GuiaTransporte.Razon, GuiaTransporte.Bultos, GuiaTransporte.Cuit, GuiaTransporte.Direccion, GuiaTransporte.Localidad, GuiaTransporte.Provincia, GuiaTransporte.Valor, GuiaTransporte.Facturas, " _
         + "Expreso.Nombre, Expreso.Direccion " _
        + "From " _
        + "Celugama.dbo.GuiaTransporte GuiaTransporte, " _
        + "Celugama.dbo.Expreso Expreso " _
        + "Where " _
        + "GuiaTransporte.Expreso = Expreso.Codigo AND " _
        + "GuiaTransporte.Codigo >= " + Codigo.Text + " AND " _
        + "GuiaTransporte.Codigo <= " + Codigo.Text
    
    Listado.GroupSelectionFormula = "{GuiaTransporte.Codigo} in " + Codigo.Text + " to " + Codigo.Text
    Listado.SelectionFormula = "{GuiaTransporte.Codigo} in " + Codigo.Text + " to " + Codigo.Text
    
    Listado.Destination = 1
    
    Listado.ReportFileName = "ImpreGuia.rpt"
    Listado.Action = 1
    
    Rem Listado.ReportFileName = "ImpreGuiaII.rpt"
    Rem Listado.Action = 1
    
    Call CmdLimpiar_Click
    
End Sub

Private Sub Cancela_click()
    PrgGuiaTransporte.Hide
    Unload Me
    Menu4.Show
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM GuiaTransporte"
            ZSql = ZSql + " Where GuiaTransporte.Codigo = " + "'" + Codigo.Text + "'"
            spGuiaTransporte = ZSql
            Set rstGuiaTransporte = db.OpenRecordset(spGuiaTransporte, dbOpenSnapshot, dbSQLPassThrough)
            If rstGuiaTransporte.RecordCount > 0 Then
                Fecha.Text = rstGuiaTransporte!Fecha
                Cliente.Text = rstGuiaTransporte!Cliente
                Razon.Text = rstGuiaTransporte!Razon
                Bultos.Text = rstGuiaTransporte!Bultos
                Cuit.Text = rstGuiaTransporte!Cuit
                Direccion.Text = Trim(rstGuiaTransporte!Direccion)
                Localidad.Text = Trim(rstGuiaTransporte!Localidad)
                Provincia.Text = rstGuiaTransporte!Provincia
                Expreso.Text = rstGuiaTransporte!Expreso
                Valor.Text = rstGuiaTransporte!Valor
                Facturas.Text = rstGuiaTransporte!Facturas
                rstGuiaTransporte.Close
            End If
            Cliente.SetFocus
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Expreso"
            ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
            spExpreso = ZSql
            Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
            If rstExpreso.RecordCount > 0 Then
                DesExpreso.Caption = rstExpreso!Nombre
                rstExpreso.Close
                    Else
                DesExpreso.Caption = ""
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Trim(Cliente.Text) <> "" Then
            Auxi = UCase(Left$(Cliente.Text, 1))
            Auxi1 = Mid$(Cliente.Text, 2, 5)
            Call Ceros(Auxi1, 3)
            Cliente.Text = Auxi + "-" + Auxi1
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            Razon.Text = Trim(rstCliente!Razon)
            Expreso.Text = rstCliente!Expreso
            ZZProvincia = rstCliente!Provincia
            Provincia.Text = ZProvincia(ZZProvincia)
            Direccion.Text = Trim(rstCliente!Direccion)
            Localidad.Text = Trim(rstCliente!Localidad)
            Cuit.Text = rstCliente!Cuit
            rstCliente.Close
            Bultos.SetFocus
        End If
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Expreso"
        ZSql = ZSql + " Where Expreso.Codigo = " + "'" + Expreso.Text + "'"
        spExpreso = ZSql
        Set rstExpreso = db.OpenRecordset(spExpreso, dbOpenSnapshot, dbSQLPassThrough)
        If rstExpreso.RecordCount > 0 Then
            DesExpreso.Caption = rstExpreso!Nombre
            rstExpreso.Close
                Else
            DesExpreso.Caption = ""
        End If
        
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM HistorialCliente"
        ZSql = ZSql + " Where HistorialCliente.Cliente = " + "'" + Cliente.Text + "'"
        spHistorialCliente = ZSql
        Set rstHistorialCliente = db.OpenRecordset(spHistorialCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstHistorialCliente.RecordCount > 0 Then
            rstHistorialCliente.Close
            ZZPasaCliente = Cliente.Text
            ZZPasaProceso = 1
            PrgHistorialClienteConsulta.Show
        End If
        
        
    End If
    
    If KeyAscii = 27 Then
        Cliente.Text = ""
    End If
    
End Sub

Private Sub Bultos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Valor.SetFocus
    End If
    If KeyAscii = 27 Then
        Bultos.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Facturas.SetFocus
    End If
    If KeyAscii = 27 Then
        Valor.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Facturas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Bultos.SetFocus
    End If
    If KeyAscii = 27 Then
        Facturas.Text = ""
    End If
End Sub

Sub Form_Load()

    ZProvincia(0) = "Capital Federal"
    ZProvincia(1) = "Buenos Aires"
    ZProvincia(2) = "Catamarca"
    ZProvincia(3) = "Cordoba"
    ZProvincia(4) = "Corrientes"
    ZProvincia(5) = "Chaco"
    ZProvincia(6) = "Chubut"
    ZProvincia(7) = "Entre Rios"
    ZProvincia(8) = "Formosa"
    ZProvincia(9) = "Jujuy"
    ZProvincia(10) = "La Pampa"
    ZProvincia(11) = "La Rioja"
    ZProvincia(12) = "Mendoza"
    ZProvincia(13) = "Misiones"
    ZProvincia(14) = "Neuquen"
    ZProvincia(15) = "Rio Negro"
    ZProvincia(16) = "Salta"
    ZProvincia(17) = "San Juan"
    ZProvincia(18) = "San Luis"
    ZProvincia(19) = "Santa Cruz"
    ZProvincia(20) = "Santa Fe"
    ZProvincia(21) = "Santiago del Estero"
    ZProvincia(22) = "Tucuman"
    ZProvincia(23) = "Tierra del Fuego"
    ZProvincia(24) = "Exterior"
    ZProvincia(25) = ""

    Codigo.Text = "1"
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cliente.Text = ""
    Razon.Text = ""
    Bultos.Text = ""
    Cuit.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Provincia.Text = ""
    Expreso.Text = ""
    Valor.Text = ""
    Facturas.Text = ""
 
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM GuiaTransporte"
    spGuiaTransporte = ZSql
    Set rstGuiaTransporte = db.OpenRecordset(spGuiaTransporte, dbOpenSnapshot, dbSQLPassThrough)
    If rstGuiaTransporte.RecordCount > 0 Then
        rstGuiaTransporte.MoveLast
        ZUltimo = IIf(IsNull(rstGuiaTransporte!CodigoMayor), "0", rstGuiaTransporte!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstGuiaTransporte.Close
    End If
    
    Frame2.Visible = True
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
    Cliente.Text = WIndice.List(Indice)
    
    Ayuda.Visible = False
    Call Cliente_KeyPress(13)
    
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
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Razon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Bultos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cuit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Direccion_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Localidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Provincia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Expreso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Valor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Facturas_KeyDown(KeyCode As Integer, Shift As Integer)
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
            Call Cancela_click
        Case Else
    End Select
End Sub













