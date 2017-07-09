VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form prgCierreStock 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Mensual de Stock"
   ClientHeight    =   2910
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   2910
   ScaleWidth      =   11790
   Visible         =   0   'False
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
      Left            =   3360
      MouseIcon       =   "cierrestock.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "cierrestock.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   720
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
      Left            =   5040
      MouseIcon       =   "cierrestock.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "cierrestock.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salida"
      Top             =   720
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   8160
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Articulo.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Clientes"
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
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
End
Attribute VB_Name = "prgCierreStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZPrecio  As Double
Dim ZZMargen As Double
Dim ZZImporte As Double
Dim ZZCosto As Double

Dim ZVector(10000) As String

Private Sub cmdAdd_Click()

    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Articulo SET "
    Rem ZSql = ZSql + " Stock = StockAnterior"
    Rem spArticulo = ZSql
    Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " StockAnterior = Stock"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Entradas = 0"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Salidas = 0"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    m$ = "El proceso ha finalizado con exito"
    a% = MsgBox(m$, 0, "Cierre de mes")
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdClose_Click()
    prgCierreStock.Hide
    Unload Me
    Menu2.Show
End Sub











































