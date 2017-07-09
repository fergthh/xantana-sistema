VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form prgPasajeCostoFuturo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Traspaso de Costo Futuro a Actual"
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
      MouseIcon       =   "PasajeCostoFuturo.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "PasajeCostoFuturo.frx":030A
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
      MouseIcon       =   "PasajeCostoFuturo.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "PasajeCostoFuturo.frx":0E56
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
Attribute VB_Name = "prgPasajeCostoFuturo"
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
    
    ZLugar = 0
    Erase ZVector

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.CostoFuturo <> 0"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstArticulo!Codigo
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZVector(Ciclo)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ZCodigo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then

            ZZPrecio = 0
            ZZCosto = rstArticulo!CostoFuturo
            
            ZZFechaCosto = rstArticulo!FechaCosto
            ZZCostoAnterior = Str$(rstArticulo!CostoAnterior)
            ZZFechaCostoAnterior = rstArticulo!FechaCostoAnterior
            ZZCostoActual = rstArticulo!Costo
            ZZMargenActual = rstArticulo!Margen
            If rstArticulo!MargenFuturo <> 0 Then
                ZZMargen = rstArticulo!MargenFuturo
                    Else
                ZZMargen = rstArticulo!Margen
            End If
            
            If ZZCosto <> 0 And ZZMargen <> 0 Then
                ZZImporte = ZZCosto * (ZZMargen / 100)
                Call Redondeo(ZZImporte)
                ZZPrecio = ZZCosto + ZZImporte
            End If
    
            If ZZCosto <> ZZCostoActual Or ZZMargen <> ZZMargenActual Then
            
                ZZFechaCostoAnterior = rstArticulo!FechaCosto
                ZZCostoAnterior = Str$(ZZCostoActual)
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
        
                rstArticulo.Close
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Articulo SET "
                ZSql = ZSql + " Margen = " + "'" + Str$(ZZMargen) + "',"
                ZSql = ZSql + " CostoAnterior = " + "'" + ZZCostoAnterior + "',"
                ZSql = ZSql + " FechaCostoAnterior = " + "'" + ZZFechaCostoAnterior + "',"
                ZSql = ZSql + " Costo = " + "'" + Str$(ZZCosto) + "',"
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "',"
                ZSql = ZSql + " Precio = " + "'" + Str$(ZZPrecio) + "',"
                ZSql = ZSql + " CostoFuturo = " + "'" + "0" + "',"
                ZSql = ZSql + " PrecioFuturo = " + "'" + "0" + "',"
                ZSql = ZSql + " MargenFuturo = " + "'" + "0" + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + ZCodigo + "'"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
    Next Ciclo
        
    m$ = "El proceso ha finalizado con exito"
    a% = MsgBox(m$, 0, "Actualizacion de Precios de Venta")
    
    Call cmdClose_Click
    
End Sub

Private Sub cmdClose_Click()
    prgPasajeCostoFuturo.Hide
    Unload Me
    Menu22.Show
End Sub











































