VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form prgActualizacionGeneralInsumos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizacion General de Insumos"
   ClientHeight    =   6525
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   11790
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   11790
   Visible         =   0   'False
   Begin VB.TextBox HastaArt 
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
      Left            =   3960
      MaxLength       =   10
      TabIndex        =   18
      Text            =   " "
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox DesdeArt 
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   17
      Text            =   " "
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox Costo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
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
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   15
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Proveedor 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
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
      Left            =   360
      MouseIcon       =   "actualizaciongeneralInsumos.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "actualizaciongeneralInsumos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Graba los Datos Ingresados"
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
      Left            =   1320
      MouseIcon       =   "actualizaciongeneralInsumos.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "actualizaciongeneralInsumos.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   10
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
      Left            =   2280
      MouseIcon       =   "actualizaciongeneralInsumos.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "actualizaciongeneralInsumos.frx":19A2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Consulta de Datos"
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
      Left            =   3240
      MouseIcon       =   "actualizaciongeneralInsumos.frx":21E4
      MousePointer    =   99  'Custom
      Picture         =   "actualizaciongeneralInsumos.frx":24EE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salida"
      Top             =   1920
      Width           =   855
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
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   8655
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   240
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Grupo 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9240
      Top             =   0
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
      Left            =   9000
      TabIndex        =   1
      Top             =   480
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
      ItemData        =   "actualizaciongeneralInsumos.frx":2D30
      Left            =   120
      List            =   "actualizaciongeneralInsumos.frx":2D37
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   8655
   End
   Begin VB.Label Label2 
      Caption         =   "Desde Hasta Insumo"
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
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label21 
      Caption         =   "% de Aumento"
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
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Proveedor"
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
      TabIndex        =   14
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label DesProveedor 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   13
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label DesGrupo 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
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
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label8 
      Caption         =   "Grupo"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   1935
   End
End
Attribute VB_Name = "prgActualizacionGeneralInsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZCosto As Double
Dim ZZImporte As Double

Dim ZVector(10000) As String

Private Sub cmdAdd_Click()
    
    ZLugar = 0
    Erase ZVector

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Insumo"
    ZSql = ZSql + " Where Insumo.Codigo >= " + "'" + DesdeArt.Text + "'"
    ZSql = ZSql + " and Insumo.Codigo <= " + "'" + HastaArt.Text + "'"
    If Val(Grupo.Text) <> 0 Then
        ZSql = ZSql + " and Insumo.Linea = " + "'" + Grupo.Text + "'"
    End If
    If Val(Proveedor.Text) <> 0 Then
        ZSql = ZSql + " and Insumo.Proveedor = " + "'" + Proveedor.Text + "'"
    End If
    spInsumo = ZSql
    Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    If rstInsumo.RecordCount > 0 Then
        With rstInsumo
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstInsumo!Codigo
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInsumo.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZVector(Ciclo)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Insumo"
        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + ZCodigo + "'"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        If rstInsumo.RecordCount > 0 Then

            ZZCosto = rstInsumo!Costo
            
            If Val(Costo.Text) <> 0 Then
                ZZImporte = ZZCosto * (Val(Costo.Text) / 100)
                Call Redondeo(ZZImporte)
                ZZCosto = ZZCosto + ZZImporte
                Call Redondeo(ZZCosto)
            End If
            
            ZZFechaCosto = rstInsumo!FechaCosto
            ZZCostoAnterior = Str$(rstInsumo!CostoAnterior)
            ZZFechaCostoAnterior = rstInsumo!FechaCostoAnterior
            ZZCostoActual = rstInsumo!Costo
            
            If ZZCosto <> ZZCostoActual Then
            
                ZZFechaCostoAnterior = rstInsumo!FechaCosto
                ZZCostoAnterior = Str$(ZZCostoActual)
                ZZFechaCosto = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                ZZOrdFechaCosto = Right$(ZZFechaCosto, 4) + Mid$(ZZFechaCosto, 4, 2) + Left$(ZZFechaCosto, 2)
        
                rstInsumo.Close
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Insumo SET "
                ZSql = ZSql + " CostoAnterior = " + "'" + ZZCostoAnterior + "',"
                ZSql = ZSql + " FechaCostoAnterior = " + "'" + ZZFechaCostoAnterior + "',"
                ZSql = ZSql + " Costo = " + "'" + Str$(ZZCosto) + "',"
                ZSql = ZSql + " FechaCosto = " + "'" + ZZFechaCosto + "',"
                ZSql = ZSql + " OrdFechaCosto = " + "'" + ZZOrdFechaCosto + "'"
                ZSql = ZSql + " Where Codigo = " + "'" + ZCodigo + "'"
                spInsumo = ZSql
                Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
    Next Ciclo
        
    m$ = "El proceso ha finalizado con exito"
    a% = MsgBox(m$, 0, "Actualizacion de Precios de Insumos")
    
    Call CmdLimpiar_Click
    Grupo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Grupo.Text = ""
    Proveedor.Text = ""
    DesGrupo.Caption = ""
    DesProveedor.Caption = ""
    DesdeArt.Text = ""
    HastaArt.Text = ""
    
    Costo.Text = ""
    
    Grupo.SetFocus
End Sub

Private Sub cmdClose_Click()
    prgActualizacionGeneralInsumos.Hide
    Unload Me
    Menu3.Show
End Sub

Private Sub Grupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Grupo.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM LineaInsumo"
            ZSql = ZSql + " Where LineaInsumo.Linea = " + "'" + Grupo.Text + "'"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineaInsumo.RecordCount > 0 Then
                DesGrupo.Caption = rstLineaInsumo!Nombre
                rstLineaInsumo.Close
                Proveedor.SetFocus
            End If
                Else
            DesGrupo.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        Grupo.Text = ""
        DesGrupo.Caption = ""
    End If
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Proveedor.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + Proveedor.Text + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                DesProveedor.Caption = rstProveedor!Nombre
                rstProveedor.Close
                DesdeArt.SetFocus
            End If
                Else
            Proveedor.Text = ""
            DesProveedor.Caption = ""
        End If
    End If
    If KeyAscii = 27 Then
        Proveedor.Text = ""
    End If
End Sub

Private Sub DesdeArt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaArt.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaArt.Text = ""
    End If
End Sub

Private Sub HastaArt_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Costo.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaArt.Text = ""
    End If
End Sub

Private Sub Costo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Costo.Text) <> 0 Then
            Costo.Text = Pusing("###,###.##", Costo.Text)
        End If
        Grupo.SetFocus
    End If
    If KeyAscii = 27 Then
        Costo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub
    
Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Grupos"
     Opcion.AddItem "Proveedores"

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
            ZSql = ZSql + " FROM LineaInsumo"
            ZSql = ZSql + " Order by LineaInsumo.Linea"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineaInsumo.RecordCount > 0 Then
                With rstLineaInsumo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstLineaInsumo!Linea) + " " + rstLineaInsumo!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstLineaInsumo!Linea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLineaInsumo.Close
            End If
            
        Case 1
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
                            IngresaItem = Str$(rstProveedor!Proveedor) + " " + rstProveedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstProveedor!Proveedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedor.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
    Rem Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Grupo.Text = WIndice.List(Indice)
            Call Grupo_KeyPress(13)
            
        Case 1
            Indice = Pantalla.ListIndex
            Proveedor.Text = WIndice.List(Indice)
            Call Proveedor_KeyPress(13)
            
        Case Else
    End Select
    
End Sub


Sub Form_Load()

    Grupo.Text = ""
    DesGrupo.Caption = ""
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Costo.Text = ""
    DesdeArt.Text = ""
    HastaArt.Text = ""
    
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
            ZSql = ZSql + " FROM LineaInsumo"
            ZSql = ZSql + " Where LineaInsumo.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by LineaInsumo.Linea"
            spLineaInsumo = ZSql
            Set rstLineaInsumo = db.OpenRecordset(spLineaInsumo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLineaInsumo.RecordCount > 0 Then
                With rstLineaInsumo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Linea) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Linea
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstLineaInsumo.Close
            End If
            
        Case 1
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
                            IngresaItem = Str$(!Proveedor) + " " + !Nombre
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
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Grupo_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub

Private Sub Proveedor_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem ""
    Opcion.AddItem ""

    Opcion.ListIndex = 1
    
    Call Opcion_Click

End Sub
Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Grupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Proveedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdeArt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaArt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Costo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub


Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case 114
            Call CmdLimpiar_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub











































