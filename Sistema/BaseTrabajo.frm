VERSION 5.00
Begin VB.Form BaseTrabajo 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "BaseTrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZArti(10000) As String

Private Producto As String
Private Costo As Double

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
    If Val(Desde.Text) = 0 Then
        Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
        Hasta.Text = "0"
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " PrecioList = Minimo - Stock"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase ZArti
    ZLugar = 0
    
    If Tipo.ListIndex = 0 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Proveedor >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Articulo.Proveedor <= " + "'" + Hasta.Text + "'"
        ZSql = ZSql + " and Articulo.PrecioList > 0"
            Else
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Grupo >= " + "'" + Desde.Text + "'"
        ZSql = ZSql + " and Articulo.Grupo <= " + "'" + Hasta.Text + "'"
        ZSql = ZSql + " and Articulo.PrecioList > 0"
    End If
    
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    ZArti(ZLugar) = rstArticulo!Codigo
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZArticulo = ZArti(Ciclo)
        ZEmbarque = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM OrdenImportacion"
        ZSql = ZSql + " Where OrdenImportacion.Articulo = " + "'" + ZArticulo + "'"
        ZSql = ZSql + " and OrdenImportacion.Estado = 0"
        spOrdenImportacion = ZSql
        Set rstOrdenImportacion = db.OpenRecordset(spOrdenImportacion, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrdenImportacion.RecordCount > 0 Then
            With rstOrdenImportacion
                .MoveFirst
                Do
                    If .EOF = False Then
                        ZEmbarque = ZEmbarque + rstOrdenImportacion!Cantidad
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstOrdenImportacion.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + " Embarque = " + "'" + Str$(ZEmbarque) + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Ciclo
    
    
    Listado.WindowTitle = "Listado de Reposicion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Tipo.ListIndex = 0 Then
    
        Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Minimo, Articulo.Stock, Articulo.PrecioList, Articulo.Embarque,  " _
                + "Auxiliar.Nombre, " _
                + "Proveedor.Nombre " _
                + "From " _
                + DSQ + ".dbo.Articulo Articulo, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Proveedor Proveedor " _
                + "Where " _
                + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "Articulo.Proveedor = Proveedor.Proveedor AND " _
                + "Articulo.Grupo <> 999 AND " _
                + "Articulo.Proveedor >= " + Desde.Text + " AND " _
                + "Articulo.Proveedor <= " + Hasta.Text + " AND " _
                + "Articulo.PrecioList > 0"
            
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
                
        Uno = "{Articulo.PrecioList} > 0"
        Dos = " and {Articulo.Proveedor} in " + Desde.Text + " to " + Hasta.Text
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
            
        Listado.ReportFileName = "ListaStockMinimoProveedor.rpt"
        Listado.Action = 1
            
                Else
    
        Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Grupo, Articulo.Proveedor, Articulo.Stock, Articulo.PrecioList, Articulo.Embarque,  " _
                    + "Auxiliar.Nombre, " _
                    + "Familia.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar, " _
                    + DSQ + ".dbo.Familia Familia, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Articulo.CodigoEmpresa = Auxiliar.Empresa AND " _
                    + "Articulo.Grupo = Familia.Codigo AND " _
                    + "Articulo.Proveedor = Proveedor.Proveedor AND " _
                    + "Articulo.Grupo >= " + Desde.Text + " AND " _
                    + "Articulo.Grupo <= " + Hasta.Text + " AND " _
                    + "Articulo.PrecioList > 0"
            
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
                
        Uno = "{Articulo.PrecioList} > 0"
        Dos = " and {Articulo.Grupo} in " + Desde.Text + " to " + Hasta.Text
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
            
        Listado.ReportFileName = "ListaStockMinimoGrupo.rpt"
        Listado.Action = 1
                
    End If
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaStockMinimo.Hide
    Unload Me
    Menu231.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Proveedor"
    Tipo.AddItem "GRupo"
    
    Tipo.ListIndex = 0
    
    NombreI.Caption = "Desde Proveedor"
    NombreII.Caption = "Hasta Proveedor"

    Desde.Text = ""
    Hasta.Text = ""
    Frame2.Visible = True
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
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_click
        Case Else
    End Select
End Sub

Private Sub Tipo_click()
    If Tipo.ListIndex = 0 Then
        NombreI.Caption = "Desde Proveedor"
        NombreII.Caption = "Hasta Proveedor"
            Else
        NombreI.Caption = "Desde Grupo"
        NombreII.Caption = "Hasta Grupo"
    End If
End Sub


