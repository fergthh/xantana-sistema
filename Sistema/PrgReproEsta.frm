VERSION 5.00
Begin VB.Form PrgReproEstA 
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
Attribute VB_Name = "PrgReproEstA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPlanilla(10000, 5) As String

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
    
    WAno = Right$(DesdeFec.Text, 4)
    WMes = Mid$(DesdeFec.Text, 4, 2)
    WDia = Left$(DesdeFec.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFec.Text, 4)
    WMes = Mid$(HastaFec.Text, 4, 2)
    WDia = Left$(HastaFec.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WTitulo = "del " + DesdeFec.Text + " al " + HastaFec.Text
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Varios = " + "'" + WTitulo + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Estadistica SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "',"
    ZSql = ZSql + " Lista = " + "'" + "" + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    ZLugar = 0
    Erase WPlanilla
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Recibos"
    ZSql = ZSql + " Where Recibos.FechaOrd >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and Recibos.FechaOrd <= " + "'" + WHasta + "'"
    ZSql = ZSql + " and Recibos.Importe1 <> 0"
    spRecibos = ZSql
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then
        With rstRecibos
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    WPlanilla(ZLugar, 1) = rstRecibos!Tipo1
                    WPlanilla(ZLugar, 2) = rstRecibos!Letra1
                    WPlanilla(ZLugar, 3) = rstRecibos!Punto1
                    WPlanilla(ZLugar, 4) = rstRecibos!Numero1
                    WPlanilla(ZLugar, 5) = rstRecibos!Cliente
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibos.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZTipo = WPlanilla(Ciclo, 1)
        ZLetra = WPlanilla(Ciclo, 2)
        ZPunto = WPlanilla(Ciclo, 3)
        ZNumero = WPlanilla(Ciclo, 4)
        ZCliente = WPlanilla(Ciclo, 5)
        ZVendedor = 0
        ZTipoComision = 0
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + ZCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            ZVendedor = Str$(rstCliente!Vendedor)
            rstCliente.Close
        End If
    
        If Val(Desde.Text) = ZVendedor Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CtaCte"
            ZSql = ZSql + " Where CtaCte.Tipo = " + "'" + ZTipo + "'"
            ZSql = ZSql + " and CtaCte.Letra = " + "'" + ZLetra + "'"
            ZSql = ZSql + " and CtaCte.Punto = " + "'" + ZPunto + "'"
            ZSql = ZSql + " and CtaCte.Numero = " + "'" + ZNumero + "'"
            spCtaCte = ZSql
            Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
            If rstCtaCte.RecordCount > 0 Then
                ZTipoComision = rstCtaCte!Comision
                rstCtaCte.Close
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Estadistica SET "
            ZSql = ZSql + " TipoComision = " + "'" + Str$(ZTipoComision) + "',"
            ZSql = ZSql + " Vendedor = " + "'" + Str$(ZVendedor) + "',"
            ZSql = ZSql + " Lista = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Estadistica.Tipo = " + "'" + Str$(Val(ZTipo)) + "'"
            ZSql = ZSql + " and Estadistica.Letra = " + "'" + ZLetra + "'"
            ZSql = ZSql + " and Estadistica.Punto = " + "'" + Str$(Val(ZPunto)) + "'"
            ZSql = ZSql + " and Estadistica.Numero = " + "'" + Str$(Val(ZNumero)) + "'"
            spEstadistica = ZSql
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Ciclo
    
    ZSql = "UPDATE Estadistica SET "
    ZSql = ZSql + " Estadistica.Comision = Articulo.Comision"
    ZSql = ZSql + " From Estadistica, Articulo"
    ZSql = ZSql + " Where Estadistica.Articulo = Articulo.Codigo"
    ZSql = ZSql + " and Estadistica.Lista = " + "'" + "S" + "'"
    spEstadistica = ZSql
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    Listado.WindowTitle = "Listado de Comisiones por Venta"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Val(Desde.Text) = 0 Then
        Desde.Text = "0"
    End If

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CtaCte.Tipo, CtaCte.Numero, CtaCte.Cliente, CtaCte.fecha, CtaCte,OrdFecha, CtaCte.Impre, CtaCte.Neto, CtaCte.Vendedor, CtaCte.Partida, " _
                + "Cliente.Razon," _
                + "Auxiliar.Nombre, Auxiliar.Varios, " _
                + "Vendedor.Nombre, Vendedor.Comision " _
                + "From " _
                + DSQ + ".dbo.CtaCte CtaCte, " _
                + DSQ + ".dbo.Cliente Cliente, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Vendedor Vendedor " _
                + "Where " _
                + "CtaCte.Cliente = Cliente.Cliente AND " _
                + "CtaCte.CodigoEmpresa = Auxiliar.Empresa AND " _
                + "CtaCte.Vendedor = Vendedor.Codigo AND " _
                + "CtaCte.Tipo >= '01' AND " _
                + "CtaCte.Tipo <= '05' AND " _
                + "CtaCte.OrdFecha >= '" + WDesde + "' AND " _
                + "CtaCte.OrdFecha <= '" + WHasta + "' AND " _
                + "CtaCte.Vendedor >= " + Desde.Text + " AND " _
                + "CtaCte.Vendedor <= " + Hasta.Text
    
    Uno = "{CtaCte.OrdFecha} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Dos = " and {CtaCte.Tipo} in " + Chr$(34) + "01" + Chr$(34) + " to " + Chr$(34) + "05" + Chr$(34)
    Tres = " and {CtaCte.Vendedor} in " + Desde.Text + " to " + Hasta.Text
        
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    Listado.SelectionFormula = Uno + Dos + Tres
    
    If Tipo.ListIndex = 0 Then
        Listado.ReportFileName = "ListaComisiones.rpt"
            Else
        Listado.ReportFileName = "ListaComisiones.rpt"
    End If
    
    Listado.Action = 1
    
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgReproEstA.Hide
    Unload Me
    Menu41.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeFec.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(DesdeFec.Text, Auxi)
        If Auxi = "S" Then
            HastaFec.SetFocus
                Else
            DesdeFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        DesdeFec.Text = "  /  /    "
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(HastaFec.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            HastaFec.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        HastaFec.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0

    Desde.Text = ""
    DesdeFec.Text = "  /  /    "
    HastaFec.Text = "  /  /    "
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

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
    End If
End Sub

Private Sub DesdeFec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
    End If
End Sub

Private Sub HastaFec_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
    End If
End Sub












