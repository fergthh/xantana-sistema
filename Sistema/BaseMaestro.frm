VERSION 5.00
Begin VB.Form BaseMaestro 
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
Attribute VB_Name = "BaseMaestro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Imprime_Nombre()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Lineas"
    ZSql = ZSql + " Where Lineas.Linea = " + "'" + Linea.Text + "'"
    spLineas = ZSql
    Set rstLineas = db.OpenRecordset(spLineas, dbOpenSnapshot, dbSQLPassThrough)
    If rstLineas.RecordCount > 0 Then
        DesLinea.Caption = rstLineas!Nombre
        rstLineas.Close
            Else
        DesLinea.Caption = ""
    End If
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cuenta"
    ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        Descripcion.Text = rstCuenta!Descripcion
        rstCuenta.Close
    End If
End Sub

Private Sub Acepta_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Cuenta SET "
    ZSql = ZSql + " CodigoEmpresa = " + "'" + WEmpresa + "'"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)


    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    
    Listado.WindowTitle = "Listado de Rubros"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        
        Listado.SQLQuery = "SELECT Cuenta.Cuenta, Cuenta.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Cuenta Cuenta " _
                    + "Where " _
                    + "Cuenta.Cuenta >= " + Desde.Text + " AND " _
                    + "Cuenta.Cuenta <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Cuenta.Cuenta} in " + Desde.Text + " to " + Hasta.Text
    Listado.SelectionFormula = "{Cuenta.Cuenta} in " + Desde.Text + " to " + Hasta.Text
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Cuenta.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Cuenta.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            rstCuenta.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Cuenta SET "
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "'"
            ZSql = ZSql + " Where Cuenta = " + "'" + Cuenta.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Cuenta ("
            ZSql = ZSql + "Cuenta ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Cuenta.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "')"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call CmdLimpiar_Click
        Cuenta.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Cuenta.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
        spCuenta = ZSql
        Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
        If rstCuenta.RecordCount > 0 Then
            rstCuenta.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE Cuenta"
                ZSql = ZSql + " Where Cuenta = " + "'" + Cuenta.Text + "'"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                
                Call CmdLimpiar_Click
            End If
        End If
    
    End If
    Cuenta.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    On Error GoTo WError
    
    Cuenta.Text = "1"
    Descripcion.Text = ""
    Cuenta.SetFocus
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Cuenta) as [CuentaMayor]"
    ZSql = ZSql + " FROM Cuenta"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        rstCuenta.MoveLast
        ZUltimo = IIf(IsNull(rstCuenta!CuentaMayor), "0", rstCuenta!CuentaMayor)
        Cuenta.Text = ZUltimo + 1
        rstCuenta.Close
    End If
    If Val(Cuenta.Text) = 0 Then
        Cuenta.Text = "1"
    End If
    
    Exit Sub
    
WError:

    Resume Next
        
    
End Sub

Private Sub cmdClose_Click()
    PrgCuenta.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    Rem If KeyAscii = 13 Then
    Rem     Cuenta.SetFocus
    Rem End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Cuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cuenta.Text = "" Then
            Call Cuenta_DblClick
                Else
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Cuenta = " + "'" + Cuenta.Text + "'"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                rstCuenta.Close
                Call Imprime_Datos
                    Else
                WCuenta = Cuenta.Text
                CmdLimpiar_Click
                Cuenta.Text = WCuenta
            End If
            
            Descripcion.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cuenta.Text = ""
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
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    Cuenta.Text = "1"
    Descripcion.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Cuenta) as [CuentaMayor]"
    ZSql = ZSql + " FROM Cuenta"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        rstCuenta.MoveLast
        ZUltimo = IIf(IsNull(rstCuenta!CuentaMayor), "0", rstCuenta!CuentaMayor)
        Cuenta.Text = ZUltimo + 1
        rstCuenta.Close
    End If
        
End Sub




Rem RUTINAS PARA AYUDA

Private Sub Consulta_Click()
    Opcion.Visible = False
    Opcion.Clear
    Opcion.AddItem "Rubros"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    Rem Call Opcion_Click
End Sub

Private Sub Cuenta_DblClick()
    Opcion.Visible = False
    Opcion.Clear
    Opcion.AddItem "Rubros"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
End Sub

Private Sub Opcion_Click()
    Ayuda.Text = ""
    Ayuda.Visible = True
    Call aYUDA_Keypress(13)
    WAyuda.Col = 1
    WAyuda.Row = 1
    WAyuda.SetFocus
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Opcion.Visible = False
        WAyuda.Visible = True
        Call Limpia_Ayuda
        Lugar = 0
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
        XIndice = Opcion.ListIndex
    
        Select Case XIndice
            Case 0
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Cuenta"
                ZSql = ZSql + " Order by Cuenta.Cuenta"
                spCuenta = ZSql
                Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
                If rstCuenta.RecordCount > 0 Then
                    With rstCuenta
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                da = Len(rstCuenta!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstCuenta!Descripcion), aa, WEspacios) Then
                                        Lugar = Lugar + 1
                                        WAyuda.Row = Lugar
                                        WAyuda.Col = 1
                                        WAyuda.Text = !Cuenta
                                        WAyuda.Col = 2
                                        WAyuda.Text = !Descripcion
                                        IngresaItem = !Cuenta
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next aa
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCuenta.Close
                End If
                
            Case Else
        End Select
    
        WAyuda.Col = 1
        WAyuda.Row = 1
        WAyuda.SetFocus
    End If
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
End Sub

Private Sub wayuda_Click()
    Call WAyuda_KeyPress(13)
End Sub

Private Sub WAyuda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Opcion.Visible = False
    WAyuda.Visible = False
    Ayuda.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Indice = WAyuda.Row - 1
            Cuenta.Text = WIndice.List(Indice)
            Call Cuenta_KeyPress(13)
                    
        Case Else
    End Select
        Else
    If KeyAscii > 48 Then
        Ayuda.Text = Chr$(KeyAscii)
        Ayuda.SelStart = 2
    End If
    Ayuda.SetFocus
    End If
End Sub

Private Sub Limpia_Ayuda()

    WAyuda.Clear

    Rem ponga la wvector1 en negritas
    WAyuda.Font.Bold = True


    ' Establesco loa Valores de la wvector1
    
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            WAyuda.FixedCols = 1
            WAyuda.Cols = 3
            WAyuda.FixedRows = 1
            WAyuda.Rows = 10001
    
            WAyuda.ColWidth(0) = 200
            WAyuda.Row = 0
            For Ciclo = 1 To WAyuda.Cols - 1
                WAyuda.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        WAyuda.Text = "Rubro"
                        WAyuda.ColWidth(Ciclo) = 1500
                        WAyuda.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        WAyuda.Text = "Descripcion"
                        WAyuda.ColWidth(Ciclo) = 6000
                        WAyuda.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
        Case Else
    End Select
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WAyuda.Cols - 1
        WAncho = WAncho + WAyuda.ColWidth(Ciclo)
    Next Ciclo
    WAyuda.Width = WAncho
    Ayuda.Width = WAncho

    WAyuda.Col = 1
    WAyuda.Row = 1
    
End Sub




Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Cuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Descripcion_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub WAyuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
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
            Call Cuenta_DblClick
        Case 116
            Call Primer_Click
        Case 117
            Call Siguiente_Click
        Case 118
            Call Anterior_Click
        Case 119
            Call Ultimo_Click
        Case 120
            Call Lista_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub


Private Sub Anterior_Click()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cuenta"
    ZSql = ZSql + " Where Cuenta.Cuenta < " + "'" + Cuenta.Text + "'"
    ZSql = ZSql + " Order by Cuenta.Cuenta"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        With rstCuenta
            .MoveLast
            Cuenta.Text = rstCuenta!Cuenta
        End With
        rstCuenta.Close
        Call Imprime_Datos
        Cuenta.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de Cuenta")
    End If
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Cuenta) as [CuentaMenor]"
    ZSql = ZSql + " FROM Cuenta"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        rstCuenta.MoveFirst
        ZUltimo = IIf(IsNull(rstCuenta!CuentaMenor), "0", rstCuenta!CuentaMenor)
        Cuenta.Text = ZUltimo
        rstCuenta.Close
        Call Imprime_Datos
        Cuenta.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Cuenta) as [CuentaMayor]"
    ZSql = ZSql + " FROM Cuenta"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        rstCuenta.MoveLast
        ZUltimo = IIf(IsNull(rstCuenta!CuentaMayor), "0", rstCuenta!CuentaMayor)
        Cuenta.Text = ZUltimo
        rstCuenta.Close
        Call Imprime_Datos
        Cuenta.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Cuenta"
    ZSql = ZSql + " Where Cuenta.Cuenta > " + "'" + Cuenta.Text + "'"
    ZSql = ZSql + " Order by Cuenta.Cuenta"
    spCuenta = ZSql
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        With rstCuenta
            .MoveFirst
            Cuenta.Text = rstCuenta!Cuenta
        End With
        rstCuenta.Close
        Call Imprime_Datos
        Cuenta.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Cuenta")
    End If
End Sub



    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Order by Cuenta.Cuenta"
            spCuenta = ZSql
            Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            If rstCuenta.RecordCount > 0 Then
                With rstCuenta
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Cuenta + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Cuenta
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCuenta.Close
            End If

