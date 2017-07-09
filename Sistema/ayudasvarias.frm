VERSION 5.00
Begin VB.Form ayudasvarias 
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
Attribute VB_Name = "ayudasvarias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aYUDA_Keypress(KeyAscii As Integer)


        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = !Codigo + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            


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
                            IngresaItem = Str$(!Cliente) + " " + !Razon
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
        

        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cuenta"
            ZSql = ZSql + " Where Cuenta.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
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
            
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Order by Vendedor.Codigo"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstVendedor!Nombre) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstVendedor!Nombre), aa, WEspacios) Then
                                    IngresaItem = Str$(!Codigo) + " " + !Nombre
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = !Codigo
                                    WIndice.AddItem IngresaItem
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
            
            
            
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoPro"
            ZSql = ZSql + " Where TipoPro.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by TipoPro.Codigo"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoPro.RecordCount > 0 Then
                With rstTipoPro
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoPro.Close
            End If
            
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
            
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Where Conceptos.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Conceptos.Concepto"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                With rstConceptos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Concepto) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Concepto
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstConceptos.Close
            End If
            
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Where Proyecto.Descripcion LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Proyecto.Codigo"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
            End If
            
            
            
        Case 0
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
            
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Where Lineas.Nombre LIKE " + "'" + "%" + ZAyuda + "%" + "'"
            ZSql = ZSql + " Order by Lineas.Linea"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                With rstLinea
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
                rstLinea.Close
            End If
            
            
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
                rstVendedor.Close
            End If
            
            
            
            

End Sub


Private Sub Opcion_Click()

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
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoPro"
            ZSql = ZSql + " Order by TipoPro.Codigo"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoPro.RecordCount > 0 Then
                With rstTipoPro
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoPro.Close
            End If
            
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
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Conceptos"
            ZSql = ZSql + " Order by Conceptos.Concepto"
            spConceptos = ZSql
            Set rstConceptos = db.OpenRecordset(spConceptos, dbOpenSnapshot, dbSQLPassThrough)
            If rstConceptos.RecordCount > 0 Then
                With rstConceptos
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Concepto) + " " + !Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Concepto
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstConceptos.Close
            End If
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proyecto"
            ZSql = ZSql + " Order by Proyecto.Codigo"
            spProyecto = ZSql
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(!Codigo) + " " + !Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProyecto.Close
                
            End If
            
        Case 0
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
            
        Case 0
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
                            IngresaItem = Str$(!Cliente) + " " + !Razon
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
        
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Order by Vendedor.Codigo"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
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
                rstVendedor.Close
            End If
            
            
            
            
            
            
            
            
            
        Case 5
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Articulo"
                ZSql = ZSql + " Order by Articulo.Codigo"
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    With rstArticulo
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                IngresaItem = !Codigo + " " + !Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = !Codigo
                                WIndice.AddItem IngresaItem
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstArticulo.Close
                End If
            
            
            
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Lineas"
            ZSql = ZSql + " Order by Lineas.Linea"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
            If rstLinea.RecordCount > 0 Then
                With rstLinea
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
                rstLinea.Close
            End If
            
            
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Vendedor"
            ZSql = ZSql + " Order by Vendedor.Codigo"
            spVendedor = ZSql
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
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
                rstVendedor.Close
            End If
            
            
            
            
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Order by Articulo.Codigo"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstArticulo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        Case Else
    End Select
End Sub














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
        End Select
    End If
    WEspacios = Len(ZAyuda)

