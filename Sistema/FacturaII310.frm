VERSION 5.00
Begin VB.Form FActuraII310 
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
Attribute VB_Name = "FActuraII310"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WNeto As Double
Private XNeto As Double
Private WIva1 As Double
Private WIva2 As Double
Private WTotal As Double
Private WImpoDto As Double
Private WCodIva As String
Private Precio As Double
Private Cantidad As Double
Private WAnterior As Integer
Private WDescri As String
Private WTipo As String
Private WTipoIva As String
Private WProvincia As String
Private WRazon As String
Private WDireccion As String
Private WLocalidad As String
Private WProv As String
Private WPostal As String
Private WImpiva As String
Private WCuit As String
Private Provincia(0 To 30) As String
Private Iva(0 To 30) As String
Private Mes(0 To 30) As String
Private XIndice As Single
Private XArticulo As String
Private XTexto1 As String
Private XTexto2 As String

Dim ZZNumero As String
Dim ZZRenglon As String
Dim ZZArticulo As String
Dim ZZCantidad As String
Dim ZZPrecio As String
Dim ZZImporte As String
Dim ZZFacturado As String
Dim ZZCliente As String
Dim ZZfecha As String
Dim ZZImporte1 As String
Dim ZZImporte2 As String
Dim ZZImporte3 As String
Dim ZZImporte4 As String
Dim ZZOrdFecha As String
Dim ZZObservaciones As String
Dim ZZFecEntrega As String
Dim ZZOrdFecEntrega As String
Dim ZZCotiza As String
Dim ZZAjuste As String
Dim ZZClave As String
Dim ZZGrupo As String
Dim ZCantidad As String
Dim ZCodigo As String
Dim ZObservaciones As String
Dim ZZPedido As String
Dim ZZVERSION As String


Dim ZVector(1000, 10) As String
Dim ZGraba(100, 30) As String


Rem para el vector

Dim WBorra(1000, 20) As String
Dim WParametros(20, 20) As Double
Dim WFormato(20) As String
Dim WControl As String


Private Sub Consulta_Click()

    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulos"

    Opcion.Visible = True
     
 End Sub

Private Sub Form_Activate()
    ZZProcesoPedido = 0
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    Opcion.Visible = False
    XIndice = Opcion.ListIndex
    
    Select Case XIndice
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
                            IngresaItem = !CLIENTE + " " + !Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !CLIENTE
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 1
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

Private Sub cmdClose_Click()
    PrgPedido.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Graba_Click()
    
    ZSql = ""
    ZSql = ZSql + "DELETE Pedido"
    ZSql = ZSql + " Where Pedido.Numero = " + "'" + Numero.Text + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    Renglon = 0
    WRenglon = 0
        
    For a = 1 To 99
        
        WRenglon = WRenglon + 1
            
        WVector1.Row = WRenglon
            
        WVector1.Col = 1
        Articulo = UCase(WVector1.Text)
                    
        WVector1.Col = 3
        Cantidad = Val(WVector1.Text)
                    
        WVector1.Col = 4
        Observaciones = WVector1.Text
        
        WVector1.Col = 5
        facturado = WVector1.Text
                    
        If Cantidad <> 0 Then
        
            Precio = 0
            ZZPosicion = "9999"
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Precio = rstArticulo!Precio
                ZZGrupo = Str$(rstArticulo!Grupo)
                If Val(rstArticulo!PosicionII) <> 0 Then
                    ZZPosicion = Str$(rstArticulo!PosicionII)
                End If
                rstArticulo.Close
            End If
            
            IMPORTE = Precio * Cantidad
                    
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(Numero.Text)
            Call Ceros(Auxi1, 8)
                    
            ZZNumero = Numero.Text
            ZZRenglon = Str$(Renglon)
            ZZArticulo = Articulo
            ZZCantidad = Str$(Cantidad)
            ZZPrecio = Str$(Precio)
            ZZCliente = CLIENTE.Text
            ZZImporte = Str$(IMPORTE)
            ZZfecha = Fecha.Text
            ZZImporte1 = "0"
            ZZImporte2 = "0"
            ZZImporte3 = "0"
            ZZImporte4 = "0"
            ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZZObservaciones = Observaciones
            ZZFecEntrega = Fecha.Text
            ZZOrdFecEntrega = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
            ZZFacturado = facturado
            ZZCotiza = "0"
            ZZAjuste = "0"
            
            ZZClave = Auxi1 + Auxi
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Pedido ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Articulo ,"
            ZSql = ZSql + "Grupo ,"
            ZSql = ZSql + "Cantidad ,"
            ZSql = ZSql + "Precio ,"
            ZSql = ZSql + "Importe ,"
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Importe1 ,"
            ZSql = ZSql + "Importe2 ,"
            ZSql = ZSql + "Importe3 ,"
            ZSql = ZSql + "Importe4 ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Posicion ,"
            ZSql = ZSql + "FecEntrega  ,"
            ZSql = ZSql + "OrdFecEntrega ,"
            ZSql = ZSql + "Facturado ,"
            ZSql = ZSql + "Ajuste )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZZClave + "',"
            ZSql = ZSql + "'" + ZZNumero + "',"
            ZSql = ZSql + "'" + ZZRenglon + "',"
            ZSql = ZSql + "'" + ZZArticulo + "',"
            ZSql = ZSql + "'" + ZZGrupo + "',"
            ZSql = ZSql + "'" + ZZCantidad + "',"
            ZSql = ZSql + "'" + ZZPrecio + "',"
            ZSql = ZSql + "'" + ZZImporte + "',"
            ZSql = ZSql + "'" + ZZCliente + "',"
            ZSql = ZSql + "'" + ZZfecha + "',"
            ZSql = ZSql + "'" + ZZImporte1 + "',"
            ZSql = ZSql + "'" + ZZImporte2 + "',"
            ZSql = ZSql + "'" + ZZImporte3 + "',"
            ZSql = ZSql + "'" + ZZImporte4 + "',"
            ZSql = ZSql + "'" + ZZOrdFecha + "',"
            ZSql = ZSql + "'" + ZZObservaciones + "',"
            ZSql = ZSql + "'" + ZZPosicion + "',"
            ZSql = ZSql + "'" + ZZFecEntrega + "',"
            ZSql = ZSql + "'" + ZZOrdFecEntrega + "',"
            ZSql = ZSql + "'" + ZZFacturado + "',"
            ZSql = ZSql + "'" + ZZAjuste + "')"
            
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
                                        
    Next a
    
    
    
    
    Rem dada
    Rem dada
    Rem dada
    
    Erase ZGraba
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Pedido"
    ZSql = ZSql + " Where Pedido.Numero = " + "'" + ZZNumero + "'"
    ZSql = ZSql + " Order by Pedido.Numero,Pedido.Posicion,Pedido.Articulo"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugar = ZLugar + 1
                    
                    ZGraba(ZLugar, 1) = rstPedido!Clave
                    ZGraba(ZLugar, 2) = Str$(rstPedido!Numero)
                    ZGraba(ZLugar, 3) = Str$(rstPedido!Renglon)
                    ZGraba(ZLugar, 4) = rstPedido!Articulo
                    ZGraba(ZLugar, 5) = Str$(rstPedido!Grupo)
                    ZGraba(ZLugar, 6) = Str$(rstPedido!Cantidad)
                    ZGraba(ZLugar, 7) = Str$(rstPedido!Precio)
                    ZGraba(ZLugar, 8) = Str$(rstPedido!IMPORTE)
                    ZGraba(ZLugar, 9) = rstPedido!CLIENTE
                    ZGraba(ZLugar, 10) = rstPedido!Fecha
                    ZGraba(ZLugar, 11) = Str$(rstPedido!Importe1)
                    ZGraba(ZLugar, 12) = Str$(rstPedido!Importe2)
                    ZGraba(ZLugar, 13) = Str$(rstPedido!Importe3)
                    ZGraba(ZLugar, 14) = Str$(rstPedido!Importe4)
                    ZGraba(ZLugar, 15) = rstPedido!ordfecha
                    ZGraba(ZLugar, 16) = rstPedido!Observaciones
                    ZGraba(ZLugar, 25) = rstPedido!FecEntrega
                    ZGraba(ZLugar, 26) = rstPedido!OrdFecEntrega
                    ZGraba(ZLugar, 27) = Str$(rstPedido!facturado)
                    ZGraba(ZLugar, 28) = Str$(rstPedido!Ajuste)
                    ZGraba(ZLugar, 29) = Str$(rstPedido!Posicion)
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstPedido.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "DELETE Pedido"
    ZSql = ZSql + " Where Pedido.Numero = " + "'" + ZZNumero + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    For a = 1 To ZLugar
        
        
        Auxi = Str$(a)
        Call Ceros(Auxi, 2)
        Auxi1 = Str$(Numero.Text)
        Call Ceros(Auxi1, 8)
        
        ZZClave = Auxi1 + Auxi
        ZZNumero = ZGraba(a, 2)
        ZZRenglon = Str$(a)
        ZZArticulo = ZGraba(a, 4)
        ZZGrupo = ZGraba(a, 5)
        ZZCantidad = ZGraba(a, 6)
        ZZPrecio = ZGraba(a, 7)
        ZZImporte = ZGraba(a, 8)
        ZZCliente = ZGraba(a, 9)
        ZZfecha = ZGraba(a, 10)
        ZZImporte1 = ZGraba(a, 11)
        ZZImporte2 = ZGraba(a, 12)
        ZZImporte3 = ZGraba(a, 13)
        ZZImporte4 = ZGraba(a, 14)
        ZZOrdFecha = ZGraba(a, 15)
        ZZObservaciones = ZGraba(a, 16)
        ZZFecEntrega = ZGraba(a, 25)
        ZZOrdFecEntrega = ZGraba(a, 26)
        ZZFacturado = ZGraba(a, 27)
        ZZAjuste = ZGraba(a, 28)
        ZZPosicion = ZGraba(a, 29)
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO Pedido ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "Renglon ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Grupo ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Precio ,"
        ZSql = ZSql + "Importe ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Importe1 ,"
        ZSql = ZSql + "Importe2 ,"
        ZSql = ZSql + "Importe3 ,"
        ZSql = ZSql + "Importe4 ,"
        ZSql = ZSql + "OrdFecha ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Posicion ,"
        ZSql = ZSql + "FecEntrega  ,"
        ZSql = ZSql + "OrdFecEntrega ,"
        ZSql = ZSql + "Facturado ,"
        ZSql = ZSql + "Ajuste )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + ZZNumero + "',"
        ZSql = ZSql + "'" + ZZRenglon + "',"
        ZSql = ZSql + "'" + ZZArticulo + "',"
        ZSql = ZSql + "'" + ZZGrupo + "',"
        ZSql = ZSql + "'" + ZZCantidad + "',"
        ZSql = ZSql + "'" + ZZPrecio + "',"
        ZSql = ZSql + "'" + ZZImporte + "',"
        ZSql = ZSql + "'" + ZZCliente + "',"
        ZSql = ZSql + "'" + ZZfecha + "',"
        ZSql = ZSql + "'" + ZZImporte1 + "',"
        ZSql = ZSql + "'" + ZZImporte2 + "',"
        ZSql = ZSql + "'" + ZZImporte3 + "',"
        ZSql = ZSql + "'" + ZZImporte4 + "',"
        ZSql = ZSql + "'" + ZZOrdFecha + "',"
        ZSql = ZSql + "'" + ZZObservaciones + "',"
        ZSql = ZSql + "'" + ZZPosicion + "',"
        ZSql = ZSql + "'" + ZZFecEntrega + "',"
        ZSql = ZSql + "'" + ZZOrdFecEntrega + "',"
        ZSql = ZSql + "'" + ZZFacturado + "',"
        ZSql = ZSql + "'" + ZZAjuste + "')"
        
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                                        
    Next a
    
    
    Rem Erase ZVector
    Rem ZLugar = 0
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select *"
    Rem ZSql = ZSql + " FROM Pedido"
    Rem ZSql = ZSql + " Where Pedido.Numero = " + "'" + ZZNumero + "'"
    Rem ZSql = ZSql + " Order by Pedido.Grupo, Pedido.Articulo"
    
    Rem spPedido = ZSql
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedido.RecordCount > 0 Then
    Rem
    Rem     With rstPedido
    Rem         .MoveFirst
    Rem         Do
    Rem             If .EOF = False Then
    Rem
    Rem                 ZLugar = ZLugar + 1
    Rem
    Rem                 ZVector(ZLugar, 1) = Trim(rstPedido!Articulo)
    Rem                 ZVector(ZLugar, 3) = Pusing("###,###", Str$(rstPedido!Cantidad))
    Rem                 ZVector(ZLugar, 4) = Trim(rstPedido!Observaciones)
    Rem                 ZVector(ZLugar, 5) = Str$(rstPedido!facturado)
    Rem
    Rem                 .MoveNext
    Rem                     Else
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem     End With
    Rem
    Rem     rstPedido.Close
    Rem End If
    
    
    
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "DELETE Pedido"
    Rem ZSql = ZSql + " Where Pedido.Numero = " + "'" + Numero.Text + "'"
    Rem spPedido = ZSql
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    Rem Renglon = 0
    Rem WRenglon = 0
        
    Rem For a = 1 To 99
    Rem
    Rem     WRenglon = WRenglon + 1
    Rem
    Rem     WVector1.Col = 1
    Rem     Articulo = ZVector(WRenglon, 1)
    Rem
    Rem     WVector1.Col = 3
    Rem     Cantidad = Val(ZVector(WRenglon, 3))
    Rem
    Rem     WVector1.Col = 4
    Rem     Observaciones = ZVector(WRenglon, 4)
    Rem
    Rem     WVector1.Col = 5
    Rem     facturado = ZVector(WRenglon, 5)
     Rem
    Rem     If Cantidad <> 0 Then
    Rem
    Rem         Precio = 0
    Rem
    Rem         ZSql = ""
    Rem         ZSql = ZSql + "Select *"
    Rem         ZSql = ZSql + " FROM Articulo"
    Rem         ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Articulo + "'"
    Rem         spArticulo = ZSql
    Rem         Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    Rem         If rstArticulo.RecordCount > 0 Then
    Rem             Precio = rstArticulo!Precio
    Rem             ZZGrupo = Str$(rstArticulo!Grupo)
    Rem             rstArticulo.Close
    Rem         End If
    Rem
    Rem         Importe = Precio * Cantidad
    Rem
    Rem         Renglon = Renglon + 1
    Rem         Auxi = Str$(Renglon)
    Rem         Call Ceros(Auxi, 2)
    Rem
    Rem         Auxi1 = Str$(Numero.Text)
    Rem         Call Ceros(Auxi1, 8)
    Rem
    Rem         ZZNumero = Numero.Text
    Rem         ZZRenglon = Str$(Renglon)
    Rem         ZZArticulo = Articulo
    Rem         ZZCantidad = Str$(Cantidad)
    Rem         ZZPrecio = Str$(Precio)
    Rem         ZZCliente = Cliente.Text
    Rem         ZZImporte = Str$(Importe)
    Rem         ZZfecha = Fecha.Text
    Rem         ZZImporte1 = "0"
    Rem         ZZImporte2 = "0"
    Rem         ZZImporte3 = "0"
    Rem         ZZImporte4 = "0"
    Rem         ZZOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    Rem         ZZObservaciones = Observaciones
    Rem         ZZFecEntrega = Fecha.Text
    Rem         ZZOrdFecEntrega = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
    Rem         ZZFacturado = facturado
    Rem         ZZCotiza = "0"
    Rem         ZZAjuste = "0"
            
    Rem         ZZClave = Auxi1 + Auxi
            
    Rem         ZSql = ""
    Rem         ZSql = ZSql + "INSERT INTO Pedido ("
    Rem         ZSql = ZSql + "Clave ,"
    Rem         ZSql = ZSql + "Numero ,"
    Rem         ZSql = ZSql + "Renglon ,"
    Rem         ZSql = ZSql + "Articulo ,"
    Rem         ZSql = ZSql + "Grupo ,"
    Rem         ZSql = ZSql + "Cantidad ,"
    Rem         ZSql = ZSql + "Precio ,"
    Rem         ZSql = ZSql + "Importe ,"
    Rem         ZSql = ZSql + "Cliente ,"
    Rem         ZSql = ZSql + "Fecha ,"
    Rem         ZSql = ZSql + "Importe1 ,"
    Rem         ZSql = ZSql + "Importe2 ,"
    Rem         ZSql = ZSql + "Importe3 ,"
    Rem         ZSql = ZSql + "Importe4 ,"
    Rem         ZSql = ZSql + "OrdFecha ,"
    Rem         ZSql = ZSql + "Observaciones ,"
    Rem         ZSql = ZSql + "FecEntrega  ,"
    Rem         ZSql = ZSql + "OrdFecEntrega ,"
    Rem         ZSql = ZSql + "Facturado ,"
    Rem         ZSql = ZSql + "Ajuste )"
    Rem         ZSql = ZSql + "Values ("
    Rem         ZSql = ZSql + "'" + ZZClave + "',"
    Rem         ZSql = ZSql + "'" + ZZNumero + "',"
    Rem         ZSql = ZSql + "'" + ZZRenglon + "',"
    Rem         ZSql = ZSql + "'" + ZZArticulo + "',"
    Rem         ZSql = ZSql + "'" + Grupo + "',"
    Rem         ZSql = ZSql + "'" + ZZCantidad + "',"
    Rem         ZSql = ZSql + "'" + ZZPrecio + "',"
    Rem         ZSql = ZSql + "'" + ZZImporte + "',"
    Rem         ZSql = ZSql + "'" + ZZCliente + "',"
    Rem         ZSql = ZSql + "'" + ZZfecha + "',"
    Rem         ZSql = ZSql + "'" + ZZImporte1 + "',"
    Rem         ZSql = ZSql + "'" + ZZImporte2 + "',"
    Rem         ZSql = ZSql + "'" + ZZImporte3 + "',"
    Rem         ZSql = ZSql + "'" + ZZImporte4 + "',"
    Rem         ZSql = ZSql + "'" + ZZOrdFecha + "',"
    Rem         ZSql = ZSql + "'" + ZZObservaciones + "',"
    Rem         ZSql = ZSql + "'" + ZZFecEntrega + "',"
    Rem         ZSql = ZSql + "'" + ZZOrdFecEntrega + "',"
    Rem         ZSql = ZSql + "'" + ZZFacturado + "',"
    Rem         ZSql = ZSql + "'" + ZZAjuste + "')"
    Rem
    Rem         spPedido = ZSql
    Rem         Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem     End If
    Rem
    Rem Next a



    T$ = "Impresion de Pedidos"
    M$ = "Desea Imprimir el Pedido"
    Respuesta% = MsgBox(M$, 32 + 4, T$)
    If Respuesta% = 6 Then
        Call Impresion
    End If
    
    Call Limpia_Click
    Numero.SetFocus
        
End Sub

Private Sub Impresion()

    Listado.WindowTitle = "Impresion de Pedido"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Pedido.Numero, Pedido.Renglon, Pedido.Articulo, Pedido.Cantidad, Pedido.Cliente, Pedido.Fecha, Pedido.Observaciones, " _
            + "Articulo.Descripcion, Articulo.MinimoVenta, Articulo.UnidadCaja, " _
            + "Cliente.Razon, Cliente.Direccion, Cliente.Localidad " _
            + "From " _
            + DSQ + ".dbo.Pedido Pedido, " _
            + DSQ + ".dbo.Articulo Articulo, " _
            + DSQ + ".dbo.Cliente Cliente " _
            + "Where " _
            + "Pedido.Articulo = Articulo.Codigo AND " _
            + "Pedido.Cliente = Cliente.Cliente AND " _
            + "Pedido.Numero >= " + Numero.Text + " AND " _
            + "Pedido.Numero <= " + Numero.Text
    
    Listado.Connect = Connect()
    
    Uno = "{Pedido.Numero} in " + Numero.Text + " to " + Numero.Text
    
    Listado.GroupSelectionFormula = Uno
    Listado.SelectionFormula = Uno
    
    Listado.Destination = 1
    Rem Listado.Destination = 0
    Listado.ReportFileName = "ImprePedido.rpt"
    
    Listado.Action = 1

End Sub

Private Sub CmdDelete_Click()

    T$ = "Baja de Comprobantes"
    M$ = "Desea Borrar el Comprobante "
    Respuesta% = MsgBox(M$, 32 + 4, T$)
    If Respuesta% = 6 Then
    
        ZSql = ""
        ZSql = ZSql + "DELETE Pedido"
        ZSql = ZSql + " Where Pedido.Numero = " + "'" + Numero.Text + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
        Call Limpia_Click
        Numero.SetFocus
        
    End If

End Sub

Private Sub Limpia_Click()

    Call Limpia_Vector

    Numero.Text = ""
    CLIENTE.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    Renglon = 0
    
    
    Numero.Text = ""
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    Rem ZSql = ZSql + " FROM Pedido"
    Rem spPedido = ZSql
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedido.RecordCount > 0 Then
    Rem     rstPedido.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstPedido!NumeroMayor), "0", rstPedido!NumeroMayor)
    Rem     Numero.Text = ZUltimo + 1
    Rem     rstPedido.Close
    Rem End If
    
    Numero.SetFocus

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Opcion.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            CLIENTE.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case 1
            WTexto1.Visible = False
            WTexto2.Visible = False
            
            Indice = Pantalla.ListIndex
            ClaveVen$ = WIndice.List(Indice)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + ClaveVen$ + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 1
                WVector1.Text = rstArticulo!Codigo
                WVector1.Col = 6
                WVector1.Text = Trim(rstArticulo!MinimoVenta)
                WVector1.Col = 7
                WVector1.Text = Str$(rstArticulo!Precio)
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                WVector1.Col = 8
                WVector1.Text = Trim(rstArticulo!Stock)
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                WVector1.Col = 3
                rstArticulo.Close
                Call StartEdit
            End If
            
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Call Limpia_Vector

    Provincia(0) = "Capital Federal"
    Provincia(1) = "Buenos Aires"
    Provincia(2) = "Catamarca"
    Provincia(3) = "Cordoba"
    Provincia(4) = "Corrientes"
    Provincia(5) = "Chaco"
    Provincia(6) = "Chubut"
    Provincia(7) = "Entre Rios"
    Provincia(8) = "Formosa"
    Provincia(9) = "Jujuy"
    Provincia(10) = "La Pampa"
    Provincia(11) = "La Rioja"
    Provincia(12) = "Mendoza"
    Provincia(13) = "Misiones"
    Provincia(14) = "Neuquen"
    Provincia(15) = "Rio Negro"
    Provincia(16) = "Salta"
    Provincia(17) = "San Juan"
    Provincia(18) = "San Luis"
    Provincia(19) = "Santa Cruz"
    Provincia(20) = "Santa Fe"
    Provincia(21) = "Santiago del Estero"
    Provincia(22) = "Tucuman"
    Provincia(23) = "Tierra del Fuego"
    Provincia(24) = "Exterior"
    Provincia(25) = ""
    
    Iva(1) = "Inscripto"
    Iva(2) = "No Inscripto"
    Iva(3) = "Inscripto"
    Iva(4) = "Inscripto"
    Iva(5) = "Inscripto"
    Iva(6) = "Inscripto"
    
    Mes(1) = "Enero"
    Mes(2) = "Febrero"
    Mes(3) = "Marzo"
    Mes(4) = "Abril"
    Mes(5) = "Mayo"
    Mes(6) = "Junio"
    Mes(7) = "Julio"
    Mes(8) = "Agosto"
    Mes(9) = "Septiembre"
    Mes(10) = "Octubre"
    Mes(11) = "Noviembre"
    Mes(12) = "Diciembre"
    
    Rem Iva(3) = "Consumidor Final"
    Rem Iva(4) = "Exento"
    Rem Iva(5) = "Monotributo"
    Rem Iva(6) = "No Catalogado"
    
    Numero.Text = ""
    CLIENTE.Text = ""
    DesCliente.Caption = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Configuracion"
    ZSql = ZSql + " Where Configuracion.Clave = 1"
    spConfiguracion = ZSql
    Set rstConfiguracion = db.OpenRecordset(spConfiguracion, dbOpenSnapshot, dbSQLPassThrough)
    If rstConfiguracion.RecordCount > 0 Then
        ConfigIva1 = rstConfiguracion!Iva1
        ConfigIva2 = rstConfiguracion!Iva2
        ConfigPercepcion = rstConfiguracion!Percepcion
        ConfigPunto = rstConfiguracion!Punto
        rstConfiguracion.Close
    End If
    
    
    Numero.Text = ""
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Numero) as [NumeroMayor]"
    Rem ZSql = ZSql + " FROM Pedido"
    Rem spPedido = ZSql
    Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstPedido.RecordCount > 0 Then
    Rem     rstPedido.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstPedido!NumeroMayor), "0", rstPedido!NumeroMayor)
    Rem     Numero.Text = ZUltimo + 1
    Rem     rstPedido.Close
    Rem End If
    
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector

    Renglon = 0
    WNeto = 0
    
    For WRenglon = 1 To 99
    
        Auxi = Numero.Text
        Call Ceros(Auxi, 8)
            
        Auxi1 = WRenglon
        Call Ceros(Auxi1, 2)
        
        WClave = Auxi + Auxi1
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Clave = " + "'" + WClave + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
        
            Canti = rstPedido!Cantidad
            
            Renglon = Renglon + 1
                    
            WVector1.Row = Renglon
                    
            WVector1.Col = 1
            WVector1.Text = Trim(rstPedido!Articulo)
            Auxi1 = rstPedido!Articulo
                
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###", Str$(rstPedido!Cantidad))
                
            WVector1.Col = 4
            WVector1.Text = Trim(rstPedido!Observaciones)
            
            WVector1.Col = 5
            WVector1.Text = Str$(rstPedido!facturado)
            
            rstPedido.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + Auxi1 + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            
        End If
    
    Next WRenglon
    
End Sub

Sub Numero_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pedido"
        ZSql = ZSql + " Where Pedido.Numero = " + "'" + Numero.Text + "'"
        spPedido = ZSql
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
    
            Fecha.Text = rstPedido!Fecha
            CLIENTE.Text = rstPedido!CLIENTE
            
            rstPedido.Close
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                CLIENTE.Text = rstCliente!CLIENTE
                DesCliente.Caption = rstCliente!Razon
                WProvincia = rstCliente!Provincia
                WCodIva = rstCliente!Iva
                WRazon = rstCliente!Razon
                WDireccion = rstCliente!Direccion
                WLocalidad = rstCliente!Localidad
                WPostal = rstCliente!Postal
                WCuit = rstCliente!Cuit
                rstCliente.Close
            End If
            
            Call Proceso_Click
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                
                Else
                    
            CLIENTE.SetFocus
               
        End If
            
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + CLIENTE.Text + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            CLIENTE.Text = rstCliente!CLIENTE
            DesCliente.Caption = rstCliente!Razon
            WProvincia = rstCliente!Provincia
            WCodIva = rstCliente!Iva
            WRazon = rstCliente!Razon
            WDireccion = rstCliente!Direccion
            WLocalidad = rstCliente!Localidad
            WPostal = rstCliente!Postal
            WCuit = rstCliente!Cuit
            rstCliente.Close
            Rem Confirma.Text = "S"
            Rem PantallaConfirma.Visible = True
            Rem Confirma.SetFocus
                Else
            CLIENTE.SetFocus
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM HistorialCliente"
        ZSql = ZSql + " Where HistorialCliente.Cliente = " + "'" + CLIENTE.Text + "'"
        spHistorialCliente = ZSql
        Set rstHistorialCliente = db.OpenRecordset(spHistorialCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstHistorialCliente.RecordCount > 0 Then
            rstHistorialCliente.Close
            ZZPasaCliente = CLIENTE.Text
            ZZPasaProceso = 2
            PrgHistorialClienteConsulta.Show
        End If
        
    End If
    If KeyAscii = 27 Then
        CLIENTE.Text = ""
        DesCliente.Caption = ""
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            WVector1.Col = 1
            WVector1.Row = 1
            Call StartEdit
                Else
            M$ = "Formato de fecha invalido"
            a% = MsgBox(M$, 0, "Emision de Comprobante varios")
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Confirma_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CONFIRMA.Text = Trim(UCase(CONFIRMA.Text))
        If CONFIRMA.Text = "S" Or CONFIRMA.Text = "N" Or CONFIRMA.Text = "/" Or CONFIRMA.Text = "?" Then
            PantallaConfirma.Visible = False
            If CONFIRMA.Text <> "N" Then
                WVector1.Col = 1
                WVector1.Row = 1
                Call StartEdit
            End If
        End If
    End If
    If KeyAscii = 27 Then
        CONFIRMA.Text = ""
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
                            IngresaItem = !CLIENTE + " " + !Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = !CLIENTE
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 1
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
            
        Case Else
    End Select
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub


Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto1.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto2.Text = WVector1.Text
            Call Ejecuta_Funcion
            
        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto3.Text = WVector1.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub




Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 3
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 1
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
        Case 1
            If Trim(WVector1.Text) <> "" Then
                ZZVeri = UCase(Left$(WVector1.Text, 1))
                If ZZVeri < "A" Or ZZVeri > "Z" Then
                    ZZVeri = Left$(WVector1.TextMatrix(WVector1.Row - 1, 1), 1)
                    WVector1.Text = ZZVeri + WVector1.Text
                End If
                Auxi = UCase(Left$(WVector1.Text, 1))
                Auxi1 = Mid$(WVector1.Text, 2, 5)
                Call Ceros(Auxi1, 5)
                WVector1.Text = Auxi + Auxi1
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WVector1.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WVector1.Col = 6
                WVector1.Text = Trim(rstArticulo!MinimoVenta)
                WVector1.Col = 7
                WVector1.Text = Str$(rstArticulo!Precio)
                WVector1.Text = Pusing("###,###.##", WVector1.Text)
                WVector1.Col = 8
                WVector1.Text = Trim(rstArticulo!Stock)
                WVector1.Col = 2
                WVector1.Text = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                WControl = "N"
            End If
            
        Case 3
            If Val(WVector1.Text) = 0 Then
                WVector1.Text = "1"
            End If
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub WVector1_DblClick()

    If WVector1.Col = 0 Or WVector1.Col = 1 Then
    
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False

    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WVector1.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector1.Rows - 1
        WVector1.Row = Ciclo
        WVector1.Col = 1
        WAuxi1 = WVector1.Text
        WVector1.Col = 3
        WAuxi3 = WVector1.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector1.Cols - 1
                WVector1.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector1.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector
    
    For Ciclo = 1 To EntraVector
        WVector1.Row = Ciclo
        For da = 1 To WVector1.Cols - 1
            WVector1.Col = da
            WVector1.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
    End If
    
End Sub

Private Sub WTexto1_DblClick()

    If WVector1.Col = 1 Then

    Opcion.Clear
    
    Opcion.AddItem "Cliente"
    Opcion.AddItem "Articulos"

    Rem Opcion.Visible = True
    
    Opcion.ListIndex = 1
    
    Call Opcion_Click
    
    End If
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 12
    WVector1.FixedRows = 1
    WVector1.Rows = 100
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Codigo"
                WVector1.ColWidth(Ciclo) = 1600
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Moneda"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Dto"
                WVector1.ColWidth(Ciclo) = 900
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 6
                WVector1.Text = "Precio $"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 7
                WVector1.Text = "Total $"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 8
                WVector1.Text = "Total U$S"
                WVector1.ColWidth(Ciclo) = 1100
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 9
                WVector1.Text = "F.Entrega"
                WVector1.ColWidth(Ciclo) = 1300
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 10
                WVector1.Text = "Observaciones"
                WVector1.ColWidth(Ciclo) = 2000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 11
                WVector1.Text = ""
                WVector1.ColWidth(Ciclo) = 10
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case Else
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho
    
    For Ciclo = 1 To 99
        WVector1.TextMatrix(Ciclo, 0) = Trim(Str$(Ciclo))
    Next Ciclo

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub

Private Sub Cliente_DblClick()

    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Articulo"

    Opcion.ListIndex = 0
    
    Call Opcion_Click

End Sub


Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Numero_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Graba_Click
        Case 113
            Call CmdDelete_Click
        Case 114
            Call Limpia_Click
        Case 115
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub



