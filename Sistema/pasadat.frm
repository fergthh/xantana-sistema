VERSION 5.00
Begin VB.Form PrgPasadat 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Clientes"
   ClientHeight    =   8280
   ClientLeft      =   510
   ClientTop       =   450
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   8280
   ScaleWidth      =   10995
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   4080
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
End
Attribute VB_Name = "PrgPasadat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Auxi As String
Dim Auxi1 As String
Dim WRenglon As String
Dim XVector(100, 30)
Dim WComprobante As String
Dim WCiclo As String
Dim Vector(1000, 20) As String
Dim WPrecio As Double
Dim WPrecio1 As Double

Private Sub Command1_Click()
    Call Pasaart
    Call cmdClose_Click
End Sub

Private Sub cmdClose_Click()
    With rstArticulo
         .Close
    End With
    With rstPasaart
        .Close
    End With
    DbsAdminis.Close
    PrgPasadat.Hide
    Unload Me
    Close
    End
End Sub

Private Sub Pasaart()

    With rstPasaart
        .Index = "Clave"
        .MoveFirst
        If .NoMatch = False Then
            Do
                .Edit
                !ordfecha = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Pasa = 0

    With rstPasaart
        .Index = "Clave"
        .MoveFirst
        If .NoMatch = False Then
            Do
                
                WFecha = IIf(IsNull(!Fecha), "", !Fecha)
                WProveedor = IIf(IsNull(!Proveedor), "", !Proveedor)
                WRazon = IIf(IsNull(!Razon), "", !Razon)
                WCantidad = IIf(IsNull(!Cantidad), "", !Cantidad)
                WArticulo1 = IIf(IsNull(!articulo1), "", !articulo1)
                WArticulo2 = IIf(IsNull(!articulo2), "", !articulo2)
                WArticulo3 = IIf(IsNull(!articulo3), "", !articulo3)
                WLinea = IIf(IsNull(!Linea), "", !Linea)
                WFamilia = IIf(IsNull(!familia), "", !familia)
                WDescripcion = IIf(IsNull(!Descripcion), "", !Descripcion)
                WColor = IIf(IsNull(!Color), "", !Color)
                WPrecioUnitario = IIf(IsNull(!precioUnitario), "", Str$(!precioUnitario))
                WPrecioTotal = IIf(IsNull(!preciototal), "", Str$(!preciototal))
                WAumento = IIf(IsNull(!aumento), "", Str$(!aumento))
                WNeto = IIf(IsNull(!Neto), "", Str$(!Neto))
                WDescuento = IIf(IsNull(!Descuento), "", Str$(!Descuento))
                WPrecioFinal = IIf(IsNull(!preciofinal), "", Str$(!preciofinal))
                WPrecioContado = IIf(IsNull(!preciocontado), "", Str$(!preciocontado))
                WOrdFecha = IIf(IsNull(!ordfecha), "", !ordfecha)
                
                If Pasa = 0 Then
                    Pasa = 1
                    WCorte1 = WProveedor
                    WCorte2 = WFecha
                    Erase Vector
                    Lugar = 0
                End If
                
                If WCorte1 <> WProveedor Or WCorte2 <> WFecha Then
                
                    With rstCompras
                        .Index = "Clave"
                        Claveven$ = "99999999"
                        .Seek "<=", Claveven$
                        If .NoMatch = False Then
                            XNumero = !Numero + 1
                                Else
                            XNumero = 1
                        End If
                    End With
                
                    For Ciclo = 1 To Lugar
                
                        XFecha = Vector(Ciclo, 1)
                        XProveedor = Vector(Ciclo, 2)
                        XRazon = Vector(Ciclo, 3)
                        XCantidad = Vector(Ciclo, 4)
                        xarticulo1 = Vector(Ciclo, 5)
                        xarticulo2 = Vector(Ciclo, 6)
                        xarticulo3 = Vector(Ciclo, 7)
                        XLinea = Vector(Ciclo, 8)
                        XFamilia = Vector(Ciclo, 9)
                        XDescripcion = Vector(Ciclo, 10)
                        XColor = Vector(Ciclo, 11)
                        XPrecioUnitario = Vector(Ciclo, 12)
                        XPrecioTotal = Vector(Ciclo, 13)
                        XAumento = Vector(Ciclo, 14)
                        XNeto = Vector(Ciclo, 15)
                        XDescuento = Vector(Ciclo, 16)
                        XPrecioFinal = Vector(Ciclo, 17)
                        XPrecioContado = Vector(Ciclo, 18)
                        XOrdFecha = Vector(Ciclo, 19)
                        
                        If Left$(xarticulo2, 1) <> "/" And xarticulo2 <> "" Then
                            xarticulo2 = "/" + xarticulo2
                        End If
                        If Left$(xarticulo3, 1) <> "/" And xarticulo3 <> "" Then
                            xarticulo3 = "/" + xarticulo3
                        End If
                        
                        XCodigo = xarticulo1 + xarticulo2 + xarticulo3
                        
                        XPrecio = Val(XPrecioUnitario)
                        If XAumento <> 0 Then
                            XPrecio = XPrecio + (XPrecio * Val(XAumento))
                        End If
                        
                        WDescuento = Val(XDescuento) * 100
                        WMargen = 94
                        WMargen1 = 76

                        WCosto = XPrecio
                        WDto = WCosto * (WDescuento / 100)
                        WCosto = WCosto - WDto
                        WSuma = WCosto * (WMargen / 100)
                        WPrecio = WCosto + WSuma
                        
                        WCosto = XPrecio
                        WDto = WCosto * (WDescuento / 100)
                        WCosto = WCosto - WDto
                        WSuma = WCosto * (WMargen1 / 100)
                        WPrecio1 = WCosto + WSuma
                        
                        Call Redondeo(WPrecio)
                        Call Redondeo(WPrecio1)
                    
                        With rstArticulo
                            .Index = "Codigo"
                            .Seek "=", XCodigo
                            If .NoMatch Then
                                .AddNew
                                !Codigo = XCodigo
                                !Descripcion = Left$(XDescripcion, 50)
                                !Linea = Val(XFamilia)
                                !Proveedor = Val(XProveedor)
                                !Costo = XPrecio
                                !Descuento = WDescuento
                                !Precio = WPrecio
                                !Margen = WMargen
                                !Precio1 = WPrecio1
                                !Margen1 = WMargen1
                                !Stock = Val(XCantidad)
                                !Observacion = XRazon
                                !Color = XColor
                                !Empresa = 1
                                .Update
                                .Bookmark = .LastModified
                                    Else
                                .Edit
                                !Codigo = XCodigo
                                !Descripcion = Left$(XDescripcion, 50)
                                !Linea = Val(XFamilia)
                                !Proveedor = Val(XProveedor)
                                !Costo = XPrecio
                                !Descuento = WDescuento
                                !Precio = WPrecio
                                !Margen = WMargen
                                !Precio1 = WPrecio1
                                !Margen1 = WMargen1
                                !Stock = !Stock + Val(XCantidad)
                                !Observacion = ""
                                !Color = XColor
                                !Empresa = 1
                                .Update
                                .Bookmark = .LastModified
                            End If
                        End With
                        
                        With rstCompras
                            .Index = "Clave"
                            Articulo = XCodigo
                            Cantidad = Val(XCantidad)
                            Costo = XPrecio
                            Auxi = Str$(Ciclo)
                            Call Ceros(Auxi, 2)
                            Auxi1 = Str$(XNumero)
                            Call Ceros(Auxi1, 6)
                            .AddNew
                            !Numero = Auxi1
                            !Renglon = Auxi
                            !Articulo = Articulo
                            !Cantidad = Cantidad
                            !Fecha = XFecha
                            !Auxiliar = 0
                            !ordfecha = Right$(XFecha, 4) + Mid$(XFecha, 4, 2) + Left$(XFecha, 2)
                            !Clave = Auxi1 + Auxi
                            !Observaciones = XRazon
                            !Color = 0
                            !Costo = Costo
                            .Update
                        End With
                    
                    Next Ciclo
                    
                    WCorte1 = !Proveedor
                    WCorte2 = !Fecha
                    Erase Vector
                    Lugar = 0
                
                End If
                
                Lugar = Lugar + 1
                
                Vector(Lugar, 1) = WFecha
                Vector(Lugar, 2) = WProveedor
                Vector(Lugar, 3) = WRazon
                Vector(Lugar, 4) = WCantidad
                Vector(Lugar, 5) = WArticulo1
                Vector(Lugar, 6) = WArticulo2
                Vector(Lugar, 7) = WArticulo3
                Vector(Lugar, 8) = WLinea
                Vector(Lugar, 9) = WFamilia
                Vector(Lugar, 10) = WDescripcion
                Vector(Lugar, 11) = WColor
                Vector(Lugar, 12) = WPrecioUnitario
                Vector(Lugar, 13) = WPrecioTotal
                Vector(Lugar, 14) = WAumento
                Vector(Lugar, 15) = WNeto
                Vector(Lugar, 16) = WDescuento
                Vector(Lugar, 17) = WPrecioFinal
                Vector(Lugar, 18) = WPrecioContado
                Vector(Lugar, 19) = WOrdFecha
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    
    
                    With rstCompras
                        .Index = "Clave"
                        Claveven$ = "99999999"
                        .Seek "<=", Claveven$
                        If .NoMatch = False Then
                            XNumero = !Numero + 1
                                Else
                            XNumero = 1
                        End If
                    End With
                
                    For Ciclo = 1 To Lugar
                
                        XFecha = Vector(Ciclo, 1)
                        XProveedor = Vector(Ciclo, 2)
                        XRazon = Vector(Ciclo, 3)
                        XCantidad = Vector(Ciclo, 4)
                        xarticulo1 = Vector(Ciclo, 5)
                        xarticulo2 = Vector(Ciclo, 6)
                        xarticulo3 = Vector(Ciclo, 7)
                        XLinea = Vector(Ciclo, 8)
                        XFamilia = Vector(Ciclo, 9)
                        XDescripcion = Vector(Ciclo, 10)
                        XColor = Vector(Ciclo, 11)
                        XPrecioUnitario = Vector(Ciclo, 12)
                        XPrecioTotal = Vector(Ciclo, 13)
                        XAumento = Vector(Ciclo, 14)
                        XNeto = Vector(Ciclo, 15)
                        XDescuento = Vector(Ciclo, 16)
                        XPrecioFinal = Vector(Ciclo, 17)
                        XPrecioContado = Vector(Ciclo, 18)
                        XOrdFecha = Vector(Ciclo, 19)
                        
                        If Left$(xarticulo2, 1) <> "/" And xarticulo2 <> "" Then
                            xarticulo2 = "/" + xarticulo2
                        End If
                        If Left$(xarticulo3, 1) <> "/" And xarticulo3 <> "" Then
                            xarticulo3 = "/" + xarticulo3
                        End If
                        
                        XCodigo = xarticulo1 + xarticulo2 + xarticulo3
                        
                        XPrecio = Val(XPrecioUnitario)
                        If XAumento <> 0 Then
                            XPrecio = XPrecio + (XPrecio * Val(XAumento))
                        End If
                        
                        WDescuento = Val(XDescuento) * 100
                        WMargen = 94
                        WMargen1 = 76

                        WCosto = XPrecio
                        WDto = WCosto * (WDescuento / 100)
                        WCosto = WCosto - WDto
                        WSuma = WCosto * (WMargen / 100)
                        WPrecio = WCosto + WSuma
                        
                        WCosto = XPrecio
                        WDto = WCosto * (WDescuento / 100)
                        WCosto = WCosto - WDto
                        WSuma = WCosto * (WMargen1 / 100)
                        WPrecio1 = WCosto + WSuma
                        
                        Call Redondeo(WPrecio)
                        Call Redondeo(WPrecio1)
                    
                        With rstArticulo
                            .Index = "Codigo"
                            .Seek "=", XCodigo
                            If .NoMatch Then
                                .AddNew
                                !Codigo = XCodigo
                                !Descripcion = Left$(XDescripcion, 50)
                                !Linea = Val(XFamilia)
                                !Proveedor = Val(XProveedor)
                                !Costo = XPrecio
                                !Descuento = WDescuento
                                !Precio = WPrecio
                                !Margen = WMargen
                                !Precio1 = WPrecio1
                                !Margen1 = WMargen1
                                !Stock = Val(XCantidad)
                                !Observacion = XRazon
                                !Color = XColor
                                !Empresa = 1
                                .Update
                                .Bookmark = .LastModified
                                    Else
                                .Edit
                                !Codigo = XCodigo
                                !Descripcion = Left$(XDescripcion, 50)
                                !Linea = Val(XFamilia)
                                !Proveedor = Val(XProveedor)
                                !Costo = XPrecio
                                !Descuento = WDescuento
                                !Precio = WPrecio
                                !Margen = WMargen
                                !Precio1 = WPrecio1
                                !Margen1 = WMargen1
                                !Stock = !Stock + Val(XCantidad)
                                !Observacion = ""
                                !Color = XColor
                                !Empresa = 1
                                .Update
                                .Bookmark = .LastModified
                            End If
                        End With
                        
                        With rstCompras
                            .Index = "Clave"
                            Articulo = XCodigo
                            Cantidad = Val(XCantidad)
                            Costo = XPrecio
                            Auxi = Str$(Ciclo)
                            Call Ceros(Auxi, 2)
                            Auxi1 = Str$(XNumero)
                            Call Ceros(Auxi1, 6)
                            .AddNew
                            !Numero = Auxi1
                            !Renglon = Auxi
                            !Articulo = Articulo
                            !Cantidad = Cantidad
                            !Fecha = XFecha
                            !Auxiliar = 0
                            !ordfecha = Right$(XFecha, 4) + Mid$(XFecha, 4, 2) + Left$(XFecha, 2)
                            !Clave = Auxi1 + Auxi
                            !Observaciones = XRazon
                            !Color = 0
                            !Costo = Costo
                            .Update
                        End With
                    
                    Next Ciclo
                    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Articulo
    OPEN_FILE_Compras
    OPEN_FILE_Articulo
    OPEN_FILE_Proveedor
    OPEN_FILE_Pasaart
End Sub

