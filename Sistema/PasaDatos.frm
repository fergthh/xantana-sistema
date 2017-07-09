VERSION 5.00
Begin VB.Form PrgPasaDatos 
   Caption         =   "Traspaso de Lista de Precios"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      Caption         =   "Control de Grabacion"
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "PrgPasaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()
    Call Proceso
    Call Cancela_click
End Sub

Private Sub Cancela_click()
    PrgPasaDatos.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Proceso()
    
    With rstPasaClientes
        .Index = "Codigo"
        .MoveFirst
        Do
            
            WCodigo = IIf(IsNull(!Codigo), "", !Codigo)
            WNombre = IIf(IsNull(!Nombre), "", !Nombre)
            WDireccion = IIf(IsNull(!Direccion), "", !Direccion)
            WLocalidad = IIf(IsNull(!Localidad), "", !Localidad)
            WProvincia = IIf(IsNull(!Provincia), "", !Provincia)
            WIva = IIf(IsNull(!Iva), "", !Iva)
            WCuit = IIf(IsNull(!Cuit), "", !Cuit)
            WTelefono = IIf(IsNull(!Telefono), "", !Telefono)
            WObservaciones = IIf(IsNull(!Observaciones), "", !Observaciones)
                
            With rstClientes
                .Index = "Cliente"
                .Seek "=", WCodigo
                If .NoMatch Then
                    .AddNew
                    !Cliente = WCodigo
                    !Razon = WNombre
                    !Direccion = WDireccion
                    !Localidad = WLocalidad
                    Select Case Left$(WProvincia, 1)
                        Case "C"
                            !Provincia = "0"
                        Case "B"
                            !Provincia = "1"
                        Case Else
                            !Provincia = "0"
                    End Select
                    Select Case Val(WIva)
                        Case 1
                            !Iva = "1"
                        Case 4
                            !Iva = "4"
                        Case 6
                            !Iva = "6"
                            !Provincia = "24"
                        Case Else
                            !Iva = "1"
                    End Select
                    !Cuit = WCuit
                    !Telefono = WTelefono
                    !Observaciones = WObservaciones
                    !Postal = ""
                    !Observaciones = ""
                    !EMail = ""
                    !fax = ""
                    !Dias = 0
                    !Vendedor = 0
                    !Descuento = 0
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Cliente = WCodigo
                    !Razon = WNombre
                    !Direccion = WDireccion
                    !Localidad = WLocalidad
                    Select Case Left$(WProvincia, 1)
                        Case "C"
                            !Provincia = "0"
                        Case "B"
                            !Provincia = "1"
                        Case Else
                            !Provincia = "0"
                    End Select
                    Select Case Val(WIva)
                        Case 1
                            !Iva = "1"
                        Case 4
                            !Iva = "4"
                        Case 6
                            !Iva = "6"
                            !Provincia = "24"
                        Case Else
                            !Iva = "1"
                    End Select
                    !Cuit = WCuit
                    !Telefono = WTelefono
                    !Observaciones = WObservaciones
                    !Postal = ""
                    !Observaciones = ""
                    !EMail = ""
                    !fax = ""
                    !Dias = 0
                    !Vendedor = 0
                    !Descuento = 0
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
            
        Loop
    End With
    
    
    With rstPasaCuentas
        .Index = "Cuenta"
        .MoveFirst
        Do
            
            WCuenta = IIf(IsNull(!Cuenta), "", !Cuenta)
            WNivel = IIf(IsNull(!Nivel), "", !Nivel)
            WDescripcion = IIf(IsNull(!Descripcion), "", !Descripcion)
            WCuenta = Left$(WCuenta, 10)
                
            With rstCuenta
                .Index = "Cuenta"
                .Seek "=", WCuenta
                If .NoMatch Then
                    .AddNew
                    !Cuenta = WCuenta
                    !Descripcion = WDescripcion
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Cuenta = WCuenta
                    !Descripcion = WDescripcion
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
            
        Loop
    End With
    
    
    
    With rstPasaProve
        .Index = "Codigo"
        .MoveFirst
        Do
            
            WCodigo = IIf(IsNull(!Codigo), "", !Codigo)
            WNombre = IIf(IsNull(!Nombre), "", !Nombre)
            WDireccion = IIf(IsNull(!Direccion), "", !Direccion)
            WLocalidad = IIf(IsNull(!Localidad), "", !Localidad)
            WObservaciones = IIf(IsNull(!Observaciones), "", !Observaciones)
            WPostal = IIf(IsNull(!Postal), "", !Postal)
            WCuit = IIf(IsNull(!Cuit), "", !Cuit)
            WTelefono = IIf(IsNull(!Telefono), "", !Telefono)
            WGanancia = IIf(IsNull(!Ganancia), "", !Ganancia)
            WPago = IIf(IsNull(!Pago), "", !Pago)
            WTipo = IIf(IsNull(!Tipo), "", !Tipo)
            WDias = IIf(IsNull(!Dias), "", !Dias)
            WServicio = IIf(IsNull(!servicio), "", !servicio)
            WObserva = IIf(IsNull(!observa), "", !observa)
                
            With RstProveedor
                .Index = "Proveedor"
                .Seek "=", WCodigo
                If .NoMatch Then
                    .AddNew
                    !Proveedor = WCodigo
                    !Nombre = WNombre
                    !Direccion = WDireccion
                    !Localidad = WLocalidad
                    !Postal = WPostal
                    !Cuit = WCuit
                    !Telefono = WTelefono
                    !EMail = ""
                    !Observaciones = WObservaciones
                    !Dias = Val(WDias)
                    !Ganancia = Val(WGanancia) - 1
                    !Iva = 3
                    !Provincia = 1
                    !Tipo = Val(WTipo)
                    !NombreCheque = ""
                    !PorceReteIva = 0
                    !Reteiva = 0
                    .Update
                    .Bookmark = .LastModified
                        Else
                    .Edit
                    !Proveedor = WCodigo
                    !Nombre = WNombre
                    !Direccion = WDireccion
                    !Localidad = WLocalidad
                    !Postal = WPostal
                    !Cuit = WCuit
                    !Telefono = WTelefono
                    !EMail = ""
                    !Observaciones = WObservaciones
                    !Dias = Val(WDias)
                    !Ganancia = Val(WGanancia) - 1
                    !Iva = 3
                    !Provincia = 1
                    !Tipo = Val(WTipo)
                    !NombreCheque = ""
                    !PorceReteIva = 0
                    !Reteiva = 0
                    .Update
                    .Bookmark = .LastModified
                End If
            End With
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
            
        Loop
    End With
    
    
    
    
    
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub

Private Sub Form_Activate()
    OPEN_FILE_PasaClientes
    OPEN_FILE_Clientes
    OPEN_FILE_PasaCuentas
    OPEN_FILE_Cuenta
    OPEN_FILE_PasaProve
    OPEN_FILE_Proveedor
End Sub

