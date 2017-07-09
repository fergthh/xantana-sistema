VERSION 5.00
Begin VB.Form PrgPasaart 
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
Attribute VB_Name = "PrgPasaart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()
    Call Proceso
    Call Cancela_click
End Sub

Private Sub Cancela_click()
    PrgPasaart.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Proceso()

    On Error GoTo Error
    
    With rstLibreria
            .Index = "Clave"
            .MoveFirst
            Do
            
                WCodigo = !Prefijo + Mid$(Str$(!Codigo), 2, 20)
                WDescripcion = !Descri
                WLinea = 1
                WProveedor = !Prove
                WCosto = !Precio2
                WDescuento = !Dto
                WPrecio = 0
                WMargen = !Margen1
                WPrecio1 = 0
                WMargen1 = !Margen2
                WStock = 0
                
                With rstArticulo
                        .Index = "Codigo"
                        .Seek "=", WCodigo
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WCodigo
                            !Descripcion = WDescripcion
                            !Linea = WLinea
                            !Proveedor = WProveedor
                            !Costo = WCosto
                            !Descuento = WDescuento
                            !Precio = WPrecio
                            !Margen = WMargen
                            !Precio1 = WPrecio1
                            !Margen1 = WMargen1
                            !Stock = WStock
                            !Empresa = 1
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WCodigo
                            !Descripcion = WDescripcion
                            !Linea = WLinea
                            !Proveedor = WProveedor
                            !Costo = WCosto
                            !Descuento = WDescuento
                            !Precio = WPrecio
                            !Margen = WMargen
                            !Precio1 = WPrecio1
                            !Margen1 = WMargen1
                            !Stock = WStock
                            !Empresa = 1
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
    OPEN_FILE_Config
    OPEN_FILE_Articulo
    OPEN_FILE_Libreria
End Sub

Private Sub Form_Load()

    With rstConfig
        .Index = "Codigo"
        .Seek "=", 1
        If .NoMatch = False Then
            WIva1 = !Iva1
            WIva2 = !Iva2
        End If
    End With

End Sub
