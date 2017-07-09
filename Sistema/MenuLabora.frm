VERSION 5.00
Begin VB.Form MenuVen 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Control de Gestion - Ventas y Stock : "
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   660
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MenuLabora.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "MenuLabora.frx":0442
   ScaleHeight     =   8310
   ScaleWidth      =   11640
   WindowState     =   2  'Maximized
   Begin VB.TextBox dadaII 
      Height          =   675
      Left            =   720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox dada 
      Height          =   675
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   840
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   0
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image Logo 
      Height          =   1005
      Left            =   8520
      Picture         =   "MenuLabora.frx":0884
      Top             =   6360
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label VentasII 
      BackColor       =   &H8000000C&
      Caption         =   "Ventas y Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   960
      MouseIcon       =   "MenuLabora.frx":1202
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Ventas y Stock"
      Top             =   7560
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   960
      Picture         =   "MenuLabora.frx":150C
      Top             =   120
      Visible         =   0   'False
      Width           =   9600
   End
   Begin VB.Menu adasdsadasdasdsad 
      Caption         =   "Laboratorio"
      Begin VB.Menu ControlPedidos 
         Caption         =   "Control de Pedidos"
      End
      Begin VB.Menu OrdenCpa 
         Caption         =   "Ingreso de Solicitud de Orden de Compra"
      End
      Begin VB.Menu Compra 
         Caption         =   "Ingreso de Remitos"
      End
      Begin VB.Menu Produccion 
         Caption         =   "Ingreso de Produccion sobre Pedidos"
      End
      Begin VB.Menu ProduccionII 
         Caption         =   "Ingreso de Produccion"
      End
      Begin VB.Menu caratula 
         Caption         =   "Impresion de Caratula"
      End
      Begin VB.Menu movstk 
         Caption         =   "Ajustes al Stock de Insumos"
      End
      Begin VB.Menu movstkii 
         Caption         =   "Ajustes al Stock de Articulos"
      End
      Begin VB.Menu dvfnghj 
         Caption         =   "--------------------------------------"
      End
      Begin VB.Menu OrdenFabricacion 
         Caption         =   "Listado de Ordenes de Fabricacion"
      End
      Begin VB.Menu ListadoSolInsumos 
         Caption         =   "Listado de Solicitud de Insumos Pendientes de Entrega"
      End
      Begin VB.Menu ListaOrdencpa 
         Caption         =   "Listado de Ordenes de Compra Pendientes"
      End
      Begin VB.Menu Listaremitos 
         Caption         =   "Listado de Remitos INgresados"
      End
      Begin VB.Menu ListaStocvGrupo 
         Caption         =   "Litado de Stock de Insumos"
      End
      Begin VB.Menu ListaStockArti 
         Caption         =   "Litado de Stock de Articulos"
      End
   End
   Begin VB.Menu fgh 
      Caption         =   "Procesos"
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "MenuVen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Registro As String


Private Sub cash_Click()
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgCash.Show
End Sub

Private Sub Banco_Click()
    PrgBanco.Show
End Sub

Private Sub Concepto_Click()
    PrgConcepto.Show
End Sub

Private Sub Cuentas_Click()
    PrgCuenta.Show
End Sub

Private Sub Agenda_Click()
    OPEN_FILE_Agenda
    PrgAgenda.Show
End Sub

Private Sub Aplica_Click()
    PrgAplica.Show
End Sub

Private Sub ActalizaCodigo_Click()
    prgActualizacionIndividual.Show
End Sub

Private Sub ActualizaCostoFuturo_Click()
    prgActualizacionCostoFuturo.Show
End Sub

Private Sub ActualizaGeneral_Click()
    prgActualizacionGeneral.Show
End Sub

Private Sub AplicaCteCte_Click()
    PrgAplicacTAcTE.Show
End Sub

Private Sub Calculadora_Click()
    Calculator.Show
End Sub

Private Sub Cotiza_Click()
    OPEN_FILE_Cotiza
    OPEN_FILE_Configuracion
    PrgCotiza.Show
End Sub

Private Sub Calidad_Click()
    If ZZNivel = 0 Then
        PrgCalidad.Show
    End If
End Sub

Private Sub caratula_Click()
    PrgCaratula.Show
End Sub

Private Sub Command1_Click()









Stop






    Open "ClientesNuevo.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCliente = Trim(Mid$(WDato, 1, 10))
        WRazon = Mid$(WDato, 12, 50)
        WFantasia = Mid$(WDato, 63, 30)
        WIb = Mid$(WDato, 94, 20)
        WCuit = Mid$(WDato, 115, 20)
        WGanancia = Mid$(WDato, 136, 20)
        WCalle = Mid$(WDato, 157, 30)
        WNumero = Mid$(WDato, 189, 7)
        WExtension = Mid$(WDato, 197, 15)
        WPostal = Mid$(WDato, 213, 8)
        WTelefono = Mid$(WDato, 222, 30)
        WFax = Mid$(WDato, 253, 20)
        WEmail = Mid$(WDato, 274, 50)
        WCateIva = Mid$(WDato, 325, 10)
        WCateGana = Mid$(WDato, 336, 10)
        WCateIb = Mid$(WDato, 347, 9)
        WFechaAlta = Mid$(WDato, 357, 10)
        WFechaBaja = Mid$(WDato, 368, 10)
        WTipoFactu = Mid$(WDato, 379, 10)
        
        WCodCta = Mid$(WDato, 390, 9)
        WCtaDeuda = Mid$(WDato, 400, 10)
        WCondVta = Mid$(WDato, 411, 9)
        WContactoI = Mid$(WDato, 421, 30)
        WTelefonoI = Mid$(WDato, 452, 30)
        WPuestoI = Mid$(WDato, 483, 15)
        WContactoII = Mid$(WDato, 499, 30)
        WTelefonoII = Mid$(WDato, 530, 15)
        WPuestoII = Mid$(WDato, 546, 30)
        WContactoIII = Mid$(WDato, 577, 30)
        WTelefonoIII = Mid$(WDato, 608, 15)
        WPuestoIII = Mid$(WDato, 624, 30)
        
        
        WM1 = Mid$(WDato, 655, 10)
        WM2 = Mid$(WDato, 666, 10)
        WM3 = Mid$(WDato, 677, 10)
        WM4 = Mid$(WDato, 688, 10)
        WM5 = Mid$(WDato, 699, 10)
        WM6 = Mid$(WDato, 710, 10)
        WM7 = Mid$(WDato, 721, 10)
        WM8 = Mid$(WDato, 732, 10)
        WM9 = Mid$(WDato, 743, 10)
        WM10 = Mid$(WDato, 754, 10)
        WM11 = Mid$(WDato, 765, 10)
        WM12 = Mid$(WDato, 776, 10)
        
        WExpreso = Mid$(WDato, 787, 40)
        WLocalidad = Mid$(WDato, 828, 30)
        WImpreProvincia = Mid$(WDato, 859, 30)
        WCabeza = Mid$(WDato, 890, 11)
        WPatro = Mid$(WDato, 902, 11)
        WCodGrupo = Mid$(WDato, 914, 20)
        WFra = Mid$(WDato, 935, 11)
        WPlan = Mid$(WDato, 947, 5)
        WLista = Mid$(WDato, 953, 6)
        WPP = Mid$(WDato, 960, 11)
        WXX1 = Mid$(WDato, 972, 10)
        WBonifica = Mid$(WDato, 983, 9)
        wxx2 = Mid$(WDato, 993, 10)
        WCalleII = Mid$(WDato, 1004, 30)
        WNumeroII = Mid$(WDato, 1035, 8)
        WExtensionII = Mid$(WDato, 1044, 15)
        WPostalII = Mid$(WDato, 1060, 9)
        WLocalidadII = Mid$(WDato, 1070, 30)
        WImpreProvinciaII = Mid$(WDato, 1101, 30)
        WTipoClie = Mid$(WDato, 1132, 8)
        WPotencial = Mid$(WDato, 1141, 10)
        WConocio = Mid$(WDato, 1152, 50)
        
        ZZDireccion = Trim(WCalle) + " " + Trim(WNumero) + " " + Trim(WExtension)
        ZZDireccionII = Trim(WCalleII) + " " + Trim(WNumeroII) + " " + Trim(WExtensionII)
        
        Select Case Trim(WCateIva)
            Case "MO"
                WIva = "5"
            Case "CF"
                WIva = "3"
            Case "Ex"
                WIva = "4"
            Case Else
                WIva = "1"
        End Select
        
        ZZFechaAlta = WFechaAlta
            
            
        WFantasia = Trim(WFantasia)
        ZZZZHasta = Len(WFantasia)
        For Ciclo = 1 To ZZZZHasta
            If Mid$(WFantasia, Ciclo, 1) = "'" Then
                WFantasia = Left$(WFantasia, Ciclo - 1) + Mid$(WFantasia, Ciclo + 1, 1000)
            End If
        Next Ciclo
            
        WRazon = Trim(WRazon)
        ZZZZHasta = Len(WRazon)
        For Ciclo = 1 To ZZZZHasta
            If Mid$(WRazon, Ciclo, 1) = "'" Then
                WRazon = Left$(WRazon, Ciclo - 1) + Mid$(WRazon, Ciclo + 1, 1000)
            End If
        Next Ciclo
            
        ZZDireccion = Trim(ZZDireccion)
        ZZZZHasta = Len(ZZDireccion)
        For Ciclo = 1 To ZZZZHasta
            If Mid$(ZZDireccion, Ciclo, 1) = "'" Then
                ZZDireccion = Left$(ZZDireccion, Ciclo - 1) + Mid$(ZZDireccion, Ciclo + 1, 1000)
            End If
        Next Ciclo
            
        WDireccionII = Trim(ZZDireccionII)
        ZZZZHasta = Len(ZZDireccionII)
        For Ciclo = 1 To ZZZZHasta
            If Mid$(ZZDireccionII, Ciclo, 1) = "'" Then
                ZZDireccionII = Left$(ZZDireccionII, Ciclo - 1) + Mid$(ZZDireccionII, Ciclo + 1, 1000)
            End If
        Next Ciclo
            
        WContactoI = Trim(WContactoI)
        ZZZZHasta = Len(WContactoI)
        For Ciclo = 1 To ZZZZHasta
            If Mid$(WContactoI, Ciclo, 1) = "'" Then
                WContactoI = Left$(WContactoI, Ciclo - 1) + Mid$(WContactoI, Ciclo + 1, 1000)
            End If
        Next Ciclo
            
                    
        WFantasia = Trim(WFantasia)
        ZZDireccion = Trim(ZZDireccion)
        WLocalidad = Trim(WLocalidad)
        WPostal = Trim(WPostal)
        WTelefono = Trim(WTelefono)
        WContactoI = Trim(WContactoI)
        WTelefonoI = Trim(WTelefonoI)
        WContactoII = Trim(WContactoII)
        WTelefonoII = Trim(WTelefonoII)
        WContactoIII = Trim(WContactoIII)
        WTelefonoIII = Trim(WTelefonoIII)
        WRazon = Trim(WRazon)
        ZZDireccionII = Trim(ZZDireccionII)
        WLocalidadII = Trim(WLocalidadII)
        WPostalII = Trim(WPostalII)
        WCuit = Trim(WCuit)
        WEmail = Trim(WEmail)
        WFax = Trim(WFax)
        WImpreProvincia = Trim(WImpreProvincia)
        WImpreProvinciaII = Trim(WImpreProvinciaII)
        WIva = Trim(WIva)
        WExpreso = Trim(WExpreso)
        WTipoClie = Trim(WTipoClie)
        WLista = Trim(WLista)
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
        
        
            rstCliente.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + " Razon = " + "'" + WFantasia + "',"
            ZSql = ZSql + " Direccion = " + "'" + ZZDireccion + "',"
            ZSql = ZSql + " Localidad = " + "'" + WLocalidad + "',"
            ZSql = ZSql + " Postal = " + "'" + WPostal + "',"
            ZSql = ZSql + " Telefono = " + "'" + WTelefono + "',"
            ZSql = ZSql + " NombreI = " + "'" + WContactoI + "',"
            ZSql = ZSql + " TelefonoI = " + "'" + WTelefonoI + "',"
            ZSql = ZSql + " NombreII = " + "'" + WContactoII + "',"
            ZSql = ZSql + " TelefonoII = " + "'" + WTelefonoII + "',"
            ZSql = ZSql + " NombreIII = " + "'" + WContactoIII + "',"
            ZSql = ZSql + " TelefonoIII = " + "'" + WTelefonoIII + "',"
            ZSql = ZSql + " Fantasia = " + "'" + WRazon + "',"
            ZSql = ZSql + " DireccionII = " + "'" + ZZDireccionII + "',"
            ZSql = ZSql + " LocalidadII = " + "'" + WLocalidadII + "',"
            ZSql = ZSql + " PostalII = " + "'" + WPostalII + "',"
            ZSql = ZSql + " Cuit = " + "'" + Left$(WCuit, 15) + "',"
            ZSql = ZSql + " Email = " + "'" + WEmail + "',"
            ZSql = ZSql + " Fax = " + "'" + WFax + "',"
            ZSql = ZSql + " ImpreProvincia = " + "'" + WImpreProvincia + "',"
            ZSql = ZSql + " ImpreProvinciaII = " + "'" + WImpreProvinciaII + "',"
            ZSql = ZSql + " Iva = " + "'" + WIva + "',"
            ZSql = ZSql + " Expreso = " + "'" + WExpreso + "',"
            Rem ZSql = ZSql + " TipoClie = " + "'" + WTipoClie + "',"
            ZSql = ZSql + " NroLista = " + "'" + WLista + "'"
            ZSql = ZSql + " Where Cliente = " + "'" + WCliente + "'"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
            
                    Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Cliente ("
            ZSql = ZSql + "Cliente ,"
            ZSql = ZSql + "Razon ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Postal ,"
            ZSql = ZSql + "ImpreProvincia ,"
            ZSql = ZSql + "Telefono ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "NombreI ,"
            ZSql = ZSql + "TelefonoI ,"
            ZSql = ZSql + "EmailI ,"
            ZSql = ZSql + "NombreII ,"
            ZSql = ZSql + "TelefonoII ,"
            ZSql = ZSql + "EmailII ,"
            ZSql = ZSql + "NombreIII ,"
            ZSql = ZSql + "TelefonoIII ,"
            ZSql = ZSql + "EmailIII ,"
            ZSql = ZSql + "Fantasia ,"
            ZSql = ZSql + "DireccionII ,"
            ZSql = ZSql + "LocalidadII ,"
            ZSql = ZSql + "PostalII ,"
            ZSql = ZSql + "ImpreProvinciaII ,"
            ZSql = ZSql + "FechaAlta ,"
            ZSql = ZSql + "Cuit ,"
            ZSql = ZSql + "Email ,"
            ZSql = ZSql + "Fax ,"
            ZSql = ZSql + "PorceIva ,"
            ZSql = ZSql + "Provincia ,"
            ZSql = ZSql + "Iva ,"
            ZSql = ZSql + "Expreso ,"
            ZSql = ZSql + "TipoClie ,"
            ZSql = ZSql + "NroLista ,"
            ZSql = ZSql + "Condicion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCliente + "',"
            ZSql = ZSql + "'" + WFantasia + "',"
            ZSql = ZSql + "'" + ZZDireccion + "',"
            ZSql = ZSql + "'" + WLocalidad + "',"
            ZSql = ZSql + "'" + WPostal + "',"
            ZSql = ZSql + "'" + WImpreProvincia + "',"
            ZSql = ZSql + "'" + WTelefono + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WContactoI + "',"
            ZSql = ZSql + "'" + WTelefonoI + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WContactoII + "',"
            ZSql = ZSql + "'" + WTelefonoII + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WContactoIII + "',"
            ZSql = ZSql + "'" + WTelefonoIII + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + WRazon + "',"
            ZSql = ZSql + "'" + ZZDireccionII + "',"
            ZSql = ZSql + "'" + WLocalidadII + "',"
            ZSql = ZSql + "'" + WPostalII + "',"
            ZSql = ZSql + "'" + WImpreProvinciaII + "',"
            ZSql = ZSql + "'" + ZZFechaAlta + "',"
            ZSql = ZSql + "'" + WCuit + "',"
            ZSql = ZSql + "'" + WEmail + "',"
            ZSql = ZSql + "'" + WFax + "',"
            ZSql = ZSql + "'" + WPlan + "',"
            ZSql = ZSql + "'" + "0" + "',"
            ZSql = ZSql + "'" + WIva + "',"
            ZSql = ZSql + "'" + WExpreso + "',"
            ZSql = ZSql + "'" + WTipoClie + "',"
            ZSql = ZSql + "'" + WLista + "',"
            ZSql = ZSql + "'" + "" + "')"
            spCliente = ZSql
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
    
        
    Loop
    
    Close #1























Stop


    Open "Bonifica.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        
        
        WCliente = Mid$(WDato, 1, 10)
        WCodigo = Mid$(WDato, 12, 15)
        WDesde = Mid$(WDato, 28, 10)
        WHasta = Mid$(WDato, 39, 10)
        WTopeI = Val(Mid$(WDato, 51, 10))
        WDtoI = Val(Mid$(WDato, 62, 10))
        WTopeII = Val(Mid$(WDato, 73, 10))
        WDtoII = Val(Mid$(WDato, 84, 10))
        WTopeIII = Val(Mid$(WDato, 95, 10))
        WDtoIII = Val(Mid$(WDato, 106, 10))
        WTopeIV = Val(Mid$(WDato, 117, 10))
        WDtoIV = Val(Mid$(WDato, 128, 10))
        
        If Val(WTopeI) > 99999 Then
            WTopeI = "99999"
        End If
        If Val(WTopeII) > 99999 Then
            WTopeII = "99999"
        End If
        If Val(WTopeIII) > 99999 Then
            WTopeIII = "99999"
        End If
        If Val(WTopeIV) > 99999 Then
            WTopeIV = "99999"
        End If
            
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
            
        WDesde = Mid$(WDesde, 4, 3) + Mid$(WDesde, 1, 3) + Right$(WDesde, 4)
        WHasta = Mid$(WHasta, 4, 3) + Mid$(WHasta, 1, 3) + Right$(WHasta, 4)

        WOrdDesde = Right$(WDesde, 4) + Mid$(WDesde, 4, 2) + Left$(WDesde, 2)
        WOrdHasta = Right$(WHasta, 4) + Mid$(WHasta, 4, 2) + Left$(WHasta, 2)

        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
        ZZClave = Trim(WCliente) + ZZCodigo

        ZSql = ""
        ZSql = ZSql + "INSERT INTO ClienteBonifica ("
        ZSql = ZSql + "Clave ,"
        ZSql = ZSql + "Cliente ,"
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Linea ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Fragancia ,"
        ZSql = ZSql + "Calidad ,"
        ZSql = ZSql + "Tamano ,"
        ZSql = ZSql + "Desde ,"
        ZSql = ZSql + "Hasta ,"
        ZSql = ZSql + "OrdDesde ,"
        ZSql = ZSql + "OrdHasta ,"
        ZSql = ZSql + "Tope1 ,"
        ZSql = ZSql + "Valor1 ,"
        ZSql = ZSql + "Tope2 ,"
        ZSql = ZSql + "Valor2 ,"
        ZSql = ZSql + "Tope3 ,"
        ZSql = ZSql + "Valor3 ,"
        ZSql = ZSql + "Tope4 ,"
        ZSql = ZSql + "Valor4 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZClave + "',"
        ZSql = ZSql + "'" + WCliente + "',"
        ZSql = ZSql + "'" + ZZCodigo + "',"
        ZSql = ZSql + "'" + WLinea + "',"
        ZSql = ZSql + "'" + WTipo + "',"
        ZSql = ZSql + "'" + WFragancia + "',"
        ZSql = ZSql + "'" + WCalidad + "',"
        ZSql = ZSql + "'" + WTamano + "',"
        ZSql = ZSql + "'" + WDesde + "',"
        ZSql = ZSql + "'" + WHasta + "',"
        ZSql = ZSql + "'" + WOrdDesde + "',"
        ZSql = ZSql + "'" + WOrdHasta + "',"
        ZSql = ZSql + "'" + Str$(WTopeI) + "',"
        ZSql = ZSql + "'" + Str$(WDtoI) + "',"
        ZSql = ZSql + "'" + Str$(WTopeII) + "',"
        ZSql = ZSql + "'" + Str$(WDtoII) + "',"
        ZSql = ZSql + "'" + Str$(WTopeIII) + "',"
        ZSql = ZSql + "'" + Str$(WDtoIII) + "',"
        ZSql = ZSql + "'" + Str$(WTopeIV) + "',"
        ZSql = ZSql + "'" + Str$(WDtoIV) + "')"
        spClienteBonifica = ZSql
        Set rstClienteBonifica = db.OpenRecordset(spClienteBonifica, dbOpenSnapshot, dbSQLPassThrough)
        
    Loop
    
    Close #1



Stop


































    Open "preciosNUEVOs.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WLista = Trim(Mid$(WDato, 3, 1))
        WCodigo = Mid$(WDato, 8, 15)
        WDesde = Mid$(WDato, 24, 10)
        WHasta = Mid$(WDato, 35, 10)
        WTope1 = Mid$(WDato, 47, 8)
        WValor1 = Mid$(WDato, 56, 7)
        WTope2 = Mid$(WDato, 64, 8)
        WValor2 = Mid$(WDato, 73, 7)
        WTope3 = Mid$(WDato, 81, 8)
        WValor3 = Mid$(WDato, 90, 7)
        WTope4 = Mid$(WDato, 98, 8)
        WValor4 = Mid$(WDato, 107, 7)
        
        WDesde = Mid$(WDesde, 4, 3) + Mid$(WDesde, 1, 3) + Right$(WDesde, 4)
        WHasta = Mid$(WHasta, 4, 3) + Mid$(WHasta, 1, 3) + Right$(WHasta, 4)
        
        If Trim(WHasta) = "" Then
            WHasta = "20991231"
        End If
        
        WOrdDesde = Right$(WDesde, 4) + Mid$(WDesde, 4, 2) + Left$(WDesde, 2)
        WOrdHasta = Right$(WHasta, 4) + Mid$(WHasta, 4, 2) + Left$(WHasta, 2)
            
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
        
        WMoneda = "0"
        
        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
        WClave = Trim(WLista) + ZZCodigo
            
        If Trim(WDesde) <> "" And Trim(WHasta) <> "" Then
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Precios"
            ZSql = ZSql + " Where Precios.Lista = " + "'" + WLista + "'"
            ZSql = ZSql + " and Precios.LInea = " + "'" + WLinea + "'"
            ZSql = ZSql + " and Precios.Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Precios.fragancia = " + "'" + WFragancia + "'"
            ZSql = ZSql + " and Precios.Calidad = " + "'" + WCalidad + "'"
            ZSql = ZSql + " and Precios.Tamano = " + "'" + WTamano + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
            
                WOrdDesdeII = rstPrecios!OrdDesde
                WOrdHastaII = rstPrecios!OrdHasta
                
                rstPrecios.Close
                    
                ZZLugar = ZZLugar + 1
                dada.Text = ZZLugar
                dadaII.Text = ZZCodigo
                
                DoEvents
                
                If WOrdDesde > WOrdDesdeII Or WOrdHasta > WOrdHastaII Then
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Precios SET "
                    ZSql = ZSql + " Desde = " + "'" + WDesde + "',"
                    ZSql = ZSql + " Hasta = " + "'" + WHasta + "',"
                    ZSql = ZSql + " OrdDesde = " + "'" + WOrdDesde + "',"
                    ZSql = ZSql + " OrdHasta = " + "'" + WOrdHasta + "',"
                    ZSql = ZSql + " Tope1 = " + "'" + WTope1 + "',"
                    ZSql = ZSql + " Valor1 = " + "'" + WValor1 + "',"
                    ZSql = ZSql + " Tope2 = " + "'" + WTope2 + "',"
                    ZSql = ZSql + " Valor2 = " + "'" + WValor2 + "',"
                    ZSql = ZSql + " Tope3 = " + "'" + WTope3 + "',"
                    ZSql = ZSql + " Valor3 = " + "'" + WValor3 + "',"
                    ZSql = ZSql + " Tope4 = " + "'" + WTope4 + "',"
                    ZSql = ZSql + " Valor4 = " + "'" + WValor4 + "'"
                    ZSql = ZSql + " Where Lista = " + "'" + WLista + "'"
                    ZSql = ZSql + " and LInea = " + "'" + WLinea + "'"
                    ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
                    ZSql = ZSql + " and Fragancia = " + "'" + WFragancia + "'"
                    ZSql = ZSql + " and Calidad = " + "'" + WCalidad + "'"
                    ZSql = ZSql + " and Tamano = " + "'" + WTamano + "'"
                    spPrecios = ZSql
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                    Else
                    
                ZZLugar = ZZLugar + 1
                dada.Text = ZZLugar
                dadaII.Text = ZZCodigo
                
                DoEvents
                    
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Precios ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Linea ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Fragancia ,"
                ZSql = ZSql + "Calidad ,"
                ZSql = ZSql + "Tamano ,"
                ZSql = ZSql + "Lista ,"
                ZSql = ZSql + "Moneda ,"
                ZSql = ZSql + "Desde ,"
                ZSql = ZSql + "Hasta ,"
                ZSql = ZSql + "OrdDesde ,"
                ZSql = ZSql + "OrdHasta ,"
                ZSql = ZSql + "Tope1 ,"
                ZSql = ZSql + "Valor1 ,"
                ZSql = ZSql + "Tope2 ,"
                ZSql = ZSql + "Valor2 ,"
                ZSql = ZSql + "Tope3 ,"
                ZSql = ZSql + "Valor3 ,"
                ZSql = ZSql + "Tope4 ,"
                ZSql = ZSql + "Valor4 )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + ZZCodigo + "',"
                ZSql = ZSql + "'" + WLinea + "',"
                ZSql = ZSql + "'" + WTipo + "',"
                ZSql = ZSql + "'" + WFragancia + "',"
                ZSql = ZSql + "'" + WCalidad + "',"
                ZSql = ZSql + "'" + WTamano + "',"
                ZSql = ZSql + "'" + WLista + "',"
                ZSql = ZSql + "'" + WMoneda + "',"
                ZSql = ZSql + "'" + WDesde + "',"
                ZSql = ZSql + "'" + WHasta + "',"
                ZSql = ZSql + "'" + WOrdDesde + "',"
                ZSql = ZSql + "'" + WOrdHasta + "',"
                ZSql = ZSql + "'" + WTope1 + "',"
                ZSql = ZSql + "'" + WValor1 + "',"
                ZSql = ZSql + "'" + WTope2 + "',"
                ZSql = ZSql + "'" + WValor2 + "',"
                ZSql = ZSql + "'" + WTope3 + "',"
                ZSql = ZSql + "'" + WValor3 + "',"
                ZSql = ZSql + "'" + WTope4 + "',"
                ZSql = ZSql + "'" + WValor4 + "')"
                spPrecios = ZSql
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
    Loop
    
    Close #1































    Open "monedaok.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WLista = Trim(Mid$(WDato, 3, 1))
        WCodigo = Mid$(WDato, 8, 15)
        ZZMoneda = Mid$(WDato, 24, 10)
            
        If Trim(WCodigo) = "ESC-PR-FE-1-7" Then Stop
            
            
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
        
        If Trim(ZZMoneda) = "U" Then
            WMoneda = "1"
                Else
            WMoneda = "0"
        End If
        
        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
        WClave = Trim(WLista) + ZZCodigo
            
                    
        ZZLugar = ZZLugar + 1
        dada.Text = ZZLugar
        dadaII.Text = ZZCodigo
        
        DoEvents
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + WLista + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + WLinea + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + WFragancia + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + WCalidad + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + WTamano + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
                
            rstPrecios.Close
            
                
            ZSql = ""
            ZSql = ZSql + "UPDATE Precios SET "
            ZSql = ZSql + " MOneda = " + "'" + WMoneda + "'"
            ZSql = ZSql + " Where Lista = " + "'" + WLista + "'"
            ZSql = ZSql + " and LInea = " + "'" + WLinea + "'"
            ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Fragancia = " + "'" + WFragancia + "'"
            ZSql = ZSql + " and Calidad = " + "'" + WCalidad + "'"
            ZSql = ZSql + " and Tamano = " + "'" + WTamano + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Loop
    
    Close #1













Stop












   GoTo da







    Open "productook.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WCodigo = Mid$(WDato, 1, 15)
        WDescripcion = Mid$(WDato, 17, 50)
        WDescripcionII = Mid$(WDato, 68, 20)
        WWFacturable = Mid$(WDato, 89, 1)
        WTipo = Mid$(WDato, 101, 2)
        WImportado = Mid$(WDato, 107, 1)
        WEstado = Mid$(WDato, 118, 1)
        WWFechaInactivo = Mid$(WDato, 126, 8)
        WSucursal = Mid$(WDato, 138, 2)
        WAreaSol = Mid$(WDato, 148, 2)
        WAreaRea = Mid$(WDato, 160, 2)
        WStock = Mid$(WDato, 172, 1)
        WSector = Mid$(WDato, 184, 2)
        WComision = Mid$(WDato, 196, 1)
        WWEtiqueta = Mid$(WDato, 208, 1)
        WInsumo = Mid$(WDato, 220, 10)
        WCodCombo = Mid$(WDato, 231, 8)
        WCosto = Mid$(WDato, 241, 10)
        WFechaCosto = Mid$(WDato, 252, 10)
            
        WWFechaInactivo = Mid$(WWFechaInactivo, 4, 3) + Mid$(WWFechaInactivo, 1, 3) + Right$(WWFechaInactivo, 4)
        WFechaCosto = Mid$(WFechaCosto, 4, 3) + Mid$(WFechaCosto, 1, 3) + Right$(WFechaCosto, 4)
            
            
        WStock = "0"
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
        
        If WEstado = "I" Then
            WActivo = "1"
            WFechaInactivo = Right$(WFechaInactivo, 2) + "/" + Mid$(WFechaInactivo, 5, 2) + "/" + Left$(WFechaInactivo, 4)
                Else
            WActivo = "0"
            WFechaInactivo = "  /  /    "
        End If
        
        If WWFacturable = "N" Then
            WFacturable = "1"
                Else
            WFacturable = "0"
        End If
        
        If WWEtiqueta = "N" Then
            WEtiqueta = "2"
                Else
            If WWEtiqueta = "N" Then
                WEtiqueta = "1"
                    Else
                WEtiqueta = "0"
            End If
        End If
                
        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
            
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.LInea = " + "'" + WLinea + "'"
        ZSql = ZSql + " and Articulo.Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Articulo.fragancia = " + "'" + WFragancia + "'"
        ZSql = ZSql + " and Articulo.Calidad = " + "'" + WCalidad + "'"
        ZSql = ZSql + " and Articulo.Tamano = " + "'" + WTamano + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
        
            rstArticulo.Close
                
            ZZLugar = ZZLugar + 1
            dada.Text = ZZLugar
            dadaII.Text = ZZCodigo
            
            DoEvents
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "',"
            ZSql = ZSql + " DescripcionII = " + "'" + WDescripcionII + "',"
            ZSql = ZSql + " Stock = " + "'" + WStock + "',"
            ZSql = ZSql + " Insumo = " + "'" + WInsumo + "',"
            ZSql = ZSql + " Sector = " + "'" + WSector + "',"
            ZSql = ZSql + " Activo = " + "'" + WActivo + "',"
            ZSql = ZSql + " FechaInactivo = " + "'" + WFechaInactivo + "',"
            ZSql = ZSql + " Facturable = " + "'" + WFacturable + "',"
            ZSql = ZSql + " Etiqueta = " + "'" + WEtiqueta + "'"
            ZSql = ZSql + " Where LInea = " + "'" + WLinea + "'"
            ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Fragancia = " + "'" + WFragancia + "'"
            ZSql = ZSql + " and Calidad = " + "'" + WCalidad + "'"
            ZSql = ZSql + " and Tamano = " + "'" + WTamano + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZZLugar = ZZLugar + 1
            dada.Text = ZZLugar
            dadaII.Text = ZZCodigo
            
            DoEvents
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Articulo ("
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Tipo ,"
            ZSql = ZSql + "Fragancia ,"
            ZSql = ZSql + "Calidad ,"
            ZSql = ZSql + "Tamano ,"
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "DescripcionII ,"
            ZSql = ZSql + "Stock ,"
            ZSql = ZSql + "Insumo ,"
            ZSql = ZSql + "Sector ,"
            ZSql = ZSql + "Activo ,"
            ZSql = ZSql + "FechaInactivo ,"
            ZSql = ZSql + "Facturable ,"
            ZSql = ZSql + "Etiqueta )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WLinea + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + WFragancia + "',"
            ZSql = ZSql + "'" + WCalidad + "',"
            ZSql = ZSql + "'" + WTamano + "',"
            ZSql = ZSql + "'" + ZZCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "',"
            ZSql = ZSql + "'" + WDescripcionII + "',"
            ZSql = ZSql + "'" + WStock + "',"
            ZSql = ZSql + "'" + WInsumo + "',"
            ZSql = ZSql + "'" + WSector + "',"
            ZSql = ZSql + "'" + WActivo + "',"
            ZSql = ZSql + "'" + WFechaInactivo + "',"
            ZSql = ZSql + "'" + WFacturable + "',"
            ZSql = ZSql + "'" + WEtiqueta + "')"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Loop
    
    Close #1



Stop





















da:


    Open "preciosok.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WLista = Trim(Mid$(WDato, 3, 1))
        WCodigo = Mid$(WDato, 8, 15)
        WDesde = Mid$(WDato, 24, 10)
        WHasta = Mid$(WDato, 35, 10)
        WTope1 = Mid$(WDato, 47, 8)
        WValor1 = Mid$(WDato, 56, 7)
        WTope2 = Mid$(WDato, 64, 8)
        WValor2 = Mid$(WDato, 73, 7)
        WTope3 = Mid$(WDato, 81, 8)
        WValor3 = Mid$(WDato, 90, 7)
        WTope4 = Mid$(WDato, 98, 8)
        WValor4 = Mid$(WDato, 107, 7)
        
        WDesde = Mid$(WDesde, 4, 3) + Mid$(WDesde, 1, 3) + Right$(WDesde, 4)
        WHasta = Mid$(WHasta, 4, 3) + Mid$(WHasta, 1, 3) + Right$(WHasta, 4)
        
        WOrdDesde = Right$(WDesde, 4) + Mid$(WDesde, 4, 2) + Left$(WDesde, 2)
        WOrdHasta = Right$(WHasta, 4) + Mid$(WHasta, 4, 2) + Left$(WHasta, 2)
            
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
        
        WMoneda = "0"
        
        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
        WClave = Trim(WLista) + ZZCodigo
            
        If Trim(WDesde) <> "" And Trim(WHasta) <> "" Then
                
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Precios"
            ZSql = ZSql + " Where Precios.Lista = " + "'" + WLista + "'"
            ZSql = ZSql + " and Precios.LInea = " + "'" + WLinea + "'"
            ZSql = ZSql + " and Precios.Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Precios.fragancia = " + "'" + WFragancia + "'"
            ZSql = ZSql + " and Precios.Calidad = " + "'" + WCalidad + "'"
            ZSql = ZSql + " and Precios.Tamano = " + "'" + WTamano + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
            
                WOrdDesdeII = rstPrecios!OrdDesde
                WOrdHastaII = rstPrecios!OrdHasta
                
                rstPrecios.Close
                    
                ZZLugar = ZZLugar + 1
                dada.Text = ZZLugar
                dadaII.Text = ZZCodigo
                
                DoEvents
                
                If WOrdDesde > WOrdDesdeII Or WOrdHasta > WOrdHastaII Then
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Precios SET "
                    ZSql = ZSql + " Desde = " + "'" + WDesde + "',"
                    ZSql = ZSql + " Hasta = " + "'" + WHasta + "',"
                    ZSql = ZSql + " OrdDesde = " + "'" + WOrdDesde + "',"
                    ZSql = ZSql + " OrdHasta = " + "'" + WOrdHasta + "',"
                    ZSql = ZSql + " Tope1 = " + "'" + WTope1 + "',"
                    ZSql = ZSql + " Valor1 = " + "'" + WValor1 + "',"
                    ZSql = ZSql + " Tope2 = " + "'" + WTope2 + "',"
                    ZSql = ZSql + " Valor2 = " + "'" + WValor2 + "',"
                    ZSql = ZSql + " Tope3 = " + "'" + WTope3 + "',"
                    ZSql = ZSql + " Valor3 = " + "'" + WValor3 + "',"
                    ZSql = ZSql + " Tope4 = " + "'" + WTope4 + "',"
                    ZSql = ZSql + " Valor4 = " + "'" + WValor4 + "'"
                    ZSql = ZSql + " Where Lista = " + "'" + WLista + "'"
                    ZSql = ZSql + " and LInea = " + "'" + WLinea + "'"
                    ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
                    ZSql = ZSql + " and Fragancia = " + "'" + WFragancia + "'"
                    ZSql = ZSql + " and Calidad = " + "'" + WCalidad + "'"
                    ZSql = ZSql + " and Tamano = " + "'" + WTamano + "'"
                    spPrecios = ZSql
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                    Else
                    
                ZZLugar = ZZLugar + 1
                dada.Text = ZZLugar
                dadaII.Text = ZZCodigo
                
                DoEvents
                    
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Precios ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "Codigo ,"
                ZSql = ZSql + "Linea ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Fragancia ,"
                ZSql = ZSql + "Calidad ,"
                ZSql = ZSql + "Tamano ,"
                ZSql = ZSql + "Lista ,"
                ZSql = ZSql + "Moneda ,"
                ZSql = ZSql + "Desde ,"
                ZSql = ZSql + "Hasta ,"
                ZSql = ZSql + "OrdDesde ,"
                ZSql = ZSql + "OrdHasta ,"
                ZSql = ZSql + "Tope1 ,"
                ZSql = ZSql + "Valor1 ,"
                ZSql = ZSql + "Tope2 ,"
                ZSql = ZSql + "Valor2 ,"
                ZSql = ZSql + "Tope3 ,"
                ZSql = ZSql + "Valor3 ,"
                ZSql = ZSql + "Tope4 ,"
                ZSql = ZSql + "Valor4 )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WClave + "',"
                ZSql = ZSql + "'" + ZZCodigo + "',"
                ZSql = ZSql + "'" + WLinea + "',"
                ZSql = ZSql + "'" + WTipo + "',"
                ZSql = ZSql + "'" + WFragancia + "',"
                ZSql = ZSql + "'" + WCalidad + "',"
                ZSql = ZSql + "'" + WTamano + "',"
                ZSql = ZSql + "'" + WLista + "',"
                ZSql = ZSql + "'" + WMoneda + "',"
                ZSql = ZSql + "'" + WDesde + "',"
                ZSql = ZSql + "'" + WHasta + "',"
                ZSql = ZSql + "'" + WOrdDesde + "',"
                ZSql = ZSql + "'" + WOrdHasta + "',"
                ZSql = ZSql + "'" + WTope1 + "',"
                ZSql = ZSql + "'" + WValor1 + "',"
                ZSql = ZSql + "'" + WTope2 + "',"
                ZSql = ZSql + "'" + WValor2 + "',"
                ZSql = ZSql + "'" + WTope3 + "',"
                ZSql = ZSql + "'" + WValor3 + "',"
                ZSql = ZSql + "'" + WTope4 + "',"
                ZSql = ZSql + "'" + WValor4 + "')"
                spPrecios = ZSql
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
        End If
        
    Loop
    
    Close #1










    Open "monedaok.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WLista = Trim(Mid$(WDato, 3, 1))
        WCodigo = Mid$(WDato, 8, 15)
        ZZMoneda = Mid$(WDato, 24, 10)
            
        WLinea = Mid$(WCodigo, 1, 3)
        WTipo = Mid$(WCodigo, 5, 2)
        WFragancia = Mid$(WCodigo, 8, 2)
        WCalidad = Mid$(WCodigo, 11, 1)
        WTamano = Mid$(WCodigo, 13, 1)
        
        If ZZMoneda = "U" Then
            WMoneda = "1"
                Else
            WMoneda = "0"
        End If
        
        ZZCodigo = Trim(WLinea) + "-" + Trim(WTipo) + "-" + Trim(WFragancia) + "-" + Trim(WCalidad) + "-" + Trim(WTamano)
        WClave = Trim(WLista) + ZZCodigo
            
                    
        ZZLugar = ZZLugar + 1
        dada.Text = ZZLugar
        dadaII.Text = ZZCodigo
        
        DoEvents
            
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Precios"
        ZSql = ZSql + " Where Precios.Lista = " + "'" + WLista + "'"
        ZSql = ZSql + " and Precios.LInea = " + "'" + WLinea + "'"
        ZSql = ZSql + " and Precios.Tipo = " + "'" + WTipo + "'"
        ZSql = ZSql + " and Precios.fragancia = " + "'" + WFragancia + "'"
        ZSql = ZSql + " and Precios.Calidad = " + "'" + WCalidad + "'"
        ZSql = ZSql + " and Precios.Tamano = " + "'" + WTamano + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
                
            rstPrecios.Close
            
                
            ZSql = ""
            ZSql = ZSql + "UPDATE Precios SET "
            ZSql = ZSql + " MOneda = " + "'" + WMoneda + "'"
            ZSql = ZSql + " Where Lista = " + "'" + WLista + "'"
            ZSql = ZSql + " and LInea = " + "'" + WLinea + "'"
            ZSql = ZSql + " and Tipo = " + "'" + WTipo + "'"
            ZSql = ZSql + " and Fragancia = " + "'" + WFragancia + "'"
            ZSql = ZSql + " and Calidad = " + "'" + WCalidad + "'"
            ZSql = ZSql + " and Tamano = " + "'" + WTamano + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Loop
    
    Close #1






















Stop






    Open "lineas.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 3)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Lineas"
        ZSql = ZSql + " Where Lineas.Codigo = " + "'" + WCodigo + "'"
        spLinea = ZSql
        Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        If rstLinea.RecordCount > 0 Then
            rstLinea.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Lineas SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Lineas ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spLinea = ZSql
            Set rstLinea = db.OpenRecordset(spLinea, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
    Loop
    
    Close #1










    Open "tipo.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoPro"
        ZSql = ZSql + " Where TipoPro.Codigo = " + "'" + WCodigo + "'"
        spTipoPro = ZSql
        Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoPro.RecordCount > 0 Then
            rstTipoPro.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE TipoPro SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO TipoPro ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spTipoPro = ZSql
            Set rstTipoPro = db.OpenRecordset(spTipoPro, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1







    Open "fragancia.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Fragancia"
        ZSql = ZSql + " Where Fragancia.Codigo = " + "'" + WCodigo + "'"
        spFragancia = ZSql
        Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
        If rstFragancia.RecordCount > 0 Then
            rstFragancia.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Fragancia SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spFragancia = ZSql
            Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Fragancia ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spFragancia = ZSql
            Set rstFragancia = db.OpenRecordset(spFragancia, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1





    Open "calidad.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Calidad"
        ZSql = ZSql + " Where Calidad.Codigo = " + "'" + WCodigo + "'"
        spCalidad = ZSql
        Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
        If rstCalidad.RecordCount > 0 Then
            rstCalidad.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Calidad SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spCalidad = ZSql
            Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Calidad ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spCalidad = ZSql
            Set rstCalidad = db.OpenRecordset(spCalidad, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1













    Open "tamaño.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Tamano"
        ZSql = ZSql + " Where Tamano.Codigo = " + "'" + WCodigo + "'"
        spTamano = ZSql
        Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
        If rstTamano.RecordCount > 0 Then
            rstTamano.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Tamano SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spTamano = ZSql
            Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Tamano ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spTamano = ZSql
            Set rstTamano = db.OpenRecordset(spTamano, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1

















    Open "Sector.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WTipoProceso = Mid$(WDato, 1, 2)
        WCodigo = Mid$(WDato, 3, 2)
        WDescripcion = Mid$(WDato, 11, 100)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Sector"
        ZSql = ZSql + " Where Sector.Codigo = " + "'" + WCodigo + "'"
        spSector = ZSql
        Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        If rstSector.RecordCount > 0 Then
            rstSector.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Sector SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Sector ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "')"
            spSector = ZSql
            Set rstSector = db.OpenRecordset(spSector, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1






    Open "Insumosok.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WProveedor = Mid$(WDato, 1, 8)
        WCodigo = Mid$(WDato, 12, 16)
        WDescripcion = Mid$(WDato, 29, 30)
        WMoneda = Mid$(WDato, 60, 1)
        WCosto = Mid$(WDato, 68, 10)
        WFecha = Mid$(WDato, 79, 10)
        WTipo = Mid$(WDato, 90, 1)
        WFechaCosto = "  /  /    "
        WOrdFechaCosto = "00000000"
        
        WFechaCosto = Mid$(WFecha, 4, 3) + Mid$(WFecha, 1, 3) + Right$(WFecha, 4)
        WOrdFechaCosto = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        
        If WMoneda = "P" Then
            WMoneda = "1"
                Else
            WMoneda = "2"
        End If
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Insumo"
        ZSql = ZSql + " Where Insumo.Codigo = " + "'" + WCodigo + "'"
        spInsumo = ZSql
        Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
        If rstInsumo.RecordCount > 0 Then
        
            rstInsumo.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Insumo SET "
            ZSql = ZSql + " Descripcion = " + "'" + WDescripcion + "',"
            ZSql = ZSql + " Linea = " + "'" + WTipo + "',"
            ZSql = ZSql + " Proveedor = " + "'" + WProveedor + "',"
            ZSql = ZSql + " Costo = " + "'" + WCosto + "',"
            ZSql = ZSql + " MOneda = " + "'" + WMoneda + "',"
            ZSql = ZSql + " FechaCosto = " + "'" + WFechaCosto + "',"
            ZSql = ZSql + " OrdFechaCosto = " + "'" + WOrdFechaCosto + "',"
            ZSql = ZSql + " Minimo = " + "'" + "" + "',"
            ZSql = ZSql + " Stock = " + "'" + "" + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + WCodigo + "'"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Insumo ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Linea ,"
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Costo ,"
            ZSql = ZSql + "Moneda ,"
            ZSql = ZSql + "FechaCosto ,"
            ZSql = ZSql + "OrdFechaCosto ,"
            ZSql = ZSql + "Minimo ,"
            ZSql = ZSql + "Stock )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WCodigo + "',"
            ZSql = ZSql + "'" + WDescripcion + "',"
            ZSql = ZSql + "'" + WTipo + "',"
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WCosto + "',"
            ZSql = ZSql + "'" + WMoneda + "',"
            ZSql = ZSql + "'" + WFechaCosto + "',"
            ZSql = ZSql + "'" + WOrdFechaCosto + "',"
            ZSql = ZSql + "'" + "" + "',"
            ZSql = ZSql + "'" + "" + "')"
            spInsumo = ZSql
            Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        
        
    Loop
    
    Close #1








Stop

    Open "proveedoresok.txt" For Input As #1
    
    Do
        Line Input #1, WDato
        If EOF(1) Then Exit Do
        
        WProveedor = Mid$(WDato, 1, 8)
        WRazon = Mid$(WDato, 12, 50)
        WComercial = Mid$(WDato, 63, 30)
        WNroIb = Mid$(WDato, 94, 20)
        WCuit = Mid$(WDato, 115, 20)
        WGananacia = Mid$(WDato, 136, 20)
        WCalle = Mid$(WDato, 157, 30)
        WNro = Mid$(WDato, 188, 8)
        WExtension = Mid$(WDato, 197, 15)
        WPostal = Mid$(WDato, 213, 8)
        WTelefono = Mid$(WDato, 222, 30)
        WFax = Mid$(WDato, 253, 20)
        WEmail = Mid$(WDato, 274, 50)
        WDireccion = Trim(WCalle) + " " + WNro
        
        WLocalidad = ""
        WDias = "0"
        WNombreCheque = ""
        WGamancia = "0"
        WIva = "0"
        WProvincia = "0"
        WObservaciones = ""
                
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            rstProveedor.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + " Nombre = " + "'" + WRazon + "',"
            ZSql = ZSql + " Direccion = " + "'" + WDireccion + "',"
            ZSql = ZSql + " Localidad = " + "'" + WLocalidad + "',"
            ZSql = ZSql + " Postal = " + "'" + WPostal + "',"
            ZSql = ZSql + " Cuit = " + "'" + WCuit + "',"
            ZSql = ZSql + " Telefono = " + "'" + WTelefono + "',"
            ZSql = ZSql + " EMail = " + "'" + WEmail + "',"
            ZSql = ZSql + " Observaciones = " + "'" + WObservaciones + "',"
            ZSql = ZSql + " Dias = " + "'" + WDias + "',"
            ZSql = ZSql + " Ganancia = " + "'" + WGanancia + "',"
            ZSql = ZSql + " Iva = " + "'" + WIva + "',"
            ZSql = ZSql + " Provincia = " + "'" + WProvincia + "',"
            ZSql = ZSql + " NombreCheque = " + "'" + WNombreCheque + "'"
            ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
            Rem spProveedor = ZSql
            Rem Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Proveedor ("
            ZSql = ZSql + "Proveedor ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "Direccion ,"
            ZSql = ZSql + "Localidad ,"
            ZSql = ZSql + "Postal ,"
            ZSql = ZSql + "Cuit ,"
            ZSql = ZSql + "Telefono ,"
            ZSql = ZSql + "Email ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Dias ,"
            ZSql = ZSql + "Ganancia ,"
            ZSql = ZSql + "Iva ,"
            ZSql = ZSql + "Provincia ,"
            ZSql = ZSql + "NombreCheque )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WProveedor + "',"
            ZSql = ZSql + "'" + WRazon + "',"
            ZSql = ZSql + "'" + WDireccion + "',"
            ZSql = ZSql + "'" + WLocalidad + "',"
            ZSql = ZSql + "'" + WPostal + "',"
            ZSql = ZSql + "'" + Left$(WCuit, 15) + "',"
            ZSql = ZSql + "'" + WTelefono + "',"
            ZSql = ZSql + "'" + WEmail + "',"
            ZSql = ZSql + "'" + WObservaciones + "',"
            ZSql = ZSql + "'" + WDias + "',"
            ZSql = ZSql + "'" + WGanancia + "',"
            ZSql = ZSql + "'" + WIva + "',"
            ZSql = ZSql + "'" + WProvincia + "',"
            ZSql = ZSql + "'" + WNombreCheque + "')"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        
    Loop
    
    Close #1





Stop














End Sub

Private Sub Compra_Click()
    If ZZNivel = 0 Then
        PrgIngresoRemito.Show
    End If
End Sub

Private Sub CondPago_Click()
    If ZZNivel = 0 Then
        PrgCondPago.Show
    End If
End Sub

Private Sub ctactefecha_Click()
    PrgCtaCtefecha.Show
End Sub

Private Sub CtaCteCampana_Click()
    PrgCtaCteCampana.Show
End Sub

Private Sub ConsultaPrecios_Click()
    prgConsultaArticulo.Show
End Sub

Private Sub Despacho_Click()
    PrgDespacho.Show
End Sub

Private Sub ControlPedidos_Click()
    If ZZNivel = 0 Then
        PrgControlPedido.Show
    End If
End Sub

Private Sub Empre_Click()
    Empresa.Show
End Sub

Private Sub Escalas_Click()
    Rem PrgEscalas.Show
End Sub

Private Sub Cliente_Click()
    If ZZNivel = 0 Then
        prgcliente.Show
    End If
End Sub

Private Sub Comiven_Click()
    PrgVentVend.Show
End Sub

Private Sub ComparaVen_Click()
    PrgComparaVen.Show
End Sub

Private Sub CompArt_Click()
    PrgCompArt.Show
End Sub

Private Sub CtaCte_Click()
    PrgCtaCte.Show
End Sub

Private Sub ctacte2_Click()
    PrgCtaCte1.Show
End Sub

Private Sub CtacteVen_Click()
    PrgCtaCteVen.Show
End Sub

Private Sub devol_Click()
    PrgDevolNuevo.Show
End Sub

Private Sub Esencia_Click()
    PrgEsencia.Show
End Sub

Private Sub Expreso_Click()
    PrgExpreso.Show
End Sub

Private Sub factu_Click()
    ZZPasaProcesoFactura = 0
    PrgFactura.Show
End Sub

Private Sub FactuExpo_Click()
    OPEN_FILE_Configuracion
    WVarios = 1
    PrgFactuExpo.Show
End Sub

Private Sub Fallados_Click()
    PrgFallados.Show
End Sub

Private Sub Factura_Click()

End Sub

Private Sub Fin_Click()
    Rem Menu.WindowState = 1
    Rem a = Shell("Sistema.exe", 1)
    Close
    End
End Sub

Private Sub Form_Load()
    If ZZNivel = "0" Then
        MenuVen.Caption = "Sistema de Laboratorio : Mc Fragancias S.A."
            Else
        MenuVen.Caption = "Sistema de Laboratorio : Mc Fragancias"
    End If
End Sub

Private Sub Formula_Click()
    If ZZNivel = 0 Then
        PrgFormula.Show
    End If
End Sub

Private Sub Fragancia_Click()
    If ZZNivel = 0 Then
        PrgFragancia.Show
    End If
End Sub

Private Sub GuiaTRansporte_Click()
    PrgGuiaTransporte.Show
End Sub

Private Sub HistorialCliente_Click()
    If ZZNivel = 0 Then
        PrgHistorialCliente.Show
    End If
End Sub

Private Sub Insumo_Click()
    If ZZNivel = 0 Then
        prgInsumo.Show
    End If
End Sub

Private Sub Linea_Click()
    If ZZNivel = 0 Then
        PrgLinea.Show
    End If
End Sub

Private Sub ListaClienteProvincia_Click()
    PrgListaClieProvincia.Show
End Sub

Private Sub ListaClienteVendedor_Click()
    PrgListaClieVende.Show
End Sub

Private Sub ListaClienteZona_Click()
    PrgListaClieZona.Show
End Sub

Private Sub ListaCostoProveedor_Click()
    PrgListaCostoProveedor.Show
End Sub

Private Sub listageneral_Click()
    PrgListaNovedadesPrecio.Show
End Sub

Private Sub ListaStockMinimo_Click()
    PrgListaStockMinimo.Show
End Sub

Private Sub ListaStockProve_Click()
    PrgListaStockProveedor.Show
End Sub

Private Sub Lista_Click()
    If ZZNivel = 0 Then
        PrgLista.Show
    End If
End Sub

Private Sub ListadoSolInsumos_Click()
    If ZZNivel = 0 Then
        PrgListaSolicitudInsumo.Show
    End If
End Sub

Private Sub ListadoSolInsumosII_Click()
    If ZZNivel = 0 Then
    End If
End Sub

Private Sub ListaOrdenCpaPendiente_Click()
    If ZZNivel = 0 Then
    End If
End Sub

Private Sub ListaOrdencpa_Click()
    PrgListaOrdenCpaPendiente.Show
End Sub

Private Sub Listaremitos_Click()
    PrgListaRemito.Show
End Sub

Private Sub ListaStockArti_Click()
    If ZZNivel = 0 Then
        PrgListaStock.Show
    End If
End Sub

Private Sub ListaStockValora_Click()
    If ZZNivel = 0 Then
        PrgListaStockValora.Show
    End If
End Sub

Private Sub ListaStocvGrupo_Click()
    If ZZNivel = 0 Then
        PrgListaMinimoInsumos.Show
    End If
End Sub

Private Sub ListaventasSemestre_Click()
    PrgListaEvolucionVentas.Show
End Sub

Private Sub NotasEnvio_Click()
    PrgNotaEnvio.Show
End Sub

Private Sub PasajecostoFuturo_Click()
    prgPasajeCostoFuturo.Show
End Sub

Private Sub movstkii_Click()
    PrgMovStkArticulo.Show
End Sub

Private Sub OrdenCpa_Click()
    If ZZNivel = 0 Then
        PrgSolicitudOrdenCompra.Show
    End If
End Sub

Private Sub OrdenCpaArti_Click()
    If ZZNivel = 0 Then
        PrgOrdenCompra.Show
    End If
End Sub

Private Sub OrdenFabricacion_Click()
    If ZZNivel = 0 Then
        PrgOrdenFabricacion.Show
    End If
End Sub

Private Sub Paridad_Click()
    If ZZNivel = 0 Then
        PrgDolar.Show
    End If
End Sub

Private Sub ProduccionII_Click()
    If ZZNivel = 0 Then
        PrgProduccionInterna.Show
    End If
End Sub

Private Sub Prove_Click()
    If ZZNivel = 0 Then
        PrgProve.Show
    End If
End Sub

Private Sub IngrePtoVend_Click()
    PrgIngrePtoVend.Show
End Sub

Private Sub IvaCompo_Click()
    PrgIvaCompo.Show
End Sub

Private Sub Ivaven_Click()
    PrgIvaven.Show
End Sub

Private Sub Lineas_Click()
    PrgZona.Show
End Sub

Private Sub ListaClieVende_Click()
    PrgListaClieVende.Show
End Sub

Private Sub Listapre_Click()
    PrgPrecioFam.Show
End Sub

Private Sub ListCoti_Click()
    PrgListCoti.Show
End Sub

Private Sub Listmov_Click()
    PrgListMov.Show
End Sub

Private Sub ListPedArt_Click()
    If ZZNivel = 0 Then
        PrgListPedaRT.Show
    End If
End Sub

Private Sub ListPedCli_Click()
    If ZZNivel = 0 Then
        PrgListPedCli.Show
    End If
End Sub

Private Sub Minimo_Click()
    PrgMinimo.Show
End Sub

Private Sub MovStk_Click()
    PrgMovStkInsumo.Show
End Sub

Private Sub lledatos_Click()
    PrgLeeDatos.Show
End Sub

Private Sub MOvvar_Click()
    PrgMovStk.Show
End Sub

Private Sub OrdenFactura_Click()
    PrgOrdenFactura.Show
End Sub

Private Sub parametro_Click()
    PrgParametro.Show
End Sub

Private Sub Pedido_Click()
    If ZZNivel = 0 Then
        ZZPasaProcesoPedido = 0
        PrgPedido.Show
    End If
End Sub

Private Sub Plantilla_Click()
    PrgPlantilla.Show
End Sub

Private Sub Prod_Click()
    If ZZNivel = 0 Then
        prgArticulo.Show
    End If
End Sub

Private Sub Proyec_Click()
    PrgProyCta.Show
End Sub

Private Sub Proyecto_Click()
    PrgProyecto.Show
End Sub

Private Sub Tipopro_Click()
    PrgTipoPro.Show
End Sub

Private Sub Produccion_Click()
    If ZZNivel = 0 Then
        ZZPasaProcesoFabrica = 0
        PrgProduccionPedido.Show
    End If
End Sub

Private Sub recibos_Click()
    PrgRecibos.Show
End Sub

Private Sub SaldosCta_Click()
    PrgSaldoCta.Show
End Sub

Private Sub Valua_Click()
    PrgValua.Show
End Sub

Private Sub sdfds_Click()
    If ZZNivel = 0 Then
        PrgListaPrecios.Show
    End If
End Sub

Private Sub Sector_Click()
    If ZZNivel = 0 Then
        PrgSector.Show
    End If
End Sub

Private Sub SubLinea_Click()
    PrgFamilia.Show
End Sub

Private Sub TotalArticulo_Click()
    PrgTotalArticulo.Show
End Sub

Private Sub Tamaño_Click()
    If ZZNivel = 0 Then
        PrgTamaño.Show
    End If
End Sub

Private Sub Tipo_Click()
    If ZZNivel = 0 Then
        PrgTipoPro.Show
    End If
End Sub

Private Sub TipoClie_Click()
    If ZZNivel = 0 Then
        PrgTipoClie.Show
    End If
End Sub

Private Sub Varios1_Click()
    WVarios = 1
    PrgVarios.Show
End Sub

Private Sub Varios2_Click()
    WVarios = 2
    PrgVarios.Show
End Sub

Private Sub Varios3_Click()
    WVarios = 3
    PrgVarios.Show
End Sub

Private Sub Vende_Click()
    PrgVendedor.Show
End Sub

Private Sub VentaCampana_Click()
    PrgVentaCampana.Show
End Sub

Private Sub VentArt_Click()
    PrgEstaart.Show
End Sub

Private Sub VentasCli_Click()
    PrgVentClie.Show
End Sub

Private Sub Ventasproy_Click()
    PrgVentProy.Show
End Sub

Private Sub ventclie_Click()
    PrgEstacli.Show
End Sub

Private Sub Form_Activate()
    Menu.Caption = "Sistema de Laboratorio : " + WNombreEmpresa
    If WEmpresa = "" Then
        Menu.Hide
        Empresa.Show
    End If
End Sub

Private Sub VentPcia_Click()
    PrgVentPcia.Show
End Sub
