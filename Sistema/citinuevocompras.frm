VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCitinuevoCompras 
   AutoRedraw      =   -1  'True
   Caption         =   "Exportacion de Datos al CITI "
   ClientHeight    =   5610
   ClientLeft      =   3345
   ClientTop       =   2250
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   ScaleHeight     =   5610
   ScaleWidth      =   5475
   Begin VB.Frame Frame2 
      Height          =   5295
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1560
         TabIndex        =   9
         Top             =   2640
         Width           =   2055
      End
      Begin VB.DriveListBox Drive 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   2160
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
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
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
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
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Destino"
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
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgCitinuevoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XParam As String
Dim XVector(10000, 20) As String
Dim WTipo As String
Dim WPunto As String
Dim WNumero As String
Dim WCuit As String
Dim WNombre As String

Private Sub Acepta_Click()


    On Error GoTo WError
    

    WDrive = Drive.Drive
    WDir = Dir1.Path
    

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    
    Erase XVector
    Renglon = 0
    
    
    
    XNombre = WDir + "\" + "REGINFO_CV_COMPRAS_CBTE" + ".txt"
    Open XNombre For Output As #1
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM IvaComp"
    spIvaComp = ZSql
    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
    If rstIvaComp.RecordCount > 0 Then
    
        With rstIvaComp
            .MoveFirst
            Do
                If .EOF = False Then
                    
                    WFecha = ""
                    WFecha = Right$(!Periodo, 4) + Mid$(!Periodo, 4, 2) + Left$(!Periodo, 2)
                    If WDesde <= WFecha And WFecha <= WHasta Then
                        If !Letra = "A" Or !Letra = "B" Or !Letra = "C" Or !Letra = "M" Then
                            If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Or Val(!Tipo) = 3 Then
                            
                                    Renglon = Renglon + 1
                                    XVector(Renglon, 1) = !Letra
                                    XVector(Renglon, 2) = !Tipo
                                    XVector(Renglon, 3) = !Punto
                                    XVector(Renglon, 4) = !Numero
                                    XVector(Renglon, 5) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                                    XVector(Renglon, 6) = !Proveedor
                                    XVector(Renglon, 7) = Str$(!Neto)
                                    XVector(Renglon, 8) = Str$(!Exento)
                                    XVector(Renglon, 9) = Str$(!Iva21)
                                    XVector(Renglon, 10) = Str$(!Iva5)
                                    XVector(Renglon, 11) = Str$(!Iva27)
                                    XVector(Renglon, 12) = Str$(!Ib)
                                    XVector(Renglon, 13) = Str$(!ImpInterno)
                                    XVector(Renglon, 14) = Str$(!ImpCombustible)
                                    XVector(Renglon, 15) = !Fecha
                                    XVector(Renglon, 16) = Str$(!Iva105)
                                
                            End If
                        End If
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCtePrv.Close
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    For Ciclo = 1 To Renglon
    
        Rem If ciclo = 51 Then Stop
        Rem If ciclo = 66 Then Stop
        Rem If ciclo = 100 Then Stop
        Rem If ciclo = 122 Then Stop
    
        WLetra = XVector(Ciclo, 1)
        WTipo = XVector(Ciclo, 2)
        WPunto = XVector(Ciclo, 3)
        WNumero = XVector(Ciclo, 4)
        WFecha = XVector(Ciclo, 5)
        WProveedor = XVector(Ciclo, 6)
        
        ZZNeto = Val(XVector(Ciclo, 7))
        ZZNeto = Int(ZZNeto * 100)
        WNeto = ZZNeto
        
        ZZExento = Val(XVector(Ciclo, 8))
        ZZExento = Int(ZZExento * 100)
        WExento = ZZExento
        
        ZZIva21 = Val(XVector(Ciclo, 9))
        ZZIva21 = Int(ZZIva21 * 100)
        WIva21 = ZZIva21
        
        ZZIva5 = Val(XVector(Ciclo, 10))
        ZZIva5 = Int(ZZIva5 * 100)
        WIva5 = ZZIva5
        
        ZZIva27 = Val(XVector(Ciclo, 11))
        ZZIva27 = Int(ZZIva27 * 100)
        WIva27 = ZZIva27
        
        ZZIva105 = Val(XVector(Ciclo, 16))
        ZZIva105 = Int(ZZIva105 * 100)
        WIva105 = ZZIva105
        
        ZZZZIb = Val(XVector(Ciclo, 12)) + Val(XVector(Ciclo, 13)) + Val(XVector(Ciclo, 14))
        ZZZZIb = Int(ZZZZIb * 100)
        WIb = ZZZZIb
        
        WResto = 0
        
        WDespacho = Space$(16)
        
        Rem WDespacho = Trim(XVector(Ciclo, 13))
        Rem If Trim(WDespacho) <> "" Then
        Rem     ZZLargo = Len(WDespacho)
        Rem     For ZZCiclo = 1 To ZZLargo
        Rem         If Mid$(WDespacho, ZZCiclo, 1) = " " Then
        Rem             WDespacho = Left$(WDespacho, ZZCiclo - 1) + "" + Mid$(WDespacho, ZZCiclo + 1, 50)
        Rem         End If
        Rem     Next ZZCiclo
        Rem End If
        Rem WDespacho = Left$(WDespacho + Space$(16), 16)
        Rem If Trim(WDespacho) <> "" Then
        Rem     WDespacho = Left$(Trim(WDespacho) + "0000000000000000", 16)
        Rem End If
        
        Rem If WNumero = 2348 Then Stop

        
        If WLetra = "B" Or WLetra = "C" Then
            If WExento <> 0 Then
                WNeto = WNeto + WExento
                WExento = 0
            End If
        End If
        
        Rem wnrointerno = XVector(Ciclo, 14)
        WFechaII = XVector(Ciclo, 15)
        
        Rem Select Case WProveedor
        Rem     Case "10065511620", "10070956507", "10065786411"
        Rem         WIva = WIva21 + WIva27 + WIva105
        Rem         WIva27 = WIva
        Rem         WIva21 = 0
        Rem         WIva105 = 0
        Rem     Case "10053718600", "10050001091", "10099924210", "10050000845"
        Rem         WIva = WIva21 + WIva27 + WIva105
        Rem         WIva105 = WIva
        Rem         WIva21 = 0
        Rem         WIva27 = 0
        Rem End Select
        
        WIva = WIva21 + WIva27 + WIva105
        If WIva = 0 Then
            WNeto = WNeto + WExento
            WExento = 0
        End If
        
        WTotal = WNeto + WExento + WIva21 + WIva5 + WIva27 + WIva105 + WIb
        If Trim(WDespacho) <> "" Then
            WImpo = Int(WIva21 / 21 * 100)
            WTotal = WTotal + WImpo
        End If
        
        If WIva = 0 Then
            WCodigoExento = "N"
            Rem z   zona de exportacion
            Rem x   exportaciones al enterior
            Rem e   operaciones exentas
            Rem C   Operaciones de canje
                Else
            WCodigoExento = " "
        End If
            
        
        
        WAlicuota = 0
        If WIva21 <> 0 Then
            WAlicuota = WAlicuota + 1
        End If
        If WIva27 <> 0 Then
            WAlicuota = WAlicuota + 1
        End If
        If WIva105 <> 0 Then
            WAlicuota = WAlicuota + 1
        End If
        If WLetra = "A" Or WLetra = "M" Then
            If WIva = 0 And WNeto <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Proveedor"
        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
        spProveedor = ZSql
        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedor.RecordCount > 0 Then
            WNombre = rstProveedor!Nombre
            WNombre = WNombre + Space$(30)
            WNombre = Left$(WNombre, 30)
            WCuit = rstProveedor!Cuit
            rstProveedor.Close
            Call Eval
        End If
        
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 5)
        Call Ceros(WNumero, 20)
        Call Ceros(WCuit, 20)
        
        Rem fecha
        WImpo1 = WFecha
        
        Rem tipo de comprobante
        Select Case WLetra
            Case "A"
                Select Case Val(WTipo)
                    Case 1
                        Wimpo2 = "001"
                    Case 2
                        Wimpo2 = "002"
                    Case 3
                        Wimpo2 = "003"
                    Case Else
                        Wimpo2 = "000"
                End Select
            Case "B"
                Select Case Val(WTipo)
                    Case 1
                        Wimpo2 = "006"
                    Case 2
                        Wimpo2 = "007"
                    Case 3
                        Wimpo2 = "008"
                    Case Else
                        Wimpo2 = "000"
                End Select
            Case "C"
                Select Case Val(WTipo)
                    Case 1
                        Wimpo2 = "011"
                    Case 2
                        Wimpo2 = "012"
                    Case 3
                        Wimpo2 = "013"
                    Case Else
                        Wimpo2 = "000"
                End Select
            Case "M"
                Select Case Val(WTipo)
                    Case 1
                        Wimpo2 = "051"
                    Case 2
                        Wimpo2 = "052"
                    Case 3
                        Wimpo2 = "053"
                    Case Else
                        Wimpo2 = "000"
                End Select
            Case Else
                Wimpo2 = "000"
        End Select
        
        If Trim(WDespacho) <> "" Then
            Wimpo2 = "066"
            WPunto = "0"
            WNumero = "0"
            Call Ceros(WPunto, 5)
            Call Ceros(WNumero, 20)
        End If
        
        Rem punto
        WImpo3 = WPunto
        
        Rem Numero
        WImpo4 = WNumero
        
        Rem despacho de importacion
        WImpo5 = WDespacho
        
        Rem tipo de doc
        WImpo6 = "80"
        
        Rem numero de doc
        WImpo7 = WCuit
        
        Rem razon social
        WImpo8 = WNombre
        
        Rem total
        If WTotal >= 0 Then
            Auxi1 = Str$(WTotal)
            Call Ceros(Auxi1, 15)
            WImpo9 = Auxi1
                Else
            Auxi1 = Str$(Abs(WTotal))
            Call Ceros(Auxi1, 14)
            WImpo9 = "0" + Auxi1
        End If
        
        Rem resto del neto
        If WResto >= 0 Then
            Auxi1 = Str$(WResto)
            Call Ceros(Auxi1, 15)
            WImpo11 = Auxi1
                Else
            Auxi1 = Str$(Abs(WResto))
            Call Ceros(Auxi1, 14)
            WImpo11 = "0" + Auxi1
        End If
        
        Rem exento
        If WExento >= 0 Then
            Auxi1 = Str$(WExento)
            Call Ceros(Auxi1, 15)
            WImpo10 = Auxi1
                Else
            Auxi1 = Str$(Abs(WExento))
            Call Ceros(Auxi1, 14)
            WImpo10 = "0" + Auxi1
        End If
            
        Rem percepcion de iva
        If WIva5 >= 0 Then
            Auxi1 = Str$(WIva5)
            Call Ceros(Auxi1, 15)
            WImpo12 = Auxi1
                Else
            Auxi1 = Str$(Abs(WIva5))
            Call Ceros(Auxi1, 14)
            WImpo12 = "0" + Auxi1
        End If
                
        
        Rem otros impuuestos nacionales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo13 = Auxi1
        
        Rem ingresos brutos
        If WIb >= 0 Then
            Auxi1 = Str$(WIb)
            Call Ceros(Auxi1, 15)
            WImpo14 = Auxi1
                Else
            Auxi1 = Str$(Abs(WIb))
            Call Ceros(Auxi1, 14)
            WImpo14 = "0" + Auxi1
        End If
        
        Rem otros impuuestos municipales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo15 = Auxi1
        
        Rem otros impuuestos internos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo16 = Auxi1
        
        Rem codigo de moneda
        WImpo17 = "PES"
        
        Rem PARIDAD
        Rem ZCAmbio = "0"
        Rem spCambios = "ConsultaCambio " + "'" + WFechaII + "'"
        Rem Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstCambios.RecordCount > 0 Then
        Rem     ZCAmbio = rstCambios!Cambio
        Rem     rstCambios.Close
        Rem             Else
        Rem     ZCAmbio = "1"
        Rem End If
        Rem Auxi1 = Str$(Int(ZCAmbio * 1000000))
        
        ZCAmbio = "1"
        Auxi1 = Str$(Int(ZCAmbio * 1000000))
        Call Ceros(Auxi1, 10)
        WImpo18 = Auxi1
        
        Rem CANTIDAD DE ALICUOTAS de iva
        If WLetra = "A" Or WLetra = "M" Then
            If WAlicuota = 0 Then
                WAlicuota = "1"
            End If
        End If
        Auxi1 = Str$(WAlicuota)
        Call Ceros(Auxi1, 1)
        WImpo19 = Auxi1
        
        Rem codigo de operacion
        WImpo20 = WCodigoExento
        
        Rem iva
        If WIva >= 0 Then
            Auxi1 = Str$(WIva)
            Call Ceros(Auxi1, 15)
            WImpo21 = Auxi1
                Else
            Auxi1 = Str$(Abs(WIva))
            Call Ceros(Auxi1, 14)
            WImpo21 = "0" + Auxi1
        End If
        
        Rem otros triburtos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo22 = Auxi1
        
        Rem cuit del emisor ????
        Auxi1 = "0"
        Call Ceros(Auxi1, 11)
        WImpo23 = Auxi1
        
        Rem nombre del emisor ????
        WImpo24 = Space$(30)
        
        Rem iva co,ision ????
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo25 = Auxi1
        
        Rem If Val(WEmpresa) = 1 Then
        Rem     WCuitII = "30549165083"
        Rem     WNombreII = Left$("SURFACTAN S.A." + Space$(25), 25)
        Rem         Else
        Rem     WCuitII = ""
        Rem     WNombreII = ""
        Rem End If
        Rem WCuitII = "00000000000"
        Rem WNombreII = Space$(25)
        
        If Val(WCuit) <> 0 Then
            ZZSuma = ZZSuma + 1
            WImpre = WImpo1 + Wimpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8 + WImpo9 + WImpo10 + WImpo11 + WImpo12 + WImpo13 + WImpo14 + WImpo15 + WImpo16 + WImpo17 + WImpo18 + WImpo19 + WImpo20 + WImpo21 + WImpo22 + WImpo23 + WImpo24 + WImpo25
            Rem WImpre = Str$(Ciclo) + " " + WNroInterno + " " + WImpo1 + WImpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8 + WImpo9 + WImpo10 + WImpo11 + WImpo12 + WImpo13 + WImpo14 + WImpo15 + WImpo16 + WImpo17 + WImpo18 + WImpo19 + WImpo20 + WImpo21 + WImpo22 + WImpo23 + WImpo24 + WImpo25
            Print #1, WImpre
            Rem Print #2, Str$(ZZSuma) + " " + wnrointerno + " " + WImpre
        End If
    
        
    Next Ciclo
    
    Close #1
    
    
    
    
    
    XNombre = WDir + "\" + "REGINFO_CV_COMPRAS_ALICUOTAS" + ".txt"
    Open XNombre For Output As #1
    
    For Ciclo = 1 To Renglon
    
        WLetra = XVector(Ciclo, 1)
        
        If WLetra = "A" Or WLetra = "M" Then
        
            WTipo = XVector(Ciclo, 2)
            WPunto = XVector(Ciclo, 3)
            WNumero = XVector(Ciclo, 4)
            WFecha = XVector(Ciclo, 5)
            WProveedor = XVector(Ciclo, 6)
            
            ZZNeto = Val(XVector(Ciclo, 7))
            ZZNeto = Int(ZZNeto * 100)
            WNeto = ZZNeto
            
            ZZExento = Val(XVector(Ciclo, 8))
            ZZExento = Int(ZZExento * 100)
            WExento = ZZExento
            
            ZZIva21 = Val(XVector(Ciclo, 9))
            ZZIva21 = Int(ZZIva21 * 100)
            WIva21 = ZZIva21
            
            ZZIva5 = Val(XVector(Ciclo, 10))
            ZZIva5 = Int(ZZIva5 * 100)
            WIva5 = ZZIva5
            
            ZZIva27 = Val(XVector(Ciclo, 11))
            ZZIva27 = Int(ZZIva27 * 100)
            WIva27 = ZZIva27
            
            ZZIva105 = Val(XVector(Ciclo, 16))
            ZZIva105 = Int(ZZIva105 * 100)
            WIva105 = ZZIva105
            
            ZZZZIb = Val(XVector(Ciclo, 12)) + Val(XVector(Ciclo, 13)) + Val(XVector(Ciclo, 14))
            ZZZZIb = Int(ZZZZIb * 100)
            WIb = ZZZZIb
        
            WDespacho = Space$(16)
            
            Rem WDespacho = Trim(XVector(Ciclo, 13))
            Rem If Trim(WDespacho) <> "" Then
            Rem     ZZLargo = Len(WDespacho)
            Rem     For ZZCiclo = 1 To ZZLargo
            Rem         If Mid$(WDespacho, ZZCiclo, 1) = " " Then
            Rem             WDespacho = Left$(WDespacho, ZZCiclo - 1) + "" + Mid$(WDespacho, ZZCiclo + 1, 50)
            Rem         End If
            Rem     Next ZZCiclo
            Rem End If
            Rem WDespacho = Left$(WDespacho + Space$(16), 16)
            Rem If Trim(WDespacho) <> "" Then
            Rem     WDespacho = Left$(Trim(WDespacho) + "0000000000000000", 16)
            Rem End If
            
            Rem wnrointerno = XVector(Ciclo, 14)
            WFechaII = XVector(Ciclo, 15)
            
            Rem Select Case WProveedor
            Rem     Case "10065511620", "10070956507", "10065786411"
            Rem         WIva = WIva21 + WIva27 + WIva105
            Rem         WIva27 = WIva
            Rem         WIva21 = 0
            Rem         WIva105 = 0
            Rem     Case "10053718600", "10050001091", "10099924210", "10050000845"
            Rem         WIva = WIva21 + WIva27 + WIva105
            Rem         WIva105 = WIva
            Rem         WIva21 = 0
            Rem         WIva27 = 0
            Rem End Select
            
            WIva = WIva21 + WIva27 + WIva105
            If WIva = 0 Then
                WNeto = WNeto + WExento
                WExento = 0
            End If
            
            WTotal = WNeto + WExento + WIva21 + WIva5 + WIva27 + WIva105 + WIb
            
            
            If WIva = 0 Then
                WCodigoExento = "N"
                Rem z   zona de exportacion
                Rem x   exportaciones al enterior
                Rem e   operaciones exentas
                Rem C   Operaciones de canje
                    Else
                WCodigoExento = " "
            End If
                
            Rem If WNumero = 2348 Then Stop
            
            
            WAlicuota = 0
            If WIva21 <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
            If WIva27 <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
            If WIva105 <> 0 Then
                WAlicuota = WAlicuota + 1
            End If
            If WLetra = "A" Or WLetra = "M" Then
                If WIva = 0 And WNeto <> 0 Then
                    WAlicuota = WAlicuota + 1
                End If
            End If
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
            spProveedor = ZSql
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                WNombre = rstProveedor!Nombre
                WNombre = WNombre + Space$(30)
                WNombre = Left$(WNombre, 30)
                WCuit = rstProveedor!Cuit
                rstProveedor.Close
                Call Eval
            End If
            
            Call Ceros(WTipo, 2)
            Call Ceros(WPunto, 5)
            Call Ceros(WNumero, 20)
            Call Ceros(WCuit, 20)
            
            If Val(WCuit) <> 0 And Trim(WDespacho) = "" Then
        
                Rem tipo de comprobante
                Select Case WLetra
                    Case "A"
                        Select Case Val(WTipo)
                            Case 1
                                WImpo1 = "001"
                            Case 2
                                WImpo1 = "002"
                            Case 3
                                WImpo1 = "003"
                            Case Else
                                WImpo1 = "000"
                        End Select
                    Case "M"
                        Select Case Val(WTipo)
                            Case 1
                                WImpo1 = "051"
                            Case 2
                                WImpo1 = "052"
                            Case 3
                                WImpo1 = "053"
                            Case Else
                                WImpo1 = "000"
                        End Select
                    Case Else
                        WImpo1 = "000"
                End Select
        
                If Trim(WDespacho) <> "" Then
                    WImpo1 = "066"
                    WPunto = "0"
                    WNumero = "0"
                    Call Ceros(WPunto, 5)
                    Call Ceros(WNumero, 20)
                End If
                
                Rem punto
                Wimpo2 = WPunto
                
                Rem Numero
                WImpo3 = WNumero
                
                Rem tipo de doc
                WImpo4 = "80"
                
                Rem numero de doc
                WImpo5 = WCuit
                
                If WIva21 <> 0 Then
                
                    If WAlicuota = 1 Then
                    
                        Rem neto
                        If WNeto >= 0 Then
                            Auxi1 = Str$(WNeto)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WNeto))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                            Else
                            
                        WImpo = Int(WIva21 / 21 * 100)
                        If WImpo >= 0 Then
                            Auxi1 = Str$(WImpo)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WImpo))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                    End If
                    
                    Rem Iva 21
                    WImpo7 = "0005"
                    
                    Rem impo iva
                    If WIva21 >= 0 Then
                        Auxi1 = Str$(WIva21)
                        Call Ceros(Auxi1, 15)
                        WImpo8 = Auxi1
                            Else
                        Auxi1 = Str$(Abs(WIva21))
                        Call Ceros(Auxi1, 14)
                        WImpo8 = "0" + Auxi1
                    End If
                    
        
                    WImpre = WImpo1 + Wimpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                End If
                
                If WIva105 <> 0 Then
                
                    If WAlicuota = 1 Then
                    
                        Rem neto
                        If WNeto >= 0 Then
                            Auxi1 = Str$(WNeto)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WNeto))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                            Else
                            
                        WImpo = Int(WIva105 / 10.5 * 100)
                        If WImpo >= 0 Then
                            Auxi1 = Str$(WImpo)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WImpo))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                    End If
                    
                    Rem Iva 10.5
                    WImpo7 = "0004"
                    
                    Rem impo iva
                    Auxi1 = Str$(WIva105)
                    Call Ceros(Auxi1, 15)
                    WImpo8 = Auxi1
        
                    WImpre = WImpo1 + Wimpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                    
                End If
                
                If WIva27 <> 0 Then
                
                    If WAlicuota = 1 Then
                    
                        Rem neto
                        If WNeto >= 0 Then
                            Auxi1 = Str$(WNeto)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WNeto))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                            Else
                            
                        WImpo = Int(WIva27 / 27 * 100)
                        If WImpo >= 0 Then
                            Auxi1 = Str$(WImpo)
                            Call Ceros(Auxi1, 15)
                            WImpo6 = Auxi1
                                Else
                            Auxi1 = Str$(Abs(WImpo))
                            Call Ceros(Auxi1, 14)
                            WImpo6 = "0" + Auxi1
                        End If
                        
                    End If
                    
                    
                    Rem Iva 10.5
                    WImpo7 = "0006"
                    
                    Rem impo iva
                    Auxi1 = Str$(WIva27)
                    Call Ceros(Auxi1, 15)
                    WImpo8 = Auxi1
        
                    WImpre = WImpo1 + Wimpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                    
                End If
                            
                If WIva = 0 And WNeto <> 0 Then
                
                    Rem neto
                    If WNeto >= 0 Then
                        Auxi1 = Str$(WNeto)
                        Call Ceros(Auxi1, 15)
                        WImpo6 = Auxi1
                            Else
                        Auxi1 = Str$(Abs(WNeto))
                        Call Ceros(Auxi1, 14)
                        WImpo6 = "0" + Auxi1
                    End If
                    
                    Rem Iva 10.5
                    WImpo7 = "0003"
                    
                    Rem impo iva
                    Auxi1 = "0"
                    Call Ceros(Auxi1, 15)
                    WImpo8 = Auxi1
                                
                    WImpre = WImpo1 + Wimpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8
                    Print #1, WImpre
                End If
            
            End If
            
        End If
        
    Next Ciclo
    
    Close #1
    
    
    
    
    m$ = "El proceso a finalizado"
    a% = MsgBox(m$, 0, "Generacion de Citi Ventas")
    
    
    Call Cancela_click
    
    
    
    

    Exit Sub
    
WError:
    Resume Next
    
    
    
    
End Sub

Private Sub Cancela_click()
    PrgCitinuevoCompras.Hide
    Unload Me
    MenuAdminis.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Drive_Change()
    Dir1.Path = Drive.Drive
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub Nombre_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Frame2.Visible = True
End Sub

Private Sub Eval()

    Es = WCuit

    x = ""
    MinusOk = 1                'a minus sign is okay only once, and only
                                'if it preceeds the first numeric character
    DecOk = 1                  'only the first decimal point is okay

    For XX = 1 To Len(Es)

        Y = Mid$(Es, XX, 1)

        If Y = "-" And MinusOk = 1 Then
               x = x + Y: MinusOk = 0

        ElseIf Y = "." And DecOk = 1 Then
               x = x + Y: DecOk = 0

        ElseIf Y >= "0" And Y <= "9" Then
               x = x + Y: MinusOk = 0

        End If

    Next

    WCuit = x

End Sub


