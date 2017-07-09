VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCitinuevoVentas 
   AutoRedraw      =   -1  'True
   Caption         =   "Exportacion de Datos al CITI "
   ClientHeight    =   4590
   ClientLeft      =   3345
   ClientTop       =   2250
   ClientWidth     =   5475
   LinkTopic       =   "Form2"
   ScaleHeight     =   4590
   ScaleWidth      =   5475
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1560
         TabIndex        =   9
         Top             =   1800
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
         Top             =   1320
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
         Top             =   1320
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
Attribute VB_Name = "PrgCitinuevoVentas"
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
    
    
    
    XNombre = WDir + "\" + "REGINFO_CV_VENTAS_CBTE" + ".txt"
    Open XNombre For Output As #1
    
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCte"
    ZSql = ZSql + " Where CtaCte.OrdFecha >= " + "'" + WDesde + "'"
    ZSql = ZSql + " and CtaCte.OrdFecha <= " + "'" + WHasta + "'"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCte.RecordCount > 0 Then
    
        With rstCtaCte
            .MoveFirst
            Do
                If .EOF = False Then
    
                    If Val(!Tipo) = 1 Or Val(!Tipo) = 2 Or Val(!Tipo) = 3 Or Val(!Tipo) = 4 Or Val(!Tipo) = 5 Then
                    
                        ZZTotal = !Neto + !Iva1 + !Iva2
                        If ZZTotal <> 0 Then
                        
                        
                            Renglon = Renglon + 1
                            
                            XVector(Renglon, 1) = !Letra
                            XVector(Renglon, 2) = !Tipo
                            XVector(Renglon, 3) = !Punto
                            XVector(Renglon, 4) = !Numero
                            XVector(Renglon, 5) = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                            XVector(Renglon, 6) = !Cliente
                            XVector(Renglon, 7) = Str$(!Neto)
                            XVector(Renglon, 8) = Str$(!Iva1)
                            XVector(Renglon, 9) = Str$(!Iva2)
                            XVector(Renglon, 10) = "0"
                            XVector(Renglon, 11) = "0"
                            XVector(Renglon, 12) = "0"
                            Select Case Val(Mid$(!Fecha, 4, 2))
                                Case 1
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "02" + "01"
                                Case 2
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "03" + "01"
                                Case 3
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "04" + "01"
                                Case 4
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "05" + "01"
                                Case 5
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "06" + "01"
                                Case 6
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "07" + "01"
                                Case 7
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "08" + "01"
                                Case 8
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "09" + "01"
                                Case 9
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "10" + "01"
                                Case 10
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "11" + "01"
                                Case 11
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "12" + "01"
                                Case 12
                                    XVector(Renglon, 13) = Right$(!Fecha, 4) + "12" + "31"
                                Case Else
                            End Select
                            
                            XVector(Renglon, 14) = !Numero
    
                        End If
    
                    End If
                
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        
        rstCtaCte.Close
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    For Ciclo = 1 To Renglon
    
        WLetra = XVector(Ciclo, 1)
        WTipo = XVector(Ciclo, 2)
        WPunto = XVector(Ciclo, 3)
        WNumero = XVector(Ciclo, 4)
        WFecha = XVector(Ciclo, 5)
        WCliente = XVector(Ciclo, 6)
        WNumeroII = XVector(Ciclo, 14)
        
        
        
        Rem If Val(WNumero) = 475 Then Stop
        
        
        ZZNeto = Val(XVector(Ciclo, 7))
        ZZNeto = Int(ZZNeto * 100)
        WNeto = ZZNeto
        
        ZZIva1 = Val(XVector(Ciclo, 8))
        ZZIva1 = Int(ZZIva1 * 100)
        WIva1 = ZZIva1
        
        ZZIva2 = Val(XVector(Ciclo, 9))
        ZZIva2 = Int(ZZIva2 * 100)
        WIva2 = ZZIva2
        
        ZZIbTucu = Val(XVector(Ciclo, 10))
        ZZIbTucu = Int(ZZIbTucu * 100)
        WIbTucu = ZZIbTucu
        
        ZZIbCiudad = Val(XVector(Ciclo, 11))
        ZZIbCiudad = Int(ZZIbCiudado * 100)
        WIbCiudad = ZZIbCiudad
        
        ZZZZIb = Val(XVector(Ciclo, 12))
        ZZZZIb = Int(ZZZZIb * 100)
        WIb = ZZZZIb
        
        WExento = 0
        
        WVencimiento = XVector(Ciclo, 13)
        
        WTotal = WNeto + WIva1 + WIva2 + WIbTucu + WIbCiudad + WIb
        WIva = WIva1 + WIva2
        
        Rem If WIva = 0 Then
        Rem     WExento = WNeto
        Rem rem End If
        
        
        If WIva = 0 Then
            WCodigoExento = "N"
            Rem z   zona de exportacion
            Rem x   exportaciones al enterior
            Rem e   operaciones exentas
            Rem C   Operaciones de canje
                Else
            WCodigoExento = " "
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WNombre = rstCliente!Razon
            WNombre = WNombre + Space$(30)
            WNombre = Left$(WNombre, 30)
            WCuit = rstCliente!Cuit
            rstCliente.Close
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
                    Case 1, 3
                        Wimpo2 = "001"
                    Case 4
                        Wimpo2 = "002"
                    Case 2, 5
                        Wimpo2 = "003"
                    Case Else
                        Wimpo2 = "000"
                End Select
            Case "B"
                Select Case Val(WTipo)
                    Case 1, 3
                        Wimpo2 = "006"
                    Case 4
                        Wimpo2 = "007"
                    Case 2, 5
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
        
        Rem punto
        WImpo3 = WPunto
        
        Rem Numero desde
        WImpo4 = WNumero
        
        Rem Numero hasta
        WImpo5 = WNumero
        
        Rem tipo de doc
        If Len(Trim(WCuit)) = 11 Then
            WImpo6 = "80"
                Else
            WImpo6 = "96"
        End If
        
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
        Rem ZZSumaResto = Str$(WIb + WIbTucu + WIbCiudad)
        ZZSumaResto = 0
        If ZZSumaResto >= 0 Then
            Auxi1 = Str$(ZZSumaResto)
            Call Ceros(Auxi1, 15)
            WImpo10 = Auxi1
                Else
            Auxi1 = Str$(Abs(ZZSumaResto))
            Call Ceros(Auxi1, 14)
            WImpo10 = "0" + Auxi1
        End If
        
        Rem percepcion a jo categorizados
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo11 = Auxi1
        
        Rem importes exentos
        If WExento >= 0 Then
            Auxi1 = Str$(WExento)
            Call Ceros(Auxi1, 15)
            WImpo12 = Auxi1
                Else
            Auxi1 = Str$(Abs(WExento))
            Call Ceros(Auxi1, 14)
            WImpo12 = "0" + Auxi1
        End If
        
        Rem percepsion p pago a cuenta de impuestos nacionales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo13 = Auxi1
        
        Rem percepciones  i.b.
        WTotalIb = WIbTucu + WIbCiudad + WIb
        If WTotalIb >= 0 Then
            Auxi1 = Str$(WTotalIb)
            Call Ceros(Auxi1, 15)
            WImpo14 = Auxi1
                Else
            Auxi1 = Str$(Abs(WTotalIb))
            Call Ceros(Auxi1, 14)
            WImpo14 = "0" + Auxi1
        End If
        
        Rem percepsion de impuestos municipales
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo15 = Auxi1
        
        Rem percepsion de impuestos internmos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo16 = Auxi1
        
        Rem codigo de moneda
        WImpo17 = "PES"
        
        Rem PARIDAD
        ZCAmbio = "1"
        Auxi1 = Str$(Int(ZCAmbio * 1000000))
        Call Ceros(Auxi1, 10)
        WImpo18 = Auxi1
        
        Rem cantidad de alicuotas
        WImpo19 = "1"
        
        Rem codigo de operacion
        WImpo20 = WCodigoExento
        
        Rem otros tributos
        Auxi1 = "0"
        Call Ceros(Auxi1, 15)
        WImpo21 = Auxi1
        
        Rem fecha
        WImpo22 = WVencimiento
        If Val(Wimpo2) = 19 Then
            WImpo22 = "00000000"
        End If
        
        WImpre = WImpo1 + Wimpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6 + WImpo7 + WImpo8 + WImpo9 + WImpo10 + WImpo11 + WImpo12 + WImpo13 + WImpo14 + WImpo15 + WImpo16 + WImpo17 + WImpo18 + WImpo19 + WImpo20 + WImpo21 + WImpo22
    
        Print #1, WImpre
        
    Next Ciclo
    
    Close #1
    
    
    
    
    
    XNombre = WDir + "\" + "REGINFO_CV_VENTAS_ALICUOTAS" + ".txt"
    Open XNombre For Output As #1
    
    For Ciclo = 1 To Renglon
    
        WLetra = XVector(Ciclo, 1)
        WTipo = XVector(Ciclo, 2)
        WPunto = XVector(Ciclo, 3)
        WNumero = XVector(Ciclo, 4)
        WFecha = XVector(Ciclo, 5)
        WCliente = XVector(Ciclo, 6)
        WNumeroII = XVector(Ciclo, 14)
        
        ZZNeto = Val(XVector(Ciclo, 7))
        ZZNeto = Int(ZZNeto * 100)
        WNeto = ZZNeto
        
        ZZIva1 = Val(XVector(Ciclo, 8))
        ZZIva1 = Int(ZZIva1 * 100)
        WIva1 = ZZIva1
        
        ZZIva2 = Val(XVector(Ciclo, 9))
        ZZIva2 = Int(ZZIva2 * 100)
        WIva2 = ZZIva2
        
        ZZIbTucu = Val(XVector(Ciclo, 10))
        ZZIbTucu = Int(ZZIbTucu * 100)
        WIbTucu = ZZIbTucu
        
        ZZIbCiudad = Val(XVector(Ciclo, 11))
        ZZIbCiudad = Int(ZZIbCiudado * 100)
        WIbCiudad = ZZIbCiudad
        
        ZZZZIb = Val(XVector(Ciclo, 12))
        ZZZZIb = Int(ZZIb * 100)
        WIb = ZZZZIb
        
        WTotal = WNeto + WIva1 + WIva2 + WIbTucu + WIbCiudad + WIb
        WIva = WIva1 + WIva2
        
        
        
        Rem If Val(WNumero) = 5756 Then Stop
        
        
        Rem If WIva = 0 Then
        Rem     WExento = WNeto
        Rem End If
        If WIva = 0 Then
            WCodigoExento = "N"
            Rem z   zona de exportacion
            Rem x   exportaciones al enterior
            Rem e   operaciones exentas
            Rem C   Operaciones de canje
                Else
            WCodigoExento = " "
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cliente"
        ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = ZSql
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WNombre = rstCliente!Razon
            WNombre = WNombre + Space$(30)
            WNombre = Left$(WNombre, 30)
            WCuit = rstCliente!Cuit
            rstCliente.Close
            Call Eval
        End If
        
        Call Ceros(WTipo, 2)
        Call Ceros(WPunto, 5)
        Call Ceros(WNumero, 20)
        Call Ceros(WCuit, 20)
            
            
            
        Rem tipo de comprobante
        Select Case WLetra
            Case "A"
                Select Case Val(WTipo)
                    Case 1, 3
                        WImpo1 = "001"
                    Case 4
                        WImpo1 = "002"
                    Case 2, 5
                        WImpo1 = "003"
                    Case Else
                        WImpo1 = "000"
                End Select
            Case "B"
                Select Case Val(WTipo)
                    Case 1, 3
                        WImpo1 = "006"
                    Case 4
                        WImpo1 = "007"
                    Case 2, 5
                        WImpo1 = "008"
                    Case Else
                        WImpo1 = "000"
                End Select
            Case "C"
                Select Case Val(WTipo)
                    Case 1
                        WImpo1 = "011"
                    Case 2
                        WImpo1 = "012"
                    Case 3
                        WImpo1 = "013"
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
                Wimpo2 = "000"
        End Select
        
        Rem punto
        Wimpo2 = WPunto
        
        Rem Numero
        WImpo3 = WNumero
        
        
        Rem neto
        If WNeto >= 0 Then
            Auxi1 = Str$(WNeto)
            Call Ceros(Auxi1, 15)
            WImpo4 = Auxi1
                Else
            Auxi1 = Str$(Abs(WNeto))
            Call Ceros(Auxi1, 14)
            WImpo4 = "0" + Auxi1
        End If
        
        If WIva <> 0 Then
            Rem Iva 21
            WImpo5 = "0005"
                Else
            Rem Iva 9
            WImpo5 = "0003"
        End If
        
        Rem impo iva
        If WIva >= 0 Then
            Auxi1 = Str$(WIva)
            Call Ceros(Auxi1, 15)
            WImpo6 = Auxi1
                Else
            Auxi1 = Str$(Abs(WIva))
            Call Ceros(Auxi1, 14)
            WImpo6 = "0" + Auxi1
        End If
        
        WImpre = WImpo1 + Wimpo2 + WImpo3 + WImpo4 + WImpo5 + WImpo6
        Print #1, WImpre
            
            
        
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
    Desde.SetFocus
    PrgCitinuevoVentas.Hide
    Unload Me
    MenuVen.Show
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
               x = x + Y: MinusOk = 0: DecOk = 0

        End If

    Next

    WCuit = x

End Sub


