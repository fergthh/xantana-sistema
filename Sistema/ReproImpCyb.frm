VERSION 5.00
Begin VB.Form PrgReproImpCyb 
   AutoRedraw      =   -1  'True
   Caption         =   "Reproceso de Imputaciones de Compras"
   ClientHeight    =   3375
   ClientLeft      =   3330
   ClientTop       =   1530
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3375
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CommandButton Proceso 
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
         Height          =   495
         Left            =   600
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
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
         Height          =   495
         Left            =   2280
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgReproImpCyb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WProveedor As String
Private WTipo As String
Private WPunto As String
Private WNumero As String
Dim WFecha As String
Dim WPeriodo As String
Dim WNeto As Double
Dim WIva21 As Double
Dim WIva5 As Double
Dim WIva27 As Double
Dim WIva105 As Double
Dim WIb As Double
Dim WExento As Double
Dim WImpre As Double
Dim WOrdFecha As String
Dim WContado As Integer
Dim WComputable As Integer
Dim WConcepto As Integer
Dim WObservaciones As String
Dim WTotal As Double

Dim XVector(100, 10) As String
Dim WRenglon As String


Private Sub Proceso_Click()

    Rem On Error GoTo WError
    
    With rstIvacomp
            .Index = "Iva"
            .MoveFirst
            Do
                
                WProveedor = !Proveedor
                Call Ceros(WProveedor, 6)
                WTipo = Str$(!Tipo)
                Call Ceros(WTipo, 2)
                WPunto = !Punto
                Call Ceros(WPunto, 4)
                WNumero = !Numero
                Call Ceros(WNumero, 8)
                WLetra = !Letra
                WRenglon = "01"
                    
                WFecha = !Fecha
                WPeriodo = !Periodo
                WNeto = !Neto
                WIva21 = !Iva21
                WIva5 = !Iva5
                WIva27 = !Iva27
                WIva105 = !Iva105
                WIb = !Ib
                WExento = !Exento
                WOrdFecha = !ordfecha
                WContado = !Contado
                WComputable = IIf(IsNull(!Computable), "0", !Computable)
                WConcepto = !Concepto
                WObservaciones = IIf(IsNull(!Observaciones), "", !Observaciones)
                WTotal = WNeto + WIva21 + WIva5 + WIva27 + WIva105 + WIb + WExento
                    
                Graba = "N"
                
                With rstImpcyb
                    .Index = "Clave"
                    .Seek "=", WProveedor + WTipo + WLetra + WPunto + WNumero + WRenglon
                    If .NoMatch = True Then
                        Graba = "S"
                    End If
                End With
                    
                If Graba = "S" Then
                    Call Graba_Asiento
                End If
                    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Call Cancela_click
    
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Graba_Asiento()

    Erase XVector
    WFila = 0
    
    With RstProveedor
        .Index = "Proveedor"
        .Seek "=", WProveedor
        If .NoMatch = False Then
            WTipoProveedor = !Tipo
        End If
    End With
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WCtaEfectivo = !CtaEfectivo
            WCtaIva21 = !CtaIva21
            WCtaIva5 = !CtaIva5
            WCtaIva27 = !CtaIva27
            WCtaIb = !CtaIb
            WCtaIva105 = !CtaIva105
            WCtaIvaNoCompu = !CtaIvaNoCompu
            If WComputable = 1 Then
                WCtaIva21 = !CtaIvaNoCompu
                WCtaIva27 = !CtaIvaNoCompu
                WCtaIva105 = !CtaIvaNoCompu
                WCtaIvaNoCompu = !CtaIvaNoCompu
            End If
        End If
    End With
                
    If WTotal <> 0 Then
        WFila = WFila + 1
        If WContado = 1 Then
            XVector(WFila, 1) = WCtaEfectivo
                Else
            With rstTipopro
                .Index = "Codigo"
                .Seek "=", WTipoProveedor
                If .NoMatch = False Then
                    XVector(WFila, 1) = !Cuenta
                End If
            End With
        End If
        Select Case Val(WTipo)
            Case 1, 2, 7
                XVector(WFila, 2) = ""
                XVector(WFila, 3) = Str$(WTotal)
            Case Else
                XVector(WFila, 2) = Str$(WTotal)
                XVector(WFila, 3) = ""
        End Select
    End If
                
    If WNeto <> 0 Or WExento <> 0 Then
        WFila = WFila + 1
        With rstConceptos
            .Index = "Concepto"
            .Seek "=", WConcepto
            If .NoMatch = False Then
                XVector(WFila, 1) = !Cuenta
            End If
        End With
        Select Case Val(WTipo)
            Case 1, 2, 7
                XVector(WFila, 2) = Str$(WNeto + WExento)
                XVector(WFila, 3) = ""
            Case Else
                XVector(WFila, 2) = ""
                XVector(WFila, 3) = Str$(WNeto + WExento)
        End Select
    End If
                
    If WIva21 <> 0 Then
        WFila = WFila + 1
        XVector(WFila, 1) = WCtaIva21
        Select Case Val(WTipo)
            Case 1, 2, 7
                XVector(WFila, 2) = Str$(WIva21)
                XVector(WFila, 3) = ""
            Case Else
                XVector(WFila, 2) = ""
                XVector(WFila, 3) = Str$(WIva21)
        End Select
    End If
    
    If WIva5 <> 0 Then
        WFila = WFila + 1
        XVector(WFila, 1) = WCtaIva5
        Select Case Val(WTipo)
            Case 1, 2, 7
                XVector(WFila, 2) = Str$(WIva5)
                XVector(WFila, 3) = ""
            Case Else
                XVector(WFila, 2) = ""
                XVector(WFila, 3) = Str$(WIva5)
        End Select
    End If
    
    If WIva27 <> 0 Then
        WFila = WFila + 1
        XVector(WFila, 1) = WCtaIva27
        Select Case Val(WTipo)
            Case 1, 2, 7
                XVector(WFila, 2) = Str$(WIva27)
                XVector(WFila, 3) = ""
            Case Else
                XVector(WFila, 2) = ""
                XVector(WFila, 3) = Str$(WIva27)
        End Select
    End If

    If WIb <> 0 Then
        WFila = WFila + 1
        XVector(WFila, 1) = WCtaIb
        Select Case Val(WTipo)
            Case 1, 2, 7
                XVector(WFila, 2) = Str$(WIb)
                XVector(WFila, 3) = ""
            Case Else
                XVector(WFila, 2) = ""
                XVector(WFila, 3) = Str$(WIb)
        End Select
    End If
    
    If WIva105 <> 0 Then
        WFila = WFila + 1
        XVector(WFila, 1) = WCtaIva105
        Select Case Val(WTipo)
            Case 1, 2, 7
                XVector(WFila, 2) = Str$(WIva105)
                XVector(WFila, 3) = ""
            Case Else
                XVector(WFila, 2) = ""
                XVector(WFila, 3) = Str$(WIva105)
        End Select
    End If

    For Ciclo = 1 To 100
        
        If Val(XVector(Ciclo, 2)) <> 0 Or Val(XVector(Ciclo, 3)) <> 0 Then
        
            WRenglon = Str$(Ciclo)
            Call Ceros(WRenglon, 2)
        
            With rstImpcyb
                .Index = "Clave"
                .Seek "=", WProveedor + WTipo + WLetra + WPunto + WNumero + WRenglon
                If .NoMatch Then
                    .AddNew
                    !Clave = WProveedor + WTipo + WLetra + WPunto + WNumero + WRenglon
                    !Proveedor = Val(WProveedor)
                    !Tipo = Val(WTipo)
                    !Letra = WLetra
                    !Punto = Val(WPunto)
                    !Numero = Val(WNumero)
                    !Renglon = Val(WRenglon)
                    !Cuenta = XVector(Ciclo, 1)
                    !Debito = Val(XVector(Ciclo, 2))
                    !Credito = Val(XVector(Ciclo, 3))
                    !Fecha = WFecha
                    !ordfecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    !Observaciones = WObservaciones
                    .Update
                End If
            End With
        End If
            
    Next Ciclo

End Sub

Private Sub Cancela_click()
    With rstIvacomp
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstImpcyb
        .Close
    End With
    With RstProveedor
        .Close
    End With
    With rstTipopro
        .Close
    End With
    With rstConceptos
        .Close
    End With
    DbsAdminis.Close
    PrgReproImpCyb.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Ivacomp
    OPEN_FILE_Empresa
    OPEN_FILE_Impcyb
    OPEN_FILE_Proveedor
    OPEN_FILE_TipoPro
    OPEN_FILE_Conceptos
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
    If KeyAscii = 27 Then
        Desde.Text = "  /  /    "
    End If
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
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Sub Form_Load()
    Frame2.Visible = True
End Sub

