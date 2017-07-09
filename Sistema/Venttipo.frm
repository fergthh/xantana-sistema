VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgVenttipo 
   Caption         =   "Listado Ventas por Forma de Pago"
   ClientHeight    =   4830
   ClientLeft      =   2130
   ClientTop       =   1965
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4830
   ScaleWidth      =   7350
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta Concepto"
      Height          =   615
      Left            =   6120
      TabIndex        =   14
      Top             =   240
      Width           =   1095
   End
   Begin VB.ListBox Pantalla 
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   3960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2655
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      Begin VB.TextBox Concepto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   255
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   327681
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2880
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.Label DesConcepto 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Concepto"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5880
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "venttipo.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cupones de Tarjeta de Credito en Cartera"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgVenttipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WInicial() As Variant ' Matriz de 2 dimensiones que contiene registros

Private Sub Acepta_Click()

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    Whasta = WAno + WMes + WDia

    With rstEmpresa
        .Index = "Empresa"
        Rem .Seek "=", Val(WEmpresa)
        .Seek "=", 1
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            .Update
        End If
    End With

    da = 0

    With rstVenta
        .Index = "Impre"
        .Seek ">=", da
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
            
    With rstCtaCte
            .Index = "Clave"
            .MoveFirst
            Do
            
                If !OrdFecha >= WDesde And !OrdFecha <= Whasta Then
                If !Codigo1 = Val(Concepto.Text) Then
            
                    WConcepto = !Codigo1
                    WFecha = !Fecha
                    WFechaOrd = !OrdFecha
                    WNumero = !Numero
                    WImporte = !Impo1
                
                    With rstVenta
                        .AddNew
                        !Concepto = WConcepto
                        !Fecha = WFecha
                        !FechaOrd = WFechaOrd
                        !Numero = WNumero
                        !Importe = WImporte
                        .Update
                    End With
                    
                End If
                
                If !Codigo2 = Val(Concepto.Text) Then
            
                    WConcepto = !Codigo2
                    WFecha = !Fecha
                    WFechaOrd = !OrdFecha
                    WNumero = !Numero
                    WImporte = !Impo2
                
                    With rstVenta
                        .AddNew
                        !Concepto = WConcepto
                        !Fecha = WFecha
                        !FechaOrd = WFechaOrd
                        !Numero = WNumero
                        !Importe = WImporte
                        .Update
                    End With
                    
                End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With

    Rem Listado.GroupSelectionFormula = "{Movban.banco} in " + DesdeBanco + " to " + HastaBanco
    Rem Listado.GroupSelectionFormula = "{Movban.banco} in 0 to 9999"
    
    Rem Listado.DataFiles(0) = WEmpresa + "admi.mdb"
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Listado.Action = 1
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    With rstVenta
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    With rstTarjeta
        .Close
    End With
    
    DbsAdminis.Close
    
    PrgVenttipo.Hide
    Unload Me
    menu.SetFocus
End Sub



Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Venta
    OPEN_FILE_Ctacte
    OPEN_FILE_Tarjeta
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

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Concepto.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub


Private Sub Concepto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstTarjeta
            .Index = "Tarjeta"
            Claveven$ = Concepto.Text
            .Seek "=", Val(Concepto.Text)
            If .NoMatch = False Then
                Concepto.Text = !Tarjeta
                DesConcepto.Caption = !Nombre
                Desde.SetFocus
            End If
        End With
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Concepto.Text = ""
    Panta.Value = False
    Impresora.Value = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    With rstTarjeta
        .Index = "Tarjeta"
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = Str$(!Tarjeta) + " " + !Nombre
                Pantalla.AddItem IngresaItem
                IngresaItem = !Tarjeta
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    Pantalla.Visible = True
    XIndice = 2

End Sub

Private Sub Pantalla_Click()
    Rem Pantalla.Visible = False
            With rstTarjeta
                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                .Index = "Tarjeta"
                .Seek "=", Val(Claveven$)
                If .NoMatch = False Then
                    If Lug = 1 Then
                        Concepto.Text = !Tarjeta
                        DesConcepto.Caption = !Nombre
                        Concepto.SetFocus
                            Else
                        Concepto.Text = !Tarjeta
                        DesConcepto.Caption = !Nombre
                        Concepto.SetFocus
                    End If
                End If
            End With
    
End Sub

