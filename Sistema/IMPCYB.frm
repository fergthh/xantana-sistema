VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImpcyb 
   Caption         =   "Listado de Imputaciones de Contables"
   ClientHeight    =   8115
   ClientLeft      =   2760
   ClientTop       =   480
   ClientWidth     =   6210
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   6210
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   4815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin VB.CommandButton Panta 
         Caption         =   "Pantalla F1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         MouseIcon       =   "IMPCYB.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "IMPCYB.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Impresion por Pantalla"
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta F4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         MouseIcon       =   "IMPCYB.frx":0B4C
         MousePointer    =   99  'Custom
         Picture         =   "IMPCYB.frx":0E56
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Consulta de Datos"
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Impre 
         Caption         =   "Listado F9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         MouseIcon       =   "IMPCYB.frx":1698
         MousePointer    =   99  'Custom
         Picture         =   "IMPCYB.frx":19A2
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Impresion x Impresora"
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Menu F10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4800
         MouseIcon       =   "IMPCYB.frx":21E4
         MousePointer    =   99  'Custom
         Picture         =   "IMPCYB.frx":24EE
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salida"
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox DesdeCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   17
         Text            =   " "
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox HastaCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   16
         Text            =   " "
         Top             =   1560
         Width           =   2295
      End
      Begin VB.ComboBox TipoList 
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
         Left            =   1920
         TabIndex        =   11
         Text            =   " "
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   1335
         Left            =   360
         TabIndex        =   7
         Top             =   2040
         Width           =   3135
         Begin VB.CheckBox Tipo5 
            Caption         =   "Ventas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox Tipo4 
            Caption         =   "Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1560
            TabIndex        =   12
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Tipo3 
            Caption         =   "Recibos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   1560
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox Tipo2 
            Caption         =   "Depositos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox Tipo1 
            Caption         =   "Pagos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label5 
         Caption         =   "Desde Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cuenta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   3720
         Width           =   1335
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1095
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5760
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Imputa.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Movimietos de Bancos"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgImpcyb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WImpoIva As Double
Dim ZImporte As Double

Dim ZZClave As String
Dim ZZRenglon As String
Dim ZZOrden As String
Dim ZZfecha As String
Dim ZZFechaOrd As String
Dim ZZObservaciones As String
Dim ZZTipoReg As String
Dim ZZTipoOrd As String
Dim ZZCuenta As String
Dim ZZProveedor As String
Dim ZZImporte1 As String
Dim ZZTipo1 As String
Dim ZZLetra1 As String
Dim ZZPunto1 As String
Dim ZZNumero1 As String
Dim ZZRetencion As String
Dim ZZRetIva As String
Dim ZZRetOtra As String
Dim ZZRetOtraII As String
Dim ZZTipo2 As String
Dim ZZBanco2 As String
Dim ZZImporte2 As String
Dim ZZNumero2 As String

Dim ZZDeposito  As String
Dim ZZBanco   As String
Dim ZZImporte   As String
Dim ZZTipo   As String
Dim ZZLetra   As String
Dim ZZPunto   As String
Dim ZZNumero   As String

Dim ZZRecibo As String
Dim ZZTipoRec As String
Dim ZZCliente As String
Dim ZZRetGanancias As String
Dim ZZRetSuss As String

Dim WWClave As String
Dim WWTipoMovi As String
Dim WWComprobante As String
Dim WWTipoComp As String
Dim WWLetraComp As String
Dim WWPuntoComp As String
Dim WWNroComp As String
Dim WWRenglon As String
Dim WWFecha As String
Dim WWObservaciones As String
Dim WWCuenta As String
Dim WWDebito As String
Dim WWCredito  As String
Dim WWfechaord  As String
Dim WWTitulo  As String
Dim WWNombre  As String
Dim WWTitulolist  As String
Dim WWImpre  As String

Dim ZVector(10000, 25) As String
Dim ZLugar As Integer

Private Sub Panta_Click()
    Listado.Destination = 0
    Call Proceso_Click
End Sub

Private Sub Impre_Click()
    Listado.Destination = 1
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    On Error GoTo Error_Programa
    

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    WTitulo = Desde.Text + " al " + Hasta.Text
    If Tipo1.Value = 1 Then
        WTitulo = WTitulo + " Pagos"
    End If
    If Tipo2.Value = 1 Then
        WTitulo = WTitulo + " Deposito"
    End If
    If Tipo3.Value = 1 Then
        WTitulo = WTitulo + " Recibos"
    End If
    If Tipo4.Value = 1 Then
        WTitulo = WTitulo + " Compras"
    End If
    If Tipo5.Value = 1 Then
        WTitulo = WTitulo + " Ventas"
    End If
    WTitulo = Left$(WTitulo, 50)
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Nombre = " + "'" + WNombreEmpresa + "',"
    ZSql = ZSql + " Actividad = " + "'" + Left$(WTitulo, 50) + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "DELETE Imputac"
    spImputac = ZSql
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem Procesa los pagos

    If Tipo1.Value = 1 Then
    
        ZLugar = 0
        Erase ZVector
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Pagos"
        ZSql = ZSql + " Where Pagos.FechaOrd >= '" + WDesde + "'"
        ZSql = ZSql + " and Pagos.FechaOrd <= '" + WHasta + "'"
        ZSql = ZSql + " Order by Clave"
        spPagos = ZSql
        Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
        If rstPagos.RecordCount > 0 Then
            With rstPagos
                .MoveFirst
                Do
                
                    ZZClave = rstPagos!Clave
                    ZZRenglon = rstPagos!Renglon
                    ZZOrden = rstPagos!Orden
                    ZZfecha = rstPagos!Fecha
                    ZZFechaOrd = rstPagos!fechaord
                    ZZObservaciones = rstPagos!Observaciones2
                    ZZTipoReg = rstPagos!Tiporeg
                    ZZTipoOrd = rstPagos!TipoOrd
                    ZZCuenta = rstPagos!Cuenta
                    ZZProveedor = Str$(rstPagos!Proveedor)
                    ZZImporte1 = Str$(rstPagos!Importe1)
                    ZZTipo1 = rstPagos!Tipo1
                    ZZLetra1 = rstPagos!Letra1
                    ZZPunto1 = rstPagos!Punto1
                    ZZNumero1 = rstPagos!Numero1
                    ZZRetencion = Str$(rstPagos!Retencion)
                    ZZTipo2 = rstPagos!Tipo2
                    ZZBanco2 = Str$(rstPagos!Banco2)
                    ZZImporte2 = Str$(rstPagos!Importe2)
                    ZZNumero2 = rstPagos!Numero2
                    ZZRetIva = Str$(rstPagos!RetIva)
                    ZZRetOtra = Str$(rstPagos!RetOtra)
                    
                    ZLugar = ZLugar + 1
                    
                    ZVector(ZLugar, 1) = ZZClave
                    ZVector(ZLugar, 2) = ZZRenglon
                    ZVector(ZLugar, 3) = ZZOrden
                    ZVector(ZLugar, 4) = ZZfecha
                    ZVector(ZLugar, 5) = ZZFechaOrd
                    ZVector(ZLugar, 6) = ZZObservaciones
                    ZVector(ZLugar, 7) = ZZTipoReg
                    ZVector(ZLugar, 8) = ZZTipoOrd
                    ZVector(ZLugar, 9) = ZZCuenta
                    ZVector(ZLugar, 10) = ZZProveedor
                    ZVector(ZLugar, 11) = ZZImporte1
                    ZVector(ZLugar, 12) = ZZTipo1
                    ZVector(ZLugar, 13) = ZZLetra1
                    ZVector(ZLugar, 14) = ZZPunto1
                    ZVector(ZLugar, 15) = ZZNumero1
                    ZVector(ZLugar, 16) = ZZRetencion
                    ZVector(ZLugar, 17) = ZZTipo2
                    ZVector(ZLugar, 18) = ZZBanco2
                    ZVector(ZLugar, 19) = ZZImporte2
                    ZVector(ZLugar, 20) = ZZNumero2
                    ZVector(ZLugar, 21) = ZZRetIva
                    ZVector(ZLugar, 22) = ZZRetOtra
                    ZVector(ZLugar, 23) = ZZChofer
                
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            rstPagos.Close
        End If
        
        Corte = ""
        
        For Ciclo = 1 To ZLugar
        
            ZZClave = ZVector(Ciclo, 1)
            ZZRenglon = ZVector(Ciclo, 2)
            ZZOrden = ZVector(Ciclo, 3)
            ZZfecha = ZVector(Ciclo, 4)
            ZZFechaOrd = ZVector(Ciclo, 5)
            ZZObservaciones = ZVector(Ciclo, 6)
            ZZTipoReg = ZVector(Ciclo, 7)
            ZZTipoOrd = ZVector(Ciclo, 8)
            ZZCuenta = ZVector(Ciclo, 9)
            ZZProveedor = ZVector(Ciclo, 10)
            ZZImporte1 = ZVector(Ciclo, 11)
            ZZTipo1 = ZVector(Ciclo, 12)
            ZZLetra1 = ZVector(Ciclo, 13)
            ZZPunto1 = ZVector(Ciclo, 14)
            ZZNumero1 = ZVector(Ciclo, 15)
            ZZRetencion = ZVector(Ciclo, 16)
            ZZTipo2 = ZVector(Ciclo, 17)
            ZZBanco2 = ZVector(Ciclo, 18)
            ZZImporte2 = ZVector(Ciclo, 19)
            ZZNumero2 = Trim(Str$(Int(Val(ZVector(Ciclo, 20)))))
            ZZRetIva = ZVector(Ciclo, 21)
            ZZRetOtra = ZVector(Ciclo, 22)
            ZZChofer = ZVector(Ciclo, 23)
            
            If Corte <> ZZOrden Then
                Corte = ZZOrden
                Renglon = 0
            End If
                    
            WOrden = ZZOrden
            WFecha = ZZfecha
            WFechaOrd = ZZFechaOrd
            WClave = ZZClave
            WObservaciones = ZZObservaciones
            
            Rem If ZZOrden = 722 Then Stop
                
            Select Case Val(ZZTipoReg)
                Case 1
                    If ZZTipoOrd = "3" Or ZZTipoOrd = "4" Or ZZTipoOrd = "5" Then
                        WObservaciones = ZZObservaciones
                        WCuenta = ZZCuenta
                            Else
                        WCuenta = WCtaProveedores
                        WProveedor = ZZProveedor
                        WObservaciones = ZZObservaciones
                        
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Proveedor"
                        ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                        spProveedor = ZSql
                        Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                        If rstProveedor.RecordCount > 0 Then
                            WTipoProveedor = rstProveedor!Tipo
                            WObservaciones = Trim(rstProveedor!Nombre) + " - " + ZZObservaciones
                            rstProveedor.Close
                        End If
                            
                    End If
                            
                    WDebito = ZZImporte1
                    WCredito = 0
                    WTipo = ZZTipo1
                    WLetra = ZZLetra1
                    WPunto = ZZPunto1
                    WNumero = ZZNumero1
                            
                    Select Case Val(ZZTipo1)
                        Case 1
                            WImpre = "FAC"
                        Case 2
                            WImpre = "N/D"
                        Case 3
                            WImpre = "N/C"
                        Case Else
                            WImpre = ""
                    End Select
                    
                    WWClave = WClave
                    WWTipoMovi = "1"
                    WWComprobante = WOrden
                    WWTipoComp = WTipo
                    WWLetraComp = WLetra
                    WWPuntoComp = WPunto
                    WWNroComp = WNumero
                    WWRenglon = "0"
                    WWFecha = WFecha
                    WWObservaciones = Left$(WObservaciones, 50)
                    WWCuenta = WCuenta
                    WWDebito = WDebito
                    WWCredito = WCredito
                    WWfechaord = WFechaOrd
                    WWTitulo = "Pagos"
                    WWNombre = WNombreEmpresa
                    WWTitulolist = WTitulo
                    WWImpre = WImpre
                    
                    If Val(WWCuenta) = 0 Then
                        WWCuenta = "0"
                    End If
    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Imputac ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "TipoMovi ,"
                    ZSql = ZSql + "Comprobante ,"
                    ZSql = ZSql + "TipoComp ,"
                    ZSql = ZSql + "LetraComp ,"
                    ZSql = ZSql + "PuntoComp ,"
                    ZSql = ZSql + "NroComp ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Debito ,"
                    ZSql = ZSql + "Credito ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "Titulo ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "TituloList ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWTipoMovi + "',"
                    ZSql = ZSql + "'" + WWComprobante + "',"
                    ZSql = ZSql + "'" + WWTipoComp + "',"
                    ZSql = ZSql + "'" + WWLetraComp + "',"
                    ZSql = ZSql + "'" + WWPuntoComp + "',"
                    ZSql = ZSql + "'" + WWNroComp + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWObservaciones + "',"
                    ZSql = ZSql + "'" + WWCuenta + "',"
                    ZSql = ZSql + "'" + WWDebito + "',"
                    ZSql = ZSql + "'" + WWCredito + "',"
                    ZSql = ZSql + "'" + WWfechaord + "',"
                    ZSql = ZSql + "'" + WWTitulo + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWTitulolist + "',"
                    ZSql = ZSql + "'" + WWImpre + "')"
                    spImputac = ZSql
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                    
                            
                    If Val(ZZRenglon) = 1 And Val(ZZRetencion) <> 0 Then
                            
                        WCredito = ZZRetencion
                        WDebito = 0
                        WTipo = 0
                        WLetra = ""
                        WPunto = 0
                        WNumero = 0
                            
                        WWClave = WClave
                        WWTipoMovi = "1"
                        WWComprobante = WOrden
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = WNumero
                        WWRenglon = "01"
                        WWFecha = WFecha
                        WWObservaciones = Left$(WObservaciones, 50)
                        WWCuenta = WCtaRetGan
                        WWDebito = WDebito
                        WWCredito = WCredito
                        WWfechaord = WFechaOrd
                        WWTitulo = "Pagos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = ""
                                                
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                            
                    If Val(ZZRenglon) = 1 And Val(ZZRetIva) <> 0 Then
                            
                        WCredito = ZZRetIva
                        WDebito = 0
                        WTipo = 0
                        WLetra = ""
                        WPunto = 0
                        WNumero = 0
                            
                        WWClave = WClave
                        WWTipoMovi = "1"
                        WWComprobante = WOrden
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = WNumero
                        WWRenglon = "01"
                        WWFecha = WFecha
                        WWObservaciones = Left$(WObservaciones, 50)
                        WWCuenta = "2130216"
                        WWDebito = WDebito
                        WWCredito = WCredito
                        WWfechaord = WFechaOrd
                        WWTitulo = "Pagos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = ""
                                                
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    If Val(ZZRenglon) = 1 And Val(ZZRetOtra) <> 0 Then
                            
                        WCredito = ZZRetOtra
                        WDebito = 0
                        WTipo = 0
                        WLetra = ""
                        WPunto = 0
                        WNumero = 0
                            
                        WWClave = WClave
                        WWTipoMovi = "1"
                        WWComprobante = WOrden
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = WNumero
                        WWRenglon = "01"
                        WWFecha = WFecha
                        WWObservaciones = Left$(WObservaciones, 50)
                        WWCuenta = "2130217"
                        WWDebito = WDebito
                        WWCredito = WCredito
                        WWfechaord = WFechaOrd
                        WWTitulo = "Pagos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = ""
                                                
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                                                            
                Case Else
                    Select Case Val(ZZTipo2)
                        Case 1
                            Rem caja
                            WCuenta = WCtaEfectivo
                            WImpre = "EFTO."
                            
                        Case 2
                            Rem banco
                            WImpre = "BCO"
                            WBanco2 = ZZBanco2
                            ZSql = ""
                            ZSql = ZSql + "Select *"
                            ZSql = ZSql + " FROM Banco"
                            ZSql = ZSql + " Where Banco.Banco = " + "'" + ZZBanco2 + "'"
                            spBanco = ZSql
                            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
                            If rstBanco.RecordCount > 0 Then
                                WCuenta = rstBanco!Cuenta
                                rstBanco.Close
                                    Else
                                WCuenta = "0"
                            End If
                            
                        Case 3
                            Rem che ter
                            WImpre = "CH.TER"
                            WCuenta = WCtaCheques
                            
                        Case Else
                            Rem documentos
                            WImpre = "VARIOS."
                            WCuenta = ZZCuenta
                            
                    End Select
                            
                    WObservaciones = ZZObservaciones
                    WDebito = "0"
                    WCredito = ZZImporte2
                    WProveedor = ZZProveedor
                    WLetra = ""
                    WPunto = "0"
                    WNumero = ZZNumero2
                    WFecha = ZZfecha
                    WFechaOrd = ZZFechaOrd
                            
                    WProveedor = ZZProveedor
                    WObservaciones = ZZObservaciones
                    
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Proveedor"
                    ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + WProveedor + "'"
                    spProveedor = ZSql
                    Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If rstProveedor.RecordCount > 0 Then
                        WObservaciones = Trim(rstProveedor!Nombre) + " - " + WObservaciones
                        rstProveedor.Close
                    End If
                            
                    WWClave = WClave
                    WWTipoMovi = "1"
                    WWComprobante = WOrden
                    WWTipoComp = WTipo
                    WWLetraComp = WLetra
                    WWPuntoComp = WPunto
                    WWNroComp = WNumero
                    WWRenglon = "0"
                    WWFecha = WFecha
                    WWObservaciones = Left$(WObservaciones, 50)
                    WWCuenta = WCuenta
                    WWDebito = WDebito
                    WWCredito = WCredito
                    WWfechaord = WFechaOrd
                    WWTitulo = "Pagos"
                    WWNombre = WNombreEmpresa
                    WWTitulolist = WTitulo
                    WWImpre = WImpre
                    
                    If Val(WWCuenta) = 0 Then
                        WWCuenta = "0"
                    End If
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Imputac ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "TipoMovi ,"
                    ZSql = ZSql + "Comprobante ,"
                    ZSql = ZSql + "TipoComp ,"
                    ZSql = ZSql + "LetraComp ,"
                    ZSql = ZSql + "PuntoComp ,"
                    ZSql = ZSql + "NroComp ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Debito ,"
                    ZSql = ZSql + "Credito ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "Titulo ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "TituloList ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWTipoMovi + "',"
                    ZSql = ZSql + "'" + WWComprobante + "',"
                    ZSql = ZSql + "'" + WWTipoComp + "',"
                    ZSql = ZSql + "'" + WWLetraComp + "',"
                    ZSql = ZSql + "'" + WWPuntoComp + "',"
                    ZSql = ZSql + "'" + WWNroComp + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWObservaciones + "',"
                    ZSql = ZSql + "'" + WWCuenta + "',"
                    ZSql = ZSql + "'" + WWDebito + "',"
                    ZSql = ZSql + "'" + WWCredito + "',"
                    ZSql = ZSql + "'" + WWfechaord + "',"
                    ZSql = ZSql + "'" + WWTitulo + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWTitulolist + "',"
                    ZSql = ZSql + "'" + WWImpre + "')"
                    spImputac = ZSql
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                    
                    If Val(ZZRenglon) = 1 And Val(ZZRetencion) <> 0 Then
                            
                        WCredito = ZZRetencion
                        WDebito = 0
                        WTipo = 0
                        WLetra = ""
                        WPunto = 0
                        WNumero = 0
                            
                        WWClave = WClave
                        WWTipoMovi = "1"
                        WWComprobante = WOrden
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = WNumero
                        WWRenglon = "01"
                        WWFecha = WFecha
                        WWObservaciones = Left$(WObservaciones, 50)
                        WWCuenta = WCtaRetGan
                        WWDebito = WDebito
                        WWCredito = WCredito
                        WWfechaord = WFechaOrd
                        WWTitulo = "Pagos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = ""
                                                
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                            
                    If Val(ZZRenglon) = 1 And Val(ZZRetIva) <> 0 Then
                            
                        WCredito = ZZRetIva
                        WDebito = 0
                        WTipo = 0
                        WLetra = ""
                        WPunto = 0
                        WNumero = 0
                            
                        WWClave = WClave
                        WWTipoMovi = "1"
                        WWComprobante = WOrden
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = WNumero
                        WWRenglon = "01"
                        WWFecha = WFecha
                        WWObservaciones = Left$(WObservaciones, 50)
                        WWCuenta = "2130216"
                        WWDebito = WDebito
                        WWCredito = WCredito
                        WWfechaord = WFechaOrd
                        WWTitulo = "Pagos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = ""
                                                
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    If Val(ZZRenglon) = 1 And Val(ZZRetOtra) <> 0 Then
                            
                        WCredito = ZZRetOtra
                        WDebito = 0
                        WTipo = 0
                        WLetra = ""
                        WPunto = 0
                        WNumero = 0
                            
                        WWClave = WClave
                        WWTipoMovi = "1"
                        WWComprobante = WOrden
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = WNumero
                        WWRenglon = "01"
                        WWFecha = WFecha
                        WWObservaciones = Left$(WObservaciones, 50)
                        WWCuenta = "2130217"
                        WWDebito = WDebito
                        WWCredito = WCredito
                        WWfechaord = WFechaOrd
                        WWTitulo = "Pagos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = ""
                                                
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                    End If
                    
                    
                    
                    
            End Select
            
        Next Ciclo
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa los depositos
    
    If Tipo2.Value = 1 Then
    
        ZLugar = 0
        Erase ZVector
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Depositos"
        ZSql = ZSql + " Where Depositos.FechaOrd >= '" + WDesde + "'"
        ZSql = ZSql + " and Depositos.FechaOrd <= '" + WHasta + "'"
        spDepositos = ZSql
        Set rstDepositos = db.OpenRecordset(spDepositos, dbOpenSnapshot, dbSQLPassThrough)
        If rstDepositos.RecordCount > 0 Then
            With rstDepositos
                .MoveFirst
                Do
                
                    ZZDeposito = rstDepositos!Deposito
                    ZZfecha = rstDepositos!Fecha
                    ZZFechaOrd = rstDepositos!fechaord
                    ZZClave = rstDepositos!Clave
                    ZZBanco = Str(rstDepositos!Banco)
                    ZZImporte = Str$(rstDepositos!Importe2)
                    ZZTipo = rstDepositos!Tipo2
                    ZZLetra = ""
                    ZZPunto = "0"
                    ZZNumero = rstDepositos!Numero2
                        
                    ZLugar = ZLugar + 1
                    
                    ZVector(ZLugar, 1) = ZZDeposito
                    ZVector(ZLugar, 2) = ZZfecha
                    ZVector(ZLugar, 3) = ZZFechaOrd
                    ZVector(ZLugar, 4) = ZZClave
                    ZVector(ZLugar, 5) = ZZBanco
                    ZVector(ZLugar, 6) = ZZImporte
                    ZVector(ZLugar, 7) = ZZTipo
                    ZVector(ZLugar, 8) = ZZLetra
                    ZVector(ZLugar, 9) = ZZPunto
                    ZVector(ZLugar, 10) = ZZNumero
                
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            rstDepositos.Close
        End If
        
        For Ciclo = 1 To ZLugar
        
            ZZDeposito = ZVector(Ciclo, 1)
            ZZfecha = ZVector(Ciclo, 2)
            ZZFechaOrd = ZVector(Ciclo, 3)
            ZZClave = ZVector(Ciclo, 4)
            ZZBanco = ZVector(Ciclo, 5)
            ZZImporte = ZVector(Ciclo, 6)
            ZZTipo = ZVector(Ciclo, 7)
            ZZLetra = ZVector(Ciclo, 8)
            ZZPunto = ZVector(Ciclo, 9)
            ZZNumero = ZVector(Ciclo, 10)
            
        
            WDeposito = ZZDeposito
            WFecha = ZZfecha
            WFechaOrd = ZZFechaOrd
            WClave = ZZClave
            WBanco = ZZBanco
            WImporte = ZZImporte
            WTipo = ZZTipo
            WLetra = ""
            WPunto = "0"
            WNumero = ZZNumero
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Banco = " + "'" + WBanco + "'"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                WCuenta = rstBanco!Cuenta
                rstBanco.Close
                    Else
                WCuenta = ""
            End If
                        
            If Val(WTipo) = 1 Then
                WObservaciones = "Deposito en Efectivo"
                WImpre = ""
                    Else
                WObservaciones = "Deposito en Cheques"
                WImpre = "Cheque"
            End If
                        
            WWClave = WClave
            WWTipoMovi = "2"
            WWComprobante = WDeposito
            WWTipoComp = WTipo
            WWLetraComp = WLetra
            WWPuntoComp = WPunto
            WWNroComp = WNumero
            WWRenglon = "0"
            WWFecha = WFecha
            WWObservaciones = Left$(WObservaciones, 50)
            WWCuenta = WCuenta
            WWDebito = WImporte
            WWCredito = "0"
            WWfechaord = WFechaOrd
            WWTitulo = "Deposito"
            WWNombre = WNombreEmpresa
            WWTitulolist = WTitulo
            WWImpre = WImpre
            If Val(WWCuenta) = 0 Then
                WWCuenta = "0"
            End If
                        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Imputac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "TipoMovi ,"
            ZSql = ZSql + "Comprobante ,"
            ZSql = ZSql + "TipoComp ,"
            ZSql = ZSql + "LetraComp ,"
            ZSql = ZSql + "PuntoComp ,"
            ZSql = ZSql + "NroComp ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Cuenta ,"
            ZSql = ZSql + "Debito ,"
            ZSql = ZSql + "Credito ,"
            ZSql = ZSql + "FechaOrd ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "TituloList ,"
            ZSql = ZSql + "Impre )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WWClave + "',"
            ZSql = ZSql + "'" + WWTipoMovi + "',"
            ZSql = ZSql + "'" + WWComprobante + "',"
            ZSql = ZSql + "'" + WWTipoComp + "',"
            ZSql = ZSql + "'" + WWLetraComp + "',"
            ZSql = ZSql + "'" + WWPuntoComp + "',"
            ZSql = ZSql + "'" + WWNroComp + "',"
            ZSql = ZSql + "'" + WWRenglon + "',"
            ZSql = ZSql + "'" + WWFecha + "',"
            ZSql = ZSql + "'" + WWObservaciones + "',"
            ZSql = ZSql + "'" + WWCuenta + "',"
            ZSql = ZSql + "'" + WWDebito + "',"
            ZSql = ZSql + "'" + WWCredito + "',"
            ZSql = ZSql + "'" + WWfechaord + "',"
            ZSql = ZSql + "'" + WWTitulo + "',"
            ZSql = ZSql + "'" + WWNombre + "',"
            ZSql = ZSql + "'" + WWTitulolist + "',"
            ZSql = ZSql + "'" + WWImpre + "')"
            spImputac = ZSql
            Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
            
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Banco"
            ZSql = ZSql + " Where Banco.Banco = " + "'" + WBanco + "'"
            spBanco = ZSql
            Set rstBanco = db.OpenRecordset(spBanco, dbOpenSnapshot, dbSQLPassThrough)
            If rstBanco.RecordCount > 0 Then
                WObservaciones = rstBanco!Nombre
                rstBanco.Close
                    Else
                WCuenta = ""
            End If
            
            Select Case Val(WTipo)
                Case 1
                    Rem EFECTIVO
                    WCuenta = WCtaEfectivo
                Case Else
                    Rem valores en cartera
                    WCuenta = WCtaCheques
            End Select
                        
            WWClave = WClave
            WWTipoMovi = "2"
            WWComprobante = WDeposito
            WWTipoComp = WTipo
            WWLetraComp = WLetra
            WWPuntoComp = WPunto
            WWNroComp = WNumero
            WWRenglon = "0"
            WWFecha = WFecha
            WWObservaciones = Left$(WObservaciones, 50)
            WWCuenta = WCuenta
            WWDebito = "0"
            WWCredito = WImporte
            WWfechaord = WFechaOrd
            WWTitulo = "Deposito"
            WWNombre = WNombreEmpresa
            WWTitulolist = WTitulo
            WWImpre = WImpre
            If Val(WWCuenta) = 0 Then
                WWCuenta = "0"
            End If
                        
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Imputac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "TipoMovi ,"
            ZSql = ZSql + "Comprobante ,"
            ZSql = ZSql + "TipoComp ,"
            ZSql = ZSql + "LetraComp ,"
            ZSql = ZSql + "PuntoComp ,"
            ZSql = ZSql + "NroComp ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "Observaciones ,"
            ZSql = ZSql + "Cuenta ,"
            ZSql = ZSql + "Debito ,"
            ZSql = ZSql + "Credito ,"
            ZSql = ZSql + "FechaOrd ,"
            ZSql = ZSql + "Titulo ,"
            ZSql = ZSql + "Nombre ,"
            ZSql = ZSql + "TituloList ,"
            ZSql = ZSql + "Impre )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WWClave + "',"
            ZSql = ZSql + "'" + WWTipoMovi + "',"
            ZSql = ZSql + "'" + WWComprobante + "',"
            ZSql = ZSql + "'" + WWTipoComp + "',"
            ZSql = ZSql + "'" + WWLetraComp + "',"
            ZSql = ZSql + "'" + WWPuntoComp + "',"
            ZSql = ZSql + "'" + WWNroComp + "',"
            ZSql = ZSql + "'" + WWRenglon + "',"
            ZSql = ZSql + "'" + WWFecha + "',"
            ZSql = ZSql + "'" + WWObservaciones + "',"
            ZSql = ZSql + "'" + WWCuenta + "',"
            ZSql = ZSql + "'" + WWDebito + "',"
            ZSql = ZSql + "'" + WWCredito + "',"
            ZSql = ZSql + "'" + WWfechaord + "',"
            ZSql = ZSql + "'" + WWTitulo + "',"
            ZSql = ZSql + "'" + WWNombre + "',"
            ZSql = ZSql + "'" + WWTitulolist + "',"
            ZSql = ZSql + "'" + WWImpre + "')"
            spImputac = ZSql
            Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
            
                        
        Next Ciclo
        
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las Cobranzas
    
    If Tipo3.Value = 1 Then
    
        ZLugar = 0
        Erase ZVector
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Recibos"
        ZSql = ZSql + " Where Recibos.FechaOrd >= '" + WDesde + "'"
        ZSql = ZSql + " and Recibos.FechaOrd <= '" + WHasta + "'"
        spRecibos = ZSql
        Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
        If rstRecibos.RecordCount > 0 Then

            With rstRecibos
                .MoveFirst
                Do
                
                    ZZClave = rstRecibos!Clave
                    ZZRecibo = rstRecibos!recibo
                    ZZRenglon = rstRecibos!Renglon
                    ZZfecha = rstRecibos!Fecha
                    ZZFechaOrd = rstRecibos!fechaord
                    ZZTipoRec = rstRecibos!TipoRec
                    ZZObservaciones = rstRecibos!Observaciones
                    ZZCliente = Str$(rstRecibos!Cliente)
                    ZZTipoReg = rstRecibos!Tiporeg
                    ZZCuenta = rstRecibos!Cuenta
                    ZZLetra1 = rstRecibos!Letra1
                    ZZTipo1 = rstRecibos!Tipo1
                    ZZPunto1 = rstRecibos!Punto1
                    ZZNumero1 = rstRecibos!Numero1
                    ZZImporte1 = Str$(rstRecibos!Importe1)
                    ZZTipo2 = rstRecibos!Tipo2
                    ZZNumero2 = rstRecibos!Numero2
                    ZZImporte2 = Str$(rstRecibos!Importe2)
                    ZZRetGanancias = Str$(rstRecibos!Retganancias)
                    ZZRetIva = Str$(rstRecibos!RetIva)
                    ZZRetOtra = Str$(rstRecibos!RetOtra)
                    ZZRetSuss = Str$(rstRecibos!RetSuss)
                    XXRetOtraII = IIf(IsNull(rstRecibos!RetOtraII), "0", rstRecibos!RetOtraII)
                    ZZRetOtraII = Str$(XXRetOtraII)
                        
                    ZLugar = ZLugar + 1
                    
                    ZVector(ZLugar, 1) = ZZClave
                    ZVector(ZLugar, 2) = ZZRecibo
                    ZVector(ZLugar, 3) = ZZRenglon
                    ZVector(ZLugar, 4) = ZZfecha
                    ZVector(ZLugar, 5) = ZZFechaOrd
                    ZVector(ZLugar, 6) = ZZTipoRec
                    ZVector(ZLugar, 7) = ZZObservaciones
                    ZVector(ZLugar, 8) = ZZCliente
                    ZVector(ZLugar, 9) = ZZTipoReg
                    ZVector(ZLugar, 10) = ZZCuenta
                    ZVector(ZLugar, 11) = ZZLetra1
                    ZVector(ZLugar, 12) = ZZTipo1
                    ZVector(ZLugar, 13) = ZZPunto1
                    ZVector(ZLugar, 14) = ZZNumero1
                    ZVector(ZLugar, 15) = ZZImporte1
                    ZVector(ZLugar, 16) = ZZTipo2
                    ZVector(ZLugar, 17) = ZZNumero2
                    ZVector(ZLugar, 18) = ZZImporte2
                    ZVector(ZLugar, 19) = ZZRetGanancias
                    ZVector(ZLugar, 20) = ZZRetIva
                    ZVector(ZLugar, 21) = ZZRetOtra
                    ZVector(ZLugar, 22) = ZZRetSuss
                    ZVector(ZLugar, 23) = ZZRetOtraII
                
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
        
            rstRecibos.Close
        End If
        
        Corte = ""
            
        For Ciclo = 1 To ZLugar
        
                ZZClave = ZVector(Ciclo, 1)
                ZZRecibo = ZVector(Ciclo, 2)
                ZZRenglon = ZVector(Ciclo, 3)
                ZZfecha = ZVector(Ciclo, 4)
                ZZFechaOrd = ZVector(Ciclo, 5)
                ZZTipoRec = ZVector(Ciclo, 6)
                ZZObservaciones = ZVector(Ciclo, 7)
                ZZCliente = ZVector(Ciclo, 8)
                ZZTipoReg = ZVector(Ciclo, 9)
                ZZCuenta = ZVector(Ciclo, 10)
                ZZLetra1 = ZVector(Ciclo, 11)
                ZZTipo1 = ZVector(Ciclo, 12)
                ZZPunto1 = ZVector(Ciclo, 13)
                ZZNumero1 = ZVector(Ciclo, 14)
                ZZImporte1 = ZVector(Ciclo, 15)
                ZZTipo2 = ZVector(Ciclo, 16)
                ZZNumero2 = ZVector(Ciclo, 17)
                ZZImporte2 = ZVector(Ciclo, 18)
                ZZRetGanancias = ZVector(Ciclo, 19)
                ZZRetIva = ZVector(Ciclo, 20)
                ZZRetOtra = ZVector(Ciclo, 21)
                ZZRetSuss = ZVector(Ciclo, 22)
                ZZRetOtraII = ZVector(Ciclo, 23)
                
                If Corte <> ZZRecibo Then
                    Corte = ZZRecibo
                    Renglon = 0
                End If
                    
                WClave = ZZClave
                WRecibo = ZZRecibo
                WRenglon = ZZRenglon
                WFecha = ZZfecha
                WFechaOrd = ZZFechaOrd
                
                If ZZTipoRec = "3" Then
                
                    WObservaciones = ZZObservaciones
                    
                        Else
                        
                    WCliente = ZZCliente
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Cliente"
                    ZSql = ZSql + " Where Cliente.Cliente = " + "'" + WCliente + "'"
                    spCliente = ZSql
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WObservaciones = rstCliente!Razon
                        rstCliente.Close
                    End If
                End If
                            
                Select Case Val(ZZTipoReg)
                    Case 1
                        If ZZTipoRec = "3" Then
                            WCuenta = ZZCuenta
                                Else
                            Rem clientes
                            WCuenta = WCtaDeudores
                        End If
                            
                        WLetra = ZZLetra1
                        WTipo = ZZTipo1
                        WPunto = ZZPunto1
                        WNumero = ZZNumero1
                        WImporte = ZZImporte1
                            
                        Select Case Val(ZZTipo1)
                            Case 3
                                WImpre = "FAC"
                            Case 4
                                WImpre = "N/D"
                            Case 5
                                WImpre = "N/C"
                            Case Else
                                WImpre = ""
                        End Select
                            
                        WWClave = WClave
                        WWTipoMovi = "3"
                        WWComprobante = WRecibo
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = WNumero
                        WWRenglon = WRenglon
                        WWFecha = WFecha
                        WWObservaciones = WObservaciones
                        WWCuenta = WCuenta
                        WWDebito = "0"
                        WWCredito = WImporte
                        WWfechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        WWTitulo = "Recibos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = WImpre
                        If Val(WWCuenta) = 0 Then
                            WWCuenta = "0"
                        End If
                    
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                                                            
                    Case Else
                        Select Case Val(ZZTipo2)
                            Case 1
                                Rem caja
                                WCuenta = WCtaEfectivo
                                WImpre = "EFTO."
                            Case 2
                                Rem cheques
                                WImpre = "CH.TER"
                                WCuenta = WCtaCheques
                            Case 4
                                WImpre = ""
                                WCuenta = ZZCuenta
                            Case Else
                                Rem documentos
                                WImpre = "DOC."
                                WCuenta = WCtaCheques
                        End Select
                                    
                        WLetra = ""
                        WTipo = ZZTipo2
                        WPunto = 0
                        WNumero = Trim(Str$(Int(Val(ZZNumero2))))
                        WImporte = ZZImporte2
                            
                        WWClave = WClave
                        WWTipoMovi = "3"
                        WWComprobante = WRecibo
                        WWTipoComp = WTipo
                        WWLetraComp = WLetra
                        WWPuntoComp = WPunto
                        WWNroComp = Trim(Str$(Int(Val(WNumero))))
                        WWRenglon = WRenglon
                        WWFecha = WFecha
                        WWObservaciones = WObservaciones
                        WWCuenta = WCuenta
                        WWDebito = WImporte
                        WWCredito = "0"
                        WWfechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        WWTitulo = "Recibos"
                        WWNombre = WNombreEmpresa
                        WWTitulolist = WTitulo
                        WWImpre = WImpre
                        If Val(WWCuenta) = 0 Then
                            WWCuenta = "0"
                        End If
                    
                    
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO Imputac ("
                        ZSql = ZSql + "Clave ,"
                        ZSql = ZSql + "TipoMovi ,"
                        ZSql = ZSql + "Comprobante ,"
                        ZSql = ZSql + "TipoComp ,"
                        ZSql = ZSql + "LetraComp ,"
                        ZSql = ZSql + "PuntoComp ,"
                        ZSql = ZSql + "NroComp ,"
                        ZSql = ZSql + "Renglon ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Observaciones ,"
                        ZSql = ZSql + "Cuenta ,"
                        ZSql = ZSql + "Debito ,"
                        ZSql = ZSql + "Credito ,"
                        ZSql = ZSql + "FechaOrd ,"
                        ZSql = ZSql + "Titulo ,"
                        ZSql = ZSql + "Nombre ,"
                        ZSql = ZSql + "TituloList ,"
                        ZSql = ZSql + "Impre )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + WWClave + "',"
                        ZSql = ZSql + "'" + WWTipoMovi + "',"
                        ZSql = ZSql + "'" + WWComprobante + "',"
                        ZSql = ZSql + "'" + WWTipoComp + "',"
                        ZSql = ZSql + "'" + WWLetraComp + "',"
                        ZSql = ZSql + "'" + WWPuntoComp + "',"
                        ZSql = ZSql + "'" + WWNroComp + "',"
                        ZSql = ZSql + "'" + WWRenglon + "',"
                        ZSql = ZSql + "'" + WWFecha + "',"
                        ZSql = ZSql + "'" + WWObservaciones + "',"
                        ZSql = ZSql + "'" + WWCuenta + "',"
                        ZSql = ZSql + "'" + WWDebito + "',"
                        ZSql = ZSql + "'" + WWCredito + "',"
                        ZSql = ZSql + "'" + WWfechaord + "',"
                        ZSql = ZSql + "'" + WWTitulo + "',"
                        ZSql = ZSql + "'" + WWNombre + "',"
                        ZSql = ZSql + "'" + WWTitulolist + "',"
                        ZSql = ZSql + "'" + WWImpre + "')"
                        spImputac = ZSql
                        Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                        
                End Select
                    
                If Val(ZZRenglon) = 1 And Val(ZZRetGanancias) <> 0 Then
                    
                    WLetra = ""
                    WTipo = 0
                    WPunto = 0
                    WNumero = 0
                    WImporte = ZZRetGanancias
                    WCuenta = WCtaRetGan
                                
                    WWClave = WClave
                    WWTipoMovi = "3"
                    WWComprobante = WRecibo
                    WWTipoComp = WTipo
                    WWLetraComp = WLetra
                    WWPuntoComp = WPunto
                    WWNroComp = WNumero
                    WWRenglon = WRenglon
                    WWFecha = WFecha
                    WWObservaciones = WObservaciones
                    WWCuenta = WCuenta
                    WWDebito = WImporte
                    WWCredito = "0"
                    WWfechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    WWTitulo = "Recibos"
                    WWNombre = WNombreEmpresa
                    WWTitulolist = WTitulo
                    WWImpre = ""
                    
                        
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Imputac ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "TipoMovi ,"
                    ZSql = ZSql + "Comprobante ,"
                    ZSql = ZSql + "TipoComp ,"
                    ZSql = ZSql + "LetraComp ,"
                    ZSql = ZSql + "PuntoComp ,"
                    ZSql = ZSql + "NroComp ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Debito ,"
                    ZSql = ZSql + "Credito ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "Titulo ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "TituloList ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWTipoMovi + "',"
                    ZSql = ZSql + "'" + WWComprobante + "',"
                    ZSql = ZSql + "'" + WWTipoComp + "',"
                    ZSql = ZSql + "'" + WWLetraComp + "',"
                    ZSql = ZSql + "'" + WWPuntoComp + "',"
                    ZSql = ZSql + "'" + WWNroComp + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWObservaciones + "',"
                    ZSql = ZSql + "'" + WWCuenta + "',"
                    ZSql = ZSql + "'" + WWDebito + "',"
                    ZSql = ZSql + "'" + WWCredito + "',"
                    ZSql = ZSql + "'" + WWfechaord + "',"
                    ZSql = ZSql + "'" + WWTitulo + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWTitulolist + "',"
                    ZSql = ZSql + "'" + WWImpre + "')"
                    spImputac = ZSql
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                If Val(ZZRenglon) = 1 And Val(ZZRetIva) <> 0 Then
                
                    WLetra = ""
                    WTipo = 0
                    WPunto = 0
                    WNumero = 0
                    WImporte = ZZRetIva
                    WCuenta = WCtaRetIva
                                
                    WWClave = WClave
                    WWTipoMovi = "3"
                    WWComprobante = WRecibo
                    WWTipoComp = WTipo
                    WWLetraComp = WLetra
                    WWPuntoComp = WPunto
                    WWNroComp = WNumero
                    WWRenglon = WRenglon
                    WWFecha = WFecha
                    WWObservaciones = WObservaciones
                    WWCuenta = WCuenta
                    WWDebito = WImporte
                    WWCredito = "0"
                    WWfechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    WWTitulo = "Recibos"
                    WWNombre = WNombreEmpresa
                    WWTitulolist = WTitulo
                        
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Imputac ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "TipoMovi ,"
                    ZSql = ZSql + "Comprobante ,"
                    ZSql = ZSql + "TipoComp ,"
                    ZSql = ZSql + "LetraComp ,"
                    ZSql = ZSql + "PuntoComp ,"
                    ZSql = ZSql + "NroComp ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Debito ,"
                    ZSql = ZSql + "Credito ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "Titulo ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "TituloList ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWTipoMovi + "',"
                    ZSql = ZSql + "'" + WWComprobante + "',"
                    ZSql = ZSql + "'" + WWTipoComp + "',"
                    ZSql = ZSql + "'" + WWLetraComp + "',"
                    ZSql = ZSql + "'" + WWPuntoComp + "',"
                    ZSql = ZSql + "'" + WWNroComp + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWObservaciones + "',"
                    ZSql = ZSql + "'" + WWCuenta + "',"
                    ZSql = ZSql + "'" + WWDebito + "',"
                    ZSql = ZSql + "'" + WWCredito + "',"
                    ZSql = ZSql + "'" + WWfechaord + "',"
                    ZSql = ZSql + "'" + WWTitulo + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWTitulolist + "',"
                    ZSql = ZSql + "'" + WWImpre + "')"
                    spImputac = ZSql
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                If Val(ZZRenglon) = 1 And Val(ZZRetOtra) <> 0 Then
                    
                    WLetra = ""
                    WTipo = 0
                    WPunto = 0
                    WNumero = 0
                    WImporte = ZZRetOtra
                    WCuenta = WCtaretOtra
                                
                    WWTipoMovi = "3"
                    WWClave = WClave
                    WWComprobante = WRecibo
                    WWTipoComp = WTipo
                    WWLetraComp = WLetra
                    WWPuntoComp = WPunto
                    WWNroComp = WNumero
                    WWRenglon = WRenglon
                    WWFecha = WFecha
                    WWObservaciones = WObservaciones
                    WWCuenta = WCuenta
                    WWDebito = WImporte
                    WWCredito = "0"
                    WWfechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    WWTitulo = "Recibos"
                    WWNombre = WNombreEmpresa
                    WWTitulolist = WTitulo
                    WWImpre = ""
                        
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Imputac ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "TipoMovi ,"
                    ZSql = ZSql + "Comprobante ,"
                    ZSql = ZSql + "TipoComp ,"
                    ZSql = ZSql + "LetraComp ,"
                    ZSql = ZSql + "PuntoComp ,"
                    ZSql = ZSql + "NroComp ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Debito ,"
                    ZSql = ZSql + "Credito ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "Titulo ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "TituloList ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWTipoMovi + "',"
                    ZSql = ZSql + "'" + WWComprobante + "',"
                    ZSql = ZSql + "'" + WWTipoComp + "',"
                    ZSql = ZSql + "'" + WWLetraComp + "',"
                    ZSql = ZSql + "'" + WWPuntoComp + "',"
                    ZSql = ZSql + "'" + WWNroComp + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWObservaciones + "',"
                    ZSql = ZSql + "'" + WWCuenta + "',"
                    ZSql = ZSql + "'" + WWDebito + "',"
                    ZSql = ZSql + "'" + WWCredito + "',"
                    ZSql = ZSql + "'" + WWfechaord + "',"
                    ZSql = ZSql + "'" + WWTitulo + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWTitulolist + "',"
                    ZSql = ZSql + "'" + WWImpre + "')"
                    spImputac = ZSql
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                If Val(ZZRenglon) = 1 And Val(ZZRetOtraII) <> 0 Then
                    
                    WLetra = ""
                    WTipo = 0
                    WPunto = 0
                    WNumero = 0
                    WImporte = ZZRetOtraII
                    WCuenta = "1145000"
                                
                    WWTipoMovi = "3"
                    WWClave = WClave
                    WWComprobante = WRecibo
                    WWTipoComp = WTipo
                    WWLetraComp = WLetra
                    WWPuntoComp = WPunto
                    WWNroComp = WNumero
                    WWRenglon = WRenglon
                    WWFecha = WFecha
                    WWObservaciones = WObservaciones
                    WWCuenta = WCuenta
                    WWDebito = WImporte
                    WWCredito = "0"
                    WWfechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    WWTitulo = "Recibos"
                    WWNombre = WNombreEmpresa
                    WWTitulolist = WTitulo
                    WWImpre = ""
                        
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Imputac ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "TipoMovi ,"
                    ZSql = ZSql + "Comprobante ,"
                    ZSql = ZSql + "TipoComp ,"
                    ZSql = ZSql + "LetraComp ,"
                    ZSql = ZSql + "PuntoComp ,"
                    ZSql = ZSql + "NroComp ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Debito ,"
                    ZSql = ZSql + "Credito ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "Titulo ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "TituloList ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWTipoMovi + "',"
                    ZSql = ZSql + "'" + WWComprobante + "',"
                    ZSql = ZSql + "'" + WWTipoComp + "',"
                    ZSql = ZSql + "'" + WWLetraComp + "',"
                    ZSql = ZSql + "'" + WWPuntoComp + "',"
                    ZSql = ZSql + "'" + WWNroComp + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWObservaciones + "',"
                    ZSql = ZSql + "'" + WWCuenta + "',"
                    ZSql = ZSql + "'" + WWDebito + "',"
                    ZSql = ZSql + "'" + WWCredito + "',"
                    ZSql = ZSql + "'" + WWfechaord + "',"
                    ZSql = ZSql + "'" + WWTitulo + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWTitulolist + "',"
                    ZSql = ZSql + "'" + WWImpre + "')"
                    spImputac = ZSql
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                If Val(ZZRenglon) = 1 And Val(ZZRetSuss) <> 0 Then
                    
                    WLetra = ""
                    WTipo = 0
                    WPunto = 0
                    WNumero = 0
                    WImporte = ZZRetSuss
                    WCuenta = WCtaRetSuss
                                
                    WWTipoMovi = "3"
                    WWClave = WClave
                    WWComprobante = WRecibo
                    WWTipoComp = WTipo
                    WWLetraComp = WLetra
                    WWPuntoComp = WPunto
                    WWNroComp = WNumero
                    WWRenglon = WRenglon
                    WWFecha = WFecha
                    WWObservaciones = WObservaciones
                    WWCuenta = WCuenta
                    WWDebito = WImporte
                    WWCredito = 0
                    WWfechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    WWTitulo = "Recibos"
                    WWNombre = WNombreEmpresa
                    WWTitulolist = WTitulo
                    WWImpre = ""
                    
                    ZSql = ""
                    ZSql = ZSql + "INSERT INTO Imputac ("
                    ZSql = ZSql + "Clave ,"
                    ZSql = ZSql + "TipoMovi ,"
                    ZSql = ZSql + "Comprobante ,"
                    ZSql = ZSql + "TipoComp ,"
                    ZSql = ZSql + "LetraComp ,"
                    ZSql = ZSql + "PuntoComp ,"
                    ZSql = ZSql + "NroComp ,"
                    ZSql = ZSql + "Renglon ,"
                    ZSql = ZSql + "Fecha ,"
                    ZSql = ZSql + "Observaciones ,"
                    ZSql = ZSql + "Cuenta ,"
                    ZSql = ZSql + "Debito ,"
                    ZSql = ZSql + "Credito ,"
                    ZSql = ZSql + "FechaOrd ,"
                    ZSql = ZSql + "Titulo ,"
                    ZSql = ZSql + "Nombre ,"
                    ZSql = ZSql + "TituloList ,"
                    ZSql = ZSql + "Impre )"
                    ZSql = ZSql + "Values ("
                    ZSql = ZSql + "'" + WWClave + "',"
                    ZSql = ZSql + "'" + WWTipoMovi + "',"
                    ZSql = ZSql + "'" + WWComprobante + "',"
                    ZSql = ZSql + "'" + WWTipoComp + "',"
                    ZSql = ZSql + "'" + WWLetraComp + "',"
                    ZSql = ZSql + "'" + WWPuntoComp + "',"
                    ZSql = ZSql + "'" + WWNroComp + "',"
                    ZSql = ZSql + "'" + WWRenglon + "',"
                    ZSql = ZSql + "'" + WWFecha + "',"
                    ZSql = ZSql + "'" + WWObservaciones + "',"
                    ZSql = ZSql + "'" + WWCuenta + "',"
                    ZSql = ZSql + "'" + WWDebito + "',"
                    ZSql = ZSql + "'" + WWCredito + "',"
                    ZSql = ZSql + "'" + WWfechaord + "',"
                    ZSql = ZSql + "'" + WWTitulo + "',"
                    ZSql = ZSql + "'" + WWNombre + "',"
                    ZSql = ZSql + "'" + WWTitulolist + "',"
                    ZSql = ZSql + "'" + WWImpre + "')"
                    spImputac = ZSql
                    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
            
        Next Ciclo
    
    End If
    
    
    
    
    
    
    
    
    
    
    Rem Procesa las Compras
    
    If Tipo4.Value = 1 Then
    
        ZLugar = 0
        Erase ZVector
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ImpCyb"
        ZSql = ZSql + " Where ImpCyb.OrdFecha >= '" + WDesde + "'"
        ZSql = ZSql + " and ImpCyb.OrdFecha <= '" + WHasta + "'"
        spImpCyb = ZSql
        Set rstImpCyb = db.OpenRecordset(spImpCyb, dbOpenSnapshot, dbSQLPassThrough)
        If rstImpCyb.RecordCount > 0 Then
    
            With rstImpCyb
                .MoveFirst
                Do
            
                    ZZOrdFecha = rstImpCyb!ordfecha
                    ZZfecha = rstImpCyb!Fecha
                    ZZProveedor = Str$(rstImpCyb!Proveedor)
                    ZZTipo = Str$(rstImpCyb!Tipo)
                    ZZLetra = rstImpCyb!Letra
                    ZZPunto = Str$(rstImpCyb!Punto)
                    ZZNumero = Str$(rstImpCyb!Numero)
                    ZZCuenta = rstImpCyb!Cuenta
                    ZZDebito = Str$(rstImpCyb!Debito)
                    ZZCredito = Str$(rstImpCyb!Credito)
                    ZZObservaciones = rstImpCyb!Observaciones
                    ZZTipo = Str$(rstImpCyb!Tipo)
                    Select Case Val(ZZTipo)
                        Case 1
                            ZZImpre = "FAC"
                        Case 2
                            ZZImpre = "N/D"
                        Case 3
                            ZZImpre = "N/C"
                        Case 7
                            ZZImpre = "TK"
                        Case 8
                            ZZImpre = "RC"
                        Case Else
                            ZZImpre = ""
                    End Select
                    ZZClave = rstImpCyb!Clave
                    
                    ZLugar = ZLugar + 1
                    
                    ZVector(ZLugar, 1) = ZZOrdFecha
                    ZVector(ZLugar, 2) = ZZfecha
                    ZVector(ZLugar, 3) = ZZProveedor
                    ZVector(ZLugar, 4) = ZZTipo
                    ZVector(ZLugar, 5) = ZZLetra
                    ZVector(ZLugar, 6) = ZZPunto
                    ZVector(ZLugar, 7) = ZZNumero
                    ZVector(ZLugar, 8) = ZZCuenta
                    ZVector(ZLugar, 9) = ZZDebito
                    ZVector(ZLugar, 10) = ZZCredito
                    ZVector(ZLugar, 11) = ZZObservaciones
                    ZVector(ZLugar, 12) = ZZTipo
                    ZVector(ZLugar, 13) = ZZImpre
                    ZVector(ZLugar, 14) = ZZClave
                    
                    .MoveNext
                        If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End With
            
            rstImpCyb.Close
        
            For Ciclo = 1 To ZLugar
                    
                ZZOrdFecha = ZVector(Ciclo, 1)
                ZZfecha = ZVector(Ciclo, 2)
                ZZProveedor = ZVector(Ciclo, 3)
                ZZTipo = ZVector(Ciclo, 4)
                ZZLetra = ZVector(Ciclo, 5)
                ZZPunto = ZVector(Ciclo, 6)
                ZZNumero = ZVector(Ciclo, 7)
                ZZCuenta = ZVector(Ciclo, 8)
                ZZDebito = ZVector(Ciclo, 9)
                ZZCredito = ZVector(Ciclo, 10)
                ZZOSBERVACIONES = ZVector(Ciclo, 11)
                ZZTipo = ZVector(Ciclo, 12)
                ZZImpre = ZVector(Ciclo, 13)
                ZZClave = ZVector(Ciclo, 14)

                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Proveedor"
                ZSql = ZSql + " Where Proveedor.Proveedor = " + "'" + ZZProveedor + "'"
                spProveedor = ZSql
                Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If rstProveedor.RecordCount > 0 Then
                    ZZObservaciones = rstProveedor!Nombre + " - " + ZZObservaciones
                    rstProveedor.Close
                End If
                
                WWClave = ZZClave
                WWTipoMovi = "4"
                WWComprobante = "0"
                WWTipoComp = ZZTipo
                WWLetraComp = ZZLetra
                WWPuntoComp = ZZPunto
                WWNroComp = ZZNumero
                WWRenglon = "0"
                WWFecha = ZZfecha
                WWObservaciones = Left$(ZZObservaciones, 50)
                WWCuenta = ZZCuenta
                WWDebito = ZZDebito
                WWCredito = ZZCredito
                WWfechaord = ZZFechaOrd
                WWTitulo = "Compras"
                WWNombre = WNombreEmpresa
                WWTitulolist = WTitulo
                WWImpre = ZZImpre
                If Val(WWCuenta) = 0 Then
                    WWCuenta = "0"
                End If
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO Imputac ("
                ZSql = ZSql + "Clave ,"
                ZSql = ZSql + "TipoMovi ,"
                ZSql = ZSql + "Comprobante ,"
                ZSql = ZSql + "TipoComp ,"
                ZSql = ZSql + "LetraComp ,"
                ZSql = ZSql + "PuntoComp ,"
                ZSql = ZSql + "NroComp ,"
                ZSql = ZSql + "Renglon ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Observaciones ,"
                ZSql = ZSql + "Cuenta ,"
                ZSql = ZSql + "Debito ,"
                ZSql = ZSql + "Credito ,"
                ZSql = ZSql + "FechaOrd ,"
                ZSql = ZSql + "Titulo ,"
                ZSql = ZSql + "Nombre ,"
                ZSql = ZSql + "TituloList ,"
                ZSql = ZSql + "Impre )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WWClave + "',"
                ZSql = ZSql + "'" + WWTipoMovi + "',"
                ZSql = ZSql + "'" + WWComprobante + "',"
                ZSql = ZSql + "'" + WWTipoComp + "',"
                ZSql = ZSql + "'" + WWLetraComp + "',"
                ZSql = ZSql + "'" + WWPuntoComp + "',"
                ZSql = ZSql + "'" + WWNroComp + "',"
                ZSql = ZSql + "'" + WWRenglon + "',"
                ZSql = ZSql + "'" + WWFecha + "',"
                ZSql = ZSql + "'" + WWObservaciones + "',"
                ZSql = ZSql + "'" + WWCuenta + "',"
                ZSql = ZSql + "'" + WWDebito + "',"
                ZSql = ZSql + "'" + WWCredito + "',"
                ZSql = ZSql + "'" + WWfechaord + "',"
                ZSql = ZSql + "'" + WWTitulo + "',"
                ZSql = ZSql + "'" + WWNombre + "',"
                ZSql = ZSql + "'" + WWTitulolist + "',"
                ZSql = ZSql + "'" + WWImpre + "')"
                spImputac = ZSql
                Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
                
            Next Ciclo
            
        End If
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "DELETE Imputac"
    ZSql = ZSql + " Where '" + DesdeCuenta.Text + "' > Cuenta"
    ZSql = ZSql + " or '" + HastaCuenta.Text + "' < Cuenta"
    spImputac = ZSql
    Set rstImputac = db.OpenRecordset(spImputac, dbOpenSnapshot, dbSQLPassThrough)
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
        
    Listado.SQLQuery = "SELECT Imputac.TipoMovi, Imputac.Comprobante, Imputac.NroComp, Imputac.Fecha, Imputac.Observaciones, Imputac.Cuenta, Imputac.Debito, Imputac.Credito, Imputac.FechaOrd, Imputac.Titulo, Imputac.Nombre, Imputac.Titulolist, Imputac.Impre, " _
            + "Cuenta.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Imputac Imputac, " _
            + DSQ + ".dbo.Cuenta Cuenta  " _
            + "Where " _
            + "Imputac.Cuenta = Cuenta.Cuenta AND " _
            + "Imputac.Cuenta >= '" + DesdeCuenta.Text + "' AND " _
            + "Imputac.Cuenta <= '" + HastaCuenta.Text + "'"
    
    Listado.Connect = Connect()
    
    If TipoList.ListIndex = 0 Then
        Listado.ReportFileName = "Imputa.rpt"
            Else
        Listado.ReportFileName = "ImputaCon.rpt"
    End If
    
    Listado.Action = 1
    
    Exit Sub
    
Error_Programa:
     Rem coderr = Err
     Rem Call Errores(coderr, "Error en el sistema", "Se produjo el error " + Str$(coderr))
     Resume Next
    
End Sub

Private Sub Cancela_Click()
    PrgImpcyb.Hide
    Unload Me
    MenuAdminis.SetFocus
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
            DesdeCuenta.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Private Sub DesdeCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCuenta.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeCuenta.Text = ""
    End If
End Sub

Private Sub HastaCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaCuenta.Text = ""
    End If
End Sub

Sub Form_Load()

    PrgImpcyb.Caption = "Listado de Imputaciones de Contables : " + WNombreEmpresa
    
    Tipo1.Value = False
    Tipo2.Value = False
    Tipo3.Value = False
    Tipo4.Value = False
    Tipo5.Value = False
    
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeCuenta.Text = ""
    HastaCuenta.Text = ""
    
    Frame2.Visible = True
    
    TipoList.Clear
    
    TipoList.AddItem "Completo"
    TipoList.AddItem "Resumido"
    
    TipoList.ListIndex = 0
    
    
End Sub

Private Sub Consulta_Click()

    On Error GoTo WError
    
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

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
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Indice = Pantalla.ListIndex
    DesdeCuenta.Text = WIndice.List(Indice)
    HastaCuenta.Text = WIndice.List(Indice)
    DesdeCuenta.SetFocus
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Cuenta"
        ZSql = ZSql + " Where Cuenta.Descripcion LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
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
    
    End If
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DesdeCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub HastaCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo3_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Tipo5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub TipoList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 115 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call Panta_Click
        Case 115
            Call Consulta_Click
        Case 120
            Call Impre_Click
        Case 121
            Call Cancela_Click
        Case Else
    End Select
End Sub


