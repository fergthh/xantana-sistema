VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCambiaCuenta 
   Caption         =   "Cambio de Cuentas Contables"
   ClientHeight    =   2610
   ClientLeft      =   2820
   ClientTop       =   1305
   ClientWidth     =   6210
   LinkTopic       =   "Form2"
   ScaleHeight     =   2610
   ScaleWidth      =   6210
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   5655
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
         Left            =   4080
         TabIndex        =   10
         Top             =   480
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
         Left            =   4080
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox CuentaNueva 
         Alignment       =   1  'Right Justify
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   3
         Text            =   " "
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox CuentaAnterior 
         Alignment       =   1  'Right Justify
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   2
         Text            =   " "
         Top             =   1080
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   600
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
         Left            =   2400
         TabIndex        =   0
         Top             =   240
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
         TabIndex        =   8
         Top             =   360
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Nueva Cuenta"
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
         TabIndex        =   6
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Cuenta Original"
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PrgCambiaCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Proceso_Click()

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    With rstImpcyb
        .Index = "Clave"
        .MoveFirst
        Do
            
            WOrdFecha = !ordfecha
            WFecha = !Fecha
            
            WClave = Left$(!Clave, 21)
            With rstIvacomp
                .Index = "Clave"
                .Seek "=", WClave
                If .NoMatch = False Then
                    WAno = Right$(!Periodo, 4)
                    WMes = Mid$(!Periodo, 4, 2)
                    WDia = Left$(!Periodo, 2)
                    WFecha = !Periodo
                    WOrdFecha = WAno + WMes + WDia
                End If
            End With
            
            If WDesde <= WOrdFecha And WOrdFecha <= WHasta Then
                If !Cuenta = CuentaAnterior.Text Then
                    .Edit
                    !Cuenta = CuentaNueva.Text
                    .Update
                End If
            End If
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    With rstPagos
        .Index = "Clave"
        .MoveFirst
        Do
            
            WOrdFecha = !fechaord
            WFecha = !Fecha
            
            If WDesde <= WOrdFecha And WOrdFecha <= WHasta Then
                If !Cuenta = CuentaAnterior.Text Then
                    .Edit
                    !Cuenta = CuentaNueva.Text
                    .Update
                End If
            End If
                
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
        Loop
    End With
    
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    With rstImpcyb
        .Close
    End With
    With rstIvacomp
        .Close
    End With
    With rstPagos
        .Close
    End With
    DbsAdminis.Close
    Desde.SetFocus
    PrgCambiaCuenta.Hide
    Unload Me
    Menu.SetFocus
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
            CuentaAnterior.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Hasta.Text = "  /  /    "
    End If
End Sub

Private Sub CuentaAnterior_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CuentaNueva.SetFocus
    End If
    If KeyAscii = 27 Then
        CuentaAnterior.Text = ""
    End If
End Sub

Private Sub CuentaNueva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        CuentaNueva.Text = ""
    End If
End Sub

Sub Form_Load()
    
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    CuentaAnterior.Text = ""
    CuentaNueva.Text = ""
    
    Frame2.Visible = True
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Impcyb
    OPEN_FILE_Ivacomp
    OPEN_FILE_Pagos
End Sub

