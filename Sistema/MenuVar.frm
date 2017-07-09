VERSION 5.00
Begin VB.Form MenuVar 
   Caption         =   "Sistema de Control de Gestion - Menu General"
   ClientHeight    =   7200
   ClientLeft      =   1455
   ClientTop       =   795
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9150
   Begin VB.Label ContaII 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   495
      Left            =   4200
      MouseIcon       =   "MenuVar.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Proveedores - Caja y Bancos"
      Top             =   4560
      Width           =   2895
   End
   Begin VB.Image Conta 
      Height          =   1365
      Left            =   2400
      Picture         =   "MenuVar.frx":030A
      Top             =   4200
      Width           =   1500
   End
   Begin VB.Image Logo 
      Height          =   1005
      Left            =   6960
      Picture         =   "MenuVar.frx":136D
      Top             =   6000
      Width           =   2025
   End
   Begin VB.Label FinII 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "Fin del Sistema"
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
      Height          =   240
      Left            =   720
      MouseIcon       =   "MenuVar.frx":1CEB
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Fin del Sistema"
      Top             =   6720
      Width           =   1635
   End
   Begin VB.Image Fin 
      Height          =   480
      Left            =   120
      MouseIcon       =   "MenuVar.frx":1FF5
      MousePointer    =   99  'Custom
      Picture         =   "MenuVar.frx":22FF
      ToolTipText     =   "Fin del Sistema"
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label VentasII 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Ventas y Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   4200
      MouseIcon       =   "MenuVar.frx":2B41
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Ventas y Stock"
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label AdmiII 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Provoveedores Caja y Bancos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   735
      Left            =   2280
      MouseIcon       =   "MenuVar.frx":2E4B
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Proveedores - Caja y Bancos"
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Image Ventas 
      Height          =   1365
      Left            =   2400
      MouseIcon       =   "MenuVar.frx":3155
      MousePointer    =   99  'Custom
      Picture         =   "MenuVar.frx":345F
      ToolTipText     =   "Ventas y Stock"
      Top             =   480
      Width           =   1500
   End
   Begin VB.Image Admi 
      Height          =   1365
      Left            =   5280
      MouseIcon       =   "MenuVar.frx":592D
      MousePointer    =   99  'Custom
      Picture         =   "MenuVar.frx":5C37
      ToolTipText     =   "Proveedores - Caja y Bancos"
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "MenuVar.frx":8105
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "MenuVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Admi_Click()
   MenuVar.WindowState = 1
   A = Shell("Adminis.exe", 0)
   Close
   End
End Sub

Private Sub AdmiII_Click()
   MenuVar.WindowState = 1
   A = Shell("Adminis.exe", 0)
   Close
   End
End Sub

Private Sub Compras_Click()
   MenuVar.WindowState = 1
   A = Shell("compras.exe", 0)
   Close
   End
End Sub

Private Sub Conta_Click()
   MenuVar.WindowState = 1
   A = Shell("conta.exe", 0)
   Close
   End
End Sub

Private Sub ContaII_Click()
   MenuVar.WindowState = 1
   A = Shell("conta.exe", 0)
   Close
   End
End Sub

Private Sub Copia_Click()
    m$ = "Coloque un Diskette an la Unidad -A-"
    A% = MsgBox(m$, 0, "Copia de Seguridad")
   A = Shell("copia.bat", 0)
End Sub

Private Sub CopiaII_Click()
   A = Shell("copia.bat", 0)
End Sub

Private Sub Fin_Click()
    Close
    End
End Sub

Private Sub FinII_Click()
    Close
    End
End Sub

Private Sub Form_Activate()

    Rem If Right$(Date$, 4) <> "2003" Then
    Rem     m$ = "El tiempo de uso para la verificacion del funcionamiento del sistema a terminado" + Chr$(13) + "Comuniquese con su proveedior para adquirir la version completa del mismo"
    Rem     a% = MsgBox(m$, 0, "Sistema de Control de gestion")
    Rem     Close
    Rem     End
    Rem End If

    If WEmpresa = "" Then
        Rem Empresa.Show
        Rem Empresa.SetFocus
        WEmpresa = "0001"
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                MenuVar.Caption = "Sistema de Control de Gestion - Menu General : " + !Nombre
            End If
        End With
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                MenuVar.Caption = "Sistema de Control de Gestion - Menu General : " + !Nombre
            End If
        End With
    End If

End Sub

Private Sub Sueldos_Click()
   MenuVar.WindowState = 1
   A = Shell("sueldos.exe", 0)
   Close
   End
End Sub

Private Sub Ventas_Click()
   MenuVar.WindowState = 1
   A = Shell("ventas.exe", 0)
   Close
   End
End Sub

Private Sub VentasII_Click()
   MenuVar.WindowState = 1
   A = Shell("ventas.exe", 0)
   Close
   End
End Sub


