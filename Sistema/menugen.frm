VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Menu General del Sistema"
   ClientHeight    =   7365
   ClientLeft      =   1155
   ClientTop       =   930
   ClientWidth     =   9660
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "menugen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
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
      Left            =   0
      Picture         =   "menugen.frx":0442
      TabIndex        =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de Contabilidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   3
      Top             =   5160
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de Provoveedores y Caja y Bancos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de Ventas y Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Image Image4 
      Height          =   1365
      Left            =   5400
      Picture         =   "menugen.frx":14A5
      Top             =   2520
      Width           =   1500
   End
   Begin VB.Image Image3 
      Height          =   1365
      Left            =   3000
      Picture         =   "menugen.frx":3973
      Top             =   720
      Width           =   1500
   End
   Begin VB.Image Image2 
      Height          =   1365
      Left            =   3000
      Picture         =   "menugen.frx":5E41
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   7200
      Left            =   0
      Picture         =   "menugen.frx":6EA4
      Top             =   0
      Width           =   9600
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Admi_Click()
   Menu.WindowState = 1
   A = Shell("Adminis.exe", 0)
   Close
   End
End Sub

Private Sub Compras_Click()
   Menu.WindowState = 1
   A = Shell("compras.exe", 0)
   Close
   End
End Sub

Private Sub Conta_Click()
   Menu.WindowState = 1
   A = Shell("conta.exe", 0)
   Close
   End
End Sub

Private Sub Fin_Click()
    Close
    End
End Sub

Private Sub Form_Activate()

    Rem If Right$(Date$, 4) <> "2002" Then
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
                Menu.Caption = "Sistema de Control de Gestion - Menu General : " + !Nombre
            End If
        End With
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Control de Gestion - Menu General : " + !Nombre
            End If
        End With
    End If

End Sub

Private Sub Sueldos_Click()
   Menu.WindowState = 1
   A = Shell("sueldos.exe", 0)
   Close
   End
End Sub

Private Sub Ventas_Click()
   Menu.WindowState = 1
   A = Shell("ventas.exe", 0)
   Close
   End
End Sub
