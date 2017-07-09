VERSION 5.00
Begin VB.Form Clave 
   Caption         =   "Ingreso de Clave de Usuario"
   ClientHeight    =   2370
   ClientLeft      =   3765
   ClientTop       =   3210
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   Begin VB.TextBox WClave 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "INGRESE SU CLAVE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Clave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub WClave_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
         
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
        Salida = "N"
        Select Case UCase(Trim(WClave.Text))
            Case "0", "VENTAS", "LABO"
                Salida = "S"
                ZZNivel = 0
            Case "020166"
                Salida = "S"
                ZZNivel = 1
            Case Else
        End Select
        
        If Salida = "S" Then
            Clave.Hide
            Unload Me
            MenuVen.Show
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Rem WClave.Text = "MAMMA"
    Rem Call WClave_Keypress(13)
End Sub

