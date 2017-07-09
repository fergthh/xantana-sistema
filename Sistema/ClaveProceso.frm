VERSION 5.00
Begin VB.Form ClaveProceso 
   Caption         =   "Ingreso de Clave de Usuario"
   ClientHeight    =   2370
   ClientLeft      =   3765
   ClientTop       =   3210
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   Begin VB.TextBox WClave 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese se Clave de Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
End
Attribute VB_Name = "ClaveProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub WClave_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        Salida = "N"
        Select Case UCase(Trim(WClave.Text))
            Case "MAMMA", "0"
                Salida = "S"
                ZZNivel = 0
            Case "BARBI"
                Salida = "S"
                ZZNivel = 1
            Case "HEMAN"
                Salida = "S"
                ZZNivel = 2
            Case Else
        End Select
        
        If Salida = "S" Then
            ClaveProceso.Hide
            Unload Me
            Select Case ZZClaveProceso
                Case 1
                    PrgRecibos.Show
                Case 2
                    PrgCtaCte.Show
                Case 3
                    PrgCtaCte1.Show
                Case 4
                    PrgSaldoCta.Show
                Case 5
                    PrgCtaCteVen.Show
                Case 6
                    PrgListaComisiones.Show
                Case 7
                    PrgEstaProv.Show
                Case 8
                    PrgEstaGrupo.Show
                Case 9
                    PrgEstaVendedor.Show
                Case 10
                    PrgEstaArticulo.Show
                Case 11
                    PrgEstaCliente.Show
                Case 12
                    PrgMixVentas.Show
                Case Else
            End Select
        End If
        
    End If
End Sub

