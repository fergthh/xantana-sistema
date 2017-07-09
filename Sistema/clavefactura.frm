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
Attribute VB_Name = "Clave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    WClave.Text = "MAMMA"
    Call WClave_Keypress(13)
End Sub

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
        
            WEmpresa = "1"
            
            txtOdbc = "Empresa"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Empresa"
            ZSql = ZSql + " Where Empresa.Empresa = " + "'" + WEmpresa + "'"
            spEmpresa = ZSql
            Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
            If rstEmpresa.RecordCount > 0 Then
                WNombreBase = Trim(rstEmpresa!NombreBase)
                WNombreEmpresa = Trim(rstEmpresa!Nombre)
                WDireccionEmpresa = rstEmpresa!Direccion
                WLocalidadEmresa = rstEmpresa!Localidad
                WCuitEmpresa = rstEmpresa!Cuit
                WCtaProveedor = rstEmpresa!CtaProveedores
                WCtaEfectivo = rstEmpresa!CtaEfectivo
                WCtaCheques = rstEmpresa!CtaCheque
                WCtaChequeRecha = rstEmpresa!CtaChequeRecha
                WCtaRetGan = rstEmpresa!CtaRetGan
                WCtaRetIva = rstEmpresa!CtaRetIva
                WCtaretOtra = rstEmpresa!CtaRetOtro
                WCtaRetSuss = rstEmpresa!CtaRetSuss
                WCtaDeudores = rstEmpresa!CtaDeudores
                WCtaDocumentos = rstEmpresa!CtaDocumentos
                rstEmpresa.Close
            End If
            
            txtOdbc = "Celugama"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            Clave.Hide
            Unload Me
            Menu.Show
        End If
        
    End If
End Sub
