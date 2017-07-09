VERSION 5.00
Begin VB.Form ClaveAdminis 
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
Attribute VB_Name = "ClaveAdminis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub WClave_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        Salida = "N"
        
        Rem txtUserName = "SA"
        Rem txtPassword = "Sw58125812"
        
        txtOdbc = "Fragancias"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZZOperador = rstOperador!operador
            Salida = "S"
            ZZNivel = 0
            rstOperador.Close
                Else
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operador"
            ZSql = ZSql + " Where Operador.ClaveII = " + "'" + WClave.Text + "'"
            spOperador = ZSql
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                ZZOperador = rstOperador!operador
                Salida = "S"
                ZZNivel = 1
                rstOperador.Close
            End If
        End If
        
        If Salida = "S" Then
        
            If ZZNivel <> 1 Then
                WEmpresa = "1"
                Rem txtUserName = "SA"
                Rem txtPassword = "Sw58125812"
                
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Empresa"
                ZSql = ZSql + " Where Empresa.Empresa = " + "'" + WEmpresa + "'"
                spEmpresa = ZSql
                Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
                If rstEmpresa.RecordCount > 0 Then
                    Rem WNombreBase = Trim(rstEmpresa!NombreBase)
                    WNombreEmpresa = Trim(rstEmpresa!Nombre)
                    Rem WDireccionEmpresa = rstEmpresa!Direccion
                    Rem WLocalidadEmresa = rstEmpresa!Localidad
                    Rem WCuitEmpresa = rstEmpresa!Cuit
                    Rem WCtaProveedor = rstEmpresa!CtaProveedores
                    Rem WCtaEfectivo = rstEmpresa!CtaEfectivo
                    Rem WCtaCheques = rstEmpresa!CtaCheque
                    Rem WCtaChequeRecha = rstEmpresa!CtaChequeRecha
                    Rem WCtaRetGan = rstEmpresa!CtaRetGan
                    Rem WCtaRetIva = rstEmpresa!CtaRetIva
                    Rem WCtaretOtra = rstEmpresa!CtaRetOtro
                    Rem WCtaRetSuss = rstEmpresa!CtaRetSuss
                    Rem WCtaDeudores = rstEmpresa!CtaDeudores
                    Rem WCtaDocumentos = rstEmpresa!CtaDocumentos
                    rstEmpresa.Close
                End If
                
                Rem txtUserName = "SA"
                Rem txtPassword = "Sw58125812"
                
                txtOdbc = "Fragancias"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        
                    Else
                    
                WEmpresa = "2"
                Rem txtUserName = "SA"
                Rem txtPassword = "Sw58125812"
                
                txtOdbc = "FraganciasII"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Empresa"
                ZSql = ZSql + " Where Empresa.Empresa = " + "'" + WEmpresa + "'"
                spEmpresa = ZSql
                Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
                If rstEmpresa.RecordCount > 0 Then
                    Rem WNombreBase = Trim(rstEmpresa!NombreBase)
                    WNombreEmpresa = Trim(rstEmpresa!Nombre)
                    Rem WDireccionEmpresa = rstEmpresa!Direccion
                    Rem WLocalidadEmresa = rstEmpresa!Localidad
                    Rem WCuitEmpresa = rstEmpresa!Cuit
                    Rem WCtaProveedor = rstEmpresa!CtaProveedores
                    Rem WCtaEfectivo = rstEmpresa!CtaEfectivo
                    Rem WCtaCheques = rstEmpresa!CtaCheque
                    Rem WCtaChequeRecha = rstEmpresa!CtaChequeRecha
                    Rem WCtaRetGan = rstEmpresa!CtaRetGan
                    Rem WCtaRetIva = rstEmpresa!CtaRetIva
                    Rem WCtaretOtra = rstEmpresa!CtaRetOtro
                    Rem WCtaRetSuss = rstEmpresa!CtaRetSuss
                    Rem WCtaDeudores = rstEmpresa!CtaDeudores
                    Rem WCtaDocumentos = rstEmpresa!CtaDocumentos
                    rstEmpresa.Close
                End If
                
                Rem txtUserName = "SA"
                Rem txtPassword = "Sw58125812"
                
                txtOdbc = "FraganciasII"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
            End If
        
        
            ClaveAdminis.Hide
            Unload Me
            MenuAdminis.Show
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Rem WClave.Text = "MAMMA"
    Rem Call WClave_Keypress(13)
End Sub


