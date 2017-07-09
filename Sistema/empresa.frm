VERSION 5.00
Begin VB.Form Empresa 
   AutoRedraw      =   -1  'True
   Caption         =   "Seleccion de Empresas"
   ClientHeight    =   4080
   ClientLeft      =   2850
   ClientTop       =   1275
   ClientWidth     =   7755
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   7755
   Begin VB.CommandButton Command1 
      Caption         =   "Acepta  Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin VB.ComboBox Selecciona 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
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
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
                
    WEmpresa = Selecciona.ListIndex + 1
    
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
    
    If Val(WEmpresa) = 1 Then
        txtOdbc = "Celugama"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Else
        txtOdbc = "CelugamaII"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End If
    
    
    Empresa.Hide
    Menu.Show
    
End Sub

Private Sub Form_Load()
    
    Selecciona.Clear
    
    txtOdbc = "Empresa"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM empresa"
    ZSql = ZSql + " Order by Empresa.Empresa"
    spEmpresa = ZSql
    Set rstEmpresa = db.OpenRecordset(spEmpresa, dbOpenSnapshot, dbSQLPassThrough)
    If rstEmpresa.RecordCount > 0 Then
        With rstEmpresa
            .MoveFirst
            Do
                If .EOF = False Then
                    Selecciona.AddItem rstEmpresa!Nombre
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEmpresa.Close
    End If

    If Val(WEmpresa) = 0 Then
        Rem WEmpresa = "0001"
        Rem
        Rem     OPEN_FILE_Empresa
        Rem     With rstEmpresa
        Rem         .Index = "Codigo"
        Rem         .Seek "=", Val(WEmpresa)
        Rem         If .NoMatch = False Then
        Rem             Menu.Caption = "Sistema de Sueldos y Jornales : " + !fantasia
        Rem         End If
        Rem     End With
        Rem
        Selecciona.ListIndex = 0
        Rem
    End If
    
End Sub

Private Sub Selecciona_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Command1_Click
    End If
End Sub
