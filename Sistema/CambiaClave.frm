VERSION 5.00
Begin VB.Form PrgCambiaClave 
   Caption         =   "Cambio de Clave de Seguridad"
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
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.TextBox ClaveNuevaII 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   7
         Text            =   " "
         Top             =   1440
         Width           =   1455
      End
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox ClaveNueva 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox ClaveAnterior 
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
         IMEMode         =   3  'DISABLE
         Left            =   2400
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Confirma  Clave"
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
         Left            =   600
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Nueva Clave"
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
         Left            =   600
         TabIndex        =   4
         Top             =   960
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
         Left            =   600
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PrgCambiaClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZNombreBase As String

Private Sub Proceso_Click()

    ZZNombreBase = WNombreBase

    txtOdbc = "Empresa"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Clave"
    ZSql = ZSql + " Where Clave.Clave = " + "'" + ClaveAnterior.Text + "'"
    spClave = ZSql
    Set rstClave = db.OpenRecordset(spClave, dbOpenSnapshot, dbSQLPassThrough)
    If rstClave.RecordCount > 0 Then
        rstClave.Close
            Else
        Call Cambia_Empresa
        m$ = "Clave actual erronea"
        A% = MsgBox(m$, 0, "Cambio de Claves")
        Exit Sub
    End If

    If ClaveNueva.Text <> ClaveNuevaII.Text Then
        m$ = "Las Clave actual no concuerda con la ratificacion de la misma"
        A% = MsgBox(m$, 0, "Cambio de Claves")
        Call Cambia_Empresa
        Exit Sub
    End If

    If ClaveNueva.Text = "" Then
        m$ = "Se debe informar clave de seguridad nueva"
        A% = MsgBox(m$, 0, "Cambio de Claves")
        Call Cambia_Empresa
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Clave SET "
    ZSql = ZSql + " Clave = " + "'" + ClaveNueva.Text + "'"
    ZSql = ZSql + " Where Clave = " + "'" + ClaveAnterior.Text + "'"
    spClave = ZSql
    Set rstClave = db.OpenRecordset(spClave, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Cambia_Empresa
    
    Call Cancela_Click
    
End Sub

Private Sub Cancela_Click()
    PrgCambiaClave.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub ClaveAnterior_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        ZZNombreBase = WNombreBase

        txtOdbc = "Empresa"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Clave"
        ZSql = ZSql + " Where Clave.Clave = " + "'" + ClaveAnterior.Text + "'"
        spClave = ZSql
        Set rstClave = db.OpenRecordset(spClave, dbOpenSnapshot, dbSQLPassThrough)
        If rstClave.RecordCount > 0 Then
            rstClave.Close
            ClaveNueva.SetFocus
                Else
            m$ = "Clave actual erronea"
            A% = MsgBox(m$, 0, "Cambio de Claves")
        End If
        
        Call Cambia_Empresa
        
    End If
    
    If KeyAscii = 27 Then
        ClaveAnterior.Text = ""
    End If
End Sub

Private Sub ClaveNueva_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ClaveNuevaII.SetFocus
    End If
    If KeyAscii = 27 Then
        ClaveNueva.Text = ""
    End If
End Sub

Private Sub ClaveNuevaII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ClaveNueva.Text <> ClaveNuevaII.Text Then
            m$ = "Las Clave actual no concuerda con la ratificacion de la misma"
            A% = MsgBox(m$, 0, "Cambio de Claves")
            ClaveNueva.SetFocus
                Else
            ClaveAnterior.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        ClaveNuevaII.Text = ""
    End If
End Sub

Sub Form_Load()
    
    ClaveAnterior.Text = ""
    ClaveNueva.Text = ""
    ClaveNuevaII.Text = ""
    
    Frame2.Visible = True
    
End Sub

Sub Cambia_Empresa()
        
    WNombreBase = ZZNombreBase
    
    txtOdbc = WNombreBase
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

End Sub

