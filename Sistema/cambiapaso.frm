VERSION 5.00
Begin VB.Form PrgCambiaCuenta 
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
Attribute VB_Name = "PrgCambiaCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Proceso_Click()

    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    With rstclave
        .Close
    End With
    DbsAdminis.Close
    PrgCambiaClave.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub ClaveAnterior_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ClaveNueva.SetFocus
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
        ClaveAnterior.SetFocus
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

Private Sub Form_Activate()
    OPEN_FILE_clave
End Sub

