VERSION 5.00
Begin VB.Form PrgClaveSis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso al Sistema"
   ClientHeight    =   2415
   ClientLeft      =   2925
   ClientTop       =   2085
   ClientWidth     =   5805
   Icon            =   "clavesis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1426.862
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5450.582
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
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
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Clave 
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
      Left            =   3240
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese su Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
End
Attribute VB_Name = "PrgClavesis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As Recordset
Dim spConsul As String
Dim gAplicacion As String
Dim strConnect As String
Dim WtxtUserName As String
Dim WtxtPassword As String
Dim txtOdbc As String
Dim txtUserName As String
Dim txtPassword As String
Dim mm As String
Dim aa As Integer
Dim ZOrdFecha As String
Dim ZVto As String
Dim A As Integer

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'establece la variable global a false
    'para indicar un fallo en el inicio de sesión
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Clave
End Sub

Private Sub Form_Load()

    OPEN_FILE_Clave
    
    WtxtUserName = "Desarrollo"
    WtxtPassword = "Desarrollo"

End Sub

Private Sub Clave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rstClave
            .Index = "Clave"
            .Seek "=", Clave
            If .NoMatch = False Then
                PrgClavesis.WindowState = 1
                A = Shell("sistema.exe", 0)
                Close
                End
            End If
        End With
    End If
    If KeyAscii = 27 Then
        Clave.Text = ""
    End If
End Sub

