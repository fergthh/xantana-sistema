VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrgAgenda 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Agenda Diaria"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   555
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   7995
   ScaleWidth      =   11880
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   120
      MaxLength       =   50
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
   Begin VB.Frame Ingresa 
      Height          =   2775
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox WObservaciones 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   11
         Top             =   1920
         Width           =   6015
      End
      Begin VB.TextBox WTelefono 
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
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1440
         Width           =   6015
      End
      Begin VB.TextBox WDireccion 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   9
         Top             =   960
         Width           =   6015
      End
      Begin VB.TextBox WNombre 
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
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   8
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "Observaciones"
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
         Left            =   360
         TabIndex        =   7
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Telefono"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Direccion"
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
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   10680
      TabIndex        =   1
      Top             =   7920
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Vector 
      Height          =   6735
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   11880
      _Version        =   327680
      Rows            =   1000
      Cols            =   5
      BackColor       =   16777088
   End
   Begin VB.Image Nuevo 
      Height          =   480
      Left            =   3960
      MouseIcon       =   "AGENDA.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "AGENDA.frx":030A
      ToolTipText     =   "Ingreso de una nueva entrada"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   5640
      MouseIcon       =   "AGENDA.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "AGENDA.frx":0E56
      ToolTipText     =   "Borra el renglon Seleccionado"
      Top             =   7320
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   7320
      MouseIcon       =   "AGENDA.frx":1698
      MousePointer    =   99  'Custom
      Picture         =   "AGENDA.frx":19A2
      ToolTipText     =   "Salida"
      Top             =   7320
      Width           =   480
   End
End
Attribute VB_Name = "PrgAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Dato As String
Private Auxi As String
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WFila As Integer
Private WColu As Integer
Private WInicio As Integer
Private XDato As Integer
Dim WLimpia As Integer
Dim WRow As Integer
Dim WCol As Integer
Dim WLugar As Integer
Dim Otro(1000) As Integer
Dim WBorra(1000, 4) As String
Dim WLL As Integer

Sub Imprime_Datos()

    On Error GoTo WError
    
    Renglon = 0
    Erase Otro
    Ayuda.Text = ""

    With rstAgenda
    
        .Index = "Nombre"
        .MoveFirst
        
        Do
            If .EOF = False Then
            
                Renglon = Renglon + 1
                
                Vector.Row = Renglon
                
                Vector.Col = 1
                Vector.Text = !Nombre
                
                Vector.Col = 2
                Vector.Text = !Direccion
                
                Vector.Col = 3
                Vector.Text = !Telefono
                
                Vector.Col = 4
                Vector.Text = !Observaciones
                
                Otro(Renglon) = !Clave
                
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
                
    Vector.TopRow = 1
    Vector.Col = 1
    Vector.Row = 1
    
    Exit Sub

WError:

    Resume Next

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Agenda
End Sub

Private Sub Graba_Click()

    On Error GoTo WError
        Rem da de baja los datos anteriores
        
        With rstAgenda
            .Index = "Clave"
            .MoveFirst
            Do
                If .EOF = False Then
                    .Delete
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
                                
        Rem GRABA los datos actuales
        
        Renglon = 0
        
        With rstAgenda
        
            .Index = "Clave"
                                        
            For a = 1 To 999
        
                Vector.Row = a
                Vector.Col = 1
                
                If Vector.Text <> "" Then
                
                    Renglon = Renglon + 1
                
                    Vector.Col = 1
                    WNombre = Vector.Text
                
                    Vector.Col = 2
                    WDireccion = Vector.Text
                
                    Vector.Col = 3
                    WTelefono = Vector.Text
                
                    Vector.Col = 4
                    WObservaciones = Vector.Text
                
                    .AddNew
                    !Clave = Renglon
                    !Nombre = WNombre
                    !Direccion = WDireccion
                    !Telefono = WTelefono
                    !Observaciones = WObservaciones
                    .Update
                        
                End If
                                        
            Next a
            
        End With
        
        If WLimpia = 0 Then
            Call CmdLimpiar_Click
            Call Imprime_Datos
        End If
        
    Exit Sub

WError:

    Resume Next
        
End Sub

Private Sub CmdDelete_Click()

    For Ciclo = 1 To Vector.Cols - 1
        Vector.Col = Ciclo
        Vector.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To Vector.Rows - 1
        Vector.Row = Ciclo
        Vector.Col = 1
        WAuxi1 = Vector.Text
        Vector.Col = 2
        WAuxi2 = Vector.Text
        Vector.Col = 3
        WAuxi3 = Vector.Text
        Vector.Col = 4
        WAuxi4 = Vector.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Or WAuxi3 <> "" Or WAuxi4 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To Vector.Cols - 1
                Vector.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = Vector.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call CmdLimpiar_Click
    
    For Ciclo = 1 To EntraVector
        Vector.Row = Ciclo
        For da = 1 To Vector.Cols - 1
            Vector.Col = da
            Vector.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
End Sub

Private Sub CmdLimpiar_Click()

    Vector.Clear

    Vector.ColWidth(0) = 100
    Vector.ColWidth(1) = 3000
    Vector.ColWidth(2) = 3000
    Vector.ColWidth(3) = 2000
    Vector.ColWidth(4) = 2800
    
    Vector.Row = 0
    
    Vector.Col = 1
    Vector.Text = "Nombre"
    Vector.ColAlignment(1) = flexAlignLeftCenter
    
    Vector.Col = 2
    Vector.Text = "Direccion"
    Vector.ColAlignment(2) = flexAlignLeftCenter
    
    Vector.Col = 3
    Vector.Text = "Telefono"
    Vector.ColAlignment(3) = flexAlignLeftCenter
    
    Vector.Col = 4
    Vector.Text = "Observaciones"
    Vector.ColAlignment(4) = flexAlignLeftCenter
    
    Vector.Col = 1
    Vector.Row = 1
    
    Vector.Row = 0
    For Ciclo = 1 To Vector.Cols - 1
        Vector.Col = Ciclo
        WTitulo(Ciclo).Text = Vector.Text
        WTitulo(Ciclo).Left = Vector.CellLeft + Vector.Left
        WTitulo(Ciclo).Top = Vector.CellTop + Vector.Top
        WTitulo(Ciclo).Width = Vector.CellWidth
        WTitulo(Ciclo).Height = Vector.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA vector
    
    WAncho = 400
    For Ciclo = 0 To Vector.Cols - 1
        WAncho = WAncho + Vector.ColWidth(Ciclo)
    Next Ciclo
    Vector.Width = WAncho
    
End Sub

Private Sub CmdClose_Click()

    WLimpia = 1
    Call Graba_Click
    WLimpia = 0

    CmdLimpiar_Click

    With rstAgenda
        .Close
    End With
    
    DbsAdminis.Close
    PrgAgenda.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Load()
    
    Call CmdLimpiar_Click
    Call Imprime_Datos
 
End Sub

Private Sub Limpia_Click()
    Ayuda.Text = ""
    Call CmdLimpiar_Click
    Call Imprime_Datos
    Ayuda.SetFocus
End Sub

Private Sub Nuevo_Click()

    WRow = 0
    WLL = 0
    
    WNombre.Text = ""
    WDireccion.Text = ""
    WTelefono.Text = ""
    WObservaciones.Text = ""
    
    Ingresa.Visible = True
    WNombre.SetFocus

End Sub

Private Sub Vector_DblClick()

    WRow = Vector.Row
    WLL = Otro(WRow)
    
    Ingresa.Visible = True
    
    Vector.Col = 1
    WNombre.Text = Vector.Text
    
    Vector.Col = 2
    WDireccion.Text = Vector.Text
    
    Vector.Col = 3
    WTelefono.Text = Vector.Text
    
    Vector.Col = 4
    WObservaciones.Text = Vector.Text
    
    WNombre.SetFocus

End Sub

Private Sub WNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WDireccion.SetFocus
    End If
End Sub

Private Sub WDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WTelefono.SetFocus
    End If
End Sub
Private Sub WTelefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WObservaciones.SetFocus
    End If
End Sub

Private Sub WObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If WRow <> 0 Then
            With rstAgenda
                .Index = "Clave"
                .Seek "=", WLL
                If .NoMatch = False Then
                    .Edit
                    !Nombre = WNombre
                    !Direccion = WDireccion
                    !Telefono = WTelefono
                    !Observaciones = WObservaciones
                    .Update
                End If
            End With
                Else
                
            With rstAgenda
                .Index = "Clave"
                Claveven$ = "999999"
                .Seek "<=", Claveven$
                If .NoMatch = False Then
                    WLL = !Clave + 1
                        Else
                    WLL = "1"
                End If
            End With
                
            With rstAgenda
                .Index = "Clave"
                .Seek "=", WLL
                If .NoMatch = True Then
                    .AddNew
                    !Clave = WLL
                    !Nombre = WNombre
                    !Direccion = WDireccion
                    !Telefono = WTelefono
                    !Observaciones = WObservaciones
                    .Update
                End If
            End With
            
        End If
        
        Ingresa.Visible = False
        Call Imprime_Datos
        
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    WEspacios = Len(Ayuda.Text)
    Call CmdLimpiar_Click
    
    Renglon = 0
    Erase Otro
    
    With rstAgenda
        .Index = "Nombre"
        .MoveFirst
        Do
            If .EOF = False Then
                da = Len(!Nombre) - WEspacios
                For aa = 1 To da
                    If UCase(Left$(Ayuda.Text, WEspacios)) = UCase(Mid$(!Nombre, aa, WEspacios)) Then
                    
                        Renglon = Renglon + 1
                
                        Vector.Row = Renglon
                
                        Vector.Col = 1
                        Vector.Text = !Nombre
                
                        Vector.Col = 2
                        Vector.Text = !Direccion
                
                        Vector.Col = 3
                        Vector.Text = !Telefono
                
                        Vector.Col = 4
                        Vector.Text = !Observaciones
                        
                        Otro(trenglon) = !Clave
                    
                        Exit For
                        
                    End If
                Next aa
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    Ayuda.SetFocus
    
    End If

End Sub





