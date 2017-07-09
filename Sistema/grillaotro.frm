VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form GrillaOtro 
   ClientHeight    =   5820
   ClientLeft      =   870
   ClientTop       =   1350
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9585
   Begin VB.TextBox WTitulo1 
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
      Index           =   5
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   4080
      Width           =   375
   End
   Begin VB.TextBox WTitulo1 
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3720
      Width           =   375
   End
   Begin VB.TextBox WTitulo1 
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox WTitulo1 
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox WTitulo1 
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox WTexto21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF00&
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
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.ComboBox WCombo11 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto11 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   375
   End
   Begin MSMask.MaskEdBox WTexto31 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSFlexGridLib.MSFlexGrid WVector11 
      Height          =   5175
      Left            =   480
      TabIndex        =   2
      Top             =   480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9128
      _Version        =   65541
      BackColor       =   16777152
   End
End
Attribute VB_Name = "GrillaOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Datos(10000, 3) As String

Dim WParametros1(10, 10) As Double
Dim WFormato1(10) As String
Dim WControl1 As String

Private Sub Form_Load()

    Call Limpia_Vector1

    Erase Datos
    
    Datos(1, 1) = "sgfdsjgfljdsfkl"
    Datos(1, 2) = "12.25"
    
    Datos(2, 1) = "GFFFFFFFFFFFFLFKGLFG"
    Datos(2, 2) = "320"
    
    Datos(3, 1) = "EOITEOPRITOPREITOPREITRE"
    Datos(3, 2) = "36.5"
    
End Sub

Private Sub GridEditText1(ByVal KeyAscii As Integer)

    XColumna = WVector11.Col
    XTipoDato = WParametros1(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto11.Left = WVector11.CellLeft + WVector11.Left
            WTexto11.Top = WVector11.CellTop + WVector11.Top
            WTexto11.Width = WVector11.CellWidth
            WTexto11.Height = WVector11.CellHeight
            WTexto11.Visible = True
            WTexto11.SetFocus
            WTexto11.MaxLength = WParametros1(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto11.Text = WVector11.Text
                    WTexto11.SelStart = Len(WTexto11.Text)
                Case Else
                    WTexto11.Text = Chr$(KeyAscii)
                    WTexto11.SelStart = 1
            End Select
        Case 1
            WTexto21.Left = WVector11.CellLeft + WVector11.Left
            WTexto21.Top = WVector11.CellTop + WVector11.Top
            WTexto21.Width = WVector11.CellWidth
            WTexto21.Height = WVector11.CellHeight
            WTexto21.Visible = True
            WTexto21.SetFocus
            WTexto21.MaxLength = WParametros1(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto21.Text = WVector11.Text
                    WTexto21.SelStart = Len(WTexto21.Text)
                Case Else
                    WTexto21.Text = Chr$(KeyAscii)
                    WTexto21.SelStart = 1
            End Select
        Case 2
            WTexto31.Left = WVector11.CellLeft + WVector11.Left
            WTexto31.Top = WVector11.CellTop + WVector11.Top
            WTexto31.Width = WVector11.CellWidth
            WTexto31.Height = WVector11.CellHeight
            WTexto31.Visible = True
            WTexto31.SetFocus
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector11.Text) = 10 Then
                        WTexto31.Text = WVector11.Text
                            Else
                        WTexto31.Text = "  /  /    "
                    End If
                    WTexto31.SelStart = 0
                Case Else
                    WTexto31.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto31.SelStart = 1
            End Select
        Case Else
    End Select

End Sub

Private Sub EndEdit1()
    Pasa = 0
    If WCombo11.Visible Then
        Pasa = 0
        WVector11.Text = WCombo11.Text
        WCombo11.Visible = False
            Else
        If WTexto11.Visible Then
            Pasa = 1
            WVector11.Text = WTexto11.Text
            WTexto11.Visible = False
                Else
            If WTexto21.Visible Then
                Pasa = 1
                WVector11.Text = WTexto21.Text
                WTexto21.Visible = False
                    Else
                If WTexto31.Visible Then
                    Pasa = 1
                    WVector11.Text = WTexto31.Text
                    WTexto31.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato1(WVector11.Col) <> "" Then
            WVector11.Text = Pusing(WFormato1(WVector11.Col), WVector11.Text)
        End If
    End If
End Sub

Private Sub GridEditCombo1()
    ' Position the ComboBox over the cell.
    WCombo11.Left = WVector11.CellLeft + WVector11.Left
    WCombo11.Top = WVector11.CellTop + WVector11.Top
    WCombo11.Width = WVector11.CellWidth
    WCombo11.Visible = True
    WCombo11.SetFocus
End Sub

Private Sub WTexto11_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto11.Text = ""
            
        Rem F1
        Case 113
            WTexto11.Text = WVector11.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                Call Control_Grilla1
            End If
            Call StartEdit1

        Case vbKeyDown
            ' Move down 1 row.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                If WVector11.Row < WVector11.Rows - 1 Then
                    WVector11.Row = WVector11.Row + 1
                End If
            End If
            Call StartEdit1

        Case vbKeyUp
            ' Move up 1 row.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                If WVector11.Row > WVector11.FixedRows Then
                    WVector11.Row = WVector11.Row - 1
                End If
            End If
            Call StartEdit1

    End Select
End Sub

Private Sub WTexto21_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto21.Text = ""
            
        Rem F1
        Case 113
            WTexto21.Text = WVector11.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                Call Control_Grilla1
            End If
            Call StartEdit1
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                If WVector11.Row < WVector11.Rows - 1 Then
                    WVector11.Row = WVector11.Row + 1
                End If
            End If
            Call StartEdit1

        Case vbKeyUp
            ' Move up 1 row.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                If WVector11.Row > WVector11.FixedRows Then
                    WVector11.Row = WVector11.Row - 1
                End If
            End If
            Call StartEdit1

    End Select
End Sub

Private Sub WTexto31_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto31.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto31.Text = WVector11.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector11.SetFocus
            Call Control_Campo1
            If WControl1 = "S" Then
                Call Control_Grilla1
            End If
            Call StartEdit1

        Case vbKeyDown
            ' Move down 1 row.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                If WVector11.Row < WVector11.Rows - 1 Then
                    WVector11.Row = WVector11.Row + 1
                End If
            End If
            Call StartEdit1

        Case vbKeyUp
            ' Move up 1 row.
            WVector11.SetFocus
            DoEvents
            Call Control_Campo1
            If WControl1 = "S" Then
                If WVector11.Row > WVector11.FixedRows Then
                    WVector11.Row = WVector11.Row - 1
                End If
            End If
            Call StartEdit1

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto11_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto21_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto31_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo11_Click()
    WVector11.SetFocus
End Sub

Private Sub WVector11_DblClick()
    StartEdit1
End Sub

Private Sub WVector11_Click()
    StartEdit1
End Sub

Private Sub WVector11_LeaveCell()
    EndEdit1
End Sub

Private Sub WVector11_GotFocus()
    EndEdit1
End Sub

Rem Desde aca empieza las rutinas a cambiar

Private Sub StartEdit1()
    Select Case WParametros1(4, WVector11.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo11.Clear
            WCombo11.AddItem "Campo1"
            WCombo11.AddItem "Campo2"
            On Error Resume Next
            WCombo11.Text = WVector11.Text
            On Error GoTo 0
            GridEditCombo1
        Case Else
            If WParametros1(2, WVector11.Col) = 0 Then
                GridEditText1 Asc(" ")
            End If
    End Select
End Sub

Private Sub WVector11_KeyPress(KeyAscii As Integer)
    XColumna = WVector11.Col
    Select Case WParametros1(4, WVector11.Col)
        Case 1
        Case Else
            If WParametros1(2, XColumna) = 0 Then
                GridEditText1 KeyAscii
            End If
    End Select
End Sub

Private Sub Control_Grilla1()
    Select Case WVector11.Col
        Case 1
            WVector11.Col = WVector11.Col + 2
        Case 4
            If WVector11.Row < WVector11.Rows - 1 Then
                WVector11.Row = WVector11.Row + 1
            End If
            WVector11.Col = 1
        Case Else
            If WVector11.Col < WVector11.Cols - 1 Then
                WVector11.Col = WVector11.Col + 1
            End If
    End Select
    WVector11.SetFocus
    GridEditText1 KeyAscii
End Sub

Private Sub Control_Campo1()
    XColumna = WVector11.Col
    XFila = WVector11.Row
    WControl1 = "S"
    Select Case XColumna
        Case 1
            WArti = Val(WVector11.Text)
            If WVector11.Text <> "" Then
                If Val(Datos(WArti, 2)) > 0 Then
                    WVector11.Col = 2
                    WVector11.Text = Datos(WArti, 1)
                    WVector11.Col = 4
                    WVector11.Text = Datos(WArti, 2)
                    WVector11.Text = Pusing(WFormato1(4), WVector11.Text)
                    WVector11.Col = 5
                    WVector11.Text = Str$(Val(WVector11.TextMatrix(XFila, 3)) * Val(WVector11.TextMatrix(XFila, 4)))
                    WVector11.Text = Pusing(WFormato1(5), WVector11.Text)
                    WVector11.Col = XColumna
                        Else
                    WControl1 = "N"
                End If
            End If
        Case 3, 4, 5
            WVector11.Col = 5
            WVector11.Text = Str$(Val(WVector11.TextMatrix(XFila, 3)) * Val(WVector11.TextMatrix(XFila, 4)))
            WVector11.Text = Pusing(WFormato1(5), WVector11.Text)
            WVector11.Col = XColumna
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector1()

    Rem ponga la grilla en negritas
    WVector11.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto11.FontName = WVector11.FontName
    WTexto11.FontSize = WVector11.FontSize
    WTexto11.Visible = False
    WTexto21.FontName = WVector11.FontName
    WTexto21.FontSize = WVector11.FontSize
    WTexto21.Visible = False
    WTexto31.FontName = WVector11.FontName
    WTexto31.FontSize = WVector11.FontSize
    WTexto31.Visible = False
    WCombo11.FontName = WVector11.FontName
    WCombo11.FontSize = WVector11.FontSize
    WCombo11.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector11.FixedCols = 1
    WVector11.Cols = 6
    WVector11.FixedRows = 1
    WVector11.Rows = 100
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector11.Text = "Articulo"
    
    Rem Longitud
    Rem WVector11.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros1(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros1(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros1(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros1(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector11.ColWidth(0) = 200
    WVector11.Row = 0
    For Ciclo = 1 To WVector11.Cols - 1
        WVector11.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector11.Text = "Articulo"
                WVector11.ColWidth(Ciclo) = 1200
                WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros1(1, Ciclo) = 4
                WParametros1(2, Ciclo) = 0
                WParametros1(3, Ciclo) = 1
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = ""
            Case 2
                WVector11.Text = "Descripcion"
                WVector11.ColWidth(Ciclo) = 3000
                WVector11.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros1(1, Ciclo) = 50
                WParametros1(2, Ciclo) = 1
                WParametros1(3, Ciclo) = 0
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = ""
            Case 3
                WVector11.Text = "Cantidad"
                WVector11.ColWidth(Ciclo) = 1150
                WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros1(1, Ciclo) = 10
                WParametros1(2, Ciclo) = 0
                WParametros1(3, Ciclo) = 1
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = "###,###.##"
            Case 4
                WVector11.Text = "Precio"
                WVector11.ColWidth(Ciclo) = 1000
                WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros1(1, Ciclo) = 10
                WParametros1(2, Ciclo) = 0
                WParametros1(3, Ciclo) = 1
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = "###,###.##"
            Case 5
                WVector11.Text = "Parcial"
                WVector11.ColWidth(Ciclo) = 1000
                WVector11.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros1(1, Ciclo) = 10
                WParametros1(2, Ciclo) = 1
                WParametros1(3, Ciclo) = 1
                WParametros1(4, Ciclo) = 0
                WFormato1(Ciclo) = "###,###.##"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector11.Row = 0
    For Ciclo = 1 To WVector11.Cols - 1
        WVector11.Col = Ciclo
        WTitulo1(Ciclo).Text = WVector11.Text
        WTitulo1(Ciclo).Left = WVector11.CellLeft + WVector11.Left
        WTitulo1(Ciclo).Top = WVector11.CellTop + WVector11.Top
        WTitulo1(Ciclo).Width = WVector11.CellWidth
        WTitulo1(Ciclo).Height = WVector11.CellHeight
        WTitulo1(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 340
    For Ciclo = 0 To WVector11.Cols - 1
        WAncho = WAncho + WVector11.ColWidth(Ciclo)
    Next Ciclo
    WVector11.Width = WAncho

    ' Size the columns.
    Font.Name = WVector11.Font.Name
    Font.Size = WVector11.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector11.AllowUserResizing = flexResizeBoth
    
    WVector11.Col = 1
    WVector11.Row = 1
    
End Sub

