VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Rem
Rem
Rem
Rem aca empieza los comandos del vector3
Rem
Rem
Rem
Rem


Private Sub GridEditText2(ByVal KeyAscii As Integer)

    XColumna = WVector22.Col
    XTipoDato = WParametros2(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto12.Left = WVector22.CellLeft + WVector22.Left
            WTexto12.Top = WVector22.CellTop + WVector22.Top
            WTexto12.Width = WVector22.CellWidth
            WTexto12.Height = WVector22.CellHeight
            WTexto12.MaxLength = WParametros2(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto12.Text = WVector22.Text
                    WTexto12.SelStart = Len(WTexto12.Text)
                Case Else
                    WTexto12.Text = Chr$(KeyAscii)
                    WTexto12.SelStart = 1
            End Select
            WTexto12.Visible = True
            WTexto12.SetFocus
        Case 1
            WTexto22.Left = WVector22.CellLeft + WVector22.Left
            WTexto22.Top = WVector22.CellTop + WVector22.Top
            WTexto22.Width = WVector22.CellWidth
            WTexto22.Height = WVector22.CellHeight
            WTexto22.MaxLength = WParametros2(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto22.Text = WVector22.Text
                    Rem WTexto22.SelStart = Len(WTexto22.Text)
                    WTexto22.SelStart = 0
                Case Else
                    WTexto22.Text = Chr$(KeyAscii)
                    WTexto22.SelStart = 1
            End Select
            WTexto22.Visible = True
            WTexto22.SetFocus
        Case 2
            WTexto32.Left = WVector22.CellLeft + WVector22.Left
            WTexto32.Top = WVector22.CellTop + WVector22.Top
            WTexto32.Width = WVector22.CellWidth
            WTexto32.Height = WVector22.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector22.Text) = 10 Then
                        WTexto32.Text = WVector22.Text
                            Else
                        WTexto32.Text = "  /  /    "
                    End If
                    WTexto32.SelStart = 0
                Case Else
                    WTexto32.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto32.SelStart = 1
            End Select
            WTexto32.Visible = True
            WTexto32.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit2()
    Pasa = 0
    If WCombo12.Visible Then
        Pasa = 0
        WVector22.Text = WCombo12.Text
        WCombo12.Visible = False
            Else
        If WTexto12.Visible Then
            Pasa = 1
            WVector22.Text = WTexto12.Text
            WTexto12.Visible = False
                Else
            If WTexto22.Visible Then
                Pasa = 1
                WVector22.Text = WTexto22.Text
                WTexto22.Visible = False
                    Else
                If WTexto32.Visible Then
                    Pasa = 1
                    WVector22.Text = WTexto32.Text
                    WTexto32.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato1(WVector22.Col) <> "" Then
            WVector22.Text = Pusing(WFormato1(WVector22.Col), WVector22.Text)
        End If
    End If
End Sub

Private Sub GridEditCombo2()
    ' Position the ComboBox over the cell.
    WCombo12.Left = WVector22.CellLeft + WVector22.Left
    WCombo12.Top = WVector22.CellTop + WVector22.Top
    WCombo12.Width = WVector22.CellWidth
    WCombo12.Visible = True
    WCombo12.SetFocus
End Sub

Private Sub WTexto12_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto12.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto12.Text = WVector22.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector22.SetFocus
            DoEvents
            Call Control_Campo2
            If WControl2 = "S" Then
                Call Control_Grilla2
            End If
            Call StartEdit2

        Case vbKeyDown
            ' Move down 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row < WVector22.Rows - 1 Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row > WVector22.FixedRows Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row - 1
                End If
            End If
            Call StartEdit2

    End Select
End Sub

Private Sub WTexto22_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto22.Text = ""
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto22.Text = WVector22.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector22.SetFocus
            DoEvents
            Call Control_Campo2
            If WControl2 = "S" Then
                Call Control_Grilla2
            End If
            Call StartEdit2
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row < WVector22.Rows - 1 Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row > WVector22.FixedRows Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row - 1
                End If
            End If
            Call StartEdit2

    End Select
End Sub

Private Sub WTexto32_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto32.Text = "  /  /    "
            
        Rem F1,F2,F3,F4,F10
        Case 112, 113, 114, 115, 121
            WFuncion = KeyCode
            WTexto32.Text = WVector22.Text
            Call Ejecuta_Funcion

        Case vbKeyReturn
            ' Finish editing.
            WVector22.SetFocus
            Call Control_Campo2
            If WControl2 = "S" Then
                Call Control_Grilla2
            End If
            Call StartEdit2

        Case vbKeyDown
            ' Move down 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row < WVector22.Rows - 1 Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row + 1
                End If
            End If
            Call StartEdit2

        Case vbKeyUp
            ' Move up 1 row.
            WVector22.SetFocus
            DoEvents
            If WVector22.Row > WVector22.FixedRows Then
                Call Control_Campo2
                If WControl2 = "S" Then
                    WVector22.Row = WVector22.Row - 1
                End If
            End If
            Call StartEdit2

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto12_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto22_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto32_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo12_Click()
    WVector22.SetFocus
End Sub

Private Sub WVector22_Click()
    StartEdit2
End Sub

Private Sub WVector22_LeaveCell()
    EndEdit2
End Sub

Private Sub WVector22_GotFocus()
    EndEdit2
End Sub

Rem Desde aca empieza las rutinas a cambiar

Private Sub StartEdit2()
    Select Case WParametros2(4, WVector22.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo12.Clear
            WCombo12.AddItem "Campo1"
            WCombo12.AddItem "Campo2"
            On Error Resume Next
            WCombo12.Text = WVector22.Text
            On Error GoTo 0
            GridEditCombo2
        Case Else
            If WParametros2(2, WVector22.Col) = 0 Then
                GridEditText2 Asc(" ")
            End If
    End Select
End Sub

Private Sub WVector22_KeyPress(KeyAscii As Integer)
    XColumna = WVector22.Col
    Select Case WParametros2(4, WVector22.Col)
        Case 1
        Case Else
            If WParametros2(2, XColumna) = 0 Then
                GridEditText2 KeyAscii
            End If
    End Select
End Sub

Private Sub Control_Grilla2()
    Select Case WVector22.Col
        Case 1
            WVector22.Col = WVector22.Col + 2
        Case 3
            If WVector22.Row < WVector22.Rows - 1 Then
                WVector22.Row = WVector22.Row + 1
            End If
            WVector22.Col = 1
        Case Else
            If WVector22.Col < WVector22.Cols - 1 Then
                WVector22.Col = WVector22.Col + 1
            End If
    End Select
    WVector22.SetFocus
    GridEditText2 KeyAscii
End Sub

Private Sub Control_Campo2()
    XColumna = WVector22.Col
    XFila = WVector22.Row
    WControl2 = "S"
    Select Case XColumna
        Case 1
            If Val(WVector22.Text) <> 0 Then
                With rstProyecto
                    .Index = "Codigo"
                    .Seek "=", WVector22.Text
                    If .NoMatch = False Then
                        WVector22.Col = 2
                        WVector22.Text = !Descripcion
                        WVector22.Col = XColumna
                            Else
                        WControl2 = "N"
                    End If
                End With
            End If
        Case 3, 4, 5
            WVector22.Col = XColumna
        Case Else
    End Select
End Sub

Private Sub Limpia_Vector2()

    WVector22.Clear
    
    Rem ponga la grilla en negritas
    WVector22.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto12.FontName = WVector22.FontName
    WTexto12.FontSize = WVector22.FontSize
    WTexto12.Visible = False
    WTexto22.FontName = WVector22.FontName
    WTexto22.FontSize = WVector22.FontSize
    WTexto22.Visible = False
    WTexto32.FontName = WVector22.FontName
    WTexto32.FontSize = WVector22.FontSize
    WTexto32.Visible = False
    WCombo12.FontName = WVector22.FontName
    WCombo12.FontSize = WVector22.FontSize
    WCombo12.Visible = False

    ' Establesco loa Valores de la Grilla
    
    WVector22.FixedCols = 1
    WVector22.Cols = 4
    WVector22.FixedRows = 1
    WVector22.Rows = 101
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector22.Text = "Articulo"
    
    Rem Longitud
    Rem WVector22.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros2(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros2(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros2(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros2(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector22.ColWidth(0) = 200
    WVector22.Row = 0
    For Ciclo = 1 To WVector22.Cols - 1
        WVector22.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector22.Text = "Concepto"
                WVector22.ColWidth(Ciclo) = 1500
                WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 10
                WParametros2(2, Ciclo) = 0
                WParametros2(3, Ciclo) = 1
                WParametros2(4, Ciclo) = 0
                WFormato1(Ciclo) = ""
            Case 2
                WVector22.Text = "Descripcion"
                WVector22.ColWidth(Ciclo) = 4000
                WVector22.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros2(1, Ciclo) = 50
                WParametros2(2, Ciclo) = 1
                WParametros2(3, Ciclo) = 0
                WParametros2(4, Ciclo) = 0
                WFormato1(Ciclo) = ""
            Case 3
                WVector22.Text = "Importe"
                WVector22.ColWidth(Ciclo) = 1500
                WVector22.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros2(1, Ciclo) = 10
                WParametros2(2, Ciclo) = 0
                WParametros2(3, Ciclo) = 1
                WParametros2(4, Ciclo) = 0
                WFormato1(Ciclo) = "###,###.##"
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector22.Row = 0
    For Ciclo = 1 To WVector22.Cols - 1
        WVector22.Col = Ciclo
        WTitulo2(Ciclo).Text = WVector22.Text
        WTitulo2(Ciclo).Left = WVector22.CellLeft + WVector22.Left
        WTitulo2(Ciclo).Top = WVector22.CellTop + WVector22.Top
        WTitulo2(Ciclo).Width = WVector22.CellWidth
        WTitulo2(Ciclo).Height = WVector22.CellHeight
        WTitulo2(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector22.Cols - 1
        WAncho = WAncho + WVector22.ColWidth(Ciclo)
    Next Ciclo
    WVector22.Width = WAncho

    ' Size the columns.
    Font.Name = WVector22.Font.Name
    Font.Size = WVector22.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector22.AllowUserResizing = flexResizeBoth
    
    WVector22.Col = 1
    WVector22.Row = 1
    
End Sub

Private Sub WVector22_Scroll()
    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False
End Sub

Private Sub WVector22_DblClick()

    If WVector22.Col = 0 Then
        Exit Sub
    End If

    WTexto12.Visible = False
    WTexto22.Visible = False
    WTexto32.Visible = False

    For Ciclo = 1 To WVector22.Cols - 1
        WVector22.Col = Ciclo
        WVector22.Text = ""
    Next Ciclo
    
    Erase WBorra
    EntraVector = 0
    
    For Ciclo = 1 To WVector22.Rows - 1
        WVector22.Row = Ciclo
        WVector22.Col = 1
        WAuxi1 = WVector22.Text
        WVector22.Col = 3
        WAuxi2 = WVector22.Text
        If WAuxi1 <> "" Or WAuxi2 <> "" Then
            EntraVector = EntraVector + 1
            For Ciclo1 = 1 To WVector22.Cols - 1
                WVector22.Col = Ciclo1
                WBorra(EntraVector, Ciclo1) = WVector22.Text
            Next Ciclo1
        End If
    Next Ciclo
    
    Call Limpia_Vector2
    
    For Ciclo = 1 To EntraVector
        WVector22.Row = Ciclo
        For da = 1 To WVector22.Cols - 1
            WVector22.Col = da
            WVector22.Text = WBorra(Ciclo, da)
        Next da
    Next Ciclo
    
End Sub

Private Sub WTexto22_DblClick()
    Select Case WVector22.Col
        Case 1
            Opcion.Clear
            Opcion.AddItem "Proveedores"
            Opcion.AddItem "Conceptos"
            Opcion.AddItem "Cuentas Contables"
            Opcion.AddItem "Proyectos"
            Rem Opcion.Visible = True
            Opcion.ListIndex = 3
    
            Call Opcion_Click
        Case Else
    End Select
End Sub


