VERSION 5.00
Begin VB.Form PrgDolar 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso del Valor Dolar"
   ClientHeight    =   2715
   ClientLeft      =   3390
   ClientTop       =   1650
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   5790
   Begin VB.TextBox ParidadII 
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
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   6
      Text            =   " "
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Paridad 
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
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   0
      Text            =   " "
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Menu F10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3000
      MouseIcon       =   "dolar.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "dolar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Graba F1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      MouseIcon       =   "dolar.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "dolar.frx":0E56
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Codigo 
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
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   1
      Text            =   " "
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Paridad II"
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
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Paridad"
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo "
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
End
Attribute VB_Name = "PrgDolar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxi As String

Sub Imprime_Nombre()
End Sub

Sub Verifica_datos()
End Sub

Sub Format_datos()
    Paridad.Text = Pusing("###,###.##", Paridad.Text)
    ParidadII.Text = Pusing("###,###.##", ParidadII.Text)
End Sub

Sub Imprime_Datos()
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + Codigo.Text + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        Paridad.Text = Trim(rstDolar!Paridad)
        ParidadII.Text = Trim(rstDolar!ParidadII)
        rstDolar.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
End Sub

Private Sub cmdAdd_Click()

    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Dolar"
        ZSql = ZSql + " Where Dolar.Codigo = " + "'" + Codigo.Text + "'"
        spDolar = ZSql
        Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
        If rstDolar.RecordCount > 0 Then
            rstDolar.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Dolar SET "
            ZSql = ZSql + " Paridad = " + "'" + Paridad.Text + "',"
            ZSql = ZSql + " ParidadII = " + "'" + ParidadII.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spDolar = ZSql
            Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Dolar ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Paridad ,"
            ZSql = ZSql + "ParidadII )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Paridad.Text + "',"
            ZSql = ZSql + "'" + ParidadII.Text + "')"
            spDolar = ZSql
            Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Rem Call CmdLimpiar_Click
    
        m$ = "Grabacion realizada"
        aaaaaa% = MsgBox(m$, 0, "Archivo de Valor de Dolar")
    
        
        
    End If
    
End Sub

Private Sub CmdDelete_Click()
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Dolar"
        ZSql = ZSql + " Where Dolar.Codigo = " + "'" + Codigo.Text + "'"
        spDolar = ZSql
        Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
        If rstDolar.RecordCount > 0 Then
            rstDolar.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuestaaaaaa% = MsgBox(m$, 32 + 4, T$)
            If Respuestaaaaaa% = 6 Then
            
                ZSql = ""
                ZSql = ZSql + "DELETE Dolar"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spDolar = ZSql
                Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
                    
                Call CmdLimpiar_Click
                
            End If
        End If
    
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    On Error GoTo WError
    
    Codigo.Text = ""
    Paridad.Text = ""
    ParidadII.Text = ""
    Codigo.SetFocus
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Linea) as [LineaMayor]"
    Rem ZSql = ZSql + " FROM Dolar"
    Rem spDolar = ZSql
    Rem Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstDolar.RecordCount > 0 Then
    Rem     rstDolar.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstDolar!CodigoMayor), "0", rstDolar!CodigoMayor)
    Rem     codigo.text = ZUltimo + 1
    Rem     rstDolar.Close
    Rem End If
    
    Exit Sub
    
WError:

    Resume Next
        
    
End Sub

Private Sub cmdClose_Click()
    PrgDolar.Hide
    Unload Me
    MenuVen.Show
End Sub

Private Sub Dolar_Click()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Paridad_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Paridad.Text = Pusing("###,###.##", Paridad.Text)
        ParidadII.SetFocus
    End If
    If KeyAscii = 27 Then
        Paridad.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub ParidadII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ParidadII.Text = Pusing("###,###.##", ParidadII.Text)
        Paridad.SetFocus
    End If
    If KeyAscii = 27 Then
        ParidadII.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            Codigo.Text = UCase(Codigo.Text)
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Dolar"
            ZSql = ZSql + " Where Dolar.Codigo = " + "'" + Codigo.Text + "'"
            spDolar = ZSql
            Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
            If rstDolar.RecordCount > 0 Then
                rstDolar.Close
                Call Imprime_Datos
                    Else
                WDolar = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WDolar
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
End Sub

Sub Form_Load()

    On Error GoTo WError
    
    Codigo.Text = "1"
    Paridad.Text = ""
    ParidadII.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Dolar"
    ZSql = ZSql + " Where Dolar.Codigo = " + "'" + Codigo.Text + "'"
    spDolar = ZSql
    Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    If rstDolar.RecordCount > 0 Then
        Paridad.Text = Trim(rstDolar!Paridad)
        ParidadII.Text = Trim(rstDolar!ParidadII)
        rstDolar.Close
        Call Format_datos
        Call Imprime_Nombre
    End If
    
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "Select Max(Linea) as [LineaMayor]"
    Rem ZSql = ZSql + " FROM Dolar"
    Rem spDolar = ZSql
    Rem Set rstDolar = db.OpenRecordset(spDolar, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstDolar.RecordCount > 0 Then
    Rem     rstDolar.MoveLast
    Rem     ZUltimo = IIf(IsNull(rstDolar!CodigoMayor), "0", rstDolar!CodigoMayor)
    Rem     codigo.text = ZUltimo + 1
    Rem     rstDolar.Close
    Rem End If
    
    Exit Sub
    
WError:

    Resume Next
        
End Sub

Rem ACA EMPIEZA LAS RUTINAS DE LAS FUNCIONES

Private Sub Codigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Paridad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub ParidadII_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ayuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Desde_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Hasta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Panta_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Impresora_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Or KeyCode = 123 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdAdd_Click
        Case Else
    End Select
End Sub



































