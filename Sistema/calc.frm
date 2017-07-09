VERSION 5.00
Begin VB.Form Calculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculadora"
   ClientHeight    =   2970
   ClientLeft      =   4380
   ClientTop       =   1995
   ClientWidth     =   3270
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2970
   ScaleWidth      =   3270
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Number 
      Caption         =   "7"
      Height          =   480
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      Height          =   480
      Index           =   8
      Left            =   720
      TabIndex        =   8
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      Height          =   480
      Index           =   9
      Left            =   1320
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "C"
      Height          =   480
      Left            =   2040
      TabIndex        =   10
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton CancelEntry 
      Caption         =   "CE"
      Height          =   480
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      Height          =   480
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      Height          =   480
      Index           =   5
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      Height          =   480
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "+"
      Height          =   480
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "-"
      Height          =   480
      Index           =   3
      Left            =   2640
      TabIndex        =   13
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      Height          =   480
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      Height          =   480
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      Height          =   480
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "X"
      Height          =   480
      Index           =   2
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "/"
      Height          =   480
      Index           =   0
      Left            =   2640
      TabIndex        =   15
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      Height          =   480
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      Caption         =   "."
      Height          =   480
      Left            =   1320
      TabIndex        =   18
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      Caption         =   "="
      Height          =   480
      Index           =   4
      Left            =   2040
      TabIndex        =   16
      Top             =   2400
      Width           =   480
   End
   Begin VB.CommandButton Percent 
      Caption         =   "%"
      Height          =   480
      Left            =   2640
      TabIndex        =   17
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   105
      Width           =   3000
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'               Copyright (C) 1994 Microsoft Corporation
'
' Tiene el derecho de utilizar, modificar, reproducir y distribuir libremente y sin cargo alguno
' los archivos de las Aplicaciones de ejemplo (y/o sus versiones modificadas) de cualquier forma
' que estime conveniente, asumiendo que Microsoft no ofrece garant�as sobre los archivos de las
' Aplicaciones de ejemplo, ni tiene ning�n tipo de obligaciones legales con respecto a ellos.
' ------------------------------------------------------------------------
Option Explicit
Dim Op1, Op2                ' Operando introducido previamente.
Dim DecimalFlag As Integer  ' �Est� presente la coma decimal?
Dim NumOps As Integer       ' N�mero de operandos.
Dim LastInput               ' Indica el tipo del �ltimo evento Keypress.
Dim OpFlag                  ' Indica la operaci�n pendiente.
Dim TempReadout

' Procedimiento de evento Click para la tecla C (cancelar).
' Restablece la presentaci�n e inicializa las variables.
Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

' Procedimiento de evento Click para la tecla CE (cancelar entrada).
Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

' Procedimiento de evento Click para la tecla de la coma decimal (,).
' Si la �ltima tecla presionada fue un operador, inicializa
' la lectura a "0". Si no, agrega la coma decimal
' a la presentaci�n.
Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

' Subrutina de inicializaci�n del formulario.
' Establece todas las variables a sus valores iniciales.
Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
    'Decimal.Caption = Format(0, ".")
End Sub

' Procedimiento de evento para las teclas de los n�meros (0-9).
' Agrega el n�mero nuevo a la cantidad presentada.
Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

' Procedimiento de evento para las teclas de operaci�n (+, -, x, /, =).
' Si la �ltima tecla presionada formaba parte de un n�mero,
' incrementa NumOps. Si ya hubiera presente un operando,
' establece Op1. Si hubiera dos, establece Op1 al resultado
' de la operaci�n sobre Op1 y la cadena de entrada actual,
' y presenta el resultado.
Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CDbl(Op1) + CDbl(Op2)
            Case "-"
                Op1 = CDbl(Op1) - CDbl(Op2)
            Case "X"
                Op1 = CDbl(Op1) * CDbl(Op2)
            Case "/"
                If Op2 = 0 Then
                   MsgBox "Imposible dividir entre cero", 48, "Calculadora"
                Else
                   Op1 = CDbl(Op1) / CDbl(Op2)
                End If
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

' Procedimiento de evento para la tecla de porcentaje (%).
' Calcula y presenta el porcentaje del primer operando.
Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub

