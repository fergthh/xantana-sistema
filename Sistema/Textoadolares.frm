VERSION 5.00
Begin VB.Form textoadolares 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "textoadolares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtLetra As String
Dim Decimales As Double
Private mvarNumero As Variant ' copia local

Private Sub Text1_Click()
    aa = Text1.Text
    Call numtoword(aa)
    Text2.Text = txtLetra
End Sub


Public Function numtoword(numstr As Variant) As String
'The best data type to feed in is
'Decimal, but it is up to you
Dim tempstr As String
Dim newstr As String
Dim mstrMoneda As String
Dim intNum As Variant
Dim strLetraFin As String

Dim txtnum As String
Dim i As Integer

mvarNumero = numstr
mvarNumero = Abs(mvarNumero)
numstr = CStr(Fix(mvarNumero))
Numero = mvarNumero

intNum = numstr
mstrMoneda = "Dollars"

numstr = CDec(numstr)
txtLetra = ""

If numstr = 0 Then
  numtoword = "zero "
  Exit Function
End If

If numstr > 10 ^ 24 Then
  numtoword = "Too big"
  Exit Function
End If

If numstr >= 10 ^ 12 Then
  newstr = numtoword(Int(numstr / 10 ^ 12))
  numstr = ((numstr / 10 ^ 12) - _
  Int(numstr / 10 ^ 12)) * 10 ^ 12
  If numstr = 0 Then
    tempstr = tempstr & newstr & "Billion "
  Else
    tempstr = tempstr & newstr & "Billion, "
  End If
End If

If numstr >= 10 ^ 6 Then
  newstr = numtoword(Int(numstr / 10 ^ 6))
  numstr = ((numstr / 10 ^ 6) - _
  Int(numstr / 10 ^ 6)) * 10 ^ 6
  If numstr = 0 Then
    tempstr = tempstr & newstr & "Million "
  Else
    tempstr = tempstr & newstr & "Million, "
  End If
End If

If numstr >= 10 ^ 3 Then
  newstr = numtoword(Int(numstr / 10 ^ 3))
  numstr = ((numstr / 10 ^ 3) - _
  Int(numstr / 10 ^ 3)) * 10 ^ 3
  If numstr = 0 Then
    tempstr = tempstr & newstr & "Thousand "
  Else
    tempstr = tempstr & newstr & "Thousand, "
  End If
End If

If numstr >= 10 ^ 2 Then
  newstr = numtoword(Int(numstr / 10 ^ 2))
  numstr = ((numstr / 10 ^ 2) - _
  Int(numstr / 10 ^ 2)) * 10 ^ 2
  If numstr = 0 Then
    tempstr = tempstr & newstr & "Hundred "
  Else
    tempstr = tempstr & newstr & "Hundred and "
  End If
End If

If numstr >= 20 Then
  Select Case Int(numstr / 10)
  Case 2
  tempstr = tempstr & "Twenty "
  Case 3
  tempstr = tempstr & "Thirty "
  Case 4
  tempstr = tempstr & "Forty "
  Case 5
  tempstr = tempstr & "Fifty "
  Case 6
  tempstr = tempstr & "Sixty "
  Case 7
  tempstr = tempstr & "Seventy "
  Case 8
  tempstr = tempstr & "Eighty "
  Case 9
  tempstr = tempstr & "Ninety "
  End Select
  numstr = ((numstr / 10) - _
  Int(numstr / 10)) * 10
End If

If numstr > 0 Then
  Select Case numstr
  Case 1
  tempstr = tempstr & "One "
  Case 2
  tempstr = tempstr & "Two "
  Case 3
  tempstr = tempstr & "Three "
  Case 4
  tempstr = tempstr & "Four "
  Case 5
  tempstr = tempstr & "Five "
  Case 6
  tempstr = tempstr & "Six "
  Case 7
  tempstr = tempstr & "Seven "
  Case 8
  tempstr = tempstr & "Eight "
  Case 9
  tempstr = tempstr & "Nine "
  Case 10
  tempstr = tempstr & "Ten "
  Case 11
  tempstr = tempstr & "Eleven "
  Case 12
  tempstr = tempstr & "Twelve "
  Case 13
  tempstr = tempstr & "Thirteen "
  Case 14
  tempstr = tempstr & "Fourteen "
  Case 15
  tempstr = tempstr & "Fifteen "
  Case 16
  tempstr = tempstr & "Sixteen "
  Case 17
  tempstr = tempstr & "Seventeen "
  Case 18
  tempstr = tempstr & "Eighteen "
  Case 19
  tempstr = tempstr & "Nineteen "
  End Select
  numstr = ((numstr / 10) - Int(numstr / 10)) * 10
End If

numtoword = tempstr

txtLetra = numtoword
strLetraFin = txtLetra

If Numero <> Fix(Numero) Then
    NumeroEnt = Fix(Numero)
    LenNumero = Len(Numero)
    LenNumeroEnt = Len(NumeroEnt)
    Dife = LenNumero - LenNumeroEnt - 1
    If Dife = 3 Then
        Decimales = (Numero - Fix(Numero)) * 1000
        Call Redondeo(Decimales)
        Wi = Decimales
        Wi = Fix(Wi * 1000)
        Wi = Left(Wi, 3)
        strLetraFin = "** " & strLetraFin & mstrMoneda & Str(Wi) & "/1000 US. **"
            Else
        Decimales = (Numero - Fix(Numero)) * 100
        Call Redondeo(Decimales)
        Wi = Decimales
        Wi = Fix(Wi * 100)
        Wi = Left(Wi, 2)
        strLetraFin = "** " & strLetraFin & mstrMoneda & Str(Wi) & "/100 US. **"
    End If
   Else
     strLetraFin = "** " & strLetraFin & mstrMoneda & " 00/100 US. **"
End If


txtLetra = strLetraFin



End Function

