Attribute VB_Name = "Funciones"

Sub Valida_fecha(Fecha As String, Validate As String)

    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    Dia = Left$(Fecha, 2)
    Mes = Mid$(Fecha, 4, 2)
    Ano = Right$(Fecha, 4)
    
    Validate = "S"
    
    If Val(Dia) <= 0 Or Val(Dia) > 31 Then
        Validate = "N"
    End If
    
    If Val(Mes) <= 0 Or Val(Mes) > 12 Then
        Validate = "N"
    End If
    
    If Val(Ano) < 1900 Then
        Validate = "N"
    End If
    
    If Val(Ano) = 0 And Val(Dia$) = 0 And Val(Mes$) = 0 Then
        Validate = "N"
    End If
    
End Sub

Sub Valida_fecha1(Fecha As String, Validate As String)

    Dim Dia As String
    Dim Mes As String
    Dim Ano As String
    
    Dia = Left$(Fecha, 2)
    Mes = Mid$(Fecha, 4, 2)
    Ano = Right$(Fecha, 4)
    
    Validate = "S"
    
    If Val(Dia) <= 0 Or Val(Dia) > 31 Then
        Validate = "N"
    End If
    
    If Val(Mes) <= 0 Or Val(Mes) > 12 Then
        Validate = "N"
    End If
    
    If Val(Ano) < 1900 Then
        Validate = "N"
    End If
    
    Rem If Val(Ano) = 0 And Val(Dia$) = 0 And Val(Mes$) = 0 Then
    Rem     Validate = "S"
    Rem End If
    
End Sub


Sub Verifica_datos(Dato As String, Valores As String, Valida As String)

    Dim Largo1 As Single
    Dim Largo2 As Single
    Dim Cicla1 As Single
    Dim Cicla2 As Single
    
    Largo1 = Len(Dato)
    Largo2 = Len(Valores)
    Valida = "N"
    
    For Cicla2 = 1 To Largo2
        If Dato = Mid$(Valores, Cicla2, 1) Then
            Valida = "S"
            Exit For
        End If
    Next Cicla2

End Sub

Sub Convierte_datos(Dato As String, Text As String)

    Text = ""

    For T = 1 To Len(Dato)
        If Mid$(Dato, T, 1) = "." Then
            Text = Text + ","
                Else
            Text = Text + Mid$(Dato, T, 1)
        End If
    Next T

End Sub

Sub Convierte1_datos(Dato As String, Text As String)

    Text = ""

    For T = 1 To Len(Dato)
        If Mid$(Dato, T, 1) = "," Then
            Text = Text + "."
                Else
            Text = Text + Mid$(Dato, T, 1)
        End If
    Next T

End Sub

Sub Conver(Dato As String, Text As String)

    Text = ""

    For T = 1 To Len(Dato)
        If Mid$(Dato, T, 1) <> "_" Then
            Text = Text + Mid$(Dato, T, 1)
        End If
    Next T
    
    Dato = Text

End Sub

Sub Calcula_vencimiento(WFecha As String, Plazo As Integer, Wvenci As String)

    Dim Dg As Integer
    Dim Ano As Integer
    Dim WAno As String
    Dim Mes As Integer
    Dim WMes As String
    Dim Dia As Integer
    Dim WDia As String
    Dim Di As Integer
    Dim aa As Integer
    Dim Ds(20) As Integer
    
    Ds(1) = 31
    Ds(2) = 28
    Ds(3) = 31
    Ds(4) = 30
    Ds(5) = 31
    Ds(6) = 30
    Ds(7) = 31
    Ds(8) = 31
    Ds(9) = 30
    Ds(10) = 31
    Ds(11) = 30
    Ds(12) = 31
    
    Rem   DATA "0101","0105","2505", , ,"0907", ,"1210", ,"2512", , , , , ,

    Dg = 0
    WAno = Mid$(WFecha, 7, 4)
    Ano = Val(WAno)
    WMes = Mid$(WFecha, 4, 2)
    Mes = Val(WMes)
    WDia = Mid$(WFecha, 1, 2)
    Dia = Val(WDia)
    
    'CANTIDAD DE DIAS HASTA LA FECHA
    
    Dg = Dia + Plazo - 1
    
    Do
        For aa = Mes To 12
            If (Ano Mod 4 = 0) And Mes = 2 Then Ds(2) = 29 Else Ds(2) = 28
            If Dg <= Ds(aa) Then Exit Do
            Dg = Dg - Ds(aa)
        Next aa
        Ano = Ano + 1
        Mes = 1
    Loop

    Dia = Dg
    WDia$ = Right$("0" + Mid$(Str$(Dia), 2, Len(Str$(Dia)) - 1), 2)

    Mes = aa
    WMes = Right$("0" + Mid$(Str$(Mes), 2, Len(Str$(Mes)) - 1), 2)
    
    WAno = Right$("0" + Mid$(Str$(Ano), 2, Len(Str$(Ano)) - 1), 4)
    
    Wvenci = WDia + "/" + WMes + "/" + WAno

End Sub

Sub Redondeo(Importe As Double)
            
    Dim B As Double
    Dim B1 As Double
    Dim Valor As Double
    Dim Redondeo As Double
    Dim Redondeo1 As Double
    Dim Dife As Double
            
    B = Importe * 100
    B1 = Importe * 10000
    Valor = Int(B)
    Redondeo = Int(B1)
    Redondeo1 = Int(B) * 100
    Dife = Redondeo - Redondeo1
    If Dife >= 50 Then Valor = Valor + 1
    Importe = Valor / 100
            
End Sub

Function Pusing(Mask As String, Text As String)

         Dim T As Single
         Dim T1 As Single
         Dim Tx As String
         Dim T1x As String
         Dim Places As Single
         Dim Auxi As String
         
         Rem For T = 1 To Len(Text)
         Rem    If Mid$(Text, T, 1) <> "," Then
         Rem        Auxi = Auxi + Mid$(Text, T, 1)
         Rem    End If
         Rem Next T
         
         Auxi = ""
 
         For T = 1 To Len(Text)
             If Mid$(Text, T, 1) = "," Then
                 Auxi = Auxi + "."
                     Else
                 If Mid$(Text, T, 1) = "." Then
                     Auxi = Auxi + "."
                         Else
                     Auxi = Auxi + Mid$(Text, T, 1)
                 End If
            End If
         Next T
         
         Text = Auxi
         
         X# = Val(Text)
         Pusing = Text

         T = Len(Mask)
         If T > 24 Then Error 5
         Tx = Space$(T)

         Places = 0
         T = InStr(Mask, ".")
        If T Then Places = Len(Mask) - T

         T1x = Mid$(Str$(Int(Abs(X#) + _
          (0.5 / 10# ^ Places))), 2 - _
          (Abs(X#) < 1#)) + _
          Mid$(".", 2 + (Places > 0)) + _
          Right$(Str$(Int((Abs(X#) + _
          10# ^ Places) * 10# ^ Places + _
          0.5)), Places)

         If Left$(T1x, 1) = "," Then T1x = "0" + T1x
         If InStr(Mask, "$") Then T1x = "$" + T1x
         If Sgn(X#) = -1 Then T1x = "-" + T1x

         If InStr(Mask, "+") And _
                 Left$(T1x, 1) <> "-" Then _
                 T1x = "+" + T1x

         If InStr(Mask, ",") Then
                 T = InStr(T1x, ".")
                 If T = 0 Then T = Len(T1x) + 1

                 For T = T - 4 To 1 Step -3
                T1 = Asc(Mid$(T1x, T))
                If T1 > 47 And T1 < 58 Then _
                        T1x = Left$(T1x, T) + "," + _
                         Mid$(T1x, T + 1)
       Next
    End If

         If Len(T1x) > Len(Tx) Then
                 Tx = "%" + T1x
         Else
                 RSet Tx = T1x
    End If

         If InStr(Mask, "*") Then
       For T = 1 To Len(Tx)
           If Mid$(Tx, T, 1) = " " Then Mid$(Tx, T, 1) = "*"
       Next
    End If
    
    Pusing = Trim(Tx)
    
    Auxi = ""
    
    Auxi = ""
    
    For T = 1 To Len(Pusing)
         If Mid$(Pusing, T, 1) = "," Then
            Auxi = Auxi + ""
                Else
            If Mid$(Pusing, T, 1) = "." Then
                Auxi = Auxi + "."
                    Else
                Auxi = Auxi + Mid$(Pusing, T, 1)
             End If
        End If
    Next T
    
    If Abs(Val(Auxi)) < 1 Then
        Pusing = "0" + Auxi
            Else
        Pusing = Auxi
    End If

End Function

Function Mascara(Mask As String, Text As String)

         Dim T As Single
         Dim T1 As Single
         Dim Tx As String
         Dim T1x As String
         Dim Places As Single
         Dim Auxi As String
         
         Rem For T = 1 To Len(Text)
         Rem    If Mid$(Text, T, 1) <> "," Then
         Rem        Auxi = Auxi + Mid$(Text, T, 1)
         Rem    End If
         Rem Next T
         
         Auxi = ""
         For T = 1 To Len(Text)
             If Mid$(Text, T, 1) = "," Then
                 Auxi = Auxi + "."
                     Else
                 If Mid$(Text, T, 1) = "." Then
                     Auxi = Auxi + "."
                         Else
                     Auxi = Auxi + Mid$(Text, T, 1)
                 End If
            End If
         Next T
         
         Text = Auxi
         
         X# = Val(Text)
         Mascara = Text

         T = Len(Mask)
         If T > 24 Then Error 5
         Tx = Space$(T)

         Places = 0
         T = InStr(Mask, ".")
        If T Then Places = Len(Mask) - T

         T1x = Mid$(Str$(Int(Abs(X#) + _
          (0.5 / 10# ^ Places))), 2 - _
          (Abs(X#) < 1#)) + _
          Mid$(".", 2 + (Places > 0)) + _
          Right$(Str$(Int((Abs(X#) + _
          10# ^ Places) * 10# ^ Places + _
          0.5)), Places)

         If Left$(T1x, 1) = "," Then T1x = "0" + T1x
         If InStr(Mask, "$") Then T1x = "$" + T1x
         If Sgn(X#) = -1 Then T1x = "-" + T1x

         If InStr(Mask, "+") And _
                 Left$(T1x, 1) <> "-" Then _
                 T1x = "+" + T1x

         If InStr(Mask, ",") Then
                 T = InStr(T1x, ".")
                 If T = 0 Then T = Len(T1x) + 1

                 For T = T - 4 To 1 Step -3
                T1 = Asc(Mid$(T1x, T))
                If T1 > 47 And T1 < 58 Then _
                        T1x = Left$(T1x, T) + "," + _
                         Mid$(T1x, T + 1)
       Next
    End If

         If Len(T1x) > Len(Tx) Then
                 Tx = "%" + T1x
         Else
                 RSet Tx = T1x
    End If

         If InStr(Mask, "*") Then
       For T = 1 To Len(Tx)
           If Mid$(Tx, T, 1) = " " Then Mid$(Tx, T, 1) = "*"
       Next
    End If
    
    Mascara = Trim(Tx)
    
    Auxi = ""
    For T = 1 To Len(Mascara)
         If Mid$(Mascara, T, 1) = "," Then
            Auxi = Auxi + ""
                Else
            If Mid$(Mascara, T, 1) = "." Then
                Auxi = Auxi + "."
                    Else
                Auxi = Auxi + Mid$(Mascara, T, 1)
             End If
        End If
    Next T
    
    If Abs(Val(Auxi)) < 1 Then
        Auxi = "0" + Auxi
    End If
    
    Auxi = "_____________________" + Auxi
    largo = Len(Mask)
    Text = Right$(Auxi, largo)

End Function

Function Alinea(Mask As String, Text As String)

         Dim T As Single
         Dim T1 As Single
         Dim Tx As String
         Dim T1x As String
         Dim Places As Single
         Dim Auxi As String
         
         Rem For T = 1 To Len(Text)
         Rem    If Mid$(Text, T, 1) <> "," Then
         Rem        Auxi = Auxi + Mid$(Text, T, 1)
         Rem    End If
         Rem Next T
         
         Auxi = ""
         For T = 1 To Len(Text)
             If Mid$(Text, T, 1) = "," Then
                 Auxi = Auxi + "."
                     Else
                 If Mid$(Text, T, 1) = "." Then
                     Auxi = Auxi + "."
                         Else
                     Auxi = Auxi + Mid$(Text, T, 1)
                 End If
            End If
         Next T
         
         Text = Auxi
         
         X# = Val(Text)
         Alinea = Text

         T = Len(Mask)
         If T > 24 Then Error 5
         Tx = Space$(T)

         Places = 0
         T = InStr(Mask, ".")
        If T Then Places = Len(Mask) - T

         T1x = Mid$(Str$(Int(Abs(X#) + _
          (0.5 / 10# ^ Places))), 2 - _
          (Abs(X#) < 1#)) + _
          Mid$(".", 2 + (Places > 0)) + _
          Right$(Str$(Int((Abs(X#) + _
          10# ^ Places) * 10# ^ Places + _
          0.5)), Places)

         If Left$(T1x, 1) = "," Then T1x = "0" + T1x
         If InStr(Mask, "$") Then T1x = "$" + T1x
         If Sgn(X#) = -1 Then T1x = "-" + T1x

         If InStr(Mask, "+") And _
                 Left$(T1x, 1) <> "-" Then _
                 T1x = "+" + T1x

         If InStr(Mask, ",") Then
                 T = InStr(T1x, ".")
                 If T = 0 Then T = Len(T1x) + 1

                 For T = T - 4 To 1 Step -3
                T1 = Asc(Mid$(T1x, T))
                If T1 > 47 And T1 < 58 Then _
                        T1x = Left$(T1x, T) + "," + _
                         Mid$(T1x, T + 1)
       Next
    End If

         If Len(T1x) > Len(Tx) Then
                 Tx = "%" + T1x
         Else
                 RSet Tx = T1x
    End If

         If InStr(Mask, "*") Then
       For T = 1 To Len(Tx)
           If Mid$(Tx, T, 1) = " " Then Mid$(Tx, T, 1) = "*"
       Next
    End If
    
    Alinea = Trim(Tx)
    
    Auxi = ""
    For T = 1 To Len(Alinea)
         If Mid$(Alinea, T, 1) = "," Then
            Auxi = Auxi + ""
                Else
            If Mid$(Alinea, T, 1) = "." Then
                Auxi = Auxi + "."
                    Else
                Auxi = Auxi + Mid$(Alinea, T, 1)
             End If
        End If
    Next T
    
    If Abs(Val(Auxi)) < 1 Then
        If Val(Auxi) < 0 Then
            Auxi = "-0" + Mid$(Auxi, 2, 20)
                Else
            Auxi = "0" + Auxi
        End If
    End If
    
    Auxi = "                       " + Auxi
    largo = Len(Mask)
    Alinea = Right$(Auxi, largo)

End Function


Public Function numtoword(numstr As Variant) As String

    'The best data type to feed in is
    'Decimal, but it is up to you
    Dim tempstr As String
    Dim newstr As String
    Dim mstrMoneda As String
    Dim intNum As Variant
    Dim strLetraFin As String
    Dim Decimales As Double

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

    txtLetraDolares = strLetraFin
    
End Function



Sub Redondeo3(Importe As Double)
            
    Dim B As Double
    Dim B1 As Double
    Dim Valor As Double
    Dim Redondeo As Double
    Dim Redondeo1 As Double
    Dim Dife As Double
            
    B = Importe * 1000
    B1 = Importe * 100000
    Valor = Int(B)
    Redondeo = Int(B1)
    Redondeo1 = Int(B) * 100
    Dife = Redondeo - Redondeo1
    If Dife >= 50 Then Valor = Valor + 1
    Importe = Valor / 1000
            
End Sub

