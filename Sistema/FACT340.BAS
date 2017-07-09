Attribute VB_Name = "Module1"
'***************************************************************************
'*                                                                         *
'* Programa      : ABM de Recibos                                          *
'* Fecha         : 02.03.90                                                *
'* Nombre        : Fact340                                                 *
'*                                                                         *
'***************************************************************************

Abre:

        Rem $Include: 'Redondeo.fn'
        Rem $Include: 'Impredate.fn'
        Rem $Include: 'Revdate.fn'
        Rem $Include: 'Prnusing.fn'

        Dim Descuento.001$(7,3),WDescuento.001$(7,3)
        Dim Vendedor.001$(7),WVendedor.001$(7)
        Dim Cobrador.001$(7),WCobrador.001$(7)
        Dim Impre$(4), Calcday%(12), Campo$(10, 5), Vector$(4), VerIfica$(10)
        Dim DCampo$(10, 5)
        Dim Numero.050$(10),WNumero.050$(10)
        Dim Mes$(12), Comprobante$(99), Iva#(10), M$(10), NomIva$(10)
        Dim Empresas$(99), Sig%(99)
        Dim Elije$(200, 2)

        Lmsg% = 23
        Cmsg% = 6

        Sig%(1) = 1
        Sig%(2) = 1
        Sig%(3) = -1
        Sig%(99) = -1

        Data "01", "02", "03"
        For Ciclo% = 1 To 3
                Read VerIfica$(Ciclo%)
        Next Ciclo%

        Mes$(1) = "Enero"
        Mes$(2) = "Febrero"
        Mes$(3) = "Marzo"
        Mes$(4) = "Abril"
        Mes$(5) = "Mayo"
        Mes$(6) = "Junio"
        Mes$(7) = "Julio:"
        Mes$(8) = "Agosto:"
        Mes$(9) = "Septiembre"
        Mes$(10) = "Octubre"
        Mes$(11) = "Noviembre"
        Mes$(12) = "Diciembre"

        Data "Monotributo"
        Data "Inscripto       "
        Data "No Inscripto    "
        Data "Exento          "
        Data "No Responsable  "

        For Ciclo% = 1 To 5
                Read NomIva$(Ciclo%)
        Next Ciclo%

        Empresas$(1) = "  MARILENE"
        Empresas$(2) = "  KLERICO"
        Empresas$(3) = "BIANCA'S CORSETERIA"
        Empresas$(4) = "BIANCA'S SECRET"
        Empresas$(5) = "PACO RABANNE"
        Empresas$(6) = "JAZMINE     "
        Empresas$(7) = "BIANCA'S MEDIAS"

        Comprobante$(1) = "Factura"
        Comprobante$(2) = "Nota debito"
        Comprobante$(3) = "Nota Credito"
        Comprobante$(99) = "Bonif."
        '
        GoSub Reseteo
        GoSub Erro
        '
        Open "I",#99,"Trab.seq"
        Input #99, Empresa$
        Input #99, NomEmpresa$
        Input #99, Mes$
        Input #99, Ano$
        Close #99
        '
        GoSub FCli001O
        If St.001% <> 0 Then
                Status% = St.001%
                Gosub Btr.err
        End If
        '
        GoSub FCta030O
        If St.030% <> 0 Then
                Status% = St.030%
                Gosub Btr.err
        End If
        '
        GoSub FDoc035O
        If St.035% <> 0 Then
                Status% = St.035%
                Gosub Btr.err
        End If
        '
        GoSub FIva031O
        If St.031% <> 0 Then
                Status% = St.031%
                Gosub Btr.err
        End If
        '
        GoSub FPar050O
        If St.050% <> 0 Then
                Status% = St.050%
                Gosub Btr.err
        End If
        '
        GoSub FRec009O
        If St.009% <> 0 Then
                Status% = St.009%
                Gosub Btr.err
        End If
        '
        GoSub FCam007O
        If St.007% <> 0 Then
                Status% = St.007%
                Gosub Btr.err
        End If

        Open "Lpt1" For Output As #99 Len = 255

        Open "r",#24,"Numero.dat",100
        Field #24,8 as WNumero$

        Width "LPT1:",230

TRONCO:
        WEmpresa$ = Space$(1)

        Do
                If WEmpresa$ = Space$(1) Then
                        LOCATE 1, 1
                        Print Chr$(255) + Chr$(255) + "Fact310e/"

                        Do
                                Call Ingreso(WEmpresa$, "N", "", 1, 0, 17, 49, 13, 0, "#", A$)
                        Loop Until Val(WEmpresa$) > 0 And Val(WEmpresa$) < 8
                End If

                Empresa% = Val(WEmpresa$)
                WEmpresa.030$ = WEmpresa$

                Cls
                Print Chr$(255) + Chr$(255) + "FACT340/"
                LOCATE 1, 1: Call QPrint("Empresa : " + Empresas$(Empresa%), 13)
                Gosub CLEAR.VARIABLES
                Gosub IMPRE.CABECERA
                LOCATE 23, 2: Print Space$(77)
                LOCATE 23, 2: Print "    F1 : Menu Principal   F2 : Reingreso de Datos  F3 : Campo Anterior "
                Intro% = 1
                CONFIRMA$ = "N"

                While CONFIRMA$ <> "S"
                        On Intro% Gosub  CLIENTE, NRORECIBO, FECHAREC,_
                                         PROVISORIO , IMPORTE, RetIva, RetGana, RetIb, CONFIRMA
                        Select Case Val(A$)
                                Case 1
                                        CONFIRMA$ = "S"

                                Case 2
                                        Gosub CLEAR.VARIABLES
                                        Gosub IMPRE.CABECERA
                                        Intro% = 1

                                Case 3
                                        If Intro% <> 2 Then
                                                Intro% = Intro% - 2
                                        End If

                                Case Else
                        End Select
                Wend
                If Val(A$) <> 1 Then
                        Gosub INGRESA.RENGLONES
                        If Val(A$) <> 10 Then
                                Gosub GRABA.DATOS
                                Gosub Pide.Documentos
                                Gosub GRABA.DOCUMENTOS
                                Imprime% = 2
                                For Ciclo% = 1 To 10
                                        If Val(Campo$(Ciclo%, 1)) = 3 Then
                                                Imprime% = 3
                                        End If
                                Next Ciclo%
                                For Counter% = 1 To 1
                                        Gosub Imprime.Recibo
                                Next Counter%
                        End If
                        A$ = ""
                End If
        Loop Until Val(A$) = 1
        GoTo FIN

PROVISORIO:

        Call Ingreso(PROVISORIO$, "N", "", 6, 0, 4, 37, 13, 0, "######", A$)
        Intro% = Intro% + 1
        Return

CLIENTE:

        Do
                LOCATE 23, 2: Print Space$(77)
                LOCATE 23, 2: Print "    F1 : Menu Principal   F2 : Reingreso de Datos  F3 : Campo Anterior "
                Call Ingreso(CLIENTE$, "N", "", 5, 0, 3, 37, 13, 0, "#####", A$)

                If Val(CLIENTE$) = 0 And Val(A$) = 0 Then
                        XDATO% = 1
                        GOSUB Selecciona.Datos
                        CLIENTE$ = XCodigo$
                        LOCATE 3, 37: Print Using; "#####"; Val(CLIENTE$)
                End If

                If Val(A$) = 0 Then
                        Op.001% = 5
                        Clave.001$ = Cliente$
                        GoSub FCli001R
                        If St.001% <> 0 Then
                                LOCATE 23, 2: Print Space$(77)
                                Status% = St.001%
                                Gosub Btr.err
                        End If
                End If
        Loop Until St.001% = 0 OR (Val(A$) < 4 And Val(A$) <> 0)
        If Val(A$) = 0 Then
                Locate 3,45 : Call Qprint (Left$(WRazon.001$,24),13)
                Locate 3,70 : Call Qprint (WVendedor.001$(Empresa%),13)
                If Val(WTipIva.001$) = 2 Or Val(WTipIva.001$) = 3 Then
                        Letra$ = "A"
                                Else
                        Letra$ = "B"
                End If
        End If
        Intro% = Intro% + 1
        Return

NRORECIBO:

   Do
        LOCATE 23, 2: Print Space$(77)
        LOCATE 23, 2: Print "    F1 : Menu Principal   F2 : Reingreso de Datos  F3 : Campo Anterior "
        LOCATE 1, 1

        Tipo$ = "C"

        Select Case Asc(Tipo$)
                Case 65, 88
                        WTipo$ = "05"

                Case 67
                        WTipo$ = "04"

                Case Else
        End Select

        'Op.050% = 5
        'Clave.050$ = "01"
        'Gosub FPar050R
        'If St.050% <> 0 Then
        '        Numero$ = "1"
        '                Else
        '        Numero$ = Str$( Val(WNumero.050$(4)) + 1 )
        'End IF

        If Letra$ = "B" Then
                LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Fill Page 1/"
                LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Letra/"
                play "l24o3ao4c#e"
                Tecla$ = Input$(1)
                LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Display Page 1/"
        End If

        If Letra$ <> "B" Then
                Get #24,2
                Numero$ = Str$(Val(WNUmero$) + 1)
                        Else
                Get #24,20
                Numero$ = Str$(Val(WNUmero$) + 1)
        End If

        Call Ingreso(Numero$, "N", "", 6, 0, 3, 14, 13, 0, "######", A$)

        If Val(A$) = 1 Or Val(A$) = 2 Or Val(A$) = 3 Then Exit Do
        '
        Op.030% = 5
        Clave.030$ = cLIENTE$ + WTipo$ + Numero$
        Kyn.030% = 0
        GoSub FCta030R
        If St.030% <> 4 Then
                If St.030% <> 0 Then
                        Status% = St.030%
                        Gosub Btr.err
                                Else
                        GoTo Baja
                End If
        End If
   Loop Until St.030% = 4
   Intro% = Intro% + 1
   Return

FECHAREC:

        If Val(FECHAREC$) = 0 Then FECHAREC$ = Mid$(Date$, 4, 2) + Left$(Date$, 2) + Right$(Date$, 2)
        Do
                Call Ingreso(FECHAREC$, "Y", "", 6, 0, 4, 14, 13, 0, "", A$)
                Call Validate(FECHAREC$, Mensaje$)
        Loop Until Mensaje$ = "Yes" Or Val(A$) = 1
        If Val(A$) = 1 Then Return
        Op.007% = 5
        Clave.007$ = FnRevdate$(Fecharec$)
        GoSub FCam007R
        If St.007% <> 0 Then
                Status% = St.007%
                Gosub Btr.err
                GoTo FECHAREC
        End If
        Intro% = Intro% + 1
        Return

IMPORTE:

        Do
                IMPORTE$ = Str$(Val(IMPORTE$) * 100)
                Call Ingreso(IMPORTE$, "N", "", 11, 2, 4, 61, 13, 0, "###,###,###.##", A$)
                IMPORTE$ = Str$(Val(IMPORTE$) / 100)
                If Tipo$ <> "X" Then WSaldo.030$ = ""
        Loop Until Tipo$ <> "X" OR (Tipo$ = "X" And Val(Importe$) <= Val(WSaldo.030$))
        Intro% = Intro% + 1
        Return

RetIva:

        RetIva$ = Str$(Val(RetIva$) * 100)
        Call Ingreso(RetIva$, "N", "", 11, 2, 5, 13, 13, 0, "###,###,###.##", A$)
        RetIva$ = Str$(Val(RetIva$) / 100)
        Intro% = Intro% + 1
        Return

RetGana:

        RetGana$ = Str$(Val(RetGana$) * 100)
        Call Ingreso(RetGana$, "N", "", 11, 2, 5, 36, 13, 0, "###,###,###.##", A$)
        RetGana$ = Str$(Val(RetGana$) / 100)
        Intro% = Intro% + 1
        Return

RetIb:

        RetIb$ = Str$(Val(RetIb$) * 100)
        Call Ingreso(RetIb$, "N", "", 11, 2, 5, 61, 13, 0, "###,###,###.##", A$)
        RetIb$ = Str$(Val(RetIb$) / 100)
        WImporte# = FnRedondeo#(Val(RetGana$) + Val(RetIva$) + Val(RetIb$) + Val(IMPORTE$))
        Intro% = Intro% + 1
        Return

CONFIRMA:

        LOCATE 23, 2: Print Space$(77)
        LOCATE 23, 2: Print "                          Confirma  S/N  "
        Do
                CONFIRMA$ = ""
                Call Ingreso(CONFIRMA$, "G", "", 1, 0, 23, 50, 13, 0, "", A$)
        Loop Until CONFIRMA$ = "S" Or CONFIRMA$ = "N"
        LOCATE 23, 2: Print Space$(77)
        Intro% = 1
        WIntro1% = 0
        Return

Clear.VARIABLES:

        Erase Campo$
        Intro% = 1
        WIntro% = 1
        Fechapro$ = "000000"
        Partida$ = ""
        CLIENTE$ = Space$(11)
        Numero$ = Space$(6)
        WTipo$ = Space$(2)
        Tipo$ = Space$(1)
        FECHAREC$ = "000000"
        IMPORTE$ = "0"
        RetIva$ = ""
        RetIb$ = ""
        RetGana$ = ""
        PROVISORIO$ = ""
        EFECTIVO$ = ""
        CHEQUE$ = ""
        Return

Impre.CABECERA:

        LOCATE 23, 2: Print Space$(77)
        LOCATE 3, 14: Call QPrint(FnPusing$("######", Val(Numero$)), 13)
        Op.001% = 5
        Clave.001$ = Cliente$
        GoSub FCli001R
        If St.001% <> 0 Then
                WRazon.001$ = ""
                WVendedor.001$(Empresa%) = ""
        End If
        LOCATE 3, 37: Call QPrint(FnPusing$("#####", Val(CLIENTE$)), 13)
        Locate 3,45 : Call Qprint (Left$(WRazon.001$,24),13)
        Locate 3,70 : Call Qprint (Left$(WVendedor.001$(Empresa%),4),13)
        Impresion$ = Left$(FECHAREC$, 2) + "/" + Mid$(FECHAREC$, 3, 2) + "/" + Right$(FECHAREC$, 2)
        LOCATE 4, 14: Call QPrint(Impresion$, 13)
        LOCATE 4, 37: Call QPrint(FnPusing$("######", Val(PROVISORIO$)), 13)
        LOCATE 4, 61: Call QPrint(FnPusing$("###,###,###.##", Val(IMPORTE$)), 13)
        LOCATE 5, 13: Call QPrint(FnPusing$("###,###,###.##", Val(RetIva$)), 13)
        LOCATE 5, 36: Call QPrint(FnPusing$("###,###,###.##", Val(RetGana$)), 13)
        LOCATE 5, 61: Call QPrint(FnPusing$("###,###,###.##", Val(RetIb$)), 13)
        Return

Ingresa.RENGLONES:

        Select Case Asc(Tipo$)
                Case 65
                        LOCATE 1, 1
                        Print Chr$(255) + Chr$(255) + "FILL PAGE 1/"
                        Print Chr$(255) + Chr$(255) + "FACT340C/"
                        CONFIRMA$ = ""
                        WIntro1% = 1
                        Do
                                CONFIRMA$ = ""
                                On WIntro1% GoSub EFECTIVO, CHEQUE, WCONFIRMA
                                WIntro1% = WIntro1% + 1
                                Total# = Val(EFECTIVO$) + Val(CHEQUE$)
                        Loop Until (CONFIRMA$ = "S" And Total# = Val(IMPORTE$)) Or Val(A$) = 10
                        LOCATE 1, 1
                        Print Chr$(255) + Chr$(255) + "DISPlay PAGE 1/"

                Case 67, 88
                        Fila% = 1
                        Columna% = 1
                        Do
                                WColor% = 13
                                Gosub IMPRE.LINEA
                                Lugar% = 9 + Fila%
                                Do
                                        GoSub CAMPO1
                                Loop Until Val(A$) = 1 Or Val(A$) = 11 Or _
                                           Val(A$) = 12 Or Val(Campo$(Fila%, Columna%)) <> 0 Or _
                                           Val(A$) = 10 Or Val(A$) = 2
                                Select Case Val(A$)
                                        Case 0
                                                WIntro1% = 2
                                                Gosub INGRESA.LINEA
                                                Total# = 0
                                                For WCiclo% = 1 To 10
                                                        If Val(Campo$(WCiclo%, 1)) = 3 Then
                                                                Total# = Total# - Val(Campo$(WCiclo%, 4))
                                                                                      Else
                                                                Total# = Total# + Val(Campo$(WCiclo%, 4))
                                                        End If
                                                Next WCiclo%
                                                WColor% = 7
                                                Gosub IMPRE.LINEA
                                                If Val(A$) = 12 Then
                                                        Fila% = Fila% - 1
                                                                Else
                                                        Fila% = Fila% + 1
                                                End If
                                                LOCATE 1, 1

                                        Case 11
                                                If Fila% < 10 Then
                                                        If Val(Campo$(Fila%, 1)) <> 0 Or Fila% = 9 Then
                                                                WColor% = 7
                                                                Gosub IMPRE.LINEA
                                                                Fila% = Fila% + 1
                                                        End If
                                                End If

                                        Case 12
                                                If Fila% > 1 Then
                                                        WColor% = 7
                                                        Gosub IMPRE.LINEA
                                                        Fila% = Fila% - 1
                                                End If

                                        Case 2
                                                Gosub Ingresa.Comprobante

                                        Case Else

                                End Select

                                Total# = FnRedondeo#(Total#)
                                WImporte# = FnRedondeo#(WImporte#)

                        Loop Until (Val(A$) = 1 And Total# = WImporte#) Or Val(A$) = 10
                Case Else
        End Select
        Return

Ingresa.LINEA:

        CONFIRMA$ = ""
        Do
                On WIntro1% GoSub CAMPO1, CAMPO2, CAMPO4, CAMPO3
                WIntro1% = WIntro1% + 1
        Loop Until WIntro1% = 5 Or Val(A$) = 1  Or Val(A$) = 10_
                                Or Val(A$) = 11 Or Val(A$) = 12
        Return

CAMPO1:

        Do
                LOCATE 23, 2: Print Space$(77)
                LOCATE 23, 2: Print "    F1 : Graba Recibo     F2 : Alta N/C o N/D     F10 : Anula Recibo "
                Salida% = 0
                Lugar% = Fila% + 9
                Call Ingreso(Campo$(Fila%, 1), "N", "", 2, 0, Lugar%, 14, 13, 0, "##", A$)
                For XCiclo% = 1 To 3
                        If VerIfica$(XCiclo%) = Campo$(Fila%, 1) Then
                                Salida% = 1
                                Exit For
                        End If
                Next XCiclo%
        Loop Until Salida% = 1 Or Val(A$) = 1 Or Val(A$) = 11 Or Val(A$) = 12 Or _
                                  Val(A$) = 10 Or Val(A$) = 2
        Return

CAMPO2:

        Lugar% = Fila% + 9
        Call Ingreso(Campo$(Fila%, 2), "N", "", 6, 0, Lugar%, 27, 13, 0, "######", A$)

        Salida% = 0
        hasta% = 10
        Op.030% = 5
        KYN.030% = 0
        Clave.030$ = Cliente$ + Campo$(Fila%,1%) + Campo$(Fila%,2)
        GoSub FCta030R
        If St.030% <> 0 Then
                LOCATE 23, 2: Print Space$(77)
                Status% = St.030%
                Gosub Btr.err
                Salida% = 1
                hasta% = 0
                WIntro1% = 0
                        Else
                If WCliente.030$ <> Cliente$ Then
                        LOCATE 23, 2: Print Space$(77)
                        LOCATE 23, 2: Print "  El cliente del comprobante no coincide con el de la cabecera"
                        Salida% = 1
                        hasta% = 0
                        WIntro1% = 0
                        GoSub Uniq
                        LOCATE 23, 2: Print Space$(77)
                End If
        End If

        For VerIfica% = 1 To hasta%
                If VerIfica% <> Fila% Then
                        If Campo$(VerIfica%, 1) + Campo$(VerIfica%, 2) = Campo$(Fila%, 1) + Campo$(Fila%, 2) Then
                                    LOCATE 23, 2: Print Space$(77)
                                    LOCATE 23, 2: Print "  El comprobante ya se encuentra dada de alta   "
                                    GoSub Uniq
                                    LOCATE 23, 2: Print Space$(77)
                                    Salida% = 1
                                    Exit For
                        End If
                End If
        Next VerIfica%

        Return

CAMPO4:

        If Campo$(Fila%, 5) = Space$(1) Or Campo$(Fila%, 5) = "" Then
                Campo$(Fila%, 5) = "T"
        End If

        Do
                Call Ingreso(Campo$(Fila%, 5), "G", "", 1, 0, Lugar%, 40, 13, 0, "", A$)
        Loop Until Campo$(Fila%, 5) = "T" Or Campo$(Fila%, 5) = "P"
        Return

CAMPO3:

        Lugar% = Fila% + 9
        WPartida$ = WPartida.030$
        If Campo$(Fila%, 5) = "T" Then
                Campo$(Fila%,3) = WSaldo.030$
                Select Case WPartida.030$
                        Case "S", "M", "E", "V"
                                If fnredondeo#(Val(WSaldo.030$)) <> fnredondeo#(Val(WImporte.030$)) Then
                                        If Val(WTipiva.001$) = 3 Then
                                                Auxiliar# = FnRedondeo#(Val(WSaldo.030$) / 1.315)
                                                Campo$(Fila%,4) = Str$( Val(WSaldo.030$) + Auxiliar# )
                                                        Else
                                                Auxiliar# = FnRedondeo#(Val(WSaldo.030$) / 1.21)
                                                Campo$(Fila%,4) = Str$( Val(WSaldo.030$) + Auxiliar# )
                                        End If
                                                         Else
                                        'Op.031% = 5
                                        'Clave.031$ = WLetra.030$ + WTipo.030$ + WNumero.030$
                                        'Gosub FIva031R
                                        'If St.031% = 0 Then
                                        '        XNeto$ = WNeto.031$
                                        '                else
                                        '        XNeto$ = WNeto.030$
                                        'End If
                                        Campo$(Fila%,4) = Str$( Val(WSaldo.030$) + Val(WNeto.030$) )
                                End If
                        Case "Z", "W"
                                If fnredondeo#(Val(WSaldo.030$)) <> fnredondeo#(Val(WImporte.030$)) Then
                                        If Val(WTipiva.001$) = 3 Then
                                                Auxiliar# = FnRedondeo#(Val(WSaldo.030$) / 1.315)
                                                Campo$(Fila%,4) = Str$( Val(WSaldo.030$) + Auxiliar# +  Auxiliar# )
                                                        Else
                                                Auxiliar# = FnRedondeo#(Val(WSaldo.030$) / 1.21)
                                                Campo$(Fila%,4) = Str$( Val(WSaldo.030$) + Auxiliar# +  Auxiliar# )
                                        End If
                                                         Else
                                        'Op.031% = 5
                                        'Clave.031$ = WLetra.030$ + WTipo.030$ + WNumero.030$
                                        'Gosub FIva031R
                                        'If St.031% = 0 Then
                                        '        XNeto$ = WNeto.031$
                                        '                else
                                        '        XNeto$ = WNeto.030$
                                        'End If
                                        Campo$(Fila%,4) = Str$( Val(WSaldo.030$) + Val(WNeto.030$) + Val(WNeto.030$) )
                                End If
                        Case Else
                                Campo$(Fila%,4) = WSaldo.030$
                End Select
                LOCATE Lugar%, 46: Print Using; "###,###,###.##"; Val(Campo$(Fila%, 4))
                LOCATE Lugar%, 64: Print Using; "###,###,###.##"; Val(Campo$(Fila%, 3))
                Return
        End If
        Campo$(Fila%, 4) = Str$(Val(Campo$(Fila%, 4)) * 100)
        Call Ingreso(Campo$(Fila%, 4), "N", "", 11, 2, Lugar%, 46, 13, 0, "###,###,###.##", A$)
        Campo$(Fila%, 4) = Str$(Val(Campo$(Fila%, 4)) / 100)
        Select Case WPartida.030$
                Case "S", "M", "E", "V"
                        If Val(WTipIva.001$) = 3 Then
                                IMPORTE# = FnRedondeo#(Val(Campo$(Fila%, 4)) * 0.568035)
                                        Else
                                IMPORTE# = FnRedondeo#(Val(Campo$(Fila%, 4)) * 0.547511)
                        End If
                Case "Z", "W"
                        If Val(WTipIva.001$) = 3 Then
                                IMPORTE# = FnRedondeo#(Val(Campo$(Fila%, 4)) * 0.396682)
                                        Else
                                IMPORTE# = FnRedondeo#(Val(Campo$(Fila%, 4)) * 0.376947)
                        End If
                Case Else
                        IMPORTE# = Val(Campo$(Fila%, 4))
        End Select
        Compara# = FnRedondeo#(Val(WSaldo.030$))
        IMPORTE# = FnRedondeo#(IMPORTE#)
        If IMPORTE# > Compara# Then
                LOCATE 23, 2: Print Space$(77)
                LOCATE 23, 2: Print " El Importe es mayor al saldo del comprobante"
                GoSub Uniq
                GoTo CAMPO3
        End If
        Campo$(Fila%, 3) = Str$(IMPORTE#)
        LOCATE Lugar%, 64: Print Using; "###,###,###.##"; Val(Campo$(Fila%, 3))
        Return

EFECTIVO:

        EFECTIVO$ = Str$(Val(EFECTIVO$) * 100)
        Call Ingreso(EFECTIVO$, "N", "", 11, 2, 10, 43, 13, 0, "###,###,###.##", A$)
        EFECTIVO$ = Str$(Val(EFECTIVO$) / 100)
        Return

CHEQUE:

        CHEQUE$ = Str$(Val(CHEQUE$) * 100)
        Call Ingreso(CHEQUE$, "N", "", 11, 2, 12, 43, 13, 0, "###,###,###.##", A$)
        CHEQUE$ = Str$(Val(CHEQUE$) / 100)
        Return


WCONFIRMA:

        Do
                CONFIRMA$ = ""
                Call Ingreso(CONFIRMA$, "G", "", 1, 0, 16, 43, 13, 0, "", A$)
        Loop Until CONFIRMA$ = "S" Or CONFIRMA$ = "N"
        WIntro1% = 0
        Return

LIMPIA.LINEA:

        For Ciclo% = 1 To 5
                Campo$(Fila%, Ciclo%) = ""
        Next Ciclo%
        Return

Impre.LINEA:

        LOCATE Fila% + 9, 14: Call QPrint(FnPusing$("##", Val(Campo$(Fila%, 1))), WColor%)
        LOCATE Fila% + 9, 27: Call QPrint(FnPusing$("######", Val(Campo$(Fila%, 2))), WColor%)
        LOCATE Fila% + 9, 40: Call QPrint(Campo$(Fila%, 5), WColor%)
        If Val(Campo$(Fila%, 3)) <> 0 Then
                If Val(Campo$(Fila%, 1)) = 3 Then
                        LOCATE Fila% + 9, 79
                        Call QPrint("-", WColor%)
                                            Else
                        LOCATE Fila% + 9, 79
                        Call QPrint("+", WColor%)
                End If
                LOCATE Fila% + 9, 46
                Call QPrint(FnPusing$("###,###,###.##", Val(Campo$(Fila%, 4))), WColor%)
                LOCATE Fila% + 9, 64
                Call QPrint(FnPusing$("###,###,###.##", Val(Campo$(Fila%, 3))), WColor%)
                                            Else
                LOCATE Fila% + 9, 46
                Call QPrint(Space$(14), WColor%)
                LOCATE Fila% + 9, 64
                Call QPrint(Space$(14), WColor%)
        End If
        Return

GRABA.Datos:

        If Letra$ <> "B" Then
                LSet WNUmero$ = Numero$
                Put #24,2
                        Else
                LSet WNUmero$ = Numero$
                Put #24,20
        End If

        WCliente.030$    = Cliente$
        WTipo.030$       = WTipo$
        WNumero.030$     = Numero$
        WFecha.030$      = FnRevdate$(Fecharec$)
        Select Case WPartida.030$
                Case "S", "E", "M", "V"
                        If Val(WTipIva.001$) = 3 Then
                                WImporte.030$    = Str$( FnRedondeo#( WImporte# * 0.568035 ) )
                                        Else
                                WImporte.030$    = Str$( FnRedondeo#( WImporte# * 0.547511 ) )
                        End If
                Case "Z", "W"
                        If Val(WTipIva.001$) = 3 Then
                                WImporte.030$    = Str$( FnRedondeo#( WImporte# * 0.396682 ) )
                                        Else
                                WImporte.030$    = Str$( FnRedondeo#( WImporte# * 0.376947 ) )
                        End If
                Case Else
                        WImporte.030$    = Str$(WImporte#)
        End Select
        WSaldo.030$      = ""
        WMovimiento.030$ = "2"
        WCotizacion.030$ = WCotizacion.007$
        Clave.030$ = WCliente.030$ + WTipo.030$ + WNumero.030$
        Op.030% = 2
        GoSub FCta030W
        If St.030% <> 0 Then
                Status% = St.030%
                Gosub Btr.err
        End If
        Gosub RESTA.FACTURAS

        Return

RESTA.FACTURAS:

        Kyn.030% = 0
        For GRABA% = 1 To 10
                If Val(Campo$(GRABA%, 1)) <> 0 And _
                   Val(Campo$(GRABA%, 1)) <> 99 Then
                        Clave.030$ = Cliente$ + Campo$(GRABA%,1) + Campo$(GRABA%,2)
                        Op.030% = 5
                        GoSub FCta030R
                        If St.030% <> 0 Then
                                Status% = St.030%
                                Gosub Btr.err
                                            Else
                                WSaldo.030$ = Str$(FnRedondeo#(Val(WSaldo.030$)-Val(Campo$(GRABA%,3))))
                                If Val(WSaldo.030$) = 0 Then WEstado.030$ = "1"
                                Op.030% = 3
                                GoSub FCta030W
                                If St.030% <> 0 Then
                                        Status% = St.030%
                                        Gosub Btr.err
                                End If
                        End If
                End If
                If Val(Campo$(GRABA%, 1)) <> 0 Then
                        WRecibo.009$ = Numero$
                        WCliente.009$ = Cliente$
                        WProvisorio.009$ = Provisorio$
                        WFecha.009$   = Fecharec$
                        WTotal.009$   = Importe$
                        WRetIva.009$  = RetIva$
                        WRetGana.009$ = RetGana$
                        WRetIb.009$   = RetIb$
                        WTipo.009$    = Campo$(Graba%,1)
                        WNumero.009$  = Campo$(Graba%,2)
                        WForma.009$   = Campo$(Graba%,5)
                        WImporte.009$ = Campo$(Graba%,4)
                        WImprec.009$  = Campo$(Graba%,3)
                        WRenglon.009$ = Str$(Graba%)
                        Call Ceros(WRenglon.009$,2)
                        Clave.009$ = WRecibo.009$ + WRenglon.009$
                        Op.009% = 2
                        GoSub FRec009W
                        If St.009% <> 0 Then
                                Status% = St.009%
                                Gosub Btr.err
                        End If
                End If
        Next GRABA%
        Return

GRABA.Documentos:

        For GRABA% = 1 To 10
                If Val(DCampo$(GRABA%, 2)) <> 0 Then
                        WCliente.035$    = Cliente$
                        WNumero.035$     = DCampo$(Graba%,1)
                        WEmision.035$    = Fecharec$
                        WImporte.035$    = DCampo$(Graba%,2)
                        WSaldo.035$      = DCampo$(Graba%,2)
                        WMovimiento.035$ = "1"
                        Clave.035$ = Cliente$+DCampo$(GRABA%,1)
                        Op.035% = 2
                        GoSub FDoc035R
                        If St.035% <> 0 Then
                                Status% = St.030%
                                Gosub Btr.err
                        End If
                End If
        Next GRABA%
        Return
FIN:

        GoSub Reseteo
        Close
        CHAIN "FACT300"

Pide.Documentos:

        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "FILL PAGE 1/"
        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "FACT340F/"
        Fila% = 1
        Columna% = 1
        Do
                WColor% = 13
                Gosub IMPRE.LINEA.DOCUMENTO
                Lugar% = 9 + Fila%
                Do
                        Gosub CAMPO.DOCUMENTO.1
                Loop Until Val(A$) = 1  OR Val(A$) = 11 OR _
                           Val(A$) = 12 OR_
                           Val (DCampo$(Fila%, Columna%)) <> 0
                Select Case Val(A$)
                        Case 0
                                WIntro1% = 2
                                Gosub INGRESA.LINEA.DOCUMENTO
                                Total# = 0
                                For WCiclo% = 1 To 10
                                        Total# = Total# + Val(DCampo$(WCiclo%, 2))
                                Next WCiclo%
                                WColor% = 7
                                Gosub IMPRE.LINEA.DOCUMENTO
                                If Val(A$) = 12 Then
                                        Fila% = Fila% - 1
                                                Else
                                        Fila% = Fila% + 1
                                End If
                                LOCATE 1, 1

                        Case 11

                                If Fila% < 10 Then
                                        If Val(DCampo$(Fila%, 1)) <> 0 Or Fila% = 9 Then
                                                WColor% = 7
                                                Gosub IMPRE.LINEA.DOCUMENTO
                                                Fila% = Fila% + 1
                                        End If
                                End If

                        Case 12

                                If Fila% > 1 Then
                                        WColor% = 7
                                        Gosub IMPRE.LINEA.DOCUMENTO
                                        Fila% = Fila% - 1
                                End If

                         Case Else

                End Select

        Loop Until Val(A$) = 1 Or Val(A$) = 10
        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "DISPLAY PAGE 1/"
        Return

Ingresa.LINEA.DOCUMENTO:

        CONFIRMA$ = ""
        Do
                ON WIntro1% Gosub CAMPO.DOCUMENTO.1, CAMPO.DOCUMENTO.2, CONFIRMA
                WIntro1% = WIntro1% + 1
        Loop Until Confirma$ = "S" or Val(A$) = 1  Or Val(A$) = 10_
                                   Or Val(A$) = 11 Or Val(A$) = 12
        Return

Campo.DOCUMENTO 0.1:

        Lugar% = Fila% + 9
        Call Ingreso(DCampo$(Fila%, 1), "N", "", 6, 0, Lugar%, 27, 13, 0, "######", A$)
        '
        Salida% = 0
        hasta% = 10
        Op.035% = 5
        Clave.035$ = Cliente$+Campo$(Fila%,1)
        GoSub FDoc035R
        If St.035% = 0 Then
                LOCATE 23, 2: Print Space$(77)
                Status% = 5
                Gosub Btr.err
                Salida% = 1
                hasta% = 0
                WIntro1% = 0
        End If
        For VerIfica% = 1 To hasta%
                If VerIfica% <> Fila% Then
                        If DCampo$(VerIfica%, 1) = DCampo$(Fila%, 1) Then
                                    LOCATE 23, 2: Print Space$(77)
                                    LOCATE 23, 2: Print "  El comprobante ya se encuentra dada de alta   "
                                    GoSub Uniq
                                    LOCATE 23, 2: Print Space$(77)
                                    Salida% = 1
                                    Exit For
                        End If
                End If
        Next VerIfica%
        Return

Campo.DOCUMENTO 0.2:

        DCampo$(Fila%, 2) = Str$(Val(DCampo$(Fila%, 2)) * 100)
        Call Ingreso(DCampo$(Fila%, 2), "N", "", 11, 2, Lugar%, 42, 13, 0, "###,###,###.##", A$)
        DCampo$(Fila%, 2) = Str$(Val(DCampo$(Fila%, 2)) / 100)
        Return

LIMPIA.LINEA.DOCUMENTO:

        For Ciclo% = 1 To 2
                DCampo$(Fila%, Ciclo%) = ""
        Next Ciclo%
        Return

Impre.LINEA.DOCUMENTO:

        LOCATE Fila% + 9, 27: Call QPrint(FnPusing$("######", Val(DCampo$(Fila%, 1))), WColor%)
        If Val(DCampo$(Fila%, 2)) <> 0 Then
                LOCATE Fila% + 9, 42
                Call QPrint(FnPusing$("###,###,###.##", Val(DCampo$(Fila%, 2))), WColor%)
                                            Else
                LOCATE Fila% + 9, 42
                Call QPrint(Space$(14), WColor%)
        End If
        Return

Imprime.Recibo:

        Impresion# = 0
        For Ciclo% = 1 To 10
                If Val(Campo$(Ciclo%, 1)) <> 0 Then
                        If Val(Campo$(Ciclo%,1)) = 3 Or_
                           Val(Campo$(Ciclo%,1)) = 99 Then
                               Impresion# = Impresion# - Val(Campo$(Ciclo%, 3))
                                                     Else
                               Impresion# = Impresion# + Val(Campo$(Ciclo%, 3))
                        End If
                End If
        Next Ciclo%

        Impresion# = FnRedondeo#(Impresion#)

        Print #99, Tab(1); Chr$(18); Empresas$(Empresa%);
        Print #99, Tab(34); FnImpredate$(FnRevdate$(FECHAREC$));
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99,Tab(30);"Vendedor : ";Val(WVendedor.001$(Empresa%))
        Print #99, ""
        Print #99,Left$(WRazon.001$,20);
        Print #99,Tab(25);WDireccion.001$
        Print #99,Left$(WLocalidad.001$,20);
        Print #99,"C.P.:";WPostal.001$
        Print #99, ""
        Print #99, ""
        Print #99, ""

        If Letra$ = "A" Then
                Print #99,Tab(1);NomIva$(Val(WTipIva.001$));
                        Else
                Select Case Val(WTipIva.001$)
                        Case 1
                                Print #99, Tab(20); "X";
                        Case 4
                                Print #99, Tab(13); "X";
                        Case 5
                                Print #99, Tab(7); "X";
                        Case Else
                End Select
        End If

        Print #99,Tab(30);Left$(WCuit.001$,15);
        Print #99, Tab(47); CLIENTE$
        Print #99, ""
        Print #99, ""
        Print #99, ""
        M# = Impresion#
        GoSub 4630
        Print #99, Tab(1); "PESOS : "; M$(1)
        Print #99, Tab(1); M$(2)
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""

        WTotal# = 0

        For Ciclo% = 1 To 9
                If Val(Campo$(Ciclo%, 1)) <> 0 Then
                        Print #99,Using "\            \ ######  ###,###,###.##";_
                                  Comprobante$(Val(campo$(Ciclo%,1)));_
                                  Val(right$(Campo$(Ciclo%,2),5));Val(Campo$(Ciclo%,3));
                        If Val(Campo$(Ciclo%, 1)) = 1 Or Val(Campo$(Ciclo%, 1)) = 2 Then
                                WTotal# = WTotal# + Val(Campo$(Ciclo%, 3))
                                        Else
                                WTotal# = WTotal# - Val(Campo$(Ciclo%, 3))
                        End If
                End If
                Print #99, ""
        Next Ciclo%

        If Val(RetIva$) <> 0 Then
                Print #99, Using; "Ret. Iva                    ###,###,###.##"; Val(RetIva$)
                        Else
                Print #99, ""
        End If

        If Val(RetGana$) <> 0 Then
                Print #99, Using; "Ret.Ganancia                ###,###,###.##"; Val(RetGana$)
                        Else
                Print #99, ""
        End If

        If Val(RetIb$) <> 0 Then
                Print #99, Using; "Ret.Ingresos Brutos         ###,###,###.##"; Val(RetIb$)
                        Else
                Print #99, ""
        End If

        'Print #99,""

        Neto# = 0
        Iva#(1) = 0
        Iva#(2) = 0
        IMPORTE# = 0

        For Ciclo% = 1 To 10
                If Val(Campo$(Ciclo%, 1)) = 99 Then
                                Signo$ = " "
                                IMPORTE# = Val(Campo$(Ciclo%, 3))
                                Select Case Val(WTipIva.001$)
                                        Case 1, 2
                                                Neto# = FnRedondeo#(IMPORTE# / 1.21)
                                                Iva#(1) = FnRedondeo#(IMPORTE# - Neto#)
                                                Iva#(2) = 0
                                        Case 3
                                                Neto# = FnRedondeo#(IMPORTE# / 1.315)
                                                Iva#(1) = FnRedondeo#(Neto# * 0.21)
                                                Iva#(2) = FnRedondeo#(IMPORTE# - Iva#(1) - Neto#)
                                        Case 4
                                                Neto# = IMPORTE#
                                                Iva#(1) = 0
                                                Iva#(2) = 0
                                        Case Else
                                End Select
                End If
        Next Ciclo%

        If Letra$ = "A" Then
                Print #99, Tab(7); Using; "###,###,###.##"; IMPORTE# - Iva#(1) - Iva#(2);
                Print #99, Tab(26); Using; " $ ###,###,###.##"; WTotal#
                Print #99, ""
                Print #99, ""
                Print #99, Tab(7); Using; "###,###,###.##"; Iva#(1)
                Print #99, ""
                Print #99, ""
                Print #99, Tab(7); Using; "###,###,###.##"; Iva#(2)
                Print #99, ""
                Print #99, Tab(7); Using; "###,###,###.##"; IMPORTE#
                        Else
                Print #99, ""
                Print #99, Tab(26); Using; " $ ###,###,###.##"; WTotal#
                Print #99, ""
                Print #99, ""
                Print #99, Tab(7); Using; "###,###,###.##"; IMPORTE# - Iva#(1) - Iva#(2)
                Print #99, ""
                Print #99, ""
                Print #99, ""
                Print #99, ""
                Print #99, ""
        End If


        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Return

Ingresa.Comprobante:

        Diferencia# = FnRedondeo#(WImporte# - Total#)

        If Diferencia# >= 0 Or Val(Campo$(Fila%, 1)) <> 0 Then Return

        Campo$(Fila%, 1) = "99"
        Campo$(Fila%, 2) = Numero$

        Campo$(Fila%, 5) = "P"

        WColor% = 13
        Gosub Impre.Linea

        'Call Ingreso (Campo$(Fila%,2), "N", "", 6, 0, Lugar%, 27, 13, 0, "######", A$)

        'Op.030% = 5
        'Kyn.030% = 0
        'Clave.030$ = Cliente$ + Campo$(Fila%,1) + Campo$(Fila%,1)
        'Gosub FCta030R
        'If St.030% <> 4 Then
        '        Status% = St.030%
        '        Gosub Btr.err
        '        Goto Ingresa.Comprobante
        'End If

        Campo$(Fila%, 4) = Str$(Abs(Diferencia#))
        WSaldo.030$     = Campo$(Fila%,4)

        GoSub CAMPO3

        WCliente.030$     = Cliente$
        WTipo.030$        = Campo$(Fila%,1)
        WNumero.030$      = Campo$(Fila%,2)
        WPartida.030$     = WPartida$
        WEmision.030$     = Fecharec$
        WVendedor.030$    = WVendedor.001$(Empresa%)
        WImporte.030$     = Campo$(Fila%,3)
        WSaldo.030$       = ""
        WEstado.030$      = "1"
        WReal.0320$       = ""
        WNeto.030$        = ""
        WIva1.030$        = ""
        WIva2.030$        = ""
        WProvincia.030$   = WProvincia.001$
        WCondicion.030$   = ""
        WVencimiento.030$ = Fecharec$
        WDescuento.030$   = ""
        WImpuesto.030$    = ""
        WPrecio.030$      = ""
        WCotizacion.030$  = WCotizacion.007$
        WOrigen.030$      = "R"
        WMovimiento.030$  = "2"

        Op.030% = 2
        Kyn.030% = 0
        Clave.030$ = WCliente.030$ + WTipo.030$ + WNumero.030$
        GoSub FCta030W
        If St.030% <> 0 Then
                Status% = St.030%
                Gosub Btr.err
        End If
        WOrigen.030$    = ""

        WLetra.031$     = Letra$
        WTipo.031$      = Campo$(Fila%,1)
        WNumero.031$    = Campo$(Fila%,2)
        WCliente.031$   = Cliente$
        WPartida.031$   = WPartida$
        WEmision.031$   = Fecharec$
        WVendedor.031$  = WVendedor.001$(Empresa%)
        WProvincia.031$ = WProvincia.001$

        Diferencia# = Val(Campo$(Fila%, 3))

        Select Case Val(WTipIva.001$)
                Case 1, 2
                        Neto# = FnRedondeo#(Diferencia# / 1.21)
                        Iva#(1) = FnRedondeo#(Diferencia# - Neto#)
                        Iva#(2) = 0
                Case 3
                        Neto# = FnRedondeo#(Diferencia# / 1.315)
                        Iva#(1) = FnRedondeo#(Neto# * 0.21)
                        Iva#(2) = FnRedondeo#(Diferencia# - Iva#(1) - Neto#)
                Case 4
                        Neto# = Diferencia#
                        Iva#(1) = 0
                        Iva#(2) = 0
                Case Else
        End Select

        Opcion% = Val(Campo$(Fila%, 1))
        WNeto.031$      = Str$( Neto# * Sig%(Opcion%) )
        WIva1.031$      = Str$( Iva#(1) * Sig%(Opcion%) )
        WIva2.031$      = Str$( Iva#(2) * Sig%(Opcion%) )

        Op.031% = 2
        Clave.031$ = Letra$ + WTipo.031$+WNumero.031$
        GoSub FIva031w
        If St.031%<> 0 Then
                Status% = St.031%
                Gosub Btr.err
        End If

        'Op.050% = 5
        'Clave.050$ = "01"
        'Gosub FPar050R
        'If St.050% = 0 Then
        '        If Val(Campo$(Fila%,1)) = 2 Then
        '                WNumero.050$(2) = Str$( Val(Campo$(Fila%,2)) + 1 )
        '                                Else
        '                WNumero.050$(3) = Str$( Val(Campo$(Fila%,2)) + 1 )
        '        End If
        '        Op.050% = 3
        '        Gosub FPar050W
        '        If St.050% <> 0 Then
        '                Status% = St.050%
        '                Gosub Btr.err
        '        End If
        'End If

        Total# = 0
        For WCiclo% = 1 To 10
                If Val(Campo$(WCiclo%,1)) = 3 OR_
                   Val(Campo$(WCiclo%,1)) = 99 Then
                        Total# = Total# - Val(Campo$(WCiclo%, 4))
                                              Else
                        Total# = Total# + Val(Campo$(WCiclo%, 4))
                End If
        Next WCiclo%
        WColor% = 7
        Gosub IMPRE.LINEA
        Fila% = Fila% + 1
        If Fila% > 10 Then
                Fila% = 10
        End If
        LOCATE 1, 1
        Return

Baja:
        Cantidad% = 0
        Clave.009$ = Numero$ + SPace$(2)
        Op.009% = 9
        GoSub FRec009R

        While St.009% = 0 And Numero$ = WRecibo.009$
                Cliente$ = WCliente.009$
                Fecharec$ = WFecha.009$
                Provisorio$ = WProvisorio.009$
                Importe$ = WTotal.009$
                RetIva$ = WRetIva.009$
                RetGana$ = WRetGana.009$
                RetIb$ = WRetIb.009$
                Cantidad% = Cantidad% + 1
                Campo$(Cantidad%,1) = WTipo.009$
                Campo$(Cantidad%,2) = WNUmero.009$
                Campo$(Cantidad%,5) = WForma.009$
                Campo$(Cantidad%,4) = WImporte.009$
                Campo$(Cantidad%,3) = WImprec.009$
                Op.009% = 6
                GoSub FRec009R
        Wend

        Gosub Impre.Cabecera
        For Fila% = 1 To 10
                If Val(Campo$(Fila%, 1)) <> 0 Then
                        WColor% = 13
                        Gosub Impre.Linea
                End If
        Next Fila%

        LOCATE 23, 2: Print Space$(77)
        LOCATE 23, 2: Print "                     Confirma la Baja S/N  "
        Do
                CONFIRMA$ = ""
                Call Ingreso(CONFIRMA$, "G", "", 1, 0, 23, 50, 13, 0, "", A$)
        Loop Until CONFIRMA$ = "S" Or CONFIRMA$ = "N" Or CONFIRMA$ = "R"
        LOCATE 23, 2: Print Space$(77)

        If CONFIRMA$ = "S" Then

                Kyn.030% = 0
                Clave.030$    = Cliente$+"04"+Numero$
                Op.030% = 5
                GoSub FCta030R
                If St.030% <> 0 Then
                        Status% = St.030%
                        Gosub Btr.err
                                Else
                        Op.030% = 4
                        GoSub FCta030W
                        If St.030% <> 0 Then
                                Status% = St.030%
                                Gosub Btr.err
                        End If
                End If

                For GRABA% = 1 To 10
                        If Val(Campo$(GRABA%, 1)) <> 0 Then
                                Clave.030$ = Cliente$+Campo$(GRABA%,1)+Campo$(GRABA%,2)
                                Op.030% = 5
                                GoSub FCta030R
                                If St.030% <> 0 Then
                                        Status% = St.030%
                                        Gosub Btr.err
                                                Else
                                        If WOrigen.030$ = "R" Then

                                                Op.030% = 4
                                                GoSub FCta030W
                                                If St.030% <> 0 Then
                                                        Status% = St.030%
                                                        Gosub Btr.err
                                                End If

                                                Op.031% = 5
                                                Clave.031$ = Campo$(GRABA%,1)+Campo$(GRABA%,2)
                                                GoSub FIva031r
                                                If St.031%<> 0 Then
                                                        Status% = St.031%
                                                        Gosub Btr.err
                                                               Else
                                                        Op.031% = 4
                                                        GoSub FIva031w
                                                        If St.031%<> 0 Then
                                                                Status% = St.031%
                                                                Gosub Btr.err
                                                        End If
                                                End If

                                                              Else

                                                WSaldo.030$ = Str$(FnRedondeo#(Val(WSaldo.030$)+Val(Campo$(GRABA%,3))))
                                                If Val(WSaldo.030$) = 0 Then
                                                        WEstado.030$ = "1"
                                                                        Else
                                                        WEstado.030$ = "0"
                                                End If

                                                Op.030% = 3
                                                GoSub FCta030W
                                                If St.030% <> 0 Then
                                                        Status% = St.030%
                                                        Gosub Btr.err
                                                End If
                                        End If

                                        Renglon$ = Str$(GRABA%)
                                        Call Ceros(Renglon$, 2)
                                        Clave.009$ = Numero$ + Renglon$
                                        Op.009% = 5
                                        GoSub FRec009R
                                        If St.009% <> 0 Then
                                                Status% = St.009%
                                                Gosub Btr.err
                                                        Else
                                                Op.009% = 4
                                                GoSub FRec009W
                                                If St.009% <> 0 Then
                                                        Status% = St.009%
                                                        Gosub Btr.err
                                                End If
                                        End If
                                End If
                        End If
                Next GRABA%
        End If
        If CONFIRMA$ = "R" Then
                Gosub Imprime.Recibo
        End If
        GoTo TRONCO



Rem $INCLUDE: 'NUMTOLET'
Rem $INCLUDE: 'ERR'
Rem $INCLUDE: 'PRUPAN2'
Rem $INCLUDE: 'Validate.SUB'
Rem $INCLUDE: 'FRec009.FDE'
Rem $INCLUDE: 'FCli001.FDE'
Rem $INCLUDE: 'FCta030.FDE'
Rem $INCLUDE: 'FPar050.FDE'
Rem $INCLUDE: 'FIva031.FDE'
Rem $INCLUDE: 'FDoc035.FDE'
Rem $INCLUDE: 'FCam007.FDE'
Rem $INCLUDE: 'RESET'





Selecciona.Datos:

        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "Fill Page 2/"
        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "Ayuda/"

        LOCATE 9, 28: Print Space$(40)
        Nombre$ = ""
        Call Ingreso(Nombre$, "F", "", 40, 0, 9, 28, 13, 0, "", A$)

        For Ciclo% = 40 To 1 Step -1
                If Mid$(Nombre$, Ciclo%, 1) <> Space$(1) Then
                        Nombre$ = Left$(Nombre$, Ciclo%)
                        Exit For
                End If
        Next Ciclo%

        Erase Elije$
        Cantidad% = 0

        Select Case XDATO%
                Case 1
                        Kyn.001% = 1
                        Op.001% = 9
                        Clave.001$ = Nombre$ + SPACE$(40)
                        GoSub FCli001R

                        WHILE st.001% = 0 AND (Nombre$ = LEFT$(WRazon.001$, Ciclo%) OR Ciclo% = 0)

                                Cantidad% = Cantidad% + 1
                                Elije$(Cantidad%, 1) = WCliente.001$
                                Elije$(Cantidad%, 2) = LEFT$(WRazon.001$, 25)

                                Op.001% = 6
                                GoSub FCli001R

                        Wend
                        Kyn.001% = 0

                Case Else
        End Select

        GOSUB Elije.Datos

        Return

Elije.Datos:

        WFila% = 1
        WColumna% = 1
        WDesde% = 1

        GOSUB Impre.Pantalla.Seleccion

        Do
                WColor% = 13
                GOSUB IMPRE.LINEA.Seleccion
                Do
                        WIngreso$ = ""
                        Lugar% = (((WFila% - 1) - Int((WFila% - 1) / 11) * 11)) + 13
                        Call Ingreso(WIngreso$, "G", "", 1, 0, Lugar%, 13, 13, 0, "", A$)
                Loop Until Val(A$) = 0 Or Val(A$) = 10 Or Val(A$) = 12 Or Val(A$) = 11 Or Val(A$) = 2 Or Val(A$) = 3
                Select Case Val(A$)
                        Case 0
                                XCodigo$ = Elije$(WFila%, 1)
                        Case 11
                                If WFila% < 200 And Cantidad% > WFila% Then
                                        WColor% = 7
                                        GOSUB IMPRE.LINEA.Seleccion
                                        WFila% = WFila% + 1
                                        Compara% = (Int((WFila% - 1) / 11) * 11) + 1
                                        If Compara% = WFila% Then
                                                WDesde% = WFila%
                                                GOSUB Impre.Pantalla.Seleccion
                                        End If
                                End If

                        Case 12
                                If WFila% > 1 Then
                                                WColor% = 7
                                                GOSUB IMPRE.LINEA.Seleccion
                                                WFila% = WFila% - 1
                                                Compara% = (Int((WFila%) / 11) * 11)
                                                If Compara% = WFila% Then
                                                        WDesde% = WFila% - 10
                                                        GOSUB Impre.Pantalla.Seleccion
                                                End If
                                End If

                        Case Else

                End Select

        Loop Until Val(A$) = 0 Or Val(A$) = 10

        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "Display Page 2/"

        A$ = ""

        Return

Impre.LINEA.Seleccion:

        Lugar% = (((WFila% - 1) - Int((WFila% - 1) / 11) * 11)) + 13

        LOCATE Lugar%, 16: Print Elije$(WFila%, 1)
        LOCATE Lugar%, 30: Print Elije$(WFila%, 2)
        Return

Impre.Pantalla.Seleccion:

        WHasta% = WDesde% + 10
        If WHasta% > 200 Then WHasta% = 200
        For WCiclo% = WDesde% To WHasta%
                Lugar% = (((WCiclo% - 1) - Int((WCiclo% - 1) / 11) * 11)) + 13
                If Elije$(WCiclo%, 1) <> "" Then
                        LOCATE Lugar%, 16: Print Elije$(WCiclo%, 1)
                        LOCATE Lugar%, 30: Print Elije$(WCiclo%, 2)
                                                Else
                        LOCATE Lugar%, 16: Print Space$(11)
                        LOCATE Lugar%, 30: Print Space$(35)
                End If
        Next WCiclo%
        Return
