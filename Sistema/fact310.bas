Attribute VB_Name = "Module3"
'*****************************************************************************
'*                                                                           *
'* Programa  : Emision de Facturas                                           *
'*                                                                           *
'* Nombre    : FACT310                                                       *
'*                                                                           *
'*****************************************************************************

Abre:
     
     
        Dim Nomiva$(10), Nomprov$(99)
        Dim Empresas$(10), NomEmp$(10)
        Dim Descuento.001$(7,3),WDescuento.001$(7,3)
        Dim Vendedor.001$(7),WVendedor.001$(7)
        Dim Cobrador.001$(7),WCobrador.001$(7)
        Dim Calcday%(12), Ds(15), Fr$(15)
        Dim Vector$(40, 5), Iva#(10), Dto#(10)
        Dim WColor.010$(10),WTalle.010$(10)
        Dim Color.010$(10),Talle.010$(10)
        Dim Stock.009$(10,10)
        Dim WStock.009$(10,10)
        Dim Cantidad.011$(10,10)
        Dim WCantidad.011$(10,10)
        Dim XFila%(10), XColumna%(10)
        Dim WStock$(40, 10, 8)
        Dim Fila$(100), Columna$(100)
        Dim Elije$(200, 2)

        XFila%(1) = 12
        XFila%(2) = 13
        XFila%(3) = 14
        XFila%(4) = 15
        XFila%(5) = 16
        XFila%(6) = 17
        XFila%(7) = 18
        XFila%(8) = 19
        XFila%(9) = 20
        XFila%(10) = 21

        XColumna%(1) = 14
        XColumna%(2) = 22
        XColumna%(3) = 30
        XColumna%(4) = 38
        XColumna%(5) = 46
        XColumna%(6) = 54
        XColumna%(7) = 62
        XColumna%(8) = 70

        Rem $Include: 'Redondeo.fn'
        Rem $Include: 'Prnusing.fn'
        Rem $Include: 'Revdate.fn'
        Rem $Include: 'Impredate.fn'

        Empresas$(1) = "  MARILENE"
        Empresas$(2) = "  KLERICO"
        Empresas$(3) = "BIANCA'S CORSETERIA"
        Empresas$(4) = "BIANCA'S SECRET"
        Empresas$(5) = "PACO RABANNE"
        Empresas$(6) = "JAZMINE     "
        Empresas$(7) = "BIANCA'S MEDIAS"

        NomEmp$(1) = "MA"
        NomEmp$(2) = "KL"
        NomEmp$(3) = "NA"
        NomEmp$(4) = "BA"
        NomEmp$(5) = "PR"
        NomEmp$(6) = "GP"
        NomEmp$(7) = "EL"

        Data "SALTA               ", "BUENOS AIRES        ", "CAPITAL FEDERAL     ", "SAN LUIS            "
        Data "ENTRE RIOS          ", "LA RIOJA            ", "GRAN BUENOS AIRES   ", "CHACO               "
        Data "                    ", "SAN JUAN            ", "CATAMARCA           ", "LA PAMPA            "
        Data "MENDOZA             ", "MISIONES            ", "SANTIAGO DEL ESTERO ", "FORMOSA             "
        Data "NEUQUEN             ", "RIO NEGRO           ", "SANTA FE            ", "TUCUMAN             "
        Data "CHUBUT              ", "TIERRA DEL FUEGO    ", "CORRIENTES          ", "CORDOBA             "
        Data "JUJUY               ", "SANTA CRUZ          "

        For Ciclo% = 1 To 26
                Read Nomprov$(Ciclo%)
        Next Ciclo%

        Data "Monotributo     ", "Inscripto       ", "No Inscripto    ", "Exento          ", "No Responsable  "

        For Ciclo% = 1 To 5
                Read Nomiva$(Ciclo%)
        Next Ciclo%

        GoSub RESETEO
        Close

        GoSub ERRO

        'Gosub FCol051O
        'If St.051% <> 0 Then
        '        Status% = St.051%
        '        Gosub BTR.ERR
        'End If

        GoSub FTal050O
        If St.050% <> 0 Then
                Status% = St.050%
                Gosub BTR.ERR
        End If

        Op.050% = 12
        Clave.050$ = Space$(40)
        GoSub FTal050R

        While St.050% = 0

                Columna$(Val(Clave.050$)) = WNombre.050$

                Op.050% = 6
                GoSub FTal050R

        Wend

        'Op.051% = 12
        'Clave.051$ = Space$(40)
        'Gosub FCol051R
        '
        'While St.051% = 0
        '
        '        Fila$(Val(Clave.051$)) = WNombre.051$
        '
        '        Op.051% = 6
        '        Gosub FCol051R
        '
        'Wend

        GoSub RESETEO
        Close

        GoSub ERRO

        Lmsg% = 24
        Cmsg% = 2

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
                Gosub BTR.ERR
        End If
        '
        GoSub FArt005O
        If St.005% <> 0 Then
                Status% = St.005%
                Gosub BTR.ERR
        End If
        '
        GoSub FArt055O
        If St.055% <> 0 Then
                Status% = St.055%
                Gosub BTR.ERR
        End If
        '
        GoSub FCta030O
        If St.030% <> 0 Then
                Status% = St.030%
                Gosub BTR.ERR
        End If
        '
        GoSub FCon004O
        If St.004% <> 0 Then
                Status% = St.004%
                Gosub BTR.ERR
        End If
        '
        GoSub FExp003O
        If St.003% <> 0 Then
                Status% = St.003%
                Gosub BTR.ERR
        End If
        '
        GoSub FPed006O
        If St.006% <> 0 Then
                Status% = St.006%
                Gosub BTR.ERR
        End If
        '
        GoSub FCam007O
        If St.007% <> 0 Then
                Status% = St.007%
                Gosub BTR.ERR
        End If
        '
        GoSub FPed015O
        If St.015% <> 0 Then
                Status% = St.015%
                Gosub BTR.ERR
        End If
        '
        GoSub FDes008O
        If St.008% <> 0 Then
                Status% = St.008%
                Gosub BTR.ERR
        End If

        GoSub FIva031O
        If St.031% <> 0 Then
                Status% = St.031%
                Gosub BTR.ERR
        End If

        Open "r",#24,"Numero.dat",100
        Field #24,8 as Numero$

        WEmpresa$ = Space$(1)

        Open "Lpt1" For Output As #99 Len = 255

        Width "Lpt1:",255

Tronco:
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
                Print Chr$(255) + Chr$(255) + "FACT310/"
                LOCATE 1, 1: Call Qprint("Empresa : " + Empresas$(Empresa%), 13)

                Gosub Clear.Variables
                Erase Vector$
                Erase WStock$
                Existe$ = "N"
                Intro% = 1
                Confirma$ = ""
                A$ = ""

                Do
                        On Intro% Gosub Cliente, Numero, Clasificacion, Despacho, Partida, Compra, Emision,_
                                        Pedido, Condicion, Descuento, Impuesto, Precio, Campana,_
                                        Confirma
                        Intro% = Intro% + 1
                Loop Until Confirma$ = "S" Or Val(A$) = 1 Or Existe$ = "S"

                If Val(A$) = 1 Then Exit Do

                If Existe$ = "N" Then

                        Gosub Ingreso.Factura

                        If Val(A$) = 10 Then
                                If Val(WPedido.030$) = 0 Then
                                        Op.015%  = 13
                                        Kyn.015% = 0
                                        Clave.015$ = Space$(10)
                                        GoSub FPed015R
                                        If St.015% <> 0 Then
                                                WPedido.030$ = "1"
                                                        Else
                                                WPedido.030$ = Str$( Val(WCodigo.015$) + 1 )
                                        End If
                                        Call ceros(WPedido.030$,6)
                                        Graba$ = "N"
                                                Else
                                        Graba$ = "S"
                                End If
                                GoSub Impresion
                                GoSub Grabacion
                                GoSub RESETEO
                                Close
                                Chain "FACT310"
                        End If
                                        Else
                        Gosub Muestra.Factura
                        LOCATE 24, 2: Call Qprint(Space$(77), 13)
                        LOCATE 24, 2: Call Qprint("F1:Avanza Pagina  F2:Retrocede Pagina  F3:Fin Consulta  F4:Reimpresion", 13)
                        Do
                                Confirma$ = ""
                                Call Ingreso(Confirma$, "F", "", 1, 0, 24, 77, 13, 0, "", A$)
                                If Val(A$) = 1 Then
                                        If Desde% < 88 Then
                                                Desde% = Desde% + 12
                                                Gosub Impre.Pantalla
                                                Confirma$ = ""
                                        End If
                                End If
                                If Val(A$) = 2 Then
                                        If Desde% > 12 Then
                                                Desde% = Desde% - 12
                                                Gosub Impre.Pantalla
                                                Confirma$ = ""
                                        End If
                                End If
                                If Val(A$) = 4 Then
                                        Op.007% = 5
                                        Clave.007$ = FnRevdate$(WEmision.030$)
                                        GoSub FCam007R
                                        Op.004% = 5
                                        Clave.004$ = WCondicion.030$
                                        GoSub FCon004R
                                        GoSub Impresion
                                End If
                        Loop Until Val(A$) = 3
                        A$ = ""
                End If

        Loop Until Val(A$) = 1

        GoTo CIERRE

Clear.Variables:

        WNumero.030$      = ""
        WPartida.030$     = ""
        WEmision.030$     = ""
        WVencimiento.030$ = ""
        WCondicion.030$   = ""
        WCliente.030$     = ""
        WExpreso.030$     = ""
        WImpuesto.030$    = ""
        WCompra.030$      = ""
        WPrecio.030$      = ""
        WPedido.030$      = ""
        WLetra.030$       = ""
        WReventa.030$     = ""
        WCampana.030$     = ""
        Erase Dto#
        Erase WDescuento.001$,WVendedor.001$
        Return

Cliente:

        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
                Call Ingreso (WCliente.030$, "N", "", 5, 0, 3, 15, 13, 0, "#####", A$)

                IF VAL(WCliente.030$) = 0 AND VAL(A$) = 0 THEN
                        XDATO% = 1
                        GOSUB Selecciona.Datos
                        WCliente.030$ = XCodigo$
                        LOCATE 3, 15: PRINT USING "#####"; VAL(WCliente.030$)
                End If

                If Val(A$) <> 1 Then
                        Op.001% = 5
                        Clave.001$ = WCliente.030$
                        GoSub FCli001R
                        If St.001% <> 0 Then
                                Status% = St.001%
                                Gosub Btr.err
                                        Else
                                If Mid$(WCuit.001$,3,1) <> "-" Then
                                        LOCATE 24, 2: Print Space$(2);
                                        LOCATE 24, 2: Print "Cuit Invalido o inexistente";
                                        GoSub Uniq
                                        St.001% = 99
                                                               Else
                                        If Val(WVendedor.001$(Empresa%)) = 0 Then
                                                LOCATE 24, 2: Print Space$(2);
                                                LOCATE 24, 2: Print "Vendedor Inexistente";
                                                GoSub Uniq
                                                St.001% = 99
                                                                             Else
                                                Locate 3,25 : Call Qprint (Left$(WRazon.001$,15),13)
                                                Locate 4,15 : Call Qprint (WPartida.001$,13)
                                                'Locate 4,33 : Call Qprint (FnPusing$("####",Val(WExpreso.001$)),13)
                                                Locate 4,53 : Call Qprint (FnPusing$("####",Val(WVendedor.001$(Empresa%))),13)
                                                WPartida.030$   = WPartida.001$
                                                WVendedor.030$  = WVendedor.001$(Empresa%)
                                                WProvincia.030$ = WProvincia.001$
                                                If Val(WTipIva.001$) = 2 Or Val(WTipIva.001$) = 3 Then
                                                        WLetra.030$ = "A"
                                                                        Else
                                                        WLetra.030$ = "B"
                                                End If
                                        End If
                                End If
                        End If
                End If
        Loop Until Val(A$) = 1 Or St.001% = 0
        Return

Numero:

        If WLetra.030$ = "B" Then
                LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Fill Page 1/"
                LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Letra/"
                play "l24o3ao4c#e"
                Tecla$ = Input$(1)
                LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Display Page 1/"
        End If

        'Op.030% = 10
        'Clave.030$ = WLetra.030$ + "999999"
        'Kyn.030% = 1
        'Gosub FCta030RO
        'If St.030% <> 0 Or WLetra.030$ <> Letra.030$ Then
        '        WNumero.030$ = ""
        '                Else
        '        WNumero.030$ = Str$(Val(Numero.030$)+1)
        'End If
        'WNumero.030$ = ""

        If WLetra.030$ <> "B" Then
                Get #24,1
                WNumero.030$ = Str$(Val(Numero$) + 1)
                        Else
                Get #24,10
                WNumero.030$ = Str$(Val(Numero$) + 1)
        End If

        LOCATE 24, 2: Call Qprint(Space$(77), 7)
        LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
        Call Ingreso (WNumero.030$, "N", "", 6, 0, 3, 53, 13, 0, "######", A$)
        Op.030% = 5
        Clave.030$ = WLetra.030$ + WNumero.030$
        Kyn.030% = 1
        GoSub Fcta030r
        Kyn.030% = 0
        If St.030% <> 0 And St.030% <> 4 Then
                Status% = St.030%
                Gosub Btr.err
        End If
        If St.030% = 0 Then
                Existe$ = "S"
                'If Val(WTipo.030$) <> 1 Then
                '        Locate 24,2  : Call Qprint (Space$(77),13)
                '        Locate 24,28 : Call Qprint ("Numero ya utilizado en otro comprobante",13)
                '        Gosub Uniq
                '        Existe$ = ""
                '        WCliente.030$ = ""
                '        Intro% = 0
                '        Return
                'End If
                       Else
                Existe$ = "N"
        End If
        Return

Clasificacion:

        LOCATE 24, 2: Call Qprint(Space$(77), 7)
        LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
        If WReventa.030$ = "" Then
                WReventa.030$ = "P"
        End If
        Do
                Call Ingreso (WReventa.030$, "F", "", 1, 0, 3, 66, 13, 0, "", A$)
                If Val(A$) <> 0 Then Exit Do
        Loop Until WReventa.030$ = "P" Or WReventa.030$ = "R"
        Return

Despacho:

        Numero0$ = ""
        Numero1$ = ""

        LOCATE 24, 2: Call Qprint(Space$(77), 7)
        LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
        Do
                Call Ingreso(WDespacho$, "N", "", 4, 0, 2, 75, 13, 0, "####", A$)
                If Val(WDespacho$) = 0 Then Exit Do
                Op.008% = 5
                Clave.008$ = WDespacho$
                GoSub FDes008R
                If St.008% <> 0 Then
                        Status% = St.088%
                        Gosub Btr.err
                                Else
                        Numero0$ = WNumero.008$
                End If
        Loop Until St.008% = 0

        LOCATE 24, 2: Call Qprint(Space$(77), 7)
        LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
        Do
                Call Ingreso(WDespacho1$, "N", "", 4, 0, 3, 75, 13, 0, "####", A$)
                If Val(WDespacho1$) = 0 Then Exit Do
                Op.008% = 5
                Clave.008$ = WDespacho1$
                GoSub FDes008R
                If St.008% <> 0 Then
                        Status% = St.088%
                        Gosub Btr.err
                                Else
                        Numero1$ = WNumero.008$
                End If
        Loop Until St.008% = 0
        Return

Compra:

        LOCATE 24, 2: Call Qprint(Space$(77), 7)
        LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
        Call Ingreso (WCompra.030$, "F", "", 8, 0, 4, 33, 13, 0, "", A$)
        Return

Emision:

        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
                WEmision.030$ = Mid$(Date$,4,2) + Left$(Date$,2)+ Right$(Date$,2)
                Call Ingreso (WEmision.030$, "Y", "", 6, 0, 4, 69, 13, 0, "", A$)
                Call Validate(WEmision.030$,Mensaje$)
        Loop Until Val(A$) = 1 Or Mensaje$ = "Yes"
        If Val(A$) = 1 Then Return
        Op.007% = 5
        Clave.007$ = FnRevdate$(WEmision.030$)
        GoSub FCam007R
        If St.007% <> 0 Then
                Status% = St.007%
                Gosub Btr.err
                GoTo Emision
        End If
        Return

Partida:

        Call Ingreso (WPartida.030$, "G", "", 1, 0, 4, 15, 13, 0, "", A$)
        Return

Condicion:

        If Val(WCondicion.030$) = 0 Then
                WCondicion.030$ = "1"
        End If
        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
                Call Ingreso (WCondicion.030$, "N", "", 4, 0, 5, 15, 13, 0, "####", A$)
                If Val(A$) <> 1 Then
                        Op.004% = 5
                        Clave.004$ = WCondicion.030$
                        GoSub FCon004R
                        If St.004% <> 0 Then
                                Status% = St.004%
                                Gosub Btr.err
                                        Else
                                WFecha$ = WEmision.030$
                                Plazo% = Val(WDias.004$) + 1
                                GoSub 44000
                                'Locate 5,33 :Call Qprint (Left$(Fvenc$,2)+"/"+Mid$(Fvenc$,3,2)+"/"+Right$(Fvenc$,2),13)
                                WVencimiento.030$ = FVenc$
                        End If
                End If
        Loop Until Val(A$) = 1 Or St.004% = 0
        Return

Pedido:

        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
                Call Ingreso (WPedido.030$, "N", "", 6, 0, 5, 33, 13, 0, "######", A$)
                If Val(WPedido.030$) = 0 Or Val(A$) = 1 Then Exit Do
                Kyn.015% = 0
                Op.015%  = 5
                Clave.015$ = WPedido.030$ + "01"
                GoSub FPed015R
                If St.015% <> 0 Then
                        Status% = St.015%
                        Gosub Btr.err
                                Else
                        If Val(WCliente.015$) <> Val(WCliente.030$) Then
                                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                                LOCATE 24, 2: Call Qprint("El Cliente del pedido no corresponde al de la factura", 13)
                                GoSub Uniq
                                St.015% = 4
                                                Else
                                WCondicion.030$ = WCondicion.015$
                                WPrecio.030$ = WLista.015$
                                WDescuento.001$(Empresa%,1) = WDto1.015$
                                WDescuento.001$(Empresa%,2) = WDto2.015$
                                WDescuento.001$(Empresa%,3) = WDto3.015$
                        End If
                End If
        Loop Until St.015% = 0
        Return


Descuento:

        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
                Call Ingreso (WDescuento.030$, "F", "", 1, 0, 5, 53, 13, 0, "", A$)
        Loop Until Val(A$) = 1 Or WDescuento.030$ = "S" Or WDescuento.030$ = "N"

        If WDescuento.030$ = "S" Then

                LOCATE 1, 1
                Print Chr$(255) + Chr$(255) + "Fill Page 1/"
                LOCATE 1, 1
                Print Chr$(255) + Chr$(255) + "Fact310d/"


                Locate 11,43 : Print Using "###.##";Val(WDescuento.001$(Empresa%,1))/100
                Locate 13,43 : Print Using "###.##";Val(WDescuento.001$(Empresa%,2))/100
                Locate 15,43 : Print Using "###.##";Val(WDescuento.001$(Empresa%,3))/100

                Call Ingreso (WDescuento.001$(Empresa%,1), "N", "", 5, 2, 11, 43, 13, 0, "###.##", A$)

                Call Ingreso (WDescuento.001$(Empresa%,2), "N", "", 5, 2, 13, 43, 13, 0, "###.##", A$)

                Call Ingreso (WDescuento.001$(Empresa%,3), "N", "", 5, 2, 15, 43, 13, 0, "###.##", A$)
2
                LOCATE 1, 1
                Print Chr$(255) + Chr$(255) + "Display Page 1/"

        End If


        Return

Impuesto:

        If WImpuesto.030$ = "" Then
                WImpuesto.030$ = "N"
        End If

        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
                Call Ingreso (WImpuesto.030$, "F", "", 1, 0, 5, 67, 13, 0, "", A$)
        Loop Until Val(A$) = 1 Or WImpuesto.030$ = "S" Or WImpuesto.030$ = "N"
        Return

Precio:

        If Val(WPrecio.030$) = 0 Then
                WPrecio.030$ = "1"
        End If

        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 27: Call Qprint("   F1 : Menu de Emision", 13)
                Call Ingreso (WPrecio.030$, "N", "", 1, 0, 5, 78, 13, 0, "#", A$)
        Loop Until Val(A$) = 1 Or WPrecio.030$ = "1" Or WPrecio.030$ = "2"_
                                           Or WPrecio.030$ = "3" Or WPrecio.030$ = "4"_
                                           Or WPrecio.030$ = "5" Or WPrecio.030$ = "6"
        Return

Campana:

        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "Fill Page 1/"
        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "Fact310q/"

        Call Ingreso (WCampana.030$, "F", "", 10, 0, 13, 36, 13, 0, "", A$)

        LOCATE 1, 1
        Print Chr$(255) + Chr$(255) + "Display Page 1/"

        Return

Confirma:

        LOCATE 24, 2: Call Qprint(Space$(77), 13)
        LOCATE 24, 28: Call Qprint("Confirma S/N  : ", 13)
        Do
                Confirma$ = ""
                Call Ingreso(Confirma$, "F", "", 1, 0, 24, 48, 13, 0, "", A$)
        Loop Until Confirma$ = "S" Or Confirma$ = "N"
        Intro% = 0
        Return

Ingreso.Factura:

        If Val(A$) = 1 Then Return

        Importe# = 0
        Erase Vector$
        Erase WStock$

        Gosub Imprime.base

        LOCATE 24, 2: Call Qprint(Space$(77), 7)
        LOCATE 24, 2: Call Qprint("    F1 : Menu Principal   F2 : Anula Factura   F10 : Confirma Pantalla", 13)

        Fila% = 1
        Columna% = 1

        Do
                WColor% = 13
                Gosub IMPRE.LINEA
                Do
                        GoSub CAMPO1
                Loop Until Val(A$) = 1  Or Val(A$) = 11 Or Val(A$) = 12 OR _
                           Val(A$) = 10 Or Val(A$) = 2 Or_
                           Vector$(Fila%,1) <> ""
                Select Case Val(A$)
                        Case 0
                                Intro2% = 2
                                Gosub INGRESA.LINEA
                                Importe# = 0
                                For Ciclo% = 1 To 40
                                        Importe# = Importe# + Val(Vector$(Ciclo%, 4))
                                Next Ciclo%
                                Gosub Imprime.base
                                WColor% = 7
                                Gosub IMPRE.LINEA
                                Fila% = Fila% + 1
                                If Fila% > 19 Then Fila% = 1
                                Compara% = (Int((Fila% - 1) / 12) * 12) + 1
                                If Compara% = Fila% Then
                                        Desde% = Fila%
                                        Gosub IMPRE.PANTALLA
                                End If
                        Case 11
                                If Fila% < 19 Then
                                        If Vector$(Fila%, 1) <> "" Then
                                                WColor% = 7
                                                Gosub IMPRE.LINEA
                                                Fila% = Fila% + 1
                                                Compara% = (Int((Fila% - 1) / 12) * 12) + 1
                                                If Compara% = Fila% Then
                                                        Desde% = Fila%
                                                        Gosub IMPRE.PANTALLA
                                                End If
                                        End If
                                End If

                        Case 12
                                If Fila% > 1 Then
                                                WColor% = 7
                                                Gosub IMPRE.LINEA
                                                Fila% = Fila% - 1
                                                Compara% = (Int((Fila%) / 12) * 12)
                                                If Compara% = Fila% Then
                                                        Desde% = Fila% - 12
                                                        Gosub IMPRE.PANTALLA
                                                End If
                                End If
                        Case Else

                End Select

        Loop Until Val(A$) = 1 Or Val(A$) = 10 Or Val(A$) = 2

        Return

Ingresa.Linea:

        Do
                On Intro2% GoSub CAMPO1, CAMPO3, Unitario, Campo4
                Intro2% = Intro2% + 1
        Loop Until Intro2% = 5
        Return


CAMPO1:

        Do
                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                LOCATE 24, 2: Call Qprint("    F1 : Menu Principal   F2 : Anula Factura   F10 : Confirma Pantalla", 13)
                Lugar% = (((Fila% - 1) - Int((Fila% - 1) / 12) * 12) + 8) + 1
                St.005% = 0
                Call Ingreso(Vector$(Fila%, 1), "F", "", 6, 0, Lugar%, 2, 13, 0, "######", A$)
                If Val(A$) <> 1 And Val(A$) <> 11 And Val(A$) <> 12 And_
                   Val(A$) <> 2 And Val(A$) <> 10 And Vector$(Fila%,1) <> "" Then
                        If Vector$(Fila%, 1) = "999999" Then
                                GoSub Descipcion
                                Vector$(Fila%, 5) = ""
                                St.005% = 0
                                                        Else
                                Op.005% = 5
                                Clave.005$ = Vector$(Fila%,1)
                                GoSub FArt005R
                                If St.005% <> 0 Then
                                        Status% = St.005%
                                        Gosub Btr.err
                                                Else
                                        If NomEmp$(Empresa%) <> WEmpresa.005$ Then
                                                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                                                LOCATE 24, 2: Call Qprint("El Articulo no pertenece a la empresa ", 13)
                                                GoSub Uniq
                                                St.005% = 99
                                                Vector$(Fila%, 1) = ""
                                        End If
                                        If WClasi.005$ <> WReventa.030$ Then
                                                LOCATE 24, 2: Call Qprint(Space$(77), 7)
                                                LOCATE 24, 2: Call Qprint("El Articulo no pertenece a la clasificaion informada ", 13)
                                                GoSub Uniq
                                                St.005% = 99
                                                Vector$(Fila%, 1) = ""
                                        End If
                                        If St.005% = 0 Then
                                                Vector$(Fila%,2) = Left$(WDescripcion.005$,25)
                                                LOCATE Lugar%, 12: Call Qprint(Vector$(Fila%, 2), 13)
                                                Select Case Val(WPrecio.030$)
                                                        Case 1
                                                                Vector$(Fila%,5) = WVenta1.005$
                                                        Case 2
                                                                Vector$(Fila%,5) = WVenta2.005$
                                                        Case 3
                                                                Vector$(Fila%,5) = WVenta3.005$
                                                        Case 4
                                                                Vector$(Fila%,5) = WVenta4.005$
                                                        Case 5
                                                                Vector$(Fila%,5) = WVenta5.005$
                                                        Case 6
                                                                Vector$(Fila%,5) = WVenta6.005$
                                                End Select
                                                LOCATE Lugar%, 51: Call Qprint(FnPusing$("###,###,###.##", Val(Vector$(Fila%, 5))), 13)
                                        End If
                                End If
                        End If
                End If
        Loop Until St.005% = 0 Or Val(A$) = 1 Or Val(A$) = 11 Or Val(A$) = 12_
                   Or Val(A$) = 2 Or Val(A$) = 10
        If St.005% <> 0 Then Vector$(Fila%,1) = Space$(6)

        Op.055% = 5
        Kyn.055% = 0
        Clave.055$ =  Vector$(Fila%,1)
        GoSub FArt055R
        If St.055% = 0 Then
                If Val(WDespacho.055$) <> 0 Then
                        If val(WDespacho$) = 0 OR Val(WDespacho$) = Val(WDespacho.055$) Then
                                WDespacho$ = WDespacho.055$
                                Op.008% = 5
                                Clave.008$ = WDespacho.055$
                                GoSub FDes008R
                                If St.008% = 0 Then
                                        Numero0$ = WNumero.008$
                                End If
                                LOCATE 2, 75: Print Using; "####"; Val(WDespacho$)
                                        Else
                                If val(WDespacho1$) = 0 OR Val(WDespacho1$) = Val(WDespacho.055$) Then
                                        WDespacho1$ = WDespacho.055$
                                        Op.008% = 5
                                        Clave.008$ = WDespacho.055$
                                        GoSub FDes008R
                                        If St.008% = 0 Then
                                                Numero1$ = WNumero.008$
                                        End If
                                        LOCATE 3, 75: Print Using; "####"; Val(WDespacho1$)
                                End If
                        End If
                End If

        End If

        If Val(A$) = 0 Then
        If Val(WPedido.030$) <> 0 Then
                Dada$ = Mid$(Vector$(Fila%, 1), 5, 2)
                If Dada$ = "/E" Then
                        Clave.015$ = Left$(Vector$(Fila%,1),4) + "E" + WPedido.030$
                                Else
                        Clave.015$ = Left$(Vector$(Fila%,1),4) + " " + WPedido.030$
                End If
                Op.015% = 5
                Kyn.015% = 1
                GoSub FPed015R
                If St.015% <> 0 Then
                        'CLS
                        'PRINT ST.015%,CLAVE.015$
                        'TECLA$ = INPUT$(1)
                        status% = St.015%
                        Gosub Btr.err
                        GoTo CAMPO1
                End If
        End If
        End If


        Return

CAMPO3:

        Lugar% = (((Fila% - 1) - Int((Fila% - 1) / 12) * 12) + 8) + 1
        Call Ingreso(Vector$(Fila%, 3), "N", "", 6, 0, Lugar%, 42, 13, 0, "###,###", A$)
        Vector$(Fila%, 4) = Str$(Val(Vector$(Fila%, 3)) * Val(Vector$(Fila%, 5)))
        Return

Descipcion:

        Lugar% = (((Fila% - 1) - Int((Fila% - 1) / 12) * 12) + 8) + 1
        Vector$(Fila%, 2) = ""
        Call Ingreso(Vector$(Fila%, 2), "F", "", 25, 0, Lugar%, 12, 13, 0, "", A$)
        Return

Unitario:

        Lugar% = (((Fila% - 1) - Int((Fila% - 1) / 12) * 12) + 8) + 1
        Vector$(Fila%, 5) = Str$(Val(Vector$(Fila%, 5)) * 100)
        Call Ingreso(Vector$(Fila%, 5), "N", "", 11, 2, Lugar%, 51, 13, 0, "###,###,###.##", A$)
        Vector$(Fila%,5) = Str$( Val(Vector$(Fila%,5)) * Val(WCotizacion.007$) / 100 )
        Vector$(Fila%, 4) = Str$(Val(Vector$(Fila%, 3)) * Val(Vector$(Fila%, 5)))
        LOCATE Lugar%, 51: Call Qprint(FnPusing$("###,###,###.##", Val(Vector$(Fila%, 5))), 13)
        LOCATE Lugar%, 66: Call Qprint(FnPusing$("###,###,###.##", Val(Vector$(Fila%, 4))), 13)
        Return

Campo4:

        If WLenceria.005$ = "S" Then

                GoSub FArt010O
                If St.010% <> 0 Then
                        Status% = St.010%
                        Gosub BTR.ERR
                End If

                LOCATE 1, 1
                Print Chr$(255) + Chr$(255) + "Fill Page 1/"

                LOCATE 1, 1
                Print Chr$(255) + Chr$(255) + "Ayuda3/"

                LOCATE 7, 26: Print Vector$(Fila%, 1)
                Locate 7,35 : Print Left$(WDescripcion.005$,20)
                LOCATE 8, 26: Print Using; "###,###"; Val(Vector$(Fila%, 3))

                Op.010% = 5
                Clave.010$ = Vector$(Fila%,1)
                GoSub fArt010R
                If St.010% = 0 Then

                        For DA% = 1 To 10
                                Locate 11+Da%,2 : Print LEFT$(Fila$(Val(WColor.010$(Da%))),10)
                        Next DA%

                        For DA% = 1 To 8
                                Locate 10,(14+((Da%-1)*8)) : Print LEFT$(Columna$(Val(WTalle.010$(Da%))),7)
                        Next DA%

                End If

                For Dada% = 1 To 10
                        For XDada% = 1 To 8
                                LOCATE XFila%(Dada%), XColumna%(XDada%): Call Qprint(FnPusing$("###,###", Val(WStock$(Fila%, Dada%, XDada%))), 13)
                        Next XDada%
                Next Dada%

                WFila% = 1
                WColumna% = 1

                Do

                        Gosub Ingresa.Stock

                        Select Case Val(A$)
                                Case 0
                                        WFila% = WFila% + 1
                                        If WFila% > 10 Then WFila% = 1
                                        A$ = ""
                                Case 11
                                        If WFila% < 10 Then
                                                WFila% = WFila% + 1
                                                        Else
                                                WFila% = 1
                                        End If
                                        A$ = ""

                                Case 12
                                        If WFila% > 1 Then
                                                WFila% = WFila% - 1
                                        End If
                                        A$ = ""

                                Case 13
                                        If WColumna% > 1 Then
                                                WColumna% = WColumna% - 1
                                        End If
                                        A$ = ""

                                Case 14
                                        If WColumna% < 8 Then
                                                WColumna% = WColumna% + 1
                                                        Else
                                                WColumna% = 1
                                        End If
                                        A$ = ""
                                Case Else

                        End Select

                Loop Until Val(A$) = 1 Or Val(A$) = 2

                LOCATE 1, 1
                Print Chr$(255) + Chr$(255) + "Display Page 1/"

                Op.010% = 1
                GoSub fArt010R
                Close #10

                A$ = ""

        End If
        Return

Ingresa.Stock:

        Call Ingreso(WStock$(Fila%, WFila%, WColumna%), "N", "", 6, 0, XFila%(WFila%), XColumna%(WColumna%), 13, 0, "###,###", A$)
        LOCATE XFila%(WFila%), XColumna%(WColumna%): Call Qprint(FnPusing$("###,###", Val(WStock$(Fila%, WFila%, WColumna%))), 13)
        Return

Limpia.Linea:

        For Ciclo% = 1 To 5
                Vector$(Fila%, Ciclo%) = ""
        Next Ciclo%
        Return

Impre.Linea:


        Lugar% = (((Fila% - 1) - Int((Fila% - 1) / 12) * 12) + 8) + 1
        LOCATE Lugar%, 2: Call Qprint(Vector$(Fila%, 1), WColor%)
        LOCATE Lugar%, 12: Call Qprint(Vector$(Fila%, 2), WColor%)
        If Val(Vector$(Fila%, 3)) <> 0 Then
                LOCATE Lugar%, 42
                Call Qprint(FnPusing$("###,###", Val(Vector$(Fila%, 3))), WColor%)
                                      Else
                LOCATE Lugar%, 42
                Call Qprint(Space$(7), WColor%)
        End If

        If Val(Vector$(Fila%, 5)) <> 0 Then
                LOCATE Lugar%, 51
                Call Qprint(FnPusing$("###,###,###.##", Val(Vector$(Fila%, 5))), WColor%)
                                      Else
                LOCATE Lugar%, 51: Call Qprint(Space$(14), WColor%)
        End If

        If Val(Vector$(Fila%, 4)) <> 0 Then
                LOCATE Lugar%, 66
                Call Qprint(FnPusing$("###,###,###.##", Val(Vector$(Fila%, 4))), WColor%)
                                      Else
                LOCATE Lugar%, 66: Call Qprint(Space$(14), WColor%)
        End If
        Return


Impre.Pantalla:

        Hasta% = Desde% + 11
        If Hasta% > 40 Then Hasta% = 40
        For Ciclo% = Desde% To Hasta%
                If Vector$(Ciclo%, 1) <> "" And Vector$(Ciclo%, 1) <> Space$(6) Then
                        Lugar% = (((Ciclo% - 1) - Int((Ciclo% - 1) / 12) * 12) + 8) + 1
                        LOCATE Lugar%, 2: Call Qprint(Vector$(Ciclo%, 1), 7)
                        LOCATE Lugar%, 12: Call Qprint(Vector$(Ciclo%, 2), 7)
                        If Val(Vector$(Ciclo%, 3)) <> 0 Then
                                LOCATE Lugar%, 42
                                Call Qprint(FnPusing$("###,###", Val(Vector$(Ciclo%, 3))), 7)
                                                       Else
                                LOCATE Lugar%, 42
                                Call Qprint(Space$(7), 7)
                        End If

                        If Val(Vector$(Ciclo%, 5)) <> 0 Then
                                LOCATE Lugar%, 51
                                Call Qprint(FnPusing$("###,###,###.##", Val(Vector$(Ciclo%, 5))), 7)
                                                         Else
                                LOCATE Lugar%, 51: Call Qprint(Space$(14), 7)
                        End If

                        If Val(Vector$(Ciclo%, 4)) <> 0 Then
                                LOCATE Lugar%, 66
                                Call Qprint(FnPusing$("###,###,###.##", Val(Vector$(Ciclo%, 4))), 7)
                                                     Else
                                LOCATE Lugar%, 66: Call Qprint(Space$(14), WColor%)
                        End If

                                           Else
                        Lugar% = (((Ciclo% - 1) - Int((Ciclo% - 1) / 12) * 12) + 8) + 1
                        LOCATE Lugar%, 1
                        Call Qprint("³         ³                            ³         ³              ³              ³", 7)
                End If
        Next Ciclo%
        Return

Imprime.base:


        Auxiliar# = Importe#

        If WDescuento.030$ = "S" Then
                For Ciclo% = 1 To 3
                        If Val(WDescuento.001$(Empresa%,Ciclo%)) <> 0 Then
                                Dto#(Ciclo%) = FnRedondeo#((Auxiliar# * Val(WDescuento.001$(Empresa%,Ciclo%)))/10000)
                                Auxiliar# = Auxiliar# - Dto#(Ciclo%)
                                                        Else
                                Dto#(Ciclo%) = 0
                        End If
                Next Ciclo%
        End If

        Select Case Val(WTipiva.001$)
                Case 1
                        Auxiliar# = FnRedondeo#(Auxiliar# * 1.21)
                        Dto#(1) = FnRedondeo#(Dto#(1) * 1.21)
                        Dto#(2) = FnRedondeo#(Dto#(2) * 1.21)
                        Dto#(3) = FnRedondeo#(Dto#(3) * 1.21)
                        Iva#(1) = 0
                        Iva#(2) = 0

                Case 2, 5
                        Iva#(1) = FnRedondeo#(Auxiliar# * 0.21)
                        Iva#(2) = 0

                Case 3
                        Iva#(1) = FnRedondeo#(Auxiliar# * 0.21)
                        Iva#(2) = FnRedondeo#(Auxiliar# * 0.105)

                Case 4
                        Iva#(1) = 0
                        Iva#(2) = 0

                Case Else
        End Select

        Neto# = Auxiliar#
        Total# = Neto# + Iva#(1) + Iva#(2)

        If WImpuesto.030$ = "S" Then
                Impuesto# = FnRedondeo#(Total# * 0.0101)
                                Else
                Impuesto# = 0
        End If

        Total# = Neto# + Iva#(1) + Iva#(2) + Impuesto#

        LOCATE 22, 7: Call Qprint(FnPusing$("###,###,###.##", Neto#), 13)
        LOCATE 22, 26: Call Qprint(FnPusing$("###,###,###.##", Dto#(1) + Dto#(2) + Dto#(3)), 13)
        LOCATE 22, 45: Call Qprint(FnPusing$("###,###,###.##", Iva#(1) + Iva#(2)), 13)
        LOCATE 22, 66: Call Qprint(FnPusing$("###,###,###.##", Total#), 13)
        Return

Grabacion:

        ' Actualiza el Saldo del Cliente

        If Val(WTipIva.001$) = 1 Then
                XImporte# = FnRedondeo#(Neto# / 1.21)
                Iva#(1) = FnRedondeo#(Neto# - XImporte#)
                Neto# = FnRedondeo#(Neto# - Iva#(1))
        End If

        '  Graba el registro Corerspondiente a la Cuenta Corriente

        WTipo.030$       = "01"
        WImporte.030$    = Str$( Total# )
        WSaldo.030$      = Str$( Total# )
        WReal.030$       = Str$( Neto# * Val(WPartida.030$) + Iva#(1) + Iva#(2) )
        WNeto.030$       = Str$( Neto# )
        WIva1.030$       = Str$( Iva#(1) )
        WIva2.030$       = Str$( Iva#(2) )
        WMovimiento.030$ = "1"
        WCotizacion.030$ = WCotizacion.007$
        WOrigen.030$     = ""
        Op.030% = 2
        Clave.030$ = WCliente.030$+WTipo.030$+WNumero.030$
        GoSub FCta030w
        If St.030%<> 0 Then
                Status% = St.030%
                Gosub Btr.err
        End If

        '  Graba el registro Corerspondiente al IVA

        WLetra.031$     = WLetra.030$
        WTipo.031$      = "01"
        WNumero.031$    = WNumero.030$
        WCliente.031$   = WCliente.030$
        WPartida.031$   = WPartida.030$
        WEmision.031$   = WEmision.030$
        WVendedor.031$  = WVendedor.030$
        WProvincia.031$ = WProvincia.030$
        WNeto.031$      = Str$( Neto# )
        WIva1.031$      = Str$( Iva#(1) )
        WIva2.031$      = Str$( Iva#(2) )
        WEmpresa.031$   = WEmpresa.030$
        WReventa.031$   = WReventa.030$
        WCampana.031$   = WCampana.030$

        Op.031% = 2
        Clave.031$ = WLetra.031$ + WTipo.031$ + WNumero.031$
        GoSub FIva031w
        If St.031%<> 0 Then
                Status% = St.031%
                Gosub Btr.err
        End If

        ' Graba los Movimiento del Pedido

        For Ciclo% = 1 To 40
                If Vector$(Ciclo%, 1) <> "" And Vector$(Ciclo%, 1) <> Space$(6) Then
                        WLetra.006$    = WLetra.030$
                        WFactura.006$  = WNumero.030$
                        Auxiliar$ = Str$(Ciclo%)
                        Call Ceros(Auxiliar$, 2)
                        WRenglon.006$  = Auxiliar$
                        WFecha.006$    = WEmision.030$
                        WCliente.006$  = WCliente.030$
                        WPago.006$     = WPago.030$
                        WArticulo.006$ = Vector$(Ciclo%,1)
                        WCantidad.006$ = Vector$(Ciclo%,3)
                        WPrecio.006$   = Vector$(Ciclo%,5)
                        WDescripcion.006$ = Vector$(Ciclo%,2)
                        WCompra.006$      = WCompra.030$
                        Op.006%           = 2
                        Clave.006$        = WLetra.006$ + WFactura.006$ + WRenglon.006$
                        GoSub FPed006W
                        If St.006% <> 0 Then
                                Status% = St.006%
                                Gosub Btr.err
                        End If

                        Cantidad# = Val(Vector$(Ciclo%, 3))
                        If WPartida.030$ = "V" Then
                                Cantidad# = Cantidad# * 2
                        End If
                        If WPartida.030$ = "W" Then
                                Cantidad# = Cantidad# * 3
                        End If

                        Op.055% = 5
                        Kyn.055% = 0
                        Clave.055$ =  Vector$(Ciclo%,1)
                        GoSub FArt055R
                        If St.055% = 0 Then
                                Op.055% = 3
                                       Else
                                Op.055% = 2
                                WArticulo.055$  = Vector$(Ciclo%,1)
                                WStock1.055$ = ""
                                WStock2.055$ = ""
                        End If

                        WStock1.055$ = Str$( Val(WStock1.055$) - Cantidad# )

                        Clave.055$ = Vector$(Ciclo%,1)
                        GoSub FArt055W
                        If St.055% <> 0 Then
                                Status% = St.055%
                                Gosub Btr.err
                        End If

                End If
        Next Ciclo%

        ' Actualiza los datos del Pedido

        If Graba$ = "S" Then

                For Ciclo% = 1 To 40

                        Dada$ = Mid$(Vector$(Ciclo%, 1), 5, 2)
                        If Dada$ = "/E" Then
                                Clave.015$ = Left$(Vector$(Ciclo%,1),4) + "E" + WPedido.030$
                                                             Else
                                Clave.015$ = Left$(Vector$(Ciclo%,1),4) + " " + WPedido.030$
                        End If
                        Op.015% = 5
                        Kyn.015% = 1
                        GoSub FPed015R
                        If St.015% = 0 Then
                                Cantidad# = Val(Vector$(Ciclo%, 3))
                                If WPartida.030$ = "V" Then
                                        'If WReventa.030$ = "P" Then
                                                Cantidad# = Cantidad# * 2
                                        'End If
                                End If
                                If WPartida.030$ = "W" Then
                                        'If WReventa.030$ = "P" Then
                                                Cantidad# = Cantidad# * 3
                                        'End If
                                End If
                                WFacturado.015$ = Str$( Val(WFacturado.015$) + Cantidad# )
                                Op.015% = 3
                                GoSub FPed015W
                                If St.015% <> 0 Then
                                        Status% = St.015%
                                        Gosub Btr.err
                                End If
                        End If

                Next Ciclo%

                                        Else

                For Ciclo% = 1 To 40

                        If Vector$(Ciclo%, 1) <> "" And Vector$(Ciclo%, 1) <> Space$(6) Then

                                Cantidad# = Val(Vector$(Ciclo%, 3))
                                If WPartida.030$ = "V" Then
                                        'If WReventa.030$ = "P" Then
                                                Cantidad# = Cantidad# * 2
                                        'End If
                                End If
                                If WPartida.030$ = "W" Then
                                        'If WReventa.030$ = "P" Then
                                                Cantidad# = Cantidad# * 3
                                        'End If
                                End If

                                WCodigo.015$  = WPedido.030$
                                WRenglon.015$ = Str$(Ciclo%)
                                Call Ceros(WRenglon.015$,2)

                                WEmpresa.015$   = WEmpresa$
                                WCliente.015$   = WCliente.030$
                                WFecha.015$     = WEmision.030$
                                WPartida.015$   = WPartida.030$
                                WLista.015$     = WPrecio.030$
                                WCondicion.015$ = WCondicion.030$
                                WDto1.015$      = ""
                                WDto2.015$      = ""
                                WDto3.015$      = ""

                                WArticulo.015$ = Left$(Vector$(Ciclo%,1),4)
                                WCantidad.015$ = Str$(Cantidad#)
                                WPrecio.015$   = ""
                                WObservaciones.015$ = "Pedido automatico"
                                If Mid$(Vector$(Ciclo%, 1), 5, 2) = "/E" Then
                                        WEspecial.015$ = "E"
                                                        Else
                                        WEspecial.015$ = ""
                                End If
                                WTalle1.015$   = ""
                                WTalle2.015$   = ""
                                WTalle3.015$   = ""
                                WTalle4.015$   = ""
                                WTalle5.0125$   = ""
                                WTalle6.015$   = ""
                                WTalle7.015$   = ""
                                WTalle8.015$   = ""
                                WTalle9.015$   = ""
                                WTalle10.015$  = ""
                                WEntregado.015$ = ""
                                WFacturado.015$ = Str$(Cantidad#)

                                Op.015%  = 2
                                Kyn.015% = 0
                                Clave.015$ = WCodigo.015$ + WRenglon.015$
                                GoSub FPed015W
                                If St.015% <> 0 Then
                                        Status% = St.015%
                                        Gosub Btr.err
                                End If

                        End If

                Next Ciclo%

        End If

        GoSub FMov011O
        If St.011% <> 0 Then
                Status% = St.011%
                Gosub BTR.ERR
        End If

        For Ciclo% = 1 To 40
                If Vector$(Ciclo%, 1) <> "" And Vector$(Ciclo%, 1) <> Space$(6) Then

                        Op.005% = 5
                        Kyn.005% = 0
                        Clave.005$ = Vector$(Ciclo%,1)
                        GoSub FArt005R
                        If St.005% = 0 Then

                                If WLenceria.005$ = "S" Then

                                        WTipo.011$     = WTipo.030$
                                        WLetra.011$    = WLetra.030$
                                        WNumero.011$   = WNumero.030$
                                        WFecha.011$    = WEmision.030$
                                        WCliente.011$  = WCliente.030$
                                        WArticulo.011$ = Vector$(Ciclo%,1)

                                        Auxiliar$ = Str$(Ciclo%)
                                        Call Ceros(Auxiliar$, 2)
                                        WRenglon.011$  = Auxiliar$

                                        For XDa% = 1 To 10
                                                For WDa% = 1 To 8
                                                        WCantidad.011$(XDa%,WDa%)  = WStock$(Ciclo%,XDa%,WDa%)
                                                Next WDa%
                                        Next XDa%

                                        Op.011% = 2
                                        Clave.011$ = WTipo.011$ + WLetra.011$ + WNumero.011$ + WRenglon.011$
                                        GoSub FMov011w
                                        If St.011% <> 0 Then
                                                Status% = St.011%
                                                Gosub Btr.err
                                        End If

                                End If

                        End If

                End If
        Next Ciclo%

        Op.011% = 1
        GoSub FMov011r
        Close #11

        GoSub FArt009O
        If St.009% <> 0 Then
                Status% = St.009%
                Gosub BTR.ERR
        End If

        For Ciclo% = 1 To 40
                If Vector$(Ciclo%, 1) <> "" And Vector$(Ciclo%, 1) <> Space$(6) Then

                        Op.005% = 5
                        Kyn.005% = 0
                        Clave.005$ = Vector$(Ciclo%,1)
                        GoSub FArt005R
                        If St.005% = 0 Then

                                If WLenceria.005$ = "S" Then

                                        Op.009% = 5
                                        Kyn.009% = 0
                                        Clave.009$ = Vector$(Ciclo%,1)
                                        GoSub FArt009r
                                        If St.009% = 0 Then
                                                Op.009% = 3
                                                        Else
                                                Op.009% = 2
                                                WArticulo.009$ = Vector$(Ciclo%,1)
                                                Erase WStock.009$
                                        End If

                                        For XDa% = 1 To 10
                                                For WDa% = 1 To 8
                                                        WStock.009$(XDa%,WDa%)  = Str$(Val(WStock.009$(XDa%,WDa%)) - Val(WStock$(Ciclo%,XDa%,WDa%)) )
                                                Next WDa%
                                        Next XDa%

                                        Op.009%    = 3
                                        GoSub FArt009W
                                        If St.009% <> 0 Then
                                                Status% = St.009%
                                                Gosub Btr.err
                                        End If

                                End If

                        End If

                End If
        Next Ciclo%

        Op.009% = 1
        GoSub FArt009r
        Close #9

        If Val(WNumero.030$) > Val(Numero$) Then
                If WLetra.030$ <> "B" Then
                        Lset Numero$ = WNumero.030$
                        Put #24,1
                                Else
                        Lset Numero$ = WNumero.030$
                        Put #24,10
                End If
        End If

        Return


Baja:

        '  Borra el registro Corerspondiente a la Cuenta Corriente

        WTipo.030$ = "01"
        Kyn.030%   = 0
        Op.030%    = 5
        Clave.030$ = WCliente.030$ + WTipo.030$ + WNumero.030$
        GoSub Fcta030r
        If St.030% <> 0 Then
                Status% = St.030%
                Gosub Btr.err
                        Else
                Op.030% =4
                GoSub FCta030w
                If St.030% <> 0 Then
                        Status% = St.030%
                        Gosub Btr.err
                End If
        End If

        '  Borra el registro Corerspondiente al IVA

        WTipo.031$   = "01"
        WNumero.031$ = WNumero.030$
        Op.031%      = 5
        Clave.031$   = WLetra.030$ + WTipo.031$ + WNumero.031$
        GoSub FIva031R
        If St.031%<> 0 Then
                Status% = St.031%
                Gosub Btr.err
                       Else
                Op.031% = 4
                GoSub FIva031w
                If St.031% <> 0 Then
                        Status% = St.031%
                        Gosub Btr.err
                End If
        End If

        ' Borra los Movimiento del Pedido

        For Ciclo% = 1 To 40
                If Vector$(Ciclo%, 1) <> "" And Vector$(Ciclo%, 1) <> Space$(6) Then
                        WFactura.006$  = WNumero.030$
                        WLetra.006$    = WLetra.030$
                        Auxiliar$ = Str$(Ciclo%)
                        Call Ceros(Auxiliar$, 2)
                        WRenglon.006$  = Auxiliar$
                        Op.006%        = 5
                        Clave.006$     = WLetra.006$ + WFactura.006$ + WRenglon.006$
                        GoSub FPed006r
                        If St.006% <> 0 Then
                                Status% = St.006%
                                Gosub Btr.err
                                        Else
                                Op.006% = 4
                                GoSub FPed006W
                                If St.006% <> 0 Then
                                        Status% = St.006%
                                        Gosub Btr.err
                                End If
                        End If
                End If
        Next Ciclo%

        ' Actualiza los datos del Pedido

        If Val(WPedido.030$) <> 0 Then
                For Ciclo% = 1 To 40
                        Dada$ = Mid$(Vector$(Ciclo%, 1), 5, 2)
                        If Dada$ = "/E" Then
                                Clave.015$ = Left$(Vector$(Ciclo%,1),4) + "E" + WPedido.030$
                                                             Else
                                Clave.015$ = Left$(Vector$(Ciclo%,1),4) + " " + WPedido.030$
                        End If
                        Op.015% = 5
                        Kyn.015% = 1
                        GoSub FPed015R
                        If St.015% = 0 Then

                                If left$(WObservaciones.015$,13) <> "Pedido automa" Then
                                        Cantidad# = Val(Vector$(Ciclo%, 3))
                                        If WPartida.030$ = "V" Then
                                                'If WReventa.030$ = "P" Then
                                                        Cantidad# = Cantidad# * 2
                                                'End If
                                        End If
                                        If WPartida.030$ = "W" Then
                                                'If WReventa.030$ = "P" Then
                                                        Cantidad# = Cantidad# * 3
                                                'End If
                                        End If
                                        WFacturado.015$ = Str$( Val(WFacturado.015$) - Cantidad# )
                                        Op.015% = 3
                                        GoSub FPed015W
                                        If St.015% <> 0 Then
                                                Status% = St.015%
                                                Gosub Btr.err
                                        End If
                                                Else
                                        Op.015% = 4
                                        GoSub FPed015W
                                        If St.015% <> 0 Then
                                                Status% = St.015%
                                                Gosub Btr.err
                                        End If
                                End If
                        End If
                Next Ciclo%
        End If
        Return


Muestra.Factura:

        Erase Vector$
        Erase WStock$

        Op.006% = 9
        Clave.006$ = WLetra.030$ + WNumero.030$ + Space$(2)
        GoSub FPed006r
        Pointer% = 0
        Importe# = 0
        While St.006% = 0 And WFactura.006$ = WNumero.030$ And WLetra.030$ = WLetra.006$
                Pointer% = Pointer% + 1
                WCompra.030$       = WCompra.006$
                Vector$(Pointer%,1) = WArticulo.006$
                Vector$(Pointer%,2) = Left$(WDescripcion.006$,25)
                Vector$(Pointer%,3) = WCantidad.006$
                Vector$(Pointer%,5) = WPrecio.006$
                Vector$(Pointer%,4) = Str$( Val(WPrecio.006$)  * Val(Vector$(Pointer%,3)) )
                Importe# = Importe# + Val(Vector$(Pointer%, 4))
                Op.006% = 6
                GoSub FPed006r
        Wend
        Clave.001$ = WCliente.030$
        Op.001%    = 5
        GoSub FCli001R
        If St.001% <> 0 Then
                Status% = St.001%
                Gosub Btr.err
        End If
        Locate 3,33 : Print Using "#####";Val(WCliente.030$)
        Locate 3,25 : Print Left$(WRazon.001$,15)
        Locate 3,66 : Print WReventa.030$
        Locate 4,15 : Print WPartida.030$
        Locate 4,33 : Print WCompra.030$
        Locate 4,53 : Print Using "####";Val(WVendedor.001$(Empresa%))
        Locate 4,69 : Print Left$(WEmision.030$,2)+"/"+Mid$(WEmision.030$,3,2)+"/"+Right$(WEmision.030$,2)
        Locate 5,15 : Print Using "####";Val(WCondicion.030$)
        Locate 5,33 : Print Using "######";Val(WPedido.030$)
        Locate 5,53 : Print WDescuento.030$
        Locate 5,67 : Print WImpuesto.030$
        Locate 5,78 : Print WPrecio.030$
        Desde% = 1
        Gosub Impre.pantalla
        Gosub Imprime.base

        WTipo.031$   = "01"
        WNumero.031$ = WNumero.030$
        Op.031%      = 5
        Clave.031$   = WLetra.030$ + WTipo.031$ + WNumero.031$
        GoSub FIva031R
        '\cls
        '\print clave.031$,op.031%,st.031%

        Neto#   = Val(WNeto.031$)
        Iva#(1) = Val(WIva1.031$)
        Iva#(2) = 0
        Total# = Neto# + Iva#(1)

        LOCATE 22, 7: Call Qprint(FnPusing$("###,###,###.##", Neto#), 13)
        LOCATE 22, 45: Call Qprint(FnPusing$("###,###,###.##", Iva#(1) + Iva#(2)), 13)
        LOCATE 22, 66: Call Qprint(FnPusing$("###,###,###.##", Total#), 13)

        Return

Impresion:

        LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Fill Page 1/"
        LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Fact310a/"
        Do
                Call Ingreso(Tecla$, "F", "", 1, 0, 12, 37, 13, 0, "", A$)
        Loop Until Tecla$ = "S" Or Tecla$ = "N"
        LOCATE 1, 1: Print Chr$(255) + Chr$(255) + "Display Page 1/"

        GoSub Encabezamiento
        Lineas% = 0
        Suma# = 0

        Op.003% = 5
        Clave.003$ = WExpreso.001$
        GoSub FExp003R

        For Ciclo% = 1 To 19
                If Vector$(Ciclo%, 1) <> "" And Vector$(Ciclo%, 1) <> Space$(6) Then

                        If Vector$(Ciclo%, 1) <> "999999" Then

                                If Empresa% <> 7 Then

                                        If Right$(Vector$(Ciclo%, 1), 2) <> Space$(2) Then
                                                Print #99, Tab(3); Left$(Vector$(Ciclo%, 1), 4);
                                                Print #99, "-";
                                                Print #99, Right$(Vector$(Ciclo%, 1), 2);
                                                        Else
                                                Print #99, Tab(3); Vector$(Ciclo%, 1);
                                        End If

                                                Else

                                        Op.005% = 5
                                        Clave.005$ = Vector$(Ciclo%,1)
                                        GoSub FArt005R
                                        If St.005% <> 0 Then
                                                Status% = St.005%
                                                Gosub Btr.err
                                        End If

                                        Print #99,Tab(3);WCodigo.005$;

                                End If

                        End If

                        Print #99, Tab(13); Left$(Vector$(Ciclo%, 2), 23);
                        Print #99, Tab(39); Using; "#####"; Val(Vector$(Ciclo%, 3));
                        If Val(WTipiva.001$) = 1 Then
                                Impre# = FnRedondeo#(Val(Vector$(Ciclo%, 5)) * 1.21)
                                Print #99, Tab(45); Using; "#####.##"; Impre#;
                                Impre# = FnRedondeo#(Val(Vector$(Ciclo%, 4)) * 1.21)
                                Print #99, Tab(56); Using; "###,###.##"; Impre#;
                                        Else
                                Print #99, Tab(45); Using; "#####.##"; Val(Vector$(Ciclo%, 5));
                                Print #99, Tab(56); Using; "###,###.##"; Val(Vector$(Ciclo%, 4));
                        End If

                        If WLetra.030$ = "A" Then

                                If Vector$(Ciclo%, 1) <> "999999" Then
                                        If Empresa% <> 7 Then
                                                If Right$(Vector$(Ciclo%, 1), 2) <> Space$(2) Then
                                                        Print #99, Tab(70); Left$(Vector$(Ciclo%, 1), 4);
                                                        Print #99, "-";
                                                        Print #99, Right$(Vector$(Ciclo%, 1), 2);
                                                                Else
                                                        Print #99, Tab(70); Vector$(Ciclo%, 1);
                                                End If
                                                        Else
                                                Print #99,Tab(70);WCodigo.005$;
                                        End If
                                End If

                                Print #99, Tab(81); Left$(Vector$(Ciclo%, 2), 23);
                                Print #99, Tab(107); Using; "#####"; Val(Vector$(Ciclo%, 3));
                                If Val(WTipiva.001$) = 1 Then
                                        Impre# = FnRedondeo#(Val(Vector$(Ciclo%, 5)) * 1.21)
                                        Print #99, Tab(112); Using; "#####.##"; Impre#;
                                        Impre# = FnRedondeo#(Val(Vector$(Ciclo%, 4)) * 1.21)
                                        'Print #99,Tab(122);using "###,###.##";Impre#;
                                        Suma# = Suma# + Impre#
                                                Else
                                        Print #99, Tab(112); Using; "#####.##"; Val(Vector$(Ciclo%, 5));
                                        If WLetra.030$ = "A" Then
                                                'Print #99,Tab(122);using "###,###.##";Val(Vector$(Ciclo%,4));
                                        End If
                                        Suma# = Suma# + Val(Vector$(Ciclo%, 4))
                                End If
                        End If

                        Lineas% = Lineas% + 1
                End If
        Next Ciclo%

        For Imprelinea% = Lineas% To 19
                Print #99, ""
        Next Imprelinea%

        If WLetra.030$ = "A" Then

                Print #99,Tab(13);"O.C.: ";WCompra.030$;
                Print #99,Tab(81);"O.C.: ";WCompra.030$

                If Dto#(1) <> 0 And WDescuento.030$ = "S" Then
                        Print #99,Tab(39);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,1))/100;
                        Print #99, Tab(56); Using; "###,###.##"; Dto#(1)
                        If WLetra.030$ = "A" Then
                                'Print #99,Tab(107);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,1))/100;
                                'Print #99,Tab(122);using "###,###.##";Dto#(1)
                        End If
                                Else
                        Print #99, ""
                End If

                If Dto#(2) <> 0 And WDescuento.030$ = "S" Then
                        Print #99,Tab(39);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,2))/100;
                        Print #99, Tab(56); Using; "###,###.##"; Dto#(2)
                        If WLetra.030$ = "A" Then
                                'Print #99,Tab(107);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,2))/100;
                                'Print #99,Tab(122);using "###,###.##";Dto#(2)
                        End If
                                Else
                        Print #99, ""
                End If

                If Tecla$ = "S" Then
                        Dolares# = Total# / Val(WCotizacion.007$)
                        Print #99, Chr$(15)
                        Print #99, Tab(8); "ESTA FACTURA EQUIVALE A U$S "; Using; "###,###,###.##"; Dolares#;
                        Print #99, " CALCULADOS AL TIPO DE CAMBIO VENDEDOR DEL ";
                        If WLetra.030$ = "A" Then
                                Print #99, Tab(120); "EQUIVALE A : U$S "; Using; "###,###.##"; Dolares#;
                                Print #99," VENCIMIENTO : ";FnImpredate$(WVencimiento.030$);". VALOR BASE 1 U$S  = ";Using "###,###,###.##";Val(WCotizacion.007$);
                        End If
                        Print #99,Tab(8);"DOLAR LIBRE A LA FECHA DE LA MISMA. SU VENCIMIENTO SERA EL ";FnImpredate$(WVencimiento.030$);" DEBIENDO SER "
                        Print #99, Tab(8); "CANCELADA EN BILLETES ESTADOUNIDENSES. EN SU DEFECTO SE CANCELARA CONVIRTIENDO ";
                        Print #99, Tab(8); "LA CANTIDAD DE DOLARES A PESOS, DE ACUERDO A LA COTIZACION TIPO VENDEDOR DEL DOLAR LIBRE";
                        Print #99, Tab(8); "QUE RIJA A ESE MOMENTO O A LA FECHA DE ACREDITACION DEL PAGO.";
                        Print #99,Tab(8);"VALOR BASE : 1 U$S = ";Using "###,###,###.##";Val(WCotizacion.007$);
                        Print #99, Chr$(18)
                                Else
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                End If

                Print #99,Tab(72);left$(WNombre.003$,20);
                Print #99,Tab(93);Left$(WDireccion.003$,17);
                Print #99,Tab(111);LEFT$(WCuit.003$,15)

                Print #99, ""

                        Else

                Print #99,Tab(13);"O.C.: ";WCompra.030$

                If Dto#(1) <> 0 And WDescuento.030$ = "S" Then
                        Print #99,Tab(39);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,1))/100;
                        Print #99, Tab(56); Using; "###,###.##"; Dto#(1)
                        If WLetra.030$ = "A" Then
                                'Print #99,Tab(107);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,1))/100;
                                'Print #99,Tab(122);using "###,###.##";Dto#(1)
                        End If
                                Else
                        Print #99, ""
                End If

                If Dto#(2) <> 0 And WDescuento.030$ = "S" Then
                        Print #99,Tab(39);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,2))/100;
                        Print #99, Tab(56); Using; "###,###.##"; Dto#(2)
                        If WLetra.030$ = "A" Then
                                'Print #99,Tab(107);"Bonif. ";Using "##.##";Val(WDescuento.001$(Empresa%,2))/100;
                                'Print #99,Tab(122);using "###,###.##";Dto#(2)
                        End If
                                Else
                        Print #99, ""
                End If

                If Tecla$ = "S" Then
                        Dolares# = Total# / Val(WCotizacion.007$)
                        Print #99, Chr$(15)
                        Print #99, Tab(8); "ESTA FACTURA EQUIVALE A U$S "; Using; "###,###,###.##"; Dolares#;
                        Print #99, " CALCULADOS AL TIPO DE CAMBIO VENDEDOR DEL ";
                        Print #99,Tab(8);"DOLAR LIBRE A LA FECHA DE LA MISMA. SU VENCIMIENTO SERA EL ";FnImpredate$(WVencimiento.030$);" DEBIENDO SER "
                        Print #99, Tab(8); "CANCELADA EN BILLETES ESTADOUNIDENSES. EN SU DEFECTO SE CANCELARA CONVIRTIENDO ";
                        Print #99, Tab(8); "LA CANTIDAD DE DOLARES A PESOS, DE ACUERDO A LA COTIZACION TIPO VENDEDOR DEL DOLAR LIBRE";
                        Print #99, Tab(8); "QUE RIJA A ESE MOMENTO O A LA FECHA DE ACREDITACION DEL PAGO.";
                        Print #99,Tab(8);"VALOR BASE : 1 U$S = ";Using "###,###,###.##";Val(WCotizacion.007$);
                        Print #99, Chr$(18)
                                Else
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                        Print #99, ""
                End If

                Print #99,Tab(8);left$(WNombre.003$,20);
                Print #99,Tab(30);Left$(WDireccion.003$,17);
                Print #99,Tab(50);LEFT$(WCuit.003$,15)
                Print #99, ""

        End If

        If WLetra.030$ = "A" Then

                Print #99, Tab(56); Using; "###,###.##"; Neto#
                'Print #99,Tab(122);Using "###,###.##";Neto#

                If WImpuesto.030$ = "S" Then
                        Print #99, Tab(48); "1%";
                        Print #99, Tab(56); Using; "###,###.##"; Impuesto#
                        'Print #99,Tab(116);"1%";
                        'Print #99,Tab(122);Using "###,###.##";Impuesto#
                                Else
                        Print #99, ""
                End If

                Print #99, Tab(56); Using; "###,###.##"; Neto# + Impuesto#
                'Print #99,Tab(122);Using "###,###.##";Neto#+Impuesto#

                If Iva#(1) <> 0 Then
                        Print #99, Tab(49); "21 %";
                        Print #99, Tab(56); Using; "###,###.##"; Iva#(1)
                        'Print #99,Tab(116);"21 %";
                        'Print #99,Tab(122);Using "###,###.##";Iva#(1)
                                Else
                        Print #99, ""
                End If


                If Iva#(2) <> 0 Then
                        Print #99, Tab(49); "10.5%";
                        Print #99, Tab(56); Using; "###,###.##"; Iva#(2)
                        'Print #99,Tab(116);"10.5%";
                        'Print #99,Tab(122);Using "###,###.##";Iva#(2)
                                Else
                        Print #99, ""
                End If

                Print #99, ""

                Print #99,Tab(5);WPartida.030$;
                Print #99,Tab(10);Right$(WVendedor.001$(Empresa%),2);
                Print #99,Tab(17);Using "######";Val(WPedido.030$);
                Print #99, Tab(56); Using; "###,###.##"; Total#;
                Print #99,Tab(72);WPartida.030$;
                Print #99,Tab(77);Right$(WVendedor.001$(Empresa%),2)
                'Print #99,Tab(122);Using "###,###.##";Total#
                Print #99,Tab(5);WCobrador.001$(Empresa%);
                Print #99,Tab(72);WCobrador.001$(Empresa%)

                Print #99, ""
                Print #99, ""

                If Val(WDespacho$) <> 0 Then
                        Print #99, Tab(5); "Despacho : "; Numero0$; " / "; Numero1$;
                        Print #99, Tab(72); "Despacho : "; Numero0$; " / "; Numero1$
                                       Else
                        Print #99, ""
                End If

                                        Else

                Print #99, ""
                Print #99, ""
                Print #99, ""
                Print #99, ""
                Print #99, ""
                Print #99, ""

                Print #99,Tab(5);WPartida.030$;
                Print #99,Tab(10);Right$(WVendedor.001$(Empresa%),2);
                Print #99,Tab(17);Using "######";Val(WPedido.030$);
                Print #99, Tab(56); Using; "###,###.##"; Total#
                Print #99,Tab(5);WCobrador.001$(Empresa%)

                Print #99, ""
                Print #99, ""

                If Val(WDespacho$) <> 0 Then
                        Print #99, Tab(5); "Despacho : "; Numero0$; " / "; Numero1$
                                       Else
                        Print #99, ""
                End If

        End If

        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""
        Print #99, ""

        Return

Encabezamiento:

        If WLetra.030$ = "A" Then
                Print #99, Tab(50); "FACTURA"
                Print #99, ""
                Print #99, ""
                Print #99, Tab(4); Empresas$(Empresa%);
                Print #99,Tab(44);FnImpredate$(FnRevdate$(WEmision.030$));
                Print #99, Tab(74); Empresas$(Empresa%);
                Print #99,Tab(114);FnImpredate$(FnRevdate$(WEmision.030$))
                Print #99, ""
                Print #99, ""
                Print #99, ""
                Print #99, ""

                If Val(WDespacho$) <> 0 Then
                        Print #99, Tab(40); "Aduana Bs. As.";
                        Print #99, Tab(104); "Aduana Bs. As."
                                        Else
                        Print #99, ""
                End If

                Print #99, ""
                Print #99, ""

                Print #99,Tab(34);WDireccion.001$;
                Print #99,Tab(104);WDireccion.001$
                Print #99,Tab(3);WRazon.001$;
                Print #99,Tab(34);WLocalidad.001$;
                Print #99,Tab(72);WRazon.001$;
                Print #99,Tab(104);WLocalidad.001$
                Print #99,Tab(34);Nomprov$(Asc(WProvincia.001$)-64);" ";WPostal.001$;
                Print #99,Tab(104);Nomprov$(Asc(WProvincia.001$)-64);" ";WPostal.001$

                Print #99, ""
                Print #99, ""

                        Else

                Print #99, Tab(50); "FACTURA"
                Print #99, ""
                Print #99, ""
                Print #99, Tab(4); Empresas$(Empresa%);
                Print #99,Tab(44);FnImpredate$(FnRevdate$(WEmision.030$))
                Print #99, ""
                Print #99, ""
                Print #99, ""
                Print #99, ""

                If Val(WDespacho$) <> 0 Then
                        Print #99, Tab(40); "Aduana Bs. As."
                                        Else
                        Print #99, ""
                End If

                Print #99, ""
                Print #99, ""

                Print #99,Tab(34);WDireccion.001$
                Print #99,Tab(3);WRazon.001$;
                Print #99,Tab(34);WLocalidad.001$
                Print #99,Tab(34);Nomprov$(Asc(WProvincia.001$)-64);" ";WPostal.001$

                Print #99, ""
                Print #99, ""
        End If

        If WLetra.030$ = "A" Then

                Print #99,Tab(11);NomIva$(Val(WTipIva.001$));
                Print #99,Tab(41);WCuit.001$;
                Print #99,Tab(61);WCliente.001$;
                Print #99,Tab(76);NomIva$(Val(WTipIva.001$));
                Print #99,Tab(108);WCuit.001$;
                Print #99,Tab(128);WCliente.001$;

                        Else

                Print #99,Tab(11);NomIva$(Val(WTipIva.001$));
                Print #99,Tab(41);WCuit.001$;
                Print #99,Tab(61);WCliente.001$;

        End If

        If WLetra.030$ = "A" Then
                Print #99, ""
                Print #99, ""

                Print #99,Tab(12);WNombre.004$;
                Print #99,Tab(59);VAL(right$(WNumero.030$,5))-8000;
                Print #99,Tab(82);WNombre.004$

                Print #99, ""
                Print #99, ""
                Print #99, ""
                        Else
                Print #99, ""
                Print #99, ""

                Print #99,Tab(12);WNombre.004$;
                Print #99,Tab(59);right$(WNumero.030$,5)

                Print #99, ""
                Print #99, ""
                Print #99, ""
        End If
        Return

CIERRE:

     GoSub RESETEO
     Close
     Chain "FACT300"

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

Impre.Linea.Seleccion:

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




Rem $Include: 'ERR2'
Rem $Include: 'FCli101.Fde'
Rem $Include: 'FArt005.Fde'
Rem $Include: 'FArt155.Fde'
Rem $Include: 'FExp103.Fde'
Rem $Include: 'FCon104.Fde'
Rem $Include: 'FCta030.Fde'
Rem $Include: 'FIva031.Fde'
Rem $include: 'FPed006.Fde'
Rem $include: 'FCam007.Fde'
Rem $include: 'FPed015.Fde'
Rem $include: 'FArt010.Fde'
Rem $include: 'FArt009.Fde'
Rem $include: 'FTal150.Fde'
Rem $Include: 'FDes108.Fde'
Rem $include: 'FMov011.Fde'
Rem $Include: 'RESET'
Rem $Include: 'PRUPAN2'
Rem $Include: 'validate.sub'
Rem $Include: 'FECVTO'


