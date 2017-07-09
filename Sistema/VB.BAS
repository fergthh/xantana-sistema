        LPRINT "                                      +---+"
        LPRINT "+-------------------------------------| A |-----------------------------------+"
        LPRINT "|                                     +---+                                   |"
  LPRINT USING "| \                                   \ | MERCADERIA EN CONSIGNACION          |"; RecordEmp.Nombre
  LPRINT USING "| \                                   \ |Factura N§         :0000-0000\  \    |"; RecordEmp.Direccion; RIGHT$(RecordIva.Numero, 4)
  LPRINT USING "| \                                   \ |Fecha              :\      \         |"; RecordEmp.Localidad; FnImpredate$(fNrEVDATE$(RecordIva.Fecha))
  LPRINT USING "|                                       |Cuit N§            :\               \|"; RecordEmp.Nombre
        LPRINT "|-----------------------------------------------------------------------------|"
        LPRINT "|                                                                             |"
        LPRINT USING "|Se¤or    :\                                                     \ - (\    \) |"; RecordCli.Nombre; RecordCli.Codigo
        LPRINT USING "|Domicilio:\                                                     \            |"; RecordCli.Direccion
        LPRINT USING "|Localidad:\                                                     \            |"; RecordCli.Localidad
        LPRINT USING "|C.U.I.T. :\                  \                                               |"; RecordCli.Cuit
        LPRINT USING "|Iva.     :\                  \                                               |"; Iva$(VAL(RecordCli.Iva))
        LPRINT "|                                                                             |"
        LPRINT "|-----------------------------------------------------------------------------|"
        LPRINT "| Codigo   |           Descripcion                | Cantidad  |  P.Unitario   |"
        LPRINT "|-----------------------------------------------------------------------------|"
        LPRINT "|          |                                      |           |               |"

        FOR Ciclo% = 1 TO 20
                IF Campo$(Ciclo%, 1) <> "" AND Campo$(Ciclo%, 1) <> SPACE$(6) THEN
                        Impre1$ = ""
                        IF Campo$(Ciclo%, 1) <> "999999" THEN
                                Impre1$ = Campo$(Ciclo%, 1)
                        END IF
                        Impre2$ = Campo$(Ciclo%, 2)
                        Impre3$ = Campo$(Ciclo%, 3)
                        Impre4$ = Campo$(Ciclo%, 5)
                        Impre6$ = Campo$(Ciclo%, 4)

                        LPRINT USING "|\        \|\                                    \|   ###,### |   ###,###.##  |"; Impre1$; Impre2$; VAL(Impre3$); VAL(Impre4$)

                        Lineas% = Lineas% + 1
                END IF
        NEXT Ciclo%

        FOR Imprelinea% = Lineas% TO 20
                LPRINT "|          |                                      |           |               |"
        NEXT Imprelinea%

        LPRINT "|-----------------------------------------------------------------------------|"

        LPRINT ""
        LPRINT ""
        LPRINT ""

        LPRINT "|-----------------------------------------------------------------------------|"
        LPRINT USING "|                                                Sub-Total    |   ###,###.##  |"; Total#
        LPRINT "|-----------------------------------------------------------------------------|"


        LPRINT ""
        LPRINT ""
        LPRINT ""
        LPRINT ""
        LPRINT ""

