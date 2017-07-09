ImpresioN:

	If val(Movimiento$) < 100000 Then

		LPRINT
		LPRINT
		LPRINT
		LPRINt
		LPRINT TAB(60);LEFT$(Fechaped$,2);
		LPRINT TAB(65);mid$(Fechaped$,3,2);
		LPRINT TAB(70);right$(Fechaped$,2)
		LPRINt
		LPRINt
		LPRINt
		LPRINt
		lprint
		LPRINt Tab(15);WNOmbre.009$;"  ";WCodigo.009$
		lprint
		lprint Tab(15);WDireccion.009$;
		lprint Tab(55);WPostal.009$
		lprint
		lprint Tab(15);WLocalidad.009$
		lprint
		lprint Tab(55);WCuit.009$
		lprint
		lprint
		lprint
		If Tipo$ = "2" Then
			lprint tAB(30);"DEVOLUCION"
					else
			LPrint
		End if
		lprint
		lprint

		Lin% = 0

		For Counter% = 1 to 35

			If Val(Campo$(Counter%,5)) <> 0 Then

				Articulo$ = Campo$(Counter%,3)
				Color$ = Campo$(Counter%,4)
				Cantidad$ = Campo$(Counter%,5)

				'Op.005% = 5
				'Clave.005$ = Articulo$
				'Gosub FArt005R
				'If St.005% <> 0 Then
				'        Impre1$ = ""
				'                else
				'        Impre1$ =  left$(WDescripcion.005$,28)
				'End If
				Impre1$ = Campo$(Counter%,1)

				Op.035% = 5
				Clave.035$ = Color$
				Gosub FCol035R
				If St.035% <> 0 Then
					Impre2$ = ""
						Else
					Impre2$ = Left$(WNombre.035$,15)
				End If


				lprint Tab(1);Using "###,###.##";Val(Cantidad$);
				lprint Tab(15);Articulo$;
				lprint Tab(30);left$(Impre1$,30);
				'lprint Tab(56);USIng "####";Val(Color$);
				lprint Tab(62);Impre2$

				Lin% = Lin% +1

			End If

		Next Counter%
		for da% = Lin% to 24
			lprint
		next da%




			else




		LPRINT CHR$(27) + CHR$(64);
		LPRINT CHR$(27) + CHR$(67) + CHR$(72);
		LPRINT CHR$(18)

		For WDa% = 1 to 2

			If WDa% = 1 Then
				LPRINT
				LPRINT "Numero : ";Val(Movimiento$)
				LPRINt
					else
				LPRINt
				'If val(Prove$) => 90 Then
				'        LPRINT "Numero : ";Val(Movimiento$)
				'end if
				LPRINt Prove$
				LPRINt WNOmbre.009$
			End if
			LPRINT TAB(60);LEFT$(Fechaped$,2);
			LPRINT TAB(65);mid$(Fechaped$,3,2);
			LPRINT TAB(70);right$(Fechaped$,2)
			If WDa% = 1 Then
				lprint
				LPRINt Tab(15);WNOmbre.009$;"  ";WCodigo.009$
				lprint
				lprint Tab(15);WDireccion.009$;
				lprint Tab(55);WPostal.009$
				lprint Tab(15);WLocalidad.009$
				lprint Tab(55);WCuit.009$
					else
				lprint
				lprint
			end if
			lprint

			Lin% = 0

			For Counter% = 1 to 35

				If Val(Campo$(Counter%,5)) <> 0 Then

					Articulo$ = Campo$(Counter%,3)
					Color$ = Campo$(Counter%,4)
					Cantidad$ = Campo$(Counter%,5)

					'Op.005% = 5
					'Clave.005$ = Articulo$
					'Gosub FArt005R
					'If St.005% <> 0 Then
					'        Impre1$ = ""
					'                else
					'        Impre1$ =  left$(WDescripcion.005$,40)
					'End If
					Impre1$ = Campo$(Counter%,1)

					Op.035% = 5
					Clave.035$ = Color$
					Gosub FCol035R
					If St.035% <> 0 Then
						Impre2$ = ""
							Else
						Impre2$ = Left$(WNombre.035$,15)
					End If


					lprint Tab(1);Using "###,###.##";Val(Cantidad$);
					lprint Tab(15);Articulo$;
					lprint Tab(30);left$(Impre1$,30);
					'lprint Tab(56);USIng "####";Val(Color$);
					lprint Tab(62);Impre2$

					Lin% = Lin% +1

				End If

			Next Counter%
			for da% = Lin% to 22
				lprint
			next da%
		Next WDa%

		lprint Chr$(12)

	End if

	return
