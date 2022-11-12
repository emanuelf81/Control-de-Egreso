REM  *****  BASIC  *****

Option Explicit


Sub ArchivarFinalizados
'ARCHIVA EN BDEXPEDICION LAS TAREAS FINALIZADAS, DEJA EL MES EN CURSO MAS UN MES ANTERIOR A MENOS QUE HAYA
'UNA TAREA PENDIENTE, EN TAL CASO CORTA ANTES DE ESA TAREA.
Dim nFilaInicial as Integer
Dim nFilaFinal as Integer
Dim nFila as Integer
Dim y, x
Dim vIdTareaInicial as Integer
Dim vIdTareaFinal as Integer
Dim vIdTareaBuscada as Integer
Dim FechaMaxArchivar
Dim RutaDestino As String
Dim Hoja As Object
Dim HojaDestino As Object  
Dim dRango As Object
Dim oDataArrayOrigen As Object
Dim ArchDestino As Object
Dim RangoO as String
Dim RangoD as String

Inicio:
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent

	If vUsuario <> "EMANUEL" then Exit Sub
	If Msgbox( "¿Esta seguro que desea archivar las tareas?", 4 + 48, "ESTA SEGURO?????" ) = 6 then
		Doc.Store()
	Else 
		Exit Sub
	End If

Paso1:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	nFilaInicial = 10
	vIdTareaInicial = 0
	vIdTareaBuscada = 0
	For y = nFilaInicial to 8000
	 	Cell = Sheet.getCellByPosition(0, y)
 		If Cell.String = "" then
			Posicionar = y
			PosicionadorCelda
 			Msgbox "Id.Tarea desconocida."+chr(13)+"Proceso suspendido.",16,"ERROR"
 			Exit For
 		End IF
	 	Cell = Sheet.getCellByPosition(1, y)
 		If Cell.String = "" then
			Posicionar = y
			PosicionadorCelda
 			Msgbox "Nro. Cliente/Destinatario desconocido."+chr(13)+"Proceso suspendido.",16,"ERROR"
 			Exit For
 		End IF
 		
		'BUSCA CORRELATIVIDAD EN EL NUMERO DE ID.TAREA
 		If vIdTareaInicial <> 0 then
 			Cell = Sheet.getCellByPosition(0, y-1)
 			vIdTareaBuscada = Cell.Value
 			Cell = Sheet.getCellByPosition(0, y)
 			If vIdTareaBuscada = Cell.Value - 1 then
 			
 			Else
 				Msgbox "No hay correlatividad en la Id.Tarea nº"+Cell.String+chr(13)+"Proceso suspendido.",16,"ERROR"
 				Exit Sub
 			End If
 		End If
 		 		
		'BUSCA EL NUMERO DE TAREA INICIAL
		If y = nFilaInicial then
		 	Cell = Sheet.getCellByPosition(8, y)
 			If Cell.String = "FINALIZADO" then
			 	Cell = Sheet.getCellByPosition(0, y)
 				vIdTareaInicial = Cell.Value
 			Else
				Posicionar = y
				PosicionadorCelda
	 			Msgbox "No es una tarea en Estado FINALIZADA."+chr(13)+"Proceso suspendido.",16,"ERROR"
				Exit For
	 		End IF
 		End If
 		
		'CORROBORA QUE LA TAREA ESTE FINALIZADA
		'COMIENZA A BUSCAR EN LA COLUMNA "A PARTIR DE" (HASTA 60 DIAS ANTES DE LA FECHA ACTUAL)
	 	Cell = Sheet.getCellByPosition(8, y)
		If Cell.String = "FINALIZADO" then
		 	Cell = Sheet.getCellByPosition(12, y)
'		 	vFechaApartir = Cell.getString
			If Date - 30 > cDate(Cell.String) then
				Posicionar = y
				PosicionadorCelda
				nFilaFinal = y
			 	Cell = Sheet.getCellByPosition(0, y)
 				vIdTareaFinal = Cell.Value
 			Else
				nFilaFinal = y - 1
				Posicionar = nFilaFinal
				PosicionadorCelda
		 		Cell = Sheet.getCellByPosition(0, nFilaFinal)
				vIdTareaFinal = Cell.Value
 				Exit For
			End If
 		Else
			nFilaFinal = y - 1
			Posicionar = nFilaFinal
			PosicionadorCelda
		 	Cell = Sheet.getCellByPosition(0, nFilaFinal)
			vIdTareaFinal = Cell.Value
			Exit For
 		End If
 	Next y
 	Msgbox "Nro.Fila Inicial: "+nFilaInicial+chr(13)+"Nro.Fila Final: "+nFilaFinal+chr(13)+chr(13)+"Id.Tarea Inicial :"+vIdTareaInicial+chr(13)+"Id.Tarea Final :"+vIdTareaFinal,64,"ARCHIVAR"
	Posicionar = nFilaFinal
	PosicionadorCelda	

	If nFilaFinal-nFilaInicial <> vIdTareaFinal-vIdTareaInicial then 
		MSGBOX "HAY DIFERENCIAS ENTRE LAS CANTIDADES DE FILAS Y LAS CANTIDADES DE TAREAS."+chr(13)+"Proceso suspendido.",16,"ERROR"
		exit sub
	End If
	
	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(11, 4)
	RutaDestino = ConvertToURL( Cell.String )

	nFilaInicial = nFilaInicial + 1
	nFilaFinal = nFilaFinal + 1
	vIdTareaFinal = vIdTareaFinal + 1
	vIdTareaInicial = vIdTareaInicial + 1
	RangoO = "A" + nFilaInicial + ":AE" + nFilaFinal
	RangoD = "A" + vIdTareaInicial + ":AE" + vIdTareaFinal

	'ORIGEN
	Hoja = Doc.Sheets.getByName("Carga de Tareas")
	dRango = Hoja.getCellRangeByName ( RangoO )
	oDataArrayOrigen = dRango.getDataArray()
	
	'DESTINO
	ArchDestino = StarDesktop.loadComponentFromURL( RutaDestino, "_blank", 0, Array() )
	HojaDestino = ArchDestino.Sheets.getByName("Expedicion")
	dRango = HojaDestino.getCellRangeByName ( RangoD )
	dRango.setDataArray(oDataArrayOrigen)
	ArchDestino.Store ()
	ArchDestino.dispose ()

	nFilaInicial = nFilaInicial - 1
	Hoja = Doc.Sheets.getByName("Carga de Tareas")
	Hoja.getRows.removeByIndex( 11, nFilaFinal - nFilaInicial )

msgbox "Finalizado."
End Sub
