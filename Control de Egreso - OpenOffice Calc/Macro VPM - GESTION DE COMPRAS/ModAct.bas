REM  *****  BASIC  *****


Option Explicit

Global RutaOrigen As String
Global ModArchivoCE As String
Global Procesando as Boolean

'USUARIO
Dim otxtPWVista As Object

Sub ActualizareGC
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	
	Dim oHoja As Object, Doc As Object

	Dim oHojaOrigen As Object, ArchOrigen As Object, dFuente As Object
	Dim oDataArrayOrg As Object
	Dim dDestino As Object, RutaArchivoActual As String
	Dim Flags As Long

	Dim RutaBaseDatosClientes as String, ModBaseDatosClientes As String, oArchivoAct As Object
	Dim oHojaAct As Object, oRangoAct As Object, oDataAct As Object
	Dim dHojaAct As Object, dRangoAct As Object
	Dim oHojaAct2 As Object, oRangoAct2 As Object, oDataAct2 As Object
	
	Dim vCEIdTarea As String, vCENroCliente As String, vCENombreClienteDireccion As String 
	Dim vCEZonaTareas As String, vCEPrioridad As String, vCEInfo As String, vCEEstadoAsignado As String

	Dim vGCIdTarea As String, vGCNroCliente As String, vGCNombreClienteDireccion As String 
	Dim vGCZonaTareas As String, vGCPrioridad As String, vGCInfo As String, vGCEstado As String, vGCFechaApartir As String

	'Contadores y Buscadores
	Dim ceFila
	Dim ctFila
	Dim gcFila
	Dim ctUltFila
	Dim RangoB As String
	
	Doc = thiscomponent

	Dim mCamposOrden(0) As New com.sun.star.table.TableSortField
	Dim mDescriptorOrden()

	'CORROBORA QUE HAYA INGRESADO LA CONTRASEÑA DE USUARIO	
	If vUsuario = "" then 
		Sheet = Doc.Sheets.getByName("Usuario")
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
	    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
		otxtPW = oFormulario.getByName("txtPW")
		otxtPWVista = Doc.getCurrentController.getControl( otxtPW )
		otxtPWVista.SetFocus()
		Procesando = False
		Exit Sub
	End If

'msgbox con opcion cancelar
	Msgbox "La Actualización puede demorarar unos minutos."+chr(13)+"Favor de esperar a que el sistema le informe que ha finalizado.", 48,"Importante"

	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator

	'RUTA DE ARCHIVOS
	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(11, 2)
	RutaOrigen = ConvertToURL( Cell.String )
	'RutaOrigen = ConvertToURL( "//Servidor/carpetas individuales/Ventas/CONTROL DE EGRESO/VPM - CONTROL DE EGRESO.ods" )  'RUTA EN VEPROMET
	'RutaOrigen = ConvertToURL( "C:/Users/Emanuel-Not/Documents/EMANUEL/VEPROMET/VPM - CONTROL DE EGRESO.ods" ) 'RUTA EN CASA
	Cell = Sheet.getCellByPosition(11, 1)
	RutaArchivoActual = ConvertToURL( Cell.String )
	'RutaArchivoActual = ConvertToURL( "//Servidor/carpetas individuales/Ventas/CONTROL DE EGRESO/VPM - GESTION DE COBROS.ods" )  'RUTA EN VEPROMET
	'RutaArchivoActual = ConvertToURL( "C:/Users/Emanuel-Not/Documents/EMANUEL/VEPROMET/VPM - GESTION DE COBROS.ods" ) 'RUTA EN CASA
'Goto SaltoPrueba

	'CORROBORA ACTUALIZACIONES EN BASE DE DATOS DE CLIENTES
	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(11, 3)
	RutaBaseDatosClientes = ConvertToURL( Cell.String )
	Cell = Sheet.getCellByPosition(12, 3)
	ModBaseDatosClientes = Cell.String
	If cDate(Left(ModBaseDatosClientes, 10)) <> cDate(Left(FileDateTime( RutaBaseDatosClientes ), 10)) then
		If cDate(Left(ModBaseDatosClientes, 10)) > cDate(Left(FileDateTime( RutaBaseDatosClientes ), 10)) then
			
		Else
			Goto Paso2A	
		End If
	else
		If cDate(Right(ModBaseDatosClientes, 8)) < cDate(Right(FileDateTime( RutaBaseDatosClientes ), 8)) then
			Goto Paso2A		
		End If
	End If
	
Goto Paso2B
Paso2A:
	'Actualiza la base de datos de los clientes.
	oBarraEstado.start( "Actualizando Clientes... ", 100 )
			
	oArchivoAct = StarDesktop.loadComponentFromURL( RutaBaseDatosClientes, "_blank", 0, Array() )
	oHojaAct = oArchivoAct.Sheets.getByName("BDClientes")
	oRangoAct = oHojaAct.getCellRangeByName("A11:F20010")
	oDataAct = oRangoAct.getDataArray()

	oHojaAct2 = oArchivoAct.Sheets.getByName("BDOtrosDestinatarios")
	oRangoAct2 = oHojaAct2.getCellRangeByName("A11:F1010")
	oDataAct2 = oRangoAct2.getDataArray()

	oArchivoAct.dispose ()

	dHojaAct = Doc.Sheets.getByName("BDClientes")
	dRangoAct = dHojaAct.getCellRangeByName("A3:F20002")	'SELLECCIONA HOJA DE DESTINO
	Flags = com.sun.star.sheet.CellFlags.STRING		'BORRA INFORMACION EN LAS CELDAS SELECCIONADAS
	dRangoAct.clearContents(Flags)
	dRangoAct.setDataArray(oDataAct)			'PEGA LA INFORMACION COPIA EN EL ARCHIVO DE ORIGEN

	dRangoAct = dHojaAct.getCellRangeByName("A20003:F21002")	'SELLECCIONA HOJA DE DESTINO
	Flags = com.sun.star.sheet.CellFlags.STRING		'BORRA INFORMACION EN LAS CELDAS SELECCIONADAS
	dRangoAct.clearContents(Flags)
	dRangoAct.setDataArray(oDataAct2)			'PEGA LA INFORMACION COPIA EN EL ARCHIVO DE ORIGEN
		
	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(12, 3)
	Cell.String = FileDateTime( RutaBaseDatosClientes )

	oBarraEstado.setValue( 50 )	
	WAIT 500
			
	'ORDENAR BDCLIENTES ALFABETICAMENTE
	Sheet = Doc.Sheets.getByName("BDClientes") 'La hoja donde esta el rango a ordenar
	oRango = Sheet.getCellRangeByName("A2:F21002") 'El rango a ordenar
	mDescriptorOrden = oRango.createSortDescriptor() 'Descriptor de ordenamiento, o sea, el "como"
	mCamposOrden(0).Field = 1 'Los campos empiezan en 0
	mCamposOrden(0).IsAscending = True
	mCamposOrden(0).IsCaseSensitive = False 'Sensible a MAYUSCULAS/minusculas
	mCamposOrden(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC 'Tipo de campo AUTOMATICO
	mDescriptorOrden(1).Name = "ContainsHeader" 'Indicamos si el rango contiene títulos de campos
	mDescriptorOrden(1).Value = True
	mDescriptorOrden(3).Name = "SortFields" 'La matriz de campos a ordenar
	mDescriptorOrden(3).Value = mCamposOrden
 	oRango.sort( mDescriptorOrden )	'Ordenamos con los parámetros establecidos

	'Borra Nros. de clientes que no esten asignados.
	oBarraEstado.start( "Actualizando Clientes... ", 100 )
	oBarraEstado.setValue( 60 )	
	dHojaAct = Doc.Sheets.getByName("BDClientes")
	For y = 2 to 20020
		Cell = Sheet.getCellByPosition(1, y)
		If Cell.String = "" then
			oBarraEstado.start( "Actualizando Clientes... ", 100 )
			oBarraEstado.setValue( 80 )	
			y = y + 1
			RangoB = "A"+y+":F21003"
			CellRange = Sheet.getCellRangeByName( RangoB )
			Flags = com.sun.star.sheet.CellFlags.STRING
			Flags = Flags + com.sun.star.sheet.CellFlags.VALUE
			CellRange.clearContents(Flags)
			CellRange.CellBackColor = RGB(255,255,255)			
			Exit for
		End If
	Next y
	oBarraEstado.setValue( 100 )
	oBarraEstado.end()		
Paso2B:			
	'COPIA INFORMACION DEL ARCHIVO DE ORIGEN
	ArchOrigen = StarDesktop.loadComponentFromURL( RutaOrigen, "_blank", 0, Array() )
	oHojaOrigen = ArchOrigen.Sheets.getByName("Expedicion-Cobros")
	dFuente = oHojaOrigen.getCellRangeByName("A6:G260")
	oDataArrayOrg = dFuente.getDataArray()
	ArchOrigen.dispose ()	

	'BORRA Y PEGA LA INFORMACION COPIADA EN EL ARCHIVO DE ORIGEN
	oHoja = Doc.Sheets.getByName("Expedicion-Cobros")
	dDestino = oHoja.getCellRangeByName ("A6:G260")	'SELLECCIONA HOJA DE DESTINO
	Flags = com.sun.star.sheet.CellFlags.STRING		'BORRA INFORMACION EN LAS CELDAS SELECCIONADAS
	dDestino.clearContents(Flags)
	'dDestino.CellBackColor = RGB(255,255,255)		'COLOR FONDO BLANCO CELDAS SELECCIONADAS
	dDestino.CellBackColor = RGB(213,231,234)		'COLOR FONDO CELESTE CLARO BLANCO CELDAS SELECCIONADAS

	dDestino.setDataArray(oDataArrayOrg)			'PEGA LA INFORMACION COPIA EN EL ARCHIVO DE ORIGEN

	'COLOREA GC EN EXPEDICION-COBROS
	oHoja = Doc.Sheets.getByName("Expedicion-Cobros")
	Sheet = oHoja.getCellRangeByName ("I6:O260")
	Flags = com.sun.star.sheet.CellFlags.STRING		'BORRA INFORMACION EN LAS CELDAS SELECCIONADAS
	Sheet.clearContents(Flags)
	Sheet.CellBackColor = RGB(255,255,255)		'COLOR FONDO BLANCO CELDAS SELECCIONADAS

	'ORDENA GC EN EXPEDICION-COBROS
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	ctUltFila = 10011
	ctFila = 0
	For ctFila = 11 to ctUltFila '10011
		Cell = Sheet.getCellByPosition(1, ctFila)
		If Cell.String = "" then	'VERIFICA SI LAS PROXIMAS 15 FILAS CONTIENEN INFORMACION
			z = 0
			For z = 1 to 15
				Cell = Sheet.getCellByPosition(1, ctFila + z)
				If Cell.String <> "" then Goto Continua2
			Next z
			ctUltFila = ctFila
			Exit For
		End If
Continua2:
		Cell = Sheet.getCellByPosition(9, ctFila)
		If Cell.String = "PENDIENTE" or Cell.String = "" then
			If Cell.String = "" then
				Msgbox	"HAY TAREAS EN GESTION DE COBROS QUE EL ESTADO ES DESCONOCIDO."+chr(13)+"FAVOR DE CORROBORAR QUE ESTA CELDA CONTENGA INFORMACION PARA QUE EL PROGRAMA FUNCIONE CORRECTAMENTE.", 48,"Importante"
			End If			
			vGCEstado = Cell.String
			Cell = Sheet.getCellByPosition(0, ctFila)
			vGCIdTarea = Cell.String
			Cell = Sheet.getCellByPosition(1, ctFila)
			vGCNroCliente = Cell.String
			Cell = Sheet.getCellByPosition(2, ctFila)
			vGCNombreClienteDireccion = Cell.String
			Cell = Sheet.getCellByPosition(3, ctFila)
			vGCNombreClienteDireccion = vGCNombreClienteDireccion + chr(13) + Cell.String
			Cell = Sheet.getCellByPosition(4, ctFila)
			vGCZonaTareas = Cell.String
			Cell = Sheet.getCellByPosition(5, ctFila)
			vGCZonaTareas = vGCZonaTareas + chr(13) + Cell.String
			Cell = Sheet.getCellByPosition(6, ctFila)
			vGCPrioridad = Cell.String
			Cell = Sheet.getCellByPosition(7, ctFila)
			vGCInfo = Cell.String
			Cell = Sheet.getCellByPosition(12, ctFila)
			vGCFechaApartir = Cell.String
			Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
			gcFila = 0
			For gcFila = 5 to 255
				Cell = Sheet.getCellByPosition(8, gcFila)
				If Cell.String = "" then
					Cell.String = vGCIdTarea
					Cell = Sheet.getCellByPosition(9, gcFila)
					Cell.String = vGCNroCliente
					Cell = Sheet.getCellByPosition(10, gcFila)
					Cell.String = vGCNombreClienteDireccion
					Cell = Sheet.getCellByPosition(11, gcFila)
					Cell.String = vGCZonaTareas
					Cell = Sheet.getCellByPosition(12, gcFila)
					Cell.String = vGCPrioridad
					Cell = Sheet.getCellByPosition(13, gcFila)
					Cell.String = vGCInfo
					Cell = Sheet.getCellByPosition(14, gcFila)
					Cell.String = vGCFechaApartir
					Sheet = Doc.Sheets.getByName("Carga de Tareas")
					Exit For
				End If
			Next gcFila
		End If	
	Next ctFila

	'BORRA INFORMACION DE LAS COLUMNAS CE EN CARGA DE TAREAS DE GC
	oHoja = Doc.Sheets.getByName("Carga de Tareas")
	Sheet = oHoja.getCellRangeByName ("K12:L10011")
	Flags = com.sun.star.sheet.CellFlags.STRING
	Sheet.clearContents(Flags)
SaltoPrueba:
	
	'CORRIGE ERRORES DE TRANSFERENCIA DE CE EN EXPEDICION-COBROS Y BUSCA TAREAS DE GESTION DE COBROS [GC]
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	ceFila = 0
	For ceFila = 5 to 255
		Cell = Sheet.getCellByPosition(0, ceFila)
		vCEIdTarea = Cell.String
		If Cell.String = "" then	'VERIFICA SI LAS PROXIMAS 20 FILAS CONTIENEN INFORMACION
			z = 0
			For z = 1 to 20
				Cell = Sheet.getCellByPosition(0, ceFila + z)
				If Cell.String <> "" then Goto Siguiente
			Next z
			Exit For
		End If
Continua:
		Cell = Sheet.getCellByPosition(1, ceFila)
		vCENroCliente = Cell.String
		Cell = Sheet.getCellByPosition(2, ceFila)
		CadResultado = Cell.String
		Cell.String = CadResultado
		vCENombreClienteDireccion = Cell.String
		Cell = Sheet.getCellByPosition(3, ceFila)
		CadResultado = Cell.String
		Cell.String = CadResultado
		vCEZonaTareas = Cell.String	
		Cell = Sheet.getCellByPosition(4, ceFila)
		vCEPrioridad = Cell.String
		Cell = Sheet.getCellByPosition(5, ceFila)
		CadResultado = Cell.String
		Cell.String = CadResultado
		vCEInfo = Cell.String
		Cell = Sheet.getCellByPosition(6, ceFila)
		CadResultado = Cell.String
		Cell.String = CadResultado
		vCEEstadoAsignado = Cell.String
		Cell = Sheet.getCellByPosition(5, ceFila)
		If Instr(Cell.String,"[GC]") > 0 Then
			z = 0
			For z = 0 to 6
				Cell = Sheet.getCellByPosition(z, ceFila)
				Cell.CellBackColor = RGB(255,102,102) 'ROJO
			Next z
		End If
		vGCIdTarea = ""
		Gosub Paso1
		If vGCIdTarea <> "" then
			z = 0
			For z = 0 to 6
				Cell = Sheet.getCellByPosition(z, ceFila)
				Cell.CellBackColor = RGB(102,255,102) 'VERDE EN CURSO				
			Next z
		End If
Siguiente:
	Next ceFila
	Goto Fin
	
Paso1:
	'BUSCA EQUIV. EN CARGA DE TAREAS DE GC
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	ctFila = 0
	For ctFila = 5 to 255
		Cell = Sheet.getCellByPosition(9, ctFila)
		If Cell.String = vCENroCliente then
			Cell = Sheet.getCellByPosition(10, ctFila)
			If Left(Cell.String, 6) = Left(vCENombreClienteDireccion, 6) then
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, ceFila)
					Cell.CellBackColor = RGB(255,255,0) 'AMARILLO
				Next z
				For z = 8 to 14
					Cell = Sheet.getCellByPosition(z, ctFila)
					Cell.CellBackColor = RGB(255,255,0) 'AMARILLO
				Next z					
				Cell = Sheet.getCellByPosition(11, ctFila)
				If Instr(vCEZonaTareas, "+") = 0 then
					CadBuscar = Mid (vCEZonaTareas, Len(vCEZonaTareas), 1)
				Else
					CadBuscar = Mid (vCEZonaTareas, InStr (vCEZonaTareas, "+") - 1, Len(vCEZonaTareas) - InStr (vCEZonaTareas, "+")+2)
				End If
				If Instr(Cell.String, "+") = 0 then
					CadBuscar2 = Mid (Cell.String, Len(Cell.String), 1)
				Else
					CadBuscar2 = Mid (Cell.String, InStr (Cell.String, "+") - 1, Len(Cell.String) - InStr (Cell.String, "+")+2)
				End If					
				If Instr(CadBuscar, "C") > 0 and Instr(CadBuscar2, "C") > 0 Then Goto SaltoTrue
				If Instr(CadBuscar, "D") > 0 and Instr(CadBuscar2, "D") > 0 Then Goto SaltoTrue
				If Instr(CadBuscar, "O") > 0 and Instr(CadBuscar2, "O") > 0 Then Goto SaltoTrue
				Goto SaltoFalse
				
				SaltoTrue:
				Cell = Sheet.getCellByPosition(13, ctFila)
				If Cell.String = Mid (vCEInfo, InStr (vCEInfo, "[GC]") +5, Len(Cell.String)) then
					Cell = Sheet.getCellByPosition(8, ctFila)
					vGCIdTarea = Cell.String
					z = 0
					For z = 8 to 14
						Cell = Sheet.getCellByPosition(z, ctFila)
						Cell.CellBackColor = RGB(102,255,102) 'VERDE EN CURSO				
					Next z
					
					Sheet = Doc.Sheets.getByName("Carga de Tareas")
					For y = 10000 to 11 Step -1
						Cell = Sheet.getCellByPosition(0, y)
						If Cell.String = vGCIdTarea then
							Cell = Sheet.getCellByPosition(10, y)
							Cell.String = vCEIdTarea
							Cell = Sheet.getCellByPosition(11, y)
							Cell.String = vCEEstadoAsignado
							Exit For
						End If
					Next y
					Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
					Exit For
				End If
			
				SaltoFalse:
			End If
		End If
	Next ctFila
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
Return

Fin:
	'INGRESA LA FECHA Y HORA DE LA ULTIMA MODIFICACION DE LOS ARCHIVOS
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	Cell = Sheet.getCellByPosition(0, 3)
	Cell.String = "Última Modificación del Archivo: " + FileDateTime( RutaOrigen )
	Cell = Sheet.getCellByPosition(8, 3)
	Cell.String = "Última Modificación del Archivo: " + FileDateTime( RutaArchivoActual )
	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(12, 2)
	ModArchivoCE = FileDateTime( RutaOrigen )
	Cell.String = ModArchivoCE
	Cell = Sheet.getCellByPosition(12, 1)
	Cell.String = FileDateTime( RutaArchivoActual )
	
	Msgbox "Actualización finalizada.", 48,"Aviso"
	Procesando = False
	Exit Sub
End Sub


