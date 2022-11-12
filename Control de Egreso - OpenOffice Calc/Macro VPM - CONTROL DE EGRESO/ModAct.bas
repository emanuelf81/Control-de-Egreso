REM  *****  BASIC  *****

Option Explicit

Global RutaOrigen As String
Global ModArchivoGC As String
Global HoraUltGuardar As Long
Global Procesando as Boolean

'Hoja de Calculo
'Dim Doc As Object
'Dim Sheet As Object
'Dim Cell As Object
Dim CellRange As Object
Private dlgCT1 as Object

Sub ActualizarCEyGC
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR UNA NUEVA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if

	Dim oHojaOrigen As Object, ArchOrigen As Object, dFuente As Object ', RutaOrigen As String 
	Dim oDataArrayOrg As Object
	Dim dDestino As Object, RutaArchivoActual As String
	Dim Flags As Long
	Dim oHoja As Object
	
	Dim RutaBaseDatosClientes as String, ModBaseDatosClientes As String, oArchivoAct As Object
	Dim oHojaAct As Object, oRangoAct As Object, oDataAct As Object
	Dim oHojaAct2 As Object, oRangoAct2 As Object, oDataAct2 As Object

	Dim dHojaAct As Object, dRangoAct As Object
	
	Dim vGCIdTarea As String, vGCNroCliente As String, vGCNombreClienteDireccion As String 
	Dim vGCZonaTareas As String, vGCPrioridad As String, vGCInfo As String, vGCEstadoAsignado As String

	Dim vIdTarea as String, vNroCliente as String, vNombre as String, vDireccion as string, vZona as String
	Dim vTarea as String, vPrioridad as String, vInfo as String, vEstado As String, vAsignado As String, vFechaApartir As String
	
	'Contadores y Buscadores
	Dim ceFila
	Dim ctFila
	Dim gcFila
	Dim RangoB as String

	Dim oRango As Object
	Dim dReporteMGR2
	
	Dim mCamposOrden(0) As New com.sun.star.table.TableSortField
	Dim mDescriptorOrden()

	Doc = thiscomponent

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

	Msgbox "La Actualización puede demorar unos minutos."+chr(13)+"Favor de esperar a que el sistema le informe que ha finalizado.", 48,"Importante"

	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator
Inicio:
	'BORRA EL RANGO DE CELDAS DE CE EN EXPEDICION-COBROS
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	CellRange = Sheet.getCellRangeByName("A6:G256")
	Flags = com.sun.star.sheet.CellFlags.STRING
	CellRange.clearContents(Flags)
	CellRange.CellBackColor = RGB(213,231,234) 'CELESTE CLARO
	
	oBarraEstado.start( "Buscando en CE... ", 100 )
	oBarraEstado.setValue( 10 )
Paso1:
	'BUSCA LAS TAREAS PENDIENTES Y EN CURSO DE CE EN CARGA DE TAREAS Y LAS TRANSFIERE A EXPEDICION-COBROS	
	'CARGA Y LIMPIA LA MATRIZ DE CONTROL DE EGRESO
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
    oRango = Sheet.getCellRangeByName("A11:J5011")
    dReporteMGR2 = oRango.getDataArray()
	For y = 0 to 4999
		If dReporteMGR2 (y) (0) = "" then
			If dReporteMGR2 (y+1) (0) = "" then
				oRango.getRows.removeByIndex(y,5000-y)
				dReporteMGR2 = oRango.getDataArray()
				Exit For
			End If
		End if
	Next y
	oBarraEstado.setValue( 20 )

	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	ceFila = 0
	For y = 0 to uBound(dReporteMGR2)
		If dReporteMGR2 (y) (8) = "PENDIENTE" or dReporteMGR2 (y) (8) = "EN CURSO" then
			If ceFila = 50 then oBarraEstado.setValue( 40 )
			If ceFila = 100 then oBarraEstado.setValue( 60 )
			If ceFila = 150 then oBarraEstado.setValue( 70 )
			If ceFila = 250 then oBarraEstado.setValue( 90 )

			Cell = Sheet.getCellByPosition(0, ceFila+5)
			Cell.String = dReporteMGR2 (y) (0)	
			Cell = Sheet.getCellByPosition(1, ceFila+5)
			Cell.String = dReporteMGR2 (y) (1)
			Cell = Sheet.getCellByPosition(2, ceFila+5)
			Cell.String = dReporteMGR2 (y) (2) + chr(13) + dReporteMGR2 (y) (3)
			Cell = Sheet.getCellByPosition(3, ceFila+5)
			Cell.String = dReporteMGR2 (y) (4) + chr(13) + dReporteMGR2 (y) (5)
			Cell = Sheet.getCellByPosition(4, ceFila+5)
			Cell.String = dReporteMGR2 (y) (6)
			Cell = Sheet.getCellByPosition(5, ceFila+5)
			Cell.String = dReporteMGR2 (y) (7)
				If Instr(Cell.String, "[GC]") > 0 then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, ceFila+5)
						Cell.CellBackColor = RGB(255,102,102) 'ROJO
					Next z	
				End If				
			Cell = Sheet.getCellByPosition(6, ceFila+5)
			Cell.String = dReporteMGR2 (y) (8) + chr(13) + dReporteMGR2 (y) (9)
			ceFila = ceFila + 1
			If ceFila > 250 then
				Msgbox "Hay más de 250 tareas Pendientes o En Curso en Control de Egreso"+ chr(13)+"Favor de informarlo a la persona encargada del sistema"
			End If 
		End If
	Next y	  	
	oBarraEstado.setValue( 100 )
	oBarraEstado.end()
Paso2:
	' ORIGEN ACTUALIZACION "GESTION DE COBROS"
	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(11, 1)
	RutaOrigen = ConvertToURL( Cell.String )
	'RutaOrigen = ConvertToURL( "//Servidor/carpetas individuales/Ventas/CONTROL DE EGRESO/VPM - GESTION DE COBROS.ods" )  'RUTA EN VEPROMET
	'RutaOrigen = ConvertToURL( "C:/Users/Emanuel-Not/Documents/EMANUEL/VEPROMET/VPM - GESTION DE COBROS.ods" ) 'RUTA EN CASA
	Cell = Sheet.getCellByPosition(11, 2)
	RutaArchivoActual = ConvertToURL( Cell.String )
	'RutaArchivoActual = ConvertToURL( "//Servidor/carpetas individuales/Ventas/CONTROL DE EGRESO/VPM - CONTROL DE EGRESO.ods" )  'RUTA EN VEPROMET
	'RutaArchivoActual = ConvertToURL( "C:/Users/Emanuel-Not/Documents/EMANUEL/VEPROMET/VPM - CONTROL DE EGRESO.ods" ) 'RUTA EN CASA

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
	'EXTRAE INFORMACION DE GESTION DE COBROS
	oBarraEstado.start( "Actualizando GC... ", 100 )	
	oBarraEstado.setValue( 10 )
	Sheet = Doc.Sheets.getByName("Datos")
	ArchOrigen = StarDesktop.loadComponentFromURL( RutaOrigen, "_blank", 0, Array() )
	oHojaOrigen = ArchOrigen.Sheets.getByName("Expedicion-Cobros")
	dFuente = oHojaOrigen.getCellRangeByName("I6:O256")
	oDataArrayOrg = dFuente.getDataArray()
	ArchOrigen.dispose ()	

	'DESTINO ACTUALIZACION 
	oHoja = Doc.Sheets.getByName("Expedicion-Cobros")
	dDestino = oHoja.getCellRangeByName ("I6:O256")	'SELLECCIONA HOJA DE DESTINO
	Flags = com.sun.star.sheet.CellFlags.STRING		'BORRA INFORMACION EN LAS CELDAS SELECCIONADAS
	dDestino.clearContents(Flags)
	dDestino.CellBackColor = RGB(255,255,255)		'COLOR FONDO BLANCO CELDAS SELECCIONADAS
	dDestino.setDataArray(oDataArrayOrg)			'PEGA LA INFORMACION COPIA EN EL ARCHIVO DE ORIGEN


	'CORRIGE ERRORES DE TRANSFERENCIA Y BUSCA TAREAS DE GC
	oBarraEstado.setValue( 30 )
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	gcFila = 0
	For gcFila = 5 to 255
		Cell = Sheet.getCellByPosition(8, gcFila)
		vGCIdTarea = Cell.String
		Cell = Sheet.getCellByPosition(9, gcFila)
		vGCNroCliente = Cell.String
		Cell = Sheet.getCellByPosition(10, gcFila)
		CadResultado = Cell.String
		Cell.String = Cadresultado
		vGCNombreClienteDireccion = Cell.String
		Cell = Sheet.getCellByPosition(11, gcFila)
		CadResultado = Cell.String
		Cell.String = Cadresultado
		vGCZonaTareas = Cell.String
		Cell = Sheet.getCellByPosition(12, gcFila)
		vGCPrioridad = Cell.String
		Cell = Sheet.getCellByPosition(13, gcFila)
		CadResultado = Cell.String
		Cell.String = Cadresultado
		vGCInfo = Cell.String
		Cell = Sheet.getCellByPosition(14, gcFila)
		CadResultado = Cell.String
		Cell.String = Cadresultado
		vGCEstadoAsignado = Cell.String
		'VERIFICA SI LAS PROXIMAS 10 FILAS CONTIENEN INFORMACION
		If vGCIdTarea = "" then
			z = 0
			For z = 1 to 10
				Cell = Sheet.getCellByPosition(8, gcFila + z)
				If Cell.String <> "" then Goto Continua
			Next z
			Exit For
		End If
		Gosub Paso3
Continua:
	If gcFila = 50 then oBarraEstado.setValue( 40 )
	If gcFila = 100 then oBarraEstado.setValue( 60 )
	If gcFila = 150 then oBarraEstado.setValue( 70 )
	If gcFila = 250 then oBarraEstado.setValue( 90 )
	Next gcFila

	'INGRESA LA FECHA Y HORA DE LA ULTIMA MODIFICACION DE LOS ARCHIVOS
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	Cell = Sheet.getCellByPosition(0, 3)
	Cell.String = "Última Modificación del Archivo: " + FileDateTime( RutaArchivoActual )
	Cell = Sheet.getCellByPosition(8, 3)
	Cell.String = "Última Modificación del Archivo: " + FileDateTime( RutaOrigen )
	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(12, 1)
	ModArchivoGC = FileDateTime( RutaOrigen )
	Cell.String = ModArchivoGC
	Cell = Sheet.getCellByPosition(12, 2)
	Cell.String = FileDateTime( RutaArchivoActual )

	Doc.Store()
	HoraUltGuardar = Timer
	oBarraEstado.setValue( 100 )
	oBarraEstado.end()	
	Procesando = False
	Msgbox "Actualización Finalizada.", 48,"Aviso"
	Exit Sub

Paso3:
	'VERIFICA LA INFORMACION RECIBIDA DE GC Y LAS BUSCA EN CE
	ceFila = 0
	For ceFila = 5 to 255
		Cell = Sheet.getCellByPosition(1, ceFila)
		If Cell.String = vGCNroCliente then
			Cell = Sheet.getCellByPosition(2, ceFila)
			If Left(Cell.String, 10) = Left(vGCNombreClienteDireccion, 10) then
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, ceFila)
					Cell.CellBackColor = RGB(255,255,0) 'AMARILLO
				Next z
				For z = 8 to 14
					Cell = Sheet.getCellByPosition(z, gcFila)
					Cell.CellBackColor = RGB(255,255,0) 'AMARILLO
				Next z
				Cell = Sheet.getCellByPosition(3, ceFila)
				If Instr(vGCZonaTareas, "+") = 0 then
					CadBuscar = Mid(vGCZonaTareas, Len(vGCZonaTareas), 1)
				Else 
					CadBuscar = Mid(vGCZonaTareas, InStr (vGCZonaTareas, "+") - 1, Len(vGCZonaTareas) - InStr (vGCZonaTareas, "+")+2)
				End If
				If Instr(Cell.String, "+") = 0 then
					CadBuscar2 = Mid(Cell.String, Len(Cell.String), 1)
				Else 
					CadBuscar2 = Mid(Cell.String, InStr(Cell.String, "+")-1, Len(Cell.String) - InStr(Cell.String, "+")+2)
				End If
				If Instr(CadBuscar, "C") > 0 and Instr(CadBuscar2, "C") > 0 Then Goto SaltoTrue
				If Instr(CadBuscar, "D") > 0 and Instr(CadBuscar2, "D") > 0 Then Goto SaltoTrue
				If Instr(CadBuscar, "O") > 0 and Instr(CadBuscar2, "O") > 0 Then Goto SaltoTrue
				Goto SaltoFalse

				SaltoTrue:
				Cell = Sheet.getCellByPosition(5, ceFila)
				If InStr (Cell.String, "[GC]") > 0 and Mid (Cell.String, InStr (Cell.String, "[GC]") +5, Len(vGCInfo)) = vGCInfo then
					'COLOREA FONDO CELDA
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, ceFila)

						Cell.CellBackColor = RGB(102,255,102) 'VERDE EN CURSO
					Next z
					z = 0
					For z = 8 to 14
						Cell = Sheet.getCellByPosition(z, gcFila)
						Cell.CellBackColor = RGB(102,255,102) 'VERDE EN CURSO
					Next z
					Exit For
				End If
				SaltoFalse:
			End If
		End If
	Next ceFila
Return
End Sub


Sub CargarTareasGC
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR UNA NUEVA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if

	Dim dlgCT1 as Object
	Dim dlgCT12 as Object
	Dim vGCIdTarea As String, vGCNroCliente As String, vGCNombreClienteDireccion As String 
	Dim vGCZonaTareas As String, vGCPrioridad As String, vGCInfo As String, vGCEstadoAsignado As String, vGCFechaApartir As String
	Dim vCEIdTarea As String, vCENroCliente As String, vCENombreClienteDireccion As String 
	Dim vCEZonaTareas As String, vCEPrioridad As String, vCEInfo As String, vCEEstadoAsignado As String
	'Contadores y Buscadores
	Dim ceFila
	Dim ctFila
	Dim gcFila
	Dim ctUltFila
	Dim FilaInicial, FilaFinal	
	
	Doc = thiscomponent
	DialogLibraries.LoadLibrary("Standard")

	'VERIFICA QUE SE HAYA INGRESADO EL USUARIO
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
	
	FilaInicial = 5
	FilaFinal = 255
	
	BuscaActualizacionesGC
Inicio:

Paso1:
	'BUSCA EN CE "[GC]" Y LO MARCA POR POSIBLES ERRORES O NO ENCONTRADO
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	ctUltFila = 0
	ceFila = 0
	For ceFila = FilaInicial to FilaFinal
		Cell = Sheet.getCellByPosition(0, ceFila)
		If Cell.String = "" then
			z = 0
			For z = 1 to 10
				Cell = Sheet.getCellByPosition(0, ceFila + Z)
				If Cell.String <> "" then Goto ContinuaControlCE
			Next z
			Exit For
		End If
		Cell = Sheet.getCellByPosition(5, ceFila)
		If Instr(Cell.String, "[GC]") > 0 then
			z = 0
			For z = 0 to 6
				Cell = Sheet.getCellByPosition(z, ceFila)
				Cell.CellBackColor = RGB(255,102,102) 'ROJO
			Next z
		else
			z = 0
			For z = 0 to 6
				Cell = Sheet.getCellByPosition(z, ceFila)
				Cell.CellBackColor = RGB(213,231,234) 'CELESTE CLARO
			Next z				
		End If
ContinuaControlCE:
	Next ceFila

	'BUSCA TAREAS DE GC EN E-C
	gcFila = 0
	For gcFila = FilaInicial to FilaFinal
		Cell = Sheet.getCellByPosition(8, gcFila)
		vGCIdTarea = Cell.String
		Cell.CellBackColor = RGB(255,255,255) 'BLANCO
		Cell = Sheet.getCellByPosition(9, gcFila)
		vGCNroCliente = Cell.String
		Cell.CellBackColor = RGB(255,255,255) 'BLANCO
		Cell = Sheet.getCellByPosition(10, gcFila)
		vGCNombreClienteDireccion = Cell.String
		Cell.CellBackColor = RGB(255,255,255) 'BLANCO
		Cell = Sheet.getCellByPosition(11, gcFila)
		vGCZonaTareas = Cell.String
		Cell.CellBackColor = RGB(255,255,255) 'BLANCO
		Cell = Sheet.getCellByPosition(12, gcFila)
		vGCPrioridad = Cell.String
		Cell.CellBackColor = RGB(255,255,255) 'BLANCO
		Cell = Sheet.getCellByPosition(13, gcFila)
		vGCInfo = Cell.String
		Cell.CellBackColor = RGB(255,255,255) 'BLANCO
		Cell = Sheet.getCellByPosition(14, gcFila)
		vGCFechaApartir = Cell.String	
		Cell.CellBackColor = RGB(255,255,255) 'BLANCO
		'VERIFICA SI LAS PROXIMAS 10 FILAS CONTIENEN INFORMACION
		If vGCIdTarea = "" then
			z = 0
			For z = 1 to 10
				Cell = Sheet.getCellByPosition(8, gcFila + z)
				If Cell.String <> "" then Goto ContinuaBuscaGC
			Next z
			Exit For
		End If
		vCEIdTarea = ""
		vIdTarea = ""
		Gosub BuscaCE
		If vCEIdTarea <> "" and vIdTarea <> "" then
			'Msgbox "TAREA CARGADA"
		End If
		If vCEIdTarea <> "" and vIdTarea = "" then gosub ModificarTarea
		If vCEIdTarea = "" and vIdTarea = "" then gosub NuevaTarea
ContinuaBuscaGC:
	Next gcFila
	Goto Final

BuscaCE:
	'CORROBORA LA INFORMACION DE GC EN CE
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	ceFila = 0
	For ceFila = FilaInicial to FilaFinal
		Cell = Sheet.getCellByPosition(1, ceFila)
		If Cell.String = vGCNroCliente then
			Cell = Sheet.getCellByPosition(2, ceFila)
			If Left(Cell.String, 10) = Left(vGCNombreClienteDireccion, 10) then
				Cell = Sheet.getCellByPosition(0, ceFila)
				vCEIdTarea = Cell.String
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, ceFila)
					Cell.CellBackColor = RGB(255,255,0) 'AMARILLO
				Next z
				For z = 8 to 14
					Cell = Sheet.getCellByPosition(z, gcFila)
					Cell.CellBackColor = RGB(255,255,0) 'AMARILLO
				Next z
				Cell = Sheet.getCellByPosition(3, ceFila)
				CadBuscar = ""
				CadBuscar2 = ""
				If Instr(vGCZonaTareas, "+") = 0 then
					CadBuscar = Mid(vGCZonaTareas, Len(vGCZonaTareas), 1)
				Else 
					CadBuscar = Mid(vGCZonaTareas, InStr (vGCZonaTareas, "+") - 1, Len(vGCZonaTareas) - InStr (vGCZonaTareas, "+")+2)
				End If
				If Instr(Cell.String, "+") = 0 then
					CadBuscar2 = Mid(Cell.String, Len(Cell.String), 1)
				Else 
					CadBuscar2 = Mid(Cell.String, InStr(Cell.String, "+")-1, Len(Cell.String) - InStr(Cell.String, "+")+2)
				End If
				If Instr(CadBuscar, "C") > 0 and Instr(CadBuscar2, "C") > 0 Then Goto SaltoTrue
				If Instr(CadBuscar, "D") > 0 and Instr(CadBuscar2, "D") > 0 Then Goto SaltoTrue
				If Instr(CadBuscar, "O") > 0 and Instr(CadBuscar2, "O") > 0 Then Goto SaltoTrue
				Goto SaltoFalse

				SaltoTrue:
				Cell = Sheet.getCellByPosition(5, ceFila)
				If InStr (Cell.String, "[GC]") > 0 and Mid (Cell.String, InStr (Cell.String, "[GC]") +5, Len(vGCInfo)) = vGCInfo then
					'COLOREA FONDO CELDA
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, ceFila)
'						Cell.CellBackColor = RGB(229,202,255) 'PURPURA CLARO
						Cell.CellBackColor = RGB(102,255,102) 'VERDE EN CURSO
					Next z
					z = 0
					For z = 8 to 14
						Cell = Sheet.getCellByPosition(z, gcFila)
'						Cell.CellBackColor = RGB(207,231,245) 'CELESTE CLARO
						Cell.CellBackColor = RGB(102,255,102) 'VERDE EN CURSO
					Next z
					vIdTarea = vCEIdTarea
					Exit For
				End If
				SaltoFalse:
			End If
		End If
	Next ceFila
Return
		
NuevaTarea:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	dlgCT1 = createUnoDialog(DialogLibraries.Standard.Dialog1)
	
	'BUSCA LA PRIMERA CELDA DISPONIBLE
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	If vProxTarea = 0 or nProxFila = 0 then
		BuscaProxTareaDisponibleCT
	Else
		If nProxFila > 0 then NuevaFilaCT
	End If
	yIDT = nProxFila	'eliminar o no
	If nProxFila = 0 then 
		Procesando = False
		Exit Sub
	End If
	If vIdTarea = "" then
		Procesando = False
		Exit Sub
	End If
	vNroCliente	= vGCNroCliente
	If vNroCliente = "0" then 
		vNombre = vGCNombreClienteDireccion	'VER SI SE PUEDE DIVIDIR
		vDireccion = ""						'VER SI SE PUEDE DIVIDIR
		vZona = ""
		If InStr(vGCZonaTareas, "CBAN") > 0 then vZona = "CBAN" 
		If InStr(vGCZonaTareas, "CBAC") > 0 then vZona = "CBAC"
		If InStr(vGCZonaTareas, "CBAS") > 0 then vZona = "CBAS"
		If InStr(vGCZonaTareas, "INT") > 0 then vZona = "INT"
	Else
		y = 0
		Sheet = Doc.Sheets.getByName("BDClientes")
		For y = 2 to 20002 'Número máximo de filas a buscar en BDClientes.
			Cell = Sheet.getCellByPosition(0, y)
			if Cell.String = vNroCliente then 
				Cell = Sheet.getCellByPosition(1, y)
				vNombre = Cell.String
				Cell = Sheet.getCellByPosition(2, y)
				vDireccion = Cell.String
				Cell = Sheet.getCellByPosition(3, y)
				vDireccion = vDireccion + ", " + Cell.String
				Cell = Sheet.getCellByPosition(4, y)
				vDireccion = vDireccion + ", " + Cell.String
				Cell = Sheet.getCellByPosition(5, y)
				vZona = Cell.String
				Exit For 
			End if
			If Cell.String = "" then
				MsgBox "Nro. de Cliente no encontrado en la base de datos"
				Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
				Return
			End If
		Next y
		Sheet = Doc.Sheets.getByName("Carga de Tareas")
	End If
	
	vTarea = ""
	CadBuscar = ""
	If Instr(vGCZonaTareas, "+") = 0 then
		CadBuscar = Mid(vGCZonaTareas, Len(vGCZonaTareas), 1)
	Else 
		CadBuscar = Mid(vGCZonaTareas, InStr (vGCZonaTareas, "+") - 1, Len(vGCZonaTareas) - InStr (vGCZonaTareas, "+")+2)
	End If
	If Instr(CadBuscar, "C") > 0 then vTarea = "C"
	If Instr(CadBuscar, "D") > 0 then vTarea = vTarea + "D"
	If Instr(CadBuscar, "O") > 0 then vTarea = vTarea + "O"

	vPrioridad = vGCPrioridad

	vInfo = "[GC] " + vGCInfo
	
	vEstado = "PENDIENTE"
	
	vFechaApartir = vGCFechaApartir
	
	vAsignado = ""
	

	'CARGA LISTBOX ASIGNADO
	olstDatos = dlgCT1.getControl("ComboBox3")	
	oHojaDatos = ThisComponent.getSheets.getByName("Datos")	
	oRango = oHojaDatos.getCellRangeByName("C3:C10") 'ASIGNADO
	data = oRango.getDataArray()'agregado
	co1 = 0
	Redim src(UBound(data))
	For Each d In data
    	src(co1) = d(0)
      	co1 = co1 + 1
	Next
   	olstDatos.addItems(src, 0)

	' Carga las Variables en el dialogo
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	dlgCT1.Model.TextField2.Text = vIdTarea
	dlgCT1.Model.TextField3.Text = vNroCliente 
	dlgCT1.Model.TextField4.Text = vNombre
'	dlgCT1.Model.TextField6.Text = vNombre
'	dlgCT1.Model.TextField7.Text = vDireccion
	dlgCT1.Model.TextField5.Text = vDireccion
	dlgCT1.Model.DateField1.text = vFechaApartir
	dlgCT1.Model.TextField1.Text = vInfo
	dlgCT1.Model.ComboBox1.text = vZona
'	dlgCT1.Model.ComboBox2.text = vZona
	dlgCT1.Model.ComboBox3.text = vAsignado

	dlgCT1.Model.CheckBox1.State = 0
	dlgCT1.Model.CheckBox2.State = 0
	dlgCT1.Model.CheckBox3.State = 0
	dlgCT1.Model.CheckBox4.State = 0
	dlgCT1.Model.CheckBox5.State = 0
	If Instr(vTarea, "C") > 0 then dlgCT1.Model.CheckBox2.State = 1
	If Instr(vTarea, "D") > 0 then dlgCT1.Model.CheckBox3.State = 1
	If Instr(vTarea, "O") > 0 then dlgCT1.Model.CheckBox4.State = 1
	
	If vPrioridad = "ALTA" then dlgCT1.Model.OptionButton1.State = 1
	If vPrioridad = "MEDIA" then dlgCT1.Model.OptionButton2.State = 1
	If vPrioridad = "BAJA" then dlgCT1.Model.OptionButton3.State = 1	

	If vEstado = "PENDIENTE" then dlgCT1.Model.OptionButton4.State = 1
	If vEstado = "EN CURSO" then dlgCT1.Model.OptionButton5.State = 1
	If vEstado = "FINALIZADO" then dlgCT1.Model.OptionButton6.State = 1	

	dlgCT1.Model.ComboBox3.text = vAsignado
	dlgCT1.Model.DateField1.text = vFechaApartir	

	'ABRE DIALOG1 PARA CARGAR LA TAREA
	dlgCT1.Model.Step = 1
	Select Case dlgCT1.Execute()
	Case 1
		vNombre = dlgCT1.Model.TextField4.Text
		vDireccion = dlgCT1.Model.TextField5.Text
		vZona = dlgCT1.Model.ComboBox1.text
		if vZona = "" then
			Msgbox "No ha especificado la zona"
		End if
		vInfo =  dlgCT1.Model.TextField1.Text
		vTarea = ""
		if dlgCT1.Model.CheckBox1.State = 1 then vTarea = vTarea + "E"
		if dlgCT1.Model.CheckBox2.State = 1 then vTarea = vTarea + "C"
		if dlgCT1.Model.CheckBox3.State = 1 then vTarea = vTarea + "D"
		if dlgCT1.Model.CheckBox4.State = 1 then vTarea = vTarea + "O"
		if dlgCT1.Model.CheckBox5.State = 1 then vTarea = vTarea + "V"
		CadBuscar = vTarea
		if vTarea = "" then
			Msgbox "No ha especificado cual es la tarea a realizar"
		End if
		if Len(CadBuscar) = 1 then vTarea = CadBuscar
		if Len(CadBuscar) = 2 then vTarea = Mid(CadBuscar, 1, 1) + "+" + Mid(CadBuscar, 2, 1)
		if Len(CadBuscar) = 3 then vTarea = Mid(CadBuscar, 1, 1) + "+" + Mid(CadBuscar, 2, 1) + "+" + Mid(CadBuscar, 3, 1)
		if Len(CadBuscar) = 4 then vTarea = Mid(CadBuscar, 1, 1) + "+" + Mid(CadBuscar, 2, 1) + "+" + Mid(CadBuscar, 3, 1) + "+" + Mid(CadBuscar, 4, 1)
		if Len(CadBuscar) = 5 then vTarea = Mid(CadBuscar, 1, 1) + "+" + Mid(CadBuscar, 2, 1) + "+" + Mid(CadBuscar, 3, 1) + "+" + Mid(CadBuscar, 4, 1) + "+" + Mid(CadBuscar, 5, 1)
		if dlgCT1.Model.OptionButton1.State = 1 then vPrioridad = "ALTA"	
		if dlgCT1.Model.OptionButton2.State = 1 then vPrioridad = "MEDIA"
		if dlgCT1.Model.OptionButton3.State = 1 then vPrioridad = "BAJA"	
		if dlgCT1.Model.OptionButton4.State = 1 then vEstado = "PENDIENTE"	
		if dlgCT1.Model.OptionButton5.State = 1 then vEstado = "EN CURSO"
		if dlgCT1.Model.OptionButton6.State = 1 then vEstado = "FINALIZADO"
		vFechaApartir = dlgCT1.Model.DateField1.text
		vAsignado = dlgCT1.Model.ComboBox3.text
	Case 0
		dlgCT1.Dispose()
		Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
		Return
	End Select
'	ctFila = 0
'	For ctFila = 11 to 10011 'Número máximo de filas a buscar.
'		Cell = Sheet.getCellByPosition(0, yIDT)
'		If Cell.String = vIdTarea then 
			Cell = Sheet.getCellByPosition(1, yIDT)
			Cell.String = vNroCliente
			Cell = Sheet.getCellByPosition(2, yIDT)
			Cell.String = vNombre
			Cell = Sheet.getCellByPosition(3, yIDT)
			Cell.String = vDireccion
			Cell = Sheet.getCellByPosition(4, yIDT)
			Cell.String = vZona
			Cell = Sheet.getCellByPosition(5, yIDT)
			Cell.String = vTarea
			Cell = Sheet.getCellByPosition(6, yIDT)
			Cell.String = vPrioridad
			Cell = Sheet.getCellByPosition(7, yIDT)
			Cell.String = vInfo
			Cell = Sheet.getCellByPosition(8, yIDT)
			Cell.String = vEstado
			Cell = Sheet.getCellByPosition(9, yIDT)
			Cell.String = vAsignado
			Cell = Sheet.getCellByPosition(10, yIDT)
			Cell.String = "" 'ver si contiene info anterior x si modifca
			Cell = Sheet.getCellByPosition(11, yIDT)
			Cell.String = ""
			If vEstado = "FINALIZADO" then Cell.String = DATE
			Cell = Sheet.getCellByPosition(12, yIDT)
			Cell.String = vFechaApartir
			Cell = Sheet.getCellByPosition(13, yIDT)
			Cell.String = Date
			Cell = Sheet.getCellByPosition(14, yIDT)
			Cell.String = Date
			Cell = Sheet.getCellByPosition(15, yIDT)
			Cell.String = vUsuario
			Cell = Sheet.getCellByPosition(16, yIDT)
			Cell.Value = 0
			Cell = Sheet.getCellByPosition(17, yIDT)
			Cell.Value = 0
			Cell = Sheet.getCellByPosition(18, yIDT)
			Cell.String = ""
			Cell = Sheet.getCellByPosition(19, yIDT)
			Cell.String = ""
			Cell = Sheet.getCellByPosition(20, yIDT)
			Cell.String = ""
			Cell = Sheet.getCellByPosition(21, yIDT)
			Cell.String = ""
			Cell = Sheet.getCellByPosition(22, yIDT)
			Cell.String = ""
			'Colorea el fondo de la celdas
			If vEstado = "PENDIENTE" Then					
				z = 0
				For z = 1 to 14
					Cell = Sheet.getCellByPosition(z, yIDT)
					Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
				Next z
			End If
			If vEstado = "EN CURSO" Then					
				z = 0
				For z = 1 to 14
					Cell = Sheet.getCellByPosition(z, yIDT)
					Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
				Next z
			End If
			If vEstado = "FINALIZADO" Then					
				z = 0
				For z = 1 to 14
					Cell = Sheet.getCellByPosition(z, yIDT)
					Cell.CellBackColor = RGB(255,102,102) 'FINALIZADO
				Next z
			End If
			z = 0
			For z = 15 to 22
				Cell = Sheet.getCellByPosition(z, yIDT)
				Cell.CellBackColor = RGB(221,221,221) 'COLOR COLUMNA CONTROLO
			Next z
'			ctUltFila = ctUltFila + 1
'			Exit For
'		End If 
'	Next ctFila

	'ACTUALIZA CE SI EL ESTADO ES PENDIENTE o EN CURSO 
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	If vEstado = "PENDIENTE" or vEstado = "EN CURSO" then
		ceFila = 0
		For ceFila = FilaInicial to FilaFinal 'Número
			Cell = Sheet.getCellByPosition(0, ceFila)
			If Cell.String = "" then
				Cell.String = vIdTarea
				vCEIdTarea = Cell.String
				Cell = Sheet.getCellByPosition(1, ceFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(2, ceFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(3, ceFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(4, ceFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(5, ceFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(6, ceFila)
				Cell.String = vEstado + chr(13) + vAsignado
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, ceFila)
					'Cell.CellBackColor = RGB(229,202,255) 'PURPURA CLARO
					Cell.CellBackColor = RGB(102,255,102) 'VERDE
				Next z
				z = 0
				For z = 8 to 14
					Cell = Sheet.getCellByPosition(z, gcFila)
					'Cell.CellBackColor = RGB(207,231,245) 'CELESTE CLARO
					Cell.CellBackColor = RGB(102,255,102) 'VERDE
				Next z
				Exit For
			End If
		Next ceFila
	End If

	dlgCT1.Dispose()
Return

ModificarTarea:
	'AGREGA UNA TAREA DE GC EN CE
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	DialogLibraries.LoadLibrary("Standard")
	dlgCT12 = createUnoDialog(DialogLibraries.Standard.Dialog12)
	ceFila = 0
	For ceFila = FilaInicial to FilaFinal
		Cell = Sheet.getCellByPosition(0, ceFila)
		If vCEIdTarea = Cell.String then
			Cell = Sheet.getCellByPosition(1, ceFila)
			vCENroCliente = Cell.String
			Cell = Sheet.getCellByPosition(2, ceFila)
			vCENombreClienteDireccion = Cell.String
			Cell = Sheet.getCellByPosition(3, ceFila)
			vCEZonaTareas = Cell.String
			Cell = Sheet.getCellByPosition(4, ceFila)
			vCEPrioridad = Cell.String
			Cell = Sheet.getCellByPosition(5, ceFila)
			vCEInfo = Cell.String
			Cell = Sheet.getCellByPosition(6, ceFila)
			vCEEstadoAsignado = Cell.String
			Exit For			
		End If
	Next ceFila
		
	dlgCT12.Model.TextField1.Text = ""
	dlgCT12.Model.TextField1.Text = "Id.Tarea: " + vCEIdTarea
	dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Nro.Cliente: " + vCENroCliente
	dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Cliente/Destinatario: " + vCENombreClienteDireccion
	If Instr(vCEZonaTareas, "+") = 0 then
		dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Zona: " + Mid (vCEZonaTareas, 1, Len(vCEZonaTareas)-2)
		dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Tarea: " + Mid (vCEZonaTareas, Len(vCEZonaTareas), 1)
	Else 
		dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Zona: " + Mid (vCEZonaTareas, 1, InStr (vCEZonaTareas, "+")-3)
		dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Tarea: " + Mid (vCEZonaTareas, InStr (vCEZonaTareas, "+") - 1, Len(vCEZonaTareas) - InStr (vCEZonaTareas, "+")+2)
	End If
	dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Prioridad: " + vCEPrioridad
	dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Comentarios:" + chr(13) + vCEInfo
	If Left(vCEEstadoAsignado, 8) = "EN CURSO" then
		dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Estado: EN CURSO"  +chr(13)+ "Asignado: " + Mid (vCEEstadoAsignado, 10, Len(vCEEstadoAsignado))
	Else
		dlgCT12.Model.TextField1.Text = dlgCT12.Model.TextField1.Text + chr(13) + "Estado: PENDIENTE"  +chr(13)+ "Asignado: " + Mid (vCEEstadoAsignado, 11, Len(vCEEstadoAsignado))
	End If
	
	dlgCT12.Model.TextField2.Text = ""
	dlgCT12.Model.TextField2.Text = "Id.Tarea: " + vGCIdTarea
	dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Nro.Cliente: " + vGCNroCliente
	dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Cliente/Destinatario: " + vGCNombreClienteDireccion
	If Instr(vGCZonaTareas, "+") = 0 then
		dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Zona: " + Mid (vGCZonaTareas, 1, Len(vGCZonaTareas)-2)
		dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Tarea: " + Mid (vGCZonaTareas, Len(vGCZonaTareas), 1)
	Else 
		dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Zona: " + Mid (vGCZonaTareas, 1, InStr (vGCZonaTareas, "+")-3)
		dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Tarea: " + Mid (vGCZonaTareas, InStr (vGCZonaTareas, "+") - 1, Len(vGCZonaTareas) - InStr (vGCZonaTareas, "+")+2)
	End If
	dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Prioridad: " + vGCPrioridad
	dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "Comentarios:" + chr(13) + vGCInfo
	dlgCT12.Model.TextField2.Text = dlgCT12.Model.TextField2.Text + chr(13) + "A partir de: " + vGCFechaApartir

	'ABRE EL DIALOGO PARA MODIFICAR UNA TAREA
	Select Case dlgCT12.Execute()
	Case 1
		If Left(vCEEstadoAsignado, 8) = "EN CURSO" then
			If Msgbox( "Se dispone a modificar una tarea EN CURSO"+chr(13)+CHR(13)+"¿Desea modificarla?"+chr(13), 4 + 32, "IMPORTANTE" ) = 7 then
				goto NoModifica
			End if
		End If
		ceFila = 0
		For ceFila = FilaInicial to FilaFinal
		Cell = Sheet.getCellByPosition(0, ceFila)
			If vCEIdTarea = Cell.String then
				vIdTarea = vCEIdTarea
				
				Cell = Sheet.getCellByPosition(3, ceFila)
				CadBuscar = ""
				If Instr(vGCZonaTareas, "+") = 0 then
					CadBuscar = Mid(vGCZonaTareas, Len(vGCZonaTareas), 1)
				Else 
					CadBuscar = Mid(vGCZonaTareas, InStr (vGCZonaTareas, "+") - 1, Len(vGCZonaTareas) - InStr (vGCZonaTareas, "+")+2)
				End If
				Cell.String = vCEZonaTareas + "+" + CadBuscar 
				vTarea = CadBuscar 
								
				Cell = Sheet.getCellByPosition(4, ceFila)
				If vGCPrioridad = "ALTA" and vCEPrioridad = "MEDIA" then Cell.String = vGCPrioridad
				If vGCPrioridad = "MEDIA" and vCEPrioridad = "BAJA" then Cell.String = vGCPrioridad
				If vGCPrioridad = "ALTA" and vCEPrioridad = "BAJA" then Cell.String = vGCPrioridad
				vPrioridad = Cell.String
				
				Cell = Sheet.getCellByPosition(5, ceFila)
				Cell.String = vCEInfo + chr(13) + "[GC] " + vGCInfo 
				vInfo = Cell.String
				
				
		
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, ceFila)
					'Cell.CellBackColor = RGB(229,202,255) 'PURPURA CLARO
					Cell.CellBackColor = RGB(102,255,102) 'VERDE
				Next z
				z = 0
				For z = 8 to 14
					Cell = Sheet.getCellByPosition(z, gcFila)
					'Cell.CellBackColor = RGB(207,231,245) 'CELESTE CLARO
					Cell.CellBackColor = RGB(102,255,102) 'VERDE
				Next z
				Sheet = Doc.Sheets.getByName("Carga de Tareas")
				If ctUltFila = 0 then
					y = 10011
				Else
					y = ctUltFila
				End If
				ctFila = 0 	
				For ctFila = y to 11 Step -1
					Cell = Sheet.getCellByPosition(0, ctFila)
					If Cell.String = vIdTarea then
						Cell = Sheet.getCellByPosition(5, ctFila)
						Cell.String = Cell.String + "+" + vTarea
						Cell = Sheet.getCellByPosition(6, ctFila)
						Cell.String = vPrioridad
						Cell = Sheet.getCellByPosition(7, ctFila)
						Cell.String = vInfo
						Cell = Sheet.getCellByPosition(12, ctFila)
						Cell.String = vGCFechaApartir
						Exit For
					End If
				Next ctFila	
				Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
				Exit For
			End If
		Next ceFila

	Case 0
		vIdTarea = ""
	End Select
NoModifica:

dlgCT12.Dispose()

If vIdTarea = "" then
	vCEIdTarea = ""
End If

Return
		

Final:
	Doc.Store()
	HoraUltGuardar = Timer
	Procesando = False
	Msgbox "Finalizado.", 48,"Aviso"
End Sub

