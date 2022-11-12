REM  *****  BASIC  *****

Option Explicit

Private dlgCT1 as Object, dlgCT2 as Object, dlgCT3 as Object, dlgCT4 as Object, dlgCT13 as Object, dlgCT6 as Object
Private dlgCT7 as Object, dlgCT8 as Object, dlgCT9 as Object, dlgCT10 as Object, dlgCT11 as Object

Dim cmdBoton as String
Dim InfoMostrar As String
Global vProxTarea As Integer, nProxFila As Integer

Dim vIdTarea as String, vNroCliente as String, vNombre as String, vDireccion as string
Dim vZona as String, vTarea as String, vPrioridad as String, vInfo as String 
Dim vEstado As String, vAsignado As String, vConcurrio As String, vObjetivo As String
Dim vFechaApartir As String, vFechaCarga As String, vFechaFinalizado as String
Dim vUltMod As String

Dim bTareaExtraHR As Boolean
Dim vAsignadoHR As String
Dim vFechaHR As String

Dim oBarraEstado As Object
Dim vProgBar As Integer

'Dim CellRangeAddress As New com.sun.star.table.CellRangeAddress
'Dim CellAddress As New com.sun.star.table.CellAddress

Dim Doc As Object
Dim Sheet As Object
Dim Cell As Object

'Variables ListBox de los dialogos
Dim oHojaDatos As Object
Dim co1 As Long
Dim oRango As Object
Dim data
Dim src
Dim d

Dim xIDT, yIDT, x, y, nFila, z
Dim CadBuscar As String, CadBuscar1 As String, CadBuscar2 As String, CadResultado As String
Dim Pos1, Pos2

Dim document   as object
Dim dispatcher as object	
Dim PosCel(0) as new com.sun.star.beans.PropertyValue
Dim Posicionador As String
Dim Posicionar

Dim chkFEFinalizado As Object
Dim chkFEPendiente As Object
Dim chkFEEnCurso As Object

Global FilaActual

Sub FiltroEstado

	Doc = thiscomponent
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
    chkFEFinalizado = oFormulario.getByName( "CVerEstado1" )
    chkFEPendiente = oFormulario.getByName( "CVerEstado2" )
    chkFEEnCurso = oFormulario.getByName( "CVerEstado3" ) 
 	
 	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator
 	
 	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	oBarraEstado.start( "Filtrando Estado ", 2000 )
	For y = 10 to 5000
		If y = 100 or y = 300 or y = 600 Then oBarraEstado.setValue( y )
		If y = 800 or y = 1000 or y = 1500 Then oBarraEstado.setValue( y )
		If y = 2000 or y = 3000 or y = 4000 Then oBarraEstado.setValue( y )

		Cell = Sheet.getCellByPosition(8, y)
		If Cell.String <> "" then
			If Cell.String = "FINALIZADO" then
				If chkFEFinalizado.State = 1 then
					Sheet.getRows.getByindex(y).IsVisible = True
				Else
					Sheet.getRows.getByindex(y).IsVisible = False
				End If
			End If
			If Cell.String = "PENDIENTE" then
				If chkFEPendiente.State = 1 then
					Sheet.getRows.getByindex(y).IsVisible = True
				Else
					Sheet.getRows.getByindex(y).IsVisible = False
				End If
			End If
			If Cell.String = "EN CURSO" then
				If chkFEEnCurso.State = 1 then
					Sheet.getRows.getByindex(y).IsVisible = True
				Else
					Sheet.getRows.getByindex(y).IsVisible = False
				End If
			End If
		Else
			Posicionar = y
			PosicionadorCelda		
			Exit For
		End If
	Next Y
	oBarraEstado.end()
End Sub

Sub FilaVisible
	If Sheet.getRows.getByindex(FilaActual).IsVisible = False then
		Sheet.getRows.getByindex(FilaActual).IsVisible = True
	End IF
End Sub

Sub FilaNoVisible
	Doc = thiscomponent
    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" )
    chkFEFinalizado = oFormulario.getByName( "CVerEstado1" )
    chkFEPendiente = oFormulario.getByName( "CVerEstado2" )
    chkFEEnCurso = oFormulario.getByName( "CVerEstado3" ) 	
 	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	Cell = Sheet.getCellByPosition(8, FilaActual)
	If chkFEFinalizado.State = 0 and Cell.String = "FINALIZADO" then	
		Sheet.getRows.getByindex(FilaActual).IsVisible = False
	End If
	If chkFEPendiente.State = 0 and Cell.String = "PENDIENTE" then	
		Sheet.getRows.getByindex(FilaActual).IsVisible = False
	End If
	If chkFEEnCurso.State = 0 and Cell.String = "EN CURSO" then	
		Sheet.getRows.getByindex(FilaActual).IsVisible = False
	End If
End Sub


Sub CursarTareasCT
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	Dim cmdBotonSig as object

Inicio:
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent
	DialogLibraries.LoadLibrary("Standard")

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
		
	BuscaActualizacionesGC
	
Paso1:
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 	
	dlgCT10 = createUnoDialog(DialogLibraries.Standard.Dialog10)
	cmdBotonSig = dlgCT10.getControl("cmdSiguiente")
Paso2:
	'Carga ListCBox dialogo asignado de Dialog10
	Sheet = Doc.Sheets.getByName("Datos")
	olstDatos = dlgCT10.getControl("lstCBox1")
  	For d = 2 to 10	
	 	Cell = Sheet.getCellByPosition(2, d) 	
	  	vAsignado = Cell.String
	  	If vAsignado <> "" then olstDatos.addItem( vAsignado, -1 )
	Next d

	'Busca las tareas que se encuentran en Estado PENDIENTE
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	yIDT = 0
	For yIDT = 10 to 8000
		Cell = Sheet.getCellByPosition(0, yIDT)
		If Cell.String = "" then
			Cell = Sheet.getCellByPosition(0, yIDT + 1)
			If Cell.String = "" then
				Msgbox "Finalizado.",64,"AVISO"
				If Msgbox( "¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información.?", 4 + 32, "Guardar" ) = 6 then 
					Doc.Store()
					HoraUltGuardar = Timer
				End If
				Procesando = False
				Exit Sub
			End If
		End If
		Cell = Sheet.getCellByPosition(8, yIDT)
		If Cell.String = "PENDIENTE" then 
			Posicionar = yIDT
			PosicionadorCelda
			' Carga los valores de la Fila xIDT a las variables
			Cell = Sheet.getCellByPosition(0, yIDT)
			vIdTarea = Cell.getString
			Cell = Sheet.getCellByPosition(1, yIDT)
			vNroCliente = Cell.getString
			Cell = Sheet.getCellByPosition(2, yIDT)
			vNombre = Cell.getString
			Cell = Sheet.getCellByPosition(3, yIDT)
			vDireccion = Cell.getString
			Cell = Sheet.getCellByPosition(4, yIDT)
			vZona = Cell.getString
			Cell = Sheet.getCellByPosition(5, yIDT)
			vTarea = Cell.getString
			Cell = Sheet.getCellByPosition(6, yIDT)
			vPrioridad = Cell.getString
			Cell = Sheet.getCellByPosition(7, yIDT)
			vInfo = Cell.getString
			Cell = Sheet.getCellByPosition(8, yIDT)
			vEstado = Cell.getString
			Cell = Sheet.getCellByPosition(9, yIDT)
			vAsignado = Cell.getString
		'	Cell = Sheet.getCellByPosition(10, yIDT)
		'	vConcurrio = Cell.getString
		'	Cell = Sheet.getCellByPosition(11, yIDT)
		'	vObjetivo = Cell.getString
			Cell = Sheet.getCellByPosition(12, yIDT)
			vFechaApartir = Cell.getString
			Cell = Sheet.getCellByPosition(13, yIDT)
			vFechaCarga = Cell.getString
		'	Cell = Sheet.getCellByPosition(14, yIDT)
		'	vFechaFinalizado = Cell.getString	

			' Carga las Variables en Dialog10
			dlgCT10.Model.TextField1.Text = vIdTarea	
			dlgCT10.Model.TextField2.Text = vNroCliente	
			dlgCT10.Model.TextField3.Text = vNombre	
			dlgCT10.Model.TextField4.Text = vDireccion	
			dlgCT10.Model.TextField5.Text = vZona	
			dlgCT10.Model.lstCBox1.Text = vAsignado 

			dlgCT10.Model.CheckBox1.State = 0
			dlgCT10.Model.CheckBox2.State = 0
			dlgCT10.Model.CheckBox3.State = 0
			dlgCT10.Model.CheckBox4.State = 0
			dlgCT10.Model.CheckBox5.State = 0
			Pos1 = Len(vTarea)
			For y = 1 to Pos1 Step 2
				CadBuscar1 = ""
				CadBuscar1 = Mid(vTarea, y, 1)
				IF CadBuscar1 = "E" THEN dlgCT10.Model.CheckBox1.State = 1
				IF CadBuscar1 = "C" THEN dlgCT10.Model.CheckBox2.State = 1
				IF CadBuscar1 = "D" THEN dlgCT10.Model.CheckBox3.State = 1
				IF CadBuscar1 = "O" THEN dlgCT10.Model.CheckBox4.State = 1
				IF CadBuscar1 = "V" THEN dlgCT10.Model.CheckBox5.State = 1
			Next		
			dlgCT10.Model.TextField6.Text = vPrioridad	
			dlgCT10.Model.TextField7.Text = vEstado	
			dlgCT10.Model.TextField9.Text = vFechaApartir	
			dlgCT10.Model.TextField8.Text = vFechaCarga	
			dlgCT10.Model.TextField10.Text = vInfo		
Paso5:
			cmdBotonSig.getModel.PushButtonType = 0
			cmdBoton = ""		
			Select Case dlgCT10.Execute()
			Case 1
				If dlgCT10.Model.lstCBox1.Text = "" then 
					Msgbox "La tarea no esta Asignada.",16,"IMPORTANTE"
					goto Paso5
				End if
				vAsignado = dlgCT10.Model.lstCBox1.Text
				Cell = Sheet.getCellByPosition(9, yIDT)
				Cell.String = vAsignado
				vTarea = ""
				if dlgCT10.Model.CheckBox1.State = 1 then vTarea = vTarea + "E"
				if dlgCT10.Model.CheckBox2.State = 1 then vTarea = vTarea + "C"
				if dlgCT10.Model.CheckBox3.State = 1 then vTarea = vTarea + "D"
				if dlgCT10.Model.CheckBox4.State = 1 then vTarea = vTarea + "O"
				if dlgCT10.Model.CheckBox5.State = 1 then vTarea = vTarea + "V"
				CadBuscar1 = vTarea
				if vTarea = "" then
					Msgbox "No ha especificado cual es la tarea a realizar.",16,"IMPORTANTE"
					goto Paso5
				End if
				vFechaApartir = dlgCT10.Model.TextField9.Text
				If Date < cDate(vFechaApartir) then
					If Msgbox( "La fecha de inicio de la tarea Nº"+vIdTarea+" es superior a la actual."+chr(13)+chr(13)+"¿Esta seguro que desea modificar la fecha y cursar la tarea?"+chr(13), 4 + 256 + 32, "IMPORTANTE" ) = 6 then
						vFechaApartir = Date
					Else
						goto Paso5					
					End If
				End If
				if Len(CadBuscar1) = 1 then vTarea = CadBuscar1
				if Len(CadBuscar1) = 2 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1)
				if Len(CadBuscar1) = 3 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1)
				if Len(CadBuscar1) = 4 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1) + "+" + Mid(CadBuscar1, 4, 1)
				if Len(CadBuscar1) = 5 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1) + "+" + Mid(CadBuscar1, 4, 1) + "+" + Mid(CadBuscar1, 5, 1)
				Cell = Sheet.getCellByPosition(5, yIDT)
				Cell.String = vTarea
				Cell = Sheet.getCellByPosition(8, yIDT)
				Cell.String = "EN CURSO"
				vInfo = dlgCT10.Model.TextField10.Text
				Cell = Sheet.getCellByPosition(7, yIDT)
				Cell.String = vInfo
				Cell = Sheet.getCellByPosition(12, yIDT)
				Cell.String = vFechaApartir

				'Colorea el fondo de las celdas cursadas
				x = 0
				For x = 1 to 14
					Cell = Sheet.getCellByPosition(x, yIDT)
					Cell.CellBackColor = RGB(102,255,102)'EN CURSO
				Next x
				'Actualiza una Tarea en Expedición Vs Cobros
				Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
				nFila = 0
				For nFila = 5 to 255 'Número
					Cell = Sheet.getCellByPosition(0, nFila)
					If Cell.String = vIdTarea then
						Cell.String = vIdTarea
						Cell = Sheet.getCellByPosition(1, nFila)
						Cell.String = vNroCliente
						Cell = Sheet.getCellByPosition(2, nFila)
						Cell.String = vNombre + chr(13) + vDireccion
						Cell = Sheet.getCellByPosition(3, nFila)
						Cell.String = vZona + chr(13) + vTarea
						Cell = Sheet.getCellByPosition(4, nFila)
						Cell.String = vPrioridad
						Cell = Sheet.getCellByPosition(5, nFila)
						Cell.String = vInfo				
						Cell = Sheet.getCellByPosition(6, nFila)
						Cell.String = "EN CURSO" + chr(13) + vAsignado
						'Colorea fondo cuando es EN CURSO
'						z = 0
'						For z = 0 to 6
'							Cell = Sheet.getCellByPosition(z, nFila)
'							Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
'						Next z
					End If			
				Next nFila
				Sheet = Doc.Sheets.getByName("Carga de Tareas") 
				goto Paso3
			Case 0
				If cmdBoton = "Siguiente" then
					goto Paso3
				End If
				dlgCT10.Dispose()
'				Doc.Store()
				Procesando = False
				Exit Sub		
			End Select		
		End If
Paso3:
	Next yIDT
	Msgbox "Finalizado.",64,"AVISO"
	If Msgbox( "¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información?", 4 + 32, "GUARDAR" ) = 6 then 
		Doc.Store()
		HoraUltGuardar = Timer
		Procesando = False
		Exit Sub
	End If
	GuardarPorTiempoTranscurrido
	Procesando = False
End Sub


Sub BuscarTareasCT
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
Dim vObservacion as String
Dim Encontrado

Inicio:
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent
	DialogLibraries.LoadLibrary("Standard")

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
	
	BuscaActualizacionesGC

Paso1:
 	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator
	'CARGA EL LISTBOX DE CLIENTES DE DIALOG9
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	dlgCT9 = createUnoDialog(DialogLibraries.Standard.Dialog9)
	olstDatos = dlgCT9.getControl("lstCBox1")	
	oHojaDatos = ThisComponent.getSheets.getByName("BDClientes")	
	oRango = oHojaDatos.getCellRangeByName("G3:G12001") 'agregado
	data = oRango.getDataArray()'agregado
	co1 = 0
	Redim src(UBound(data))
	For Each d In data
    	src(co1) = d(0)
      	co1 = co1 + 1
	Next
   	olstDatos.addItems(src, 0)
	'CARGA EL LISTBOX DE ASIGNADO DE DIALOG9
	Sheet = Doc.Sheets.getByName("Datos")
	olstDatos = dlgCT9.getControl("lstCBox2")
  	For d = 1 to 10	
	 	Cell = Sheet.getCellByPosition(2, d) 	
	  	vAsignado = Cell.String
	  	If vAsignado <> "" then olstDatos.addItem( vAsignado, -1 )
	Next d
Paso2:
	'Abre el Dialogo para que ingrese la información a buscar
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	dlgCT9.Model.TextField1.Text = ""
	dlgCT9.Model.TextField2.Text = ""
	dlgCT9.Model.lstCBox1.Text = "" 
	dlgCT9.Model.lstCBox2.Text = "TODOS" 
	Select Case dlgCT9.Execute()
	Case 1
		vNroCliente = dlgCT9.Model.TextField1.Text
		vObservacion = dlgCT9.Model.TextField2.Text		
		vNombre =  dlgCT9.Model.lstCBox1.Text
		if vNombre <> "" then
			CadBuscar1 = "(Nº"
			Pos1 = InStr (vNombre, CadBuscar1)
			Pos1 = Pos1 + 3
			CadBuscar1 = ")"
			Pos2 = InStr (Pos1, vNombre, CadBuscar1)
			Pos2 = Pos2 - Pos1
			vNroCliente = Mid(vNombre, Pos1, Pos2)
		End if		
		vAsignado =  dlgCT9.Model.lstCBox2.Text
		vZona = ""
		if dlgCT9.Model.CheckBox1.State = 1 then vZona = vZona + "N"
		if dlgCT9.Model.CheckBox2.State = 1 then vZona = vZona + "C"
		if dlgCT9.Model.CheckBox3.State = 1 then vZona = vZona + "S"
		if dlgCT9.Model.CheckBox4.State = 1 then vZona = vZona + "I"
		vTarea = ""
		if dlgCT9.Model.CheckBox5.State = 1 then vTarea = vTarea + "E"
		if dlgCT9.Model.CheckBox6.State = 1 then vTarea = vTarea + "C"
		if dlgCT9.Model.CheckBox7.State = 1 then vTarea = vTarea + "D"
		if dlgCT9.Model.CheckBox8.State = 1 then vTarea = vTarea + "O"
		if dlgCT9.Model.CheckBox9.State = 1 then vTarea = vTarea + "V"
		vEstado = ""
		if dlgCT9.Model.CheckBox10.State = 1 then vEstado = vEstado + "P"
		if dlgCT9.Model.CheckBox11.State = 1 then vEstado = vEstado + "E"
		if dlgCT9.Model.CheckBox12.State = 1 then vEstado = vEstado + "F"
		goto Paso4
	Case 0
		dlgCT9.Dispose()
		Procesando = False
		Exit Sub		
	End Select
Paso4:
	'Busca 
	oBarraEstado.start( "Buscando... ", 2000 )
	yIDT = 0
	For yIDT = 10 to 8000 'Número máximo de Filas en Carga de Tareas
		oBarraEstado.setValue( yIDT )
		Cell = Sheet.getCellByPosition(0, yIDT)
		If Cell.String = "" Then
			oBarraEstado.setValue( 2000 )
			msgbox "Busqueda finalizada.",64,"AVISO"
			oBarraEstado.end()
			Procesando = False 
			Exit Sub
		End if
		If vObservacion <> "" then
			Cell = Sheet.getCellByPosition(7, yIDT)
			Encontrado = 0
			Encontrado = Instr(Cell.String,vObservacion)
			If Encontrado = 0 then goto Siguiente
		End If
		If vNroCliente <> "" then
			Cell = Sheet.getCellByPosition(1, yIDT)
			If vNroCliente <> Cell.String then goto Siguiente
		End If
		If vEstado <> "" then
			Cell = Sheet.getCellByPosition(8, yIDT)
			If Cell.String = "PENDIENTE" and dlgCT9.Model.CheckBox10.State <> 1 then goto Siguiente 
			If Cell.String = "EN CURSO" and dlgCT9.Model.CheckBox11.State <> 1 then goto Siguiente 
			If Cell.String = "FINALIZADO" and dlgCT9.Model.CheckBox12.State <> 1 then goto Siguiente
		End If
		If vAsignado <> "TODOS" then
			Cell = Sheet.getCellByPosition(9, yIDT)
			If vAsignado <> Cell.String then goto Siguiente
		End If
		If vZona <> "" then
			Cell = Sheet.getCellByPosition(4, yIDT)
			If Cell.String = "CBAN" and dlgCT9.Model.CheckBox1.State <> 1 then goto Siguiente 
			If Cell.String = "CBAC" and dlgCT9.Model.CheckBox2.State <> 1 then goto Siguiente 
			If Cell.String = "CBAS" and dlgCT9.Model.CheckBox3.State <> 1 then goto Siguiente 
			If Cell.String = "INT" and dlgCT9.Model.CheckBox4.State <> 1 then goto Siguiente 
		End if
		If vTarea <> "" then
			Cell = Sheet.getCellByPosition(5, yIDT)
			Pos1 = 0
			CadBuscar1 = "E"
			CadBuscar2 = Cell.getString
			Pos1 = InStr (CadBuscar2, CadBuscar1)
			If Pos1 > 0 and dlgCT9.Model.CheckBox5.State <> 1 then goto Siguiente 
			Pos1 = 0
			CadBuscar1 = "C"
			Pos1 = InStr (CadBuscar2, CadBuscar1)
			If Pos1 > 0 and dlgCT9.Model.CheckBox6.State <> 1 then goto Siguiente 
			Pos1 = 0
			CadBuscar1 = "D"
			Pos1 = InStr (CadBuscar2, CadBuscar1)
			If Pos1 > 0 and dlgCT9.Model.CheckBox7.State <> 1 then goto Siguiente 
			Pos1 = 0
			CadBuscar1 = "O"
			Pos1 = InStr (CadBuscar2, CadBuscar1)
			If Pos1 > 0 and dlgCT9.Model.CheckBox8.State <> 1 then goto Siguiente 
			Pos1 = 0
			CadBuscar1 = "V"
			Pos1 = InStr (CadBuscar2, CadBuscar1)
			If Pos1 > 0 and dlgCT9.Model.CheckBox9.State <> 1 then goto Siguiente 
		End if
	
		Posicionar = yIDT
		PosicionadorCelda
		FilaActual = yIDT
		FilaVisible
		Cell = Sheet.getCellByPosition(0, yIDT)
		If Msgbox( "¿Ha encontrado lo que buscaba.?", 4 + 32 + 256, "Id.Tarea: " + Cell.String ) = 6 then
			oBarraEstado.setValue( 2000 )
			FilaNoVisible
			oBarraEstado.end()
			Procesando = False		
			Exit Sub
		End if
		FilaNoVisible
Siguiente:
	Next yIDT
	oBarraEstado.setValue( 2000 )
	dlgCT9.Dispose()
	oBarraEstado.end()
	Procesando = False
End Sub

Sub ModificarTareasCT

Inicio:
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent
	DialogLibraries.LoadLibrary("Standard")

	'CORROBORA QUE HAYA INGRESADO LA CONTRASEÑA DE USUARIO	
	If vUsuario = "" then 
		Sheet = Doc.Sheets.getByName("Usuario")
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
	    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
		otxtPW = oFormulario.getByName("txtPW")
		otxtPWVista = Doc.getCurrentController.getControl( otxtPW )
		otxtPWVista.SetFocus()
		Exit Sub
	End If
		
	BuscaActualizacionesGC	

Paso1:
	'INICIA DIALOG7 PARA QUE INGRESE EL ID.TAREA
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	dlgCT7 = createUnoDialog(DialogLibraries.Standard.Dialog7)
	dlgCT8 = createUnoDialog(DialogLibraries.Standard.Dialog8)
	dlgCT7.Model.TextField1.Text = ""	
	Select Case dlgCT7.Execute()
	Case 1
		vIdTarea = dlgCT7.Model.TextField1.Text
		goto Paso2
	Case 0
		dlgCT7.Dispose()
		Exit Sub		
	End Select
Paso2:
	'BUSCA ID.TAREA EN CARGA DE TAREAS
	yIDT = 0
	For yIDT = 10 to 10011 'Número máximo de Filas en Carga de Tareas
		Cell = Sheet.getCellByPosition(0, yIDT)
		If Cell.String = vIdTarea then goto Paso3
		If Cell.String = "" then
				Exit For
		End if 
	Next yIDT
	Msgbox "Id.Tarea no encontrada.",16,"IMPORTANTE"
	goto Paso1	
	Exit Sub
Paso3:
	'CORROBORA EL ESTADO DE LA TAREA
	Posicionar = yIDT
	PosicionadorCelda
	FilaActual = yIDT
	FilaVisible
	Cell = Sheet.getCellByPosition(8, yIDT)
	If Cell.String = "PENDIENTE" then
		goto Paso4
	End If
	If Cell.String = "EN CURSO" then
		If Msgbox( "La Tarea se encuentra EN CURSO."+chr(13)+"¿Desea modificarla.?"+chr(13)+"Id.Tarea: "+vIdTarea, 4 + 32, "IMPORTANTE" ) = 6 then
			Msgbox( "Recuerde informar a la persona Asignada.", 64, "IMPORTANTE" )
			goto Paso4	
		End If
		FilaNoVisible
		Goto Paso1
	End If
	If Cell.String = "FINALIZADO" then
		If Msgbox( "La Tarea se encuentra FINALIZADA."+chr(13)+"¿Desea modificarla.?"+chr(13)+"Id.Tarea: "+vIdTarea, 4 + 32, "IMPORTANTE" ) = 6 then
			goto Paso4	
		End If
		FilaNoVisible
		Goto Paso1
	End if
	Msgbox "El ""ESTADO"" de la Tarea es desconocido.",16,"IMPORTANTE"
	FilaNoVisible
	Goto Paso1
Paso4:
	'CARGA LA INFORMACION DE LA FILA yIDT
	Cell = Sheet.getCellByPosition(1, yIDT)
	vNroCliente = Cell.getString
	Cell = Sheet.getCellByPosition(2, yIDT)
	vNombre = Cell.getString
	Cell = Sheet.getCellByPosition(3, yIDT)
	vDireccion = Cell.getString
	Cell = Sheet.getCellByPosition(4, yIDT)
	vZona = Cell.getString
	Cell = Sheet.getCellByPosition(5, yIDT)
	vTarea = Cell.getString
	Cell = Sheet.getCellByPosition(6, yIDT)
	vPrioridad = Cell.getString
	Cell = Sheet.getCellByPosition(7, yIDT)
	vInfo = Cell.getString
	Cell = Sheet.getCellByPosition(8, yIDT)
	vEstado = Cell.getString
	Cell = Sheet.getCellByPosition(9, yIDT)
	vAsignado = Cell.getString
'	Cell = Sheet.getCellByPosition(10, yIDT)
'	vConcurrio = Cell.getString
	Cell = Sheet.getCellByPosition(11, yIDT)
	vFechaFinalizado = Cell.getString
	
	Cell = Sheet.getCellByPosition(12, yIDT)
	vFechaApartir = Cell.getString
	Cell = Sheet.getCellByPosition(13, yIDT)
	vFechaCarga = Cell.getString
	Cell = Sheet.getCellByPosition(14, yIDT)
	vUltMod = Cell.getString
	
	'CARGA EL LISTBOX DE ASIGNADO A DIALOG8
	olstDatos = dlgCT8.getControl("ComboBox2")	
	oHojaDatos = ThisComponent.getSheets.getByName("Datos")	
	oRango = oHojaDatos.getCellRangeByName("C3:C10") 'ASIGNADO
	data = oRango.getDataArray()
	co1 = 0
	Redim src(UBound(data))
	For Each d In data
    	src(co1) = d(0)
      	co1 = co1 + 1
	Next
   	olstDatos.addItems(src, 0)

	' Carga las Variables en Dialog8
	dlgCT8.Model.TextField3.Text = vIdTarea
	dlgCT8.Model.TextField5.Text = vNroCliente 
	dlgCT8.Model.TextField1.Text = vNombre
	dlgCT8.Model.TextField7.Text = vDireccion
	dlgCT8.Model.ComboBox1.text = vZona
	y = 0
	dlgCT8.Model.CheckBox1.State = 0
	dlgCT8.Model.CheckBox2.State = 0
	dlgCT8.Model.CheckBox3.State = 0
	dlgCT8.Model.CheckBox4.State = 0
	dlgCT8.Model.CheckBox5.State = 0
	Pos1 = Len(vTarea)
	For y = 1 to Pos1 Step 2
		CadBuscar1 = ""
		CadBuscar1 = Mid(vTarea, y, 1)
		IF CadBuscar1 = "E" THEN dlgCT8.Model.CheckBox1.State = 1
		IF CadBuscar1 = "C" THEN dlgCT8.Model.CheckBox2.State = 1
		IF CadBuscar1 = "D" THEN dlgCT8.Model.CheckBox3.State = 1
		IF CadBuscar1 = "O" THEN dlgCT8.Model.CheckBox4.State = 1
		IF CadBuscar1 = "V" THEN dlgCT8.Model.CheckBox5.State = 1
	Next	
	If vPrioridad = "ALTA" then dlgCT8.Model.OptionButton1.State = 1
	If vPrioridad = "MEDIA" then dlgCT8.Model.OptionButton2.State = 1
	If vPrioridad = "BAJA" then dlgCT8.Model.OptionButton3.State = 1	
	If vEstado = "PENDIENTE" then dlgCT8.Model.OptionButton4.State = 1
	If vEstado = "EN CURSO" then dlgCT8.Model.OptionButton5.State = 1
	If vEstado = "FINALIZADO" then dlgCT8.Model.OptionButton6.State = 1	
	dlgCT8.Model.TextField4.Text = vInfo
	dlgCT8.Model.ComboBox2.text = vAsignado
	dlgCT8.Model.DateField1.text = vFechaApartir
	
Paso5:
	'Abre Dialogo8 para modificarlo
	Select Case dlgCT8.Execute()
	Case 1
		'Extrae la información ingresada en el Dialogo
		'vIdTarea = dlgCT8.Model.TextField3.Text
		'vNroCliente = dlgCT8.Model.TextField5.Text 
		'vNombre = dlgCT8.Model.TextField1.Text
		vDireccion = dlgCT8.Model.TextField7.Text
		vZona = dlgCT8.Model.ComboBox1.text
		If vZona = "" then
			Msgbox "No ha especificado la zona",16,"IMPORTANTE"
			goto Paso4
		End if
		vTarea = ""
		if dlgCT8.Model.CheckBox1.State = 1 then vTarea = vTarea + "E"
		if dlgCT8.Model.CheckBox2.State = 1 then vTarea = vTarea + "C"
		if dlgCT8.Model.CheckBox3.State = 1 then vTarea = vTarea + "D"
		if dlgCT8.Model.CheckBox4.State = 1 then vTarea = vTarea + "O"
		if dlgCT8.Model.CheckBox5.State = 1 then vTarea = vTarea + "V"
		CadBuscar1 = vTarea
		if vTarea = "" then
			Msgbox "No ha especificado cual es la tarea a realizar.",16,"IMPORTANTE"
			goto Paso4
		End if
		if Len(CadBuscar1) = 1 then vTarea = CadBuscar1
		if Len(CadBuscar1) = 2 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1)
		if Len(CadBuscar1) = 3 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1)
		if Len(CadBuscar1) = 4 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1) + "+" + Mid(CadBuscar1, 4, 1)
		if Len(CadBuscar1) = 5 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1) + "+" + Mid(CadBuscar1, 4, 1) + "+" + Mid(CadBuscar1, 5, 1)
		if dlgCT8.Model.OptionButton1.State = 1 then vPrioridad = "ALTA"	
		if dlgCT8.Model.OptionButton2.State = 1 then vPrioridad = "MEDIA"
		if dlgCT8.Model.OptionButton3.State = 1 then vPrioridad = "BAJA"	
		if dlgCT8.Model.OptionButton4.State = 1 then vEstado = "PENDIENTE"	
		if dlgCT8.Model.OptionButton5.State = 1 then vEstado = "EN CURSO"
		if dlgCT8.Model.OptionButton6.State = 1 then vEstado = "FINALIZADO"
		vInfo =  dlgCT8.Model.TextField4.Text
		vAsignado = dlgCT8.Model.ComboBox2.text
		if vAsignado = "" then
			Msgbox "La Tarea no ha sido ASIGNADA a una persona.",16,"IMPORTANTE"
			goto Paso4
		End if
		If dlgCT8.Model.DateField1.text <> "" then
			vFechaApartir = dlgCT8.Model.DateField1.text
		End If
		If dlgCT8.Model.DateField1.text = "" then
			vFechaApartir = Date
		End If
	Case 0
		dlgCT8.Dispose()
		Exit Sub
	End Select
Paso7:
	'Ingresa los datos en la planilla de calculo
'	Cell = Sheet.getCellByPosition(1, yIDT)
'	Cell.String = vNroCliente
'	Cell = Sheet.getCellByPosition(2, yIDT)
'	Cell.String = vNombre
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
	Cell.String = "" 'vFechaFinalizado
	Cell = Sheet.getCellByPosition(12, yIDT)
	Cell.String = vFechaApartir
	Cell = Sheet.getCellByPosition(13, yIDT)
	Cell.String = vFechaCarga
	Cell = Sheet.getCellByPosition(14, yIDT)
	Cell.String = Date
	Cell = Sheet.getCellByPosition(15, yIDT)
	If Right(Cell.String,Len(vUsuario)) = vUsuario then

	Else
		If Cell.String = "" then
			Cell.String = vUsuario
		Else
			Cell.String = Cell.String + "/" + vUsuario
		End If
	End If
	Cell = Sheet.getCellByPosition(16, yIDT)
	If Cell.String = "" then  Cell.Value = 0
	Cell = Sheet.getCellByPosition(17, yIDT)
	If Cell.String = "" then  Cell.Value = 0

	'Colorea el fondo de la fila xIDT
	x = 0
	For x = 1 to 14
		Cell = Sheet.getCellByPosition(x, yIDT)
		Cell.CellBackColor = RGB(255,255,255)'BLANCO
	Next x
	'En caso que el ESTADO sea FINALIZADO
	If vEstado = "FINALIZADO" then
		Cell = Sheet.getCellByPosition(11, yIDT)
		Cell.String = DATE
		'Colorea el fondo de la fila xIDT
		x = 0
		For x = 1 to 14
			Cell = Sheet.getCellByPosition(x, yIDT)
			Cell.CellBackColor = RGB(255,102,102)'FINALIZADO
		Next x
	End if
	If vEstado = "EN CURSO" then
		'Colorea el fondo de la fila xIDT
		x = 0
		For x = 1 to 14
			Cell = Sheet.getCellByPosition(x, yIDT)
			Cell.CellBackColor = RGB(102,255,102)'EN CURSO
		Next x
	End if
	FilaActual = yIDT
	FilaNoVisible
	
	'Actualiza una Tarea PENDIENTE o EN CURSO en Expedición Vs Cobros y Borra una FINALIZADA
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	If vEstado = "PENDIENTE" or vEstado = "EN CURSO" then
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = vIdTarea then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(1, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(2, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(3, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(4, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(5, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(6, nFila)
				Cell.String = vEstado + chr(13) + vAsignado
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255)'BLANCO 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
				If vEstado = "EN CURSO" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255)'BLANCO 'EN CURSO
					Next z
				End If			
				Goto FinActualizacion
			End If
		Next nFila
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = "" then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(1, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(2, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(3, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(4, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(5, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(6, nFila)
				Cell.String = vEstado + chr(13) + vAsignado
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255)'BLANCO 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
				If vEstado = "EN CURSO" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255)'BLANCO 'EN CURSO
					Next z
				End If			
				Goto FinActualizacion
			End If
		Next nFila
	End If
	If vEstado = "FINALIZADO" then
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = vIdTarea then
				'Borra y colorea el fondo
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, nFila)
					Cell.String = ""
					Cell.CellBackColor = RGB(213,231,234) 'CELESTE CLARO 'VACIA
				Next z
				Goto FinActualizacion
			End If
		Next nFila
	End If
FinActualizacion:
	dlgCT7.Dispose()
	dlgCT8.Dispose()
	If Msgbox( "¿Desea modificar otra tarea.?", 4 + 32, "" ) = 6 then goto Paso1
	If Msgbox( "¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información?", 4 + 32, "Guardar" ) = 6 then 
		Doc.Store()
		HoraUltGuardar = Timer
		Exit Sub
	End If
	GuardarPorTiempoTranscurrido
End Sub



Sub UnificarTareasPendientesCT
	Dim vIdTarea2 as String, vNroCliente2 as String, vNombre2 as String, vDireccion2 as string, vZona2 as String, vTarea2 as String
	Dim vPrioridad2 as String, vInfo2 as String, vEstado2 As String, vAsignado2 As String, vFechaApartir2 As String
	Dim yIDT2, y1, y2

Inicio:
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent
	DialogLibraries.LoadLibrary("Standard")
	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator

	'CORROBORA QUE HAYA INGRESADO LA CONTRASEÑA DE USUARIO	
	If vUsuario = "" then 
		Sheet = Doc.Sheets.getByName("Usuario")
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
	    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
		otxtPW = oFormulario.getByName("txtPW")
		otxtPWVista = Doc.getCurrentController.getControl( otxtPW )
		otxtPWVista.SetFocus()
		Exit Sub
	End If
	
	BuscaActualizacionesGC			
	
Paso1:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	dlgCT6 = createUnoDialog(DialogLibraries.Standard.Dialog6)
	If Msgbox( "Este comando solo busca y unifica Tareas de un mismo Cliente/Destinatario que se encuentren en estado PENDIENTE y que al menos una de ellas no haya sido asignada con anterioridad."+chr(13)+chr(13)+"Se recomienda realizar al final del día, luego de ingresar las Hojas de Ruta o al principio del día antes de emitir nuevas Hojas de Ruta."+chr(13)+chr(13)+"Espere a que el sistema le informe que ha finalizado.", 1 + 64, "INFORMACION" ) = 1 then 
		
	Else
		Exit Sub
	End If
Paso2:
	'Busca en la columna de Estado los PENDIENTES.
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	For yIDT = 10 to 10000
		Cell = Sheet.getCellByPosition(0, yIDT)
		If Cell.String = "" then
			Cell = Sheet.getCellByPosition(0, yIDT + 1)
			If Cell.String = "" then
				vProgBar = yIDT
				Exit For
			End If
		End If
	Next yIDT 	
	oBarraEstado.start( "Unificando Tareas ", vProgBar )

	For yIDT = 10 to 10000
		oBarraEstado.setValue( yIDT )
		Cell = Sheet.getCellByPosition(8, yIDT)
		If Cell.String = "PENDIENTE" then 
			Cell = Sheet.getCellByPosition(0, yIDT)
			If Cell.String = "" then
				Posicionar = yIDT
				PosicionadorCelda
				MsgBox "Id.Tarea desconocido."+chr(13)+chr(13)+"El proceso se detendra para que corrija el estado.", 16, "IMPORTANTE" 
				oBarraEstado.end()
				Exit Sub
			End If
			vIdTarea = Cell.getString
			Cell = Sheet.getCellByPosition(1, yIDT)
			If Cell.String = "" then
				Posicionar = yIDT
				PosicionadorCelda
				MsgBox "Nro. de Cliente/Destinatario desconocido."+chr(13)+chr(13)+"El proceso se detendra para que corrija el estado.", 16, "IMPORTANTE" 
				oBarraEstado.end()
				Exit Sub
			End If
			vNroCliente = Cell.getString
			Cell = Sheet.getCellByPosition(2, yIDT)
			If Cell.String = "" then
				Posicionar = yIDT
				PosicionadorCelda
				MsgBox "Cliente/Destinatario desconocido."+chr(13)+chr(13)+"El proceso se detendra para que corrija el estado.", 16, "IMPORTANTE" 
				oBarraEstado.end()
				Exit Sub
			End If
			vNombre = Cell.getString
			For yIDT2 =	yIDT + 1 to 10000
				Cell = Sheet.getCellByPosition(8, yIDT2)
				If Cell.String = "PENDIENTE" then 
					Cell = Sheet.getCellByPosition(0, yIDT2)
					If Cell.String <> "" then
						Cell = Sheet.getCellByPosition(1, yIDT2)
						If Cell.String = vNroCliente then
							Cell = Sheet.getCellByPosition(2, yIDT2)
							If Cell.String = vNombre then
								Gosub Paso3
								Exit For
							End if								
						End if
					End if
				End if
				If vProgBar + 1 < yIDT2 then Exit For
			Next yIDT2
		End if

		Cell = Sheet.getCellByPosition(8, yIDT)
		If Cell.String = "" then
			Cell = Sheet.getCellByPosition(0, yIDT)
			If Cell.String <> "" then
				Cell = Sheet.getCellByPosition(1, yIDT)
				If Cell.String <> "" then
					Cell = Sheet.getCellByPosition(2, yIDT)
					If Cell.String <> "" then
						Posicionar = yIDT
						PosicionadorCelda
						Cell = Sheet.getCellByPosition(0, yIDT)
						MsgBox "El estado de Id.Tarea Nº "+Cell.String+" es desconocido."+chr(13)+chr(13)+"El proceso se detendra para que corrija el estado.", 16, "IMPORTANTE" 
						oBarraEstado.end()
						Exit Sub
					End If
				End If
			End If
		End If
		If vProgBar + 1 < yIDT then Exit For
	Next yIDT
	oBarraEstado.end()
	dlgCT6.Dispose()
	If Msgbox( "Finalizado."+chr(13)+chr(13)+"¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información?", 4 + 32, "Guardar" ) = 6 then 
		Doc.Store()
		HoraUltGuardar = Timer
		Exit Sub
	End If
	Exit Sub
Paso3:
	'VERIFICA QUE AL MENOS UNO NO TENGA ASIGNACIONES
	Cell = Sheet.getCellByPosition(16, yIDT)
	If CInt(Cell.String) = CInt(0) then
		Cell = Sheet.getCellByPosition(16, yIDT2)
		If CInt(Cell.String) = CInt(0) then
			y1 = yIDT
			y2 = yIDT2
		Else	
			y2 = yIDT
			y1 = yIDT2
		End If
	Else
		Cell = Sheet.getCellByPosition(16, yIDT2)
		If CInt(Cell.String) = CInt(0) then
			y1 = yIDT
			y2 = yIDT2
		Else
			Return	
		End If		
	End If

	Posicionar = y1 
	PosicionadorCelda
	'CARGA LA INFORMACION A UNIFICAR
	Cell = Sheet.getCellByPosition(0, y1)
	vIdTarea = Cell.getString
	Cell = Sheet.getCellByPosition(1, y1)
	vNroCliente = Cell.getString
	Cell = Sheet.getCellByPosition(2, y1)
	vNombre = Cell.getString
	Cell = Sheet.getCellByPosition(3, y1)
	vDireccion = Cell.getString
	Cell = Sheet.getCellByPosition(4, y1)
	vZona = Cell.getString
	Cell = Sheet.getCellByPosition(5, y1)
	vTarea = Cell.getString
	Cell = Sheet.getCellByPosition(6, y1)
	vPrioridad = Cell.getString
	Cell = Sheet.getCellByPosition(7, y1)
	vInfo = Cell.getString
	Cell = Sheet.getCellByPosition(8, y1)
	vEstado = Cell.getString
	Cell = Sheet.getCellByPosition(9, y1)
	vAsignado = Cell.getString
	Cell = Sheet.getCellByPosition(12, y1)
	vFechaApartir = Cell.getString
	
	Cell = Sheet.getCellByPosition(0, y2)
	vIdTarea2 = Cell.getString
	Cell = Sheet.getCellByPosition(1, y2)
	vNroCliente2 = Cell.getString
	Cell = Sheet.getCellByPosition(2, y2)
	vNombre2 = Cell.getString
	Cell = Sheet.getCellByPosition(3, y2)
	vDireccion2 = Cell.getString
	Cell = Sheet.getCellByPosition(4, y2)
	vZona2 = Cell.getString
	Cell = Sheet.getCellByPosition(5, y2)
	vTarea2 = Cell.getString
	Cell = Sheet.getCellByPosition(6, y2)
	vPrioridad2 = Cell.getString
	Cell = Sheet.getCellByPosition(7, y2)
	vInfo2 = Cell.getString
	Cell = Sheet.getCellByPosition(8, y2)
	vEstado2 = Cell.getString
	Cell = Sheet.getCellByPosition(9, y2)
	vAsignado2 = Cell.getString
	Cell = Sheet.getCellByPosition(12, y2)
	vFechaApartir2 = Cell.getString
	
	dlgCT6.Model.TextField1.Text = "Id.Tarea Nro.: " + vIdTarea + chr(13) + "Destinatario: " + vNombre + chr(13) + "Dirección: " + vDireccion + chr(13) + "Zona: " + vZona + chr(13) + "Tarea: " + vTarea + chr(13) + "Prioridad: " + vPrioridad + chr(13) + "Comentarios: " + vInfo + chr(13) + "Estado: " + vEstado + chr(13) + "Asignado: " + vAsignado + chr(13) + "A partir de: " + vFechaApartir
	dlgCT6.Model.TextField2.Text = "Id.Tarea Nro.: " + vIdTarea2 + chr(13) + "Destinatario: " + vNombre2 + chr(13) + "Dirección: " + vDireccion2 + chr(13) + "Zona: " + vZona2 + chr(13) + "Tarea: " + vTarea2 + chr(13) + "Prioridad: " + vPrioridad2 + chr(13) + "Comentarios: " + vInfo2 + chr(13) + "Estado: " + vEstado2 + chr(13) + "Asignado: " + vAsignado2 + chr(13) + "A partir de: " + vFechaApartir2

	' Abre dialogo
	Select Case dlgCT6.Execute()
	Case 1
		Cell = Sheet.getCellByPosition(5, y1)
		If vTarea <> vTarea2 then
			Cell.String = vTarea + "+" + vTarea2
			vTarea = Cell.getString
		end if
		Cell = Sheet.getCellByPosition(6, y1)
		if vPrioridad = "BAJA" then
			Cell.String = "MEDIA"
			vPrioridad = Cell.getString
		end if
		if vPrioridad2 = "ALTA" then
			Cell.String = "ALTA"
			vPrioridad = Cell.getString
		end if

		Cell = Sheet.getCellByPosition(7, y1)
		If Instr(vInfo, "[GC]") > 0 and Instr(vInfo2, "[GC]") = 0 then
			Cell.String = vInfo2 + chr(13) + vInfo
		End If
		If Instr(vInfo, "[GC]") = 0 and Instr(vInfo2, "[GC]") > 0 then
			Cell.String = vInfo + chr(13) + vInfo2
		End If
		If Instr(vInfo, "[GC]") = 0 and Instr(vInfo2, "[GC]") = 0 then
			Cell.String = vInfo + " " + vInfo2
		End If
		vInfo = Cell.getString

		If cDate(vFechaApartir2) > cDate(vFechaApartir) then
			vFechaApartir = vFechaApartir2
		End If
		Cell = Sheet.getCellByPosition(12, y1)
		Cell.String = vFechaApartir

		'ELIMINA LA TAREA
		For x = 1 to 30
			Cell = Sheet.getCellByPosition(x, y2)
			Cell.String = ""
			If x > 0 and x < 15 then
				Cell.CellBackColor = RGB(255,255,255) 'BLANCO = PENDIENTE
			End If
		Next x
		'Elimina y actualiza una Tareas Expedición Vs Cobros
		Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = vIdTarea2 then
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, nFila)
					Cell.String = ""
					Cell.CellBackColor = RGB(213,231,234) 'CELESTE CLARO 'VACIA
				Next z
			End If
		Next nFila
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = vIdTarea then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(3, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(4, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(5, nFila)
				Cell.String = vInfo				
			End If
		Next nFila
		Sheet = Doc.Sheets.getByName("Carga de Tareas")
		Return
	Case 0
		Return
	End Select
End Sub

Sub CargarTareaCT
	Dim vCadena as String	
	
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent

	'CORROBORA QUE HAYA INGRESADO LA CONTRASEÑA DE USUARIO	
	If vUsuario = "" then 
		Sheet = Doc.Sheets.getByName("Usuario")
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
	    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
		otxtPW = oFormulario.getByName("txtPW")
		otxtPWVista = Doc.getCurrentController.getControl( otxtPW )
		otxtPWVista.SetFocus()
		Exit Sub
	End If

	If bTareaExtraHR = 0 then BuscaActualizacionesGC

Inicio:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	DialogLibraries.LoadLibrary("Standard")
	dlgCT1 = createUnoDialog(DialogLibraries.Standard.Dialog1)
		
Paso1:
Paso2:
	'Busca la primera celda vacia de Nro. de Cliente en Columna B.
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	If vProxTarea = 0 or nProxFila = 0 then
		BuscaProxTareaDisponibleCT
	Else
		If nProxFila > 0 then NuevaFilaCT
	End If
	yIDT = nProxFila	'eliminar o no
	If nProxFila = 0 then Exit Sub
	If vIdTarea = "" then Exit Sub
Paso3:
	Posicionar = nProxFila 
	PosicionadorCelda
	'CARGA LISTADO DE CLIENTES EN EL LISTBOX DE DIALOG2
	dlgCT2 = createUnoDialog(DialogLibraries.Standard.Dialog2)	
	olstDatos = dlgCT2.getControl("lstCBox1")	
	oHojaDatos = ThisComponent.getSheets.getByName("BDClientes")	
	oRango = oHojaDatos.getCellRangeByName("G3:G12001") 'agregado
	data = oRango.getDataArray()'agregado
	co1 = 0
	Redim src(UBound(data))
	For Each d In data
    	src(co1) = d(0)
      	co1 = co1 + 1
	Next
   	olstDatos.addItems(src, 0)
Paso4:
	'INICIA DIALOG2 PARA INGRESAR EL CLIENTE/DESTINATARIO
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	CadBuscar = ""
	CadBuscar2 = ""
	CadResultado = ""
	Pos1 = 1
	Pos2 = 1
	vNroCliente = ""
	vNombre = ""
	dlgCT2.Model.TextField1.Text = ""
	dlgCT2.Model.TextField2.Text = vIdTarea
	dlgCT2.Model.lstCBox1.Text = "" 
	Select Case dlgCT2.Execute()
	Case 1
		vNroCliente =  dlgCT2.Model.TextField1.Text		
		vNombre =  dlgCT2.Model.lstCBox1.Text
		if vNroCliente = "" then
			if vNombre = "" then 
				Msgbox "No ha ingresado el Nro. de Cliente"
				Goto Paso4
			End if
			CadBuscar = "(Nº"
			Pos1 = InStr (vNombre, CadBuscar)
			Pos1 = Pos1 + 3
			CadBuscar = ")"
			Pos2 = InStr (Pos1, vNombre, CadBuscar)
			Pos2 = Pos2 - Pos1
			vNroCliente = Mid(vNombre, Pos1, Pos2)
			Goto Paso5
		End if
		if vNroCliente = "0" then
			if vNombre = "" then 
				Msgbox "No ha ingresado el Nombre del Destinatario"
				Goto Paso4
			End if
			Goto Paso5
		End if
		if vNombre <> "" then
			if vNroCliente <> "" then 
				CadBuscar = "(Nº"
				Pos1 = InStr (vNombre, CadBuscar)
				Pos1 = Pos1 + 3
				CadBuscar = ")"
				Pos2 = InStr (Pos1, vNombre, CadBuscar)
				Pos2 = Pos2 - Pos1
				CadResultado = Mid(vNombre, Pos1, Pos2)
				if vNroCliente = CadResultado then
					goto Paso5
				End if
				Msgbox "No coincide el Nro. de Cliente con el Cliente seleccionado."
				Goto Paso4
			End if
			CadBuscar = "(Nº"
			Pos1 = InStr (vNombre, CadBuscar)
			Pos1 = Pos1 + 3
			CadBuscar = ")"
			Pos2 = InStr (Pos1, vNombre, CadBuscar)
			Pos2 = Pos2 - Pos1
			vNroCliente = Mid(vNombre, Pos1, Pos2)
			Goto Paso5
		End if
	Case 0
		dlgCT2.Dispose()
		Exit Sub
	End Select
Paso5:
	'BUSCA QUE EL CLIENTE/DESTINATARIO NO POSEA UNA TAREA PENDIENTE O EN CURSO.
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	If bTareaExtraHR <> 0 then goto Paso6
	Y = 0
	For Y = 10 to 8000
		Cell = Sheet.getCellByPosition(1, y)
		If vNroCliente = "0" and vNroCliente = Cell.String then
			Cell = Sheet.getCellByPosition(2, y)
			If Left(vNombre, 5) = Left(Cell.String, 5) then
				Cell = Sheet.getCellByPosition(8, y) 
				If Cell.String = "EN CURSO" then
					Posicionar = y 
					PosicionadorCelda
					y = y - 10
					If Msgbox( "El Destinatario podría tener una tarea EN CURSO"+chr(13)+"Corroborar la Id.Tarea:"+ y, 1 + 16, "IMPORTANTE" ) = 1 then
						Msgbox( "Recuerde informar a la persona Asignada cualquier modificación que realice", 0, "IMPORTANTE" )
						Exit Sub
					End If
					y = y + 10
					goto Paso5B
				End If
				If Cell.String = "PENDIENTE" then
					Posicionar = y 
					PosicionadorCelda
					y = y - 10
					If Msgbox( "El Destinatario podría tener una tarea PENDIENTE"+chr(13)+"Corroborar la Id.Tarea:"+ y, 1 + 16, "IMPORTANTE" ) = 1 then
						Exit Sub
					End If
					y = y + 10
					goto Paso5B
				End If
			End If
			goto Paso5B
		End If
		Cell = Sheet.getCellByPosition(1, y)
		If Cell.String = "" then 
			Cell = Sheet.getCellByPosition(0, y)
			If Cell.String = "" then 
				goto Paso6
			End If
		End If
		Cell = Sheet.getCellByPosition(1, y)
		If vNroCliente = Cell.String then 
			Cell = Sheet.getCellByPosition(8, y)
			If Cell.String = "EN CURSO" then
				Posicionar = y 
				PosicionadorCelda
				y = y - 10
				If Msgbox( "El Destinatario posee una tarea EN CURSO"+chr(13)+"¿Desea modificarla?"+chr(13)+"Id.Tarea:"+ y, 4 + 32, "IMPORTANTE" ) = 6 then
					Msgbox( "Recuerde informar a la persona Asignada", 0, "IMPORTANTE" )
					yIDT = y + 10
					goto Paso9
				End if
				y = y + 10
				goto Paso5B
			End If
			If Cell.String = "PENDIENTE" then
				Posicionar = y 
				PosicionadorCelda
				y = y - 10
				If Msgbox( "El Destinatario posee una tarea PENDIENTE"+chr(13)+"¿Desea modificarla?"+chr(13)+"Id.Tarea:"+ y, 4 + 32, "IMPORTANTE" ) = 6 then
					yIDT = y + 10
					goto Paso9
				End if
				y = y + 10
				goto Paso5B
			End If
		end if
Paso5B: 
	Next y

Paso6:
	'BUSCA EL CLIENTE/DESTINATARIO EN BDCLIENTES Y CARGA LA INFORMACION DEL MISMO.
	Posicionar = nProxFila 
	PosicionadorCelda

	vDireccion = ""
	vZona = ""
	vTarea = ""
	vCadena = ""
	vPrioridad = ""
	vInfo = ""
	vEstado = ""
	vFechaApartir = date
	vAsignado = ""
	y = 0
	If vNroCliente = "0" then 
		dlgCT1.Model.Step = 2
		goto Paso7 
	End If
	Sheet = Doc.Sheets.getByName("BDClientes")
	For y = 2 to 12005 'Número máximo de filas a buscar en BDClientes.
		Cell = Sheet.getCellByPosition(0, y)
		If Cell.String = vNroCliente then 
			Cell = Sheet.getCellByPosition(1, y)
			vNombre = Cell.getString
			Cell = Sheet.getCellByPosition(2, y)
			vDireccion = Cell.getString
			Cell = Sheet.getCellByPosition(3, y)
			vDireccion = vDireccion + ", " + Cell.getString
			Cell = Sheet.getCellByPosition(4, y)
			vDireccion = vDireccion + ", " + Cell.getString
			Cell = Sheet.getCellByPosition(5, y)
			vZona = Cell.getString
			Exit For
		End if
		If y = 12005 then
			MsgBox "Nro. de Cliente no encontrado en la base de datos"
			Goto Paso4 
		End If
	Next y
Paso7:
	'CARGA LISTADO DE ASIGNADO EN EL LISTBOX DE DIALOG1
	dlgCT1 = createUnoDialog(DialogLibraries.Standard.Dialog1)
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
	dlgCT1.Model.TextField1.Text = vInfo
	dlgCT1.Model.TextField2.Text = vIdTarea
	dlgCT1.Model.TextField3.Text = vNroCliente 
	dlgCT1.Model.TextField4.Text = vNombre
	dlgCT1.Model.TextField6.Text = vNombre
	dlgCT1.Model.TextField5.Text = vDireccion
	dlgCT1.Model.ComboBox1.text = vZona
	If bTareaExtraHR = 0 then
		dlgCT1.Model.DateField1.text = vFechaApartir
		dlgCT1.Model.ComboBox3.text = vAsignado
	Else
		dlgCT1.Model.DateField1.text = vFechaHR
		dlgCT1.Model.ComboBox3.text = vAsignadoHR
	End If
	dlgCT1.Model.CheckBox1.State = 1
	dlgCT1.Model.CheckBox2.State = 0
	dlgCT1.Model.CheckBox3.State = 0
	dlgCT1.Model.CheckBox4.State = 0
	dlgCT1.Model.CheckBox5.State = 0
	dlgCT1.Model.OptionButton2.State = 1
	If bTareaExtraHR = 0 then
		dlgCT1.Model.OptionButton4.State = 1
	Else
		dlgCT1.Model.OptionButton5.State = 1
	End If
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

Paso8:
	'INICIA DIALOG1 PARA INGRESAR LA TAREA
	dlgCT1.Model.Step = 1
	'modificación para que permita agregar dirección a cliente 0 y .O
	If vNroCliente = "0" or Right(vNroCliente,2) = ".O" then 
		dlgCT1.Model.Step = 2
	End If
	Select Case dlgCT1.Execute()
	Case 1
		'modificación para que permita agregar dirección a cliente 0 y .O
		If vNroCliente = "0" or Right(vNroCliente,2) = ".O" then 
			vNombre = dlgCT1.Model.TextField6.Text
			vDireccion = dlgCT1.Model.TextField7.Text
			vZona = dlgCT1.Model.ComboBox2.text
		End if
		if vZona = "" then
			Msgbox "No ha especificado la zona"
			goto Paso8
		End if
		vInfo =  dlgCT1.Model.TextField1.Text
		vTarea = ""
		if dlgCT1.Model.CheckBox1.State = 1 then vTarea = vTarea + "E"
		if dlgCT1.Model.CheckBox2.State = 1 then vTarea = vTarea + "C"
		if dlgCT1.Model.CheckBox3.State = 1 then vTarea = vTarea + "D"
		if dlgCT1.Model.CheckBox4.State = 1 then vTarea = vTarea + "O"
		if dlgCT1.Model.CheckBox5.State = 1 then vTarea = vTarea + "V"
		vCadena = vTarea
		if vTarea = "" then
			Msgbox "No ha especificado cual es la tarea a realizar"
			goto Paso8
		End if
		if Len(vCadena) = 1 then vTarea = vCadena
		if Len(vCadena) = 2 then vTarea = Mid(vCadena, 1, 1) + "+" + Mid(vCadena, 2, 1)
		if Len(vCadena) = 3 then vTarea = Mid(vCadena, 1, 1) + "+" + Mid(vCadena, 2, 1) + "+" + Mid(vCadena, 3, 1)
		if Len(vCadena) = 4 then vTarea = Mid(vCadena, 1, 1) + "+" + Mid(vCadena, 2, 1) + "+" + Mid(vCadena, 3, 1) + "+" + Mid(vCadena, 4, 1)
		if Len(vCadena) = 5 then vTarea = Mid(vCadena, 1, 1) + "+" + Mid(vCadena, 2, 1) + "+" + Mid(vCadena, 3, 1) + "+" + Mid(vCadena, 4, 1) + "+" + Mid(vCadena, 5, 1)
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
		Exit Sub
	End Select
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
	Cell.String = "" 'vacio
	Cell = Sheet.getCellByPosition(11, yIDT)
	Cell.String = ""
	If vEstado = "FINALIZADO" then Cell.String = Date
	Cell = Sheet.getCellByPosition(12, yIDT)
	Cell.String = vFechaApartir
	Cell = Sheet.getCellByPosition(13, yIDT)
	If bTareaExtraHR = 0 then
		Cell.String = DATE
	Else
		Cell.String = vFechaHR
	End If
	Cell = Sheet.getCellByPosition(14, yIDT)
	Cell.String = Date
	Cell = Sheet.getCellByPosition(15, yIDT)
	If Right(Cell.String,Len(vUsuario)) = vUsuario then
	
	Else
		If Cell.String = "" then
			Cell.String = vUsuario
		Else
			Cell.String = Cell.String + "/" + vUsuario
		End If
	End If
	Cell = Sheet.getCellByPosition(16, yIDT)
	If Cell.getString = "" then  Cell.Value = 0
	Cell = Sheet.getCellByPosition(17, yIDT)
	If Cell.getString = "" then  Cell.Value = 0
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
	FilaActual = yIDT
	FilaNoVisible
	
	'Actualiza una Tarea PENDIENTE o EN CURSO en Expedición Vs Cobros y Borra una FINALIZADA
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	If vEstado = "PENDIENTE" or vEstado = "EN CURSO" then
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = vIdTarea then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(1, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(2, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(3, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(4, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(5, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(6, nFila)
				Cell.String = vEstado + chr(13) + vAsignado
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
				If vEstado = "EN CURSO" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'EN CURSO
					Next z
				End If			
				Goto FinActualizacion
			End If
		Next nFila
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = "" then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(1, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(2, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(3, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(4, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(5, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(6, nFila)
				Cell.String = vEstado + chr(13) + vAsignado
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
				If vEstado = "EN CURSO" then
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'EN CURSO
					Next z
				End If			
				Goto FinActualizacion
			End If
		Next nFila		
	End If
	If vEstado = "FINALIZADO" then
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(0, nFila)
			If Cell.String = vIdTarea then
				'Borra y colorea el fondo
				z = 0
				For z = 0 to 6
					Cell = Sheet.getCellByPosition(z, nFila)
					Cell.String = ""
					Cell.CellBackColor = RGB(213,231,234) 'CELESTE CLARO 'VACIA
				Next z
				Goto FinActualizacion
			End If
		Next nFila
	End If
FinActualizacion:

	dlgCT1.Dispose()
	dlgCT2.Dispose()
	If yIDT = nProxFila then
		vProxTarea = vProxTarea + 1
		nProxFila = nProxFila + 1
	End If
	If bTareaExtraHR <> 0 then Exit Sub
	If Msgbox( "¿Desea cargar otra tarea?", 4 + 32, "" ) = 6 then goto Inicio
	If Msgbox( "¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información?", 4 + 32, "Guardar" ) = 6 then 
		Doc.Store()
		HoraUltGuardar = Timer
		Exit Sub
	End If
	GuardarPorTiempoTranscurrido
	Exit Sub
	
Paso9:
	'Modifica una tarea PENDIENTE o EN CURSO
	vIdTarea = ""
	vDireccion = ""
	vZona = ""
	vTarea = ""
	vCadena = ""
	vPrioridad = ""
	vInfo = ""
	vEstado = ""
	vFechaApartir = ""
	vAsignado = ""
	Cell = Sheet.getCellByPosition(0, yIDT)
	vIdTarea = Cell.getString
	Cell = Sheet.getCellByPosition(2, yIDT)
	vNombre = Cell.getString
	Cell = Sheet.getCellByPosition(3, yIDT)
	vDireccion = Cell.getString
	Cell = Sheet.getCellByPosition(4, yIDT)
	vZona = Cell.getString
	Cell = Sheet.getCellByPosition(5, yIDT)
	vTarea = Cell.getString
	Cell = Sheet.getCellByPosition(6, yIDT)
	vPrioridad = Cell.getString
	Cell = Sheet.getCellByPosition(7, yIDT)
	vInfo = Cell.getString
	Cell = Sheet.getCellByPosition(8, yIDT)
	vEstado = Cell.getString
	Cell = Sheet.getCellByPosition(9, yIDT)
	vAsignado = Cell.getString
	Cell = Sheet.getCellByPosition(12, yIDT)
	vFechaApartir = Cell.getString

	' Carga el listbox de Asignado
	dlgCT1 = createUnoDialog(DialogLibraries.Standard.Dialog1)
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
	dlgCT1.Model.TextField5.Text = vDireccion
	dlgCT1.Model.ComboBox1.text = vZona
	y = 0
	dlgCT1.Model.CheckBox1.State = 0
	dlgCT1.Model.CheckBox2.State = 0
	dlgCT1.Model.CheckBox3.State = 0
	dlgCT1.Model.CheckBox4.State = 0
	dlgCT1.Model.CheckBox5.State = 0
	Pos1 = Len(vTarea)
	For y = 1 to Pos1 Step 2
		CadBuscar = ""
		CadBuscar = Mid(vTarea, y, 1)
		IF CadBuscar = "E" THEN dlgCT1.Model.CheckBox1.State = 1
		IF CadBuscar = "C" THEN dlgCT1.Model.CheckBox2.State = 1
		IF CadBuscar = "D" THEN dlgCT1.Model.CheckBox3.State = 1
		IF CadBuscar = "O" THEN dlgCT1.Model.CheckBox4.State = 1
		IF CadBuscar = "V" THEN dlgCT1.Model.CheckBox5.State = 1
	Next	
	If vPrioridad = "ALTA" then dlgCT1.Model.OptionButton1.State = 1
	If vPrioridad = "MEDIA" then dlgCT1.Model.OptionButton2.State = 1
	If vPrioridad = "BAJA" then dlgCT1.Model.OptionButton3.State = 1	
	If vEstado = "PENDIENTE" then dlgCT1.Model.OptionButton4.State = 1
	If vEstado = "EN CURSO" then dlgCT1.Model.OptionButton5.State = 1
	If vEstado = "FINALIZADO" then dlgCT1.Model.OptionButton6.State = 1	
	dlgCT1.Model.TextField1.Text = vInfo
	dlgCT1.Model.ComboBox3.text = vAsignado
	dlgCT1.Model.DateField1.text = vFechaApartir
	Goto Paso8


End Sub

Sub CursarTareasRapidoCT
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	Dim dAsignado As String
	
'	Dim cmdBotonSig as object

	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
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
	
	BuscaActualizacionesGC
	
	'UnificarTareasPENDIENTES

Inicio:
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 	
	DialogLibraries.LoadLibrary("Standard")
Paso1:
	'CARGA lstCBox1 CORRESPONDIENTE A "ASIGNADO A" DE Dialog13
	Sheet = Doc.Sheets.getByName("Datos")
	dlgCT13 = createUnoDialog(DialogLibraries.Standard.Dialog13)
	olstDatos = dlgCT13.getControl("lstCBox1")
  	For d = 2 to 10	
	 	Cell = Sheet.getCellByPosition(2, d) 	
	  	dAsignado = Cell.String
	  	If dAsignado <> "" then olstDatos.addItem( dAsignado, -1 )
	Next d

Paso1A:
	'ABRE Dialog13 PARA QUE INGRESAR LA INFORMACION BUSCAR
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	Select Case dlgCT13.Execute()
	Case 1
		If dlgCT13.Model.lstCBox1.Text = "" then 
			Msgbox "No ha seleccionado ""Asignada a:"".",48,"Importante"
			goto Paso1A
		End if
		dAsignado = dlgCT13.Model.lstCBox1.Text
	Case 0
		dlgCT13.Dispose()
		Procesando = False
		Exit Sub		
	End Select

Paso2:
	'BUSCA LAS TAREAS EN ESTADO "PENDIENTE"
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	yIDT = 0
	For yIDT = 10 to 10011
		Cell = Sheet.getCellByPosition(1, yIDT)
		If Cell.String = "" then
		'VERIFICA SI LAS PROXIMAS 10 FILAS CONTIENEN INFORMACION
			z = 0
			For z = 1 to 10
				Cell = Sheet.getCellByPosition(1, yIDT + z)
				If Cell.String <> "" then Goto ContinuaBusqueda
			Next z
			Exit For
		End If
ContinuaBusqueda:
		Cell = Sheet.getCellByPosition(8, yIDT)
		If Cell.String = "PENDIENTE" then 
			Cell = Sheet.getCellByPosition(9, yIDT)
			If Cell.String = dAsignado then 
				Posicionar = yIDT
				PosicionadorCelda
				'EXTRAE LA INFORMACION
				Cell = Sheet.getCellByPosition(0, yIDT)
				vIdTarea = Cell.getString
				Cell = Sheet.getCellByPosition(1, yIDT)
				vNroCliente = Cell.getString
				Cell = Sheet.getCellByPosition(2, yIDT)
				vNombre = Cell.getString
				Cell = Sheet.getCellByPosition(3, yIDT)
				vDireccion = Cell.getString
				Cell = Sheet.getCellByPosition(4, yIDT)
				vZona = Cell.getString
				Cell = Sheet.getCellByPosition(5, yIDT)
				vTarea = Cell.getString
				Cell = Sheet.getCellByPosition(6, yIDT)
				vPrioridad = Cell.getString
				Cell = Sheet.getCellByPosition(7, yIDT)
				vInfo = Cell.getString
				Cell = Sheet.getCellByPosition(8, yIDT)
				vEstado = Cell.getString
				Cell = Sheet.getCellByPosition(9, yIDT)
				vAsignado = Cell.getString
			'	Cell = Sheet.getCellByPosition(10, xIDT)
			'	vConcurrio = Cell.getString
				Cell = Sheet.getCellByPosition(11, yIDT)
				vFechaFinalizado = Cell.getString	
				Cell = Sheet.getCellByPosition(12, yIDT)
				vFechaApartir = Cell.getString
				Cell = Sheet.getCellByPosition(13, yIDT)
				vFechaCarga = Cell.getString

				If Date < cDate(vFechaApartir) then
					InfoTareaCE
					If Msgbox( InfoMostrar+chr(13)+chr(13)+"LA FECHA DE INICIO DE LA TAREA Nº "+vIdTarea+" ES SUPERIOR A LA ACTUAL.   "+chr(13)+chr(13)+"¿DESEA MODIFICAR LA FECHA DE INICIO Y CURSAR LA TAREA?"+chr(13), 4 + 256 + 32, "IMPORTANTE" ) = 6 then
						vFechaApartir = Date
					Else
						Goto SiguienteTarea
					End If
				End If		
				If Instr(vInfo, "[GC]") > 0 and Instr(vInfo, "{") > 0 and Instr(vInfo, "}") > 0 then
					If Instr(Mid(vInfo,Instr(vInfo, "{"),Instr(vInfo, "}")-Instr(vInfo, "{")), "Lun") > 0 and Weekday(Date) = 2 then goto ContinuaBusqueda2
					If Instr(Mid(vInfo,Instr(vInfo, "{"),Instr(vInfo, "}")-Instr(vInfo, "{")), "Mar") > 0 and Weekday(Date) = 3 then goto ContinuaBusqueda2
					If Instr(Mid(vInfo,Instr(vInfo, "{"),Instr(vInfo, "}")-Instr(vInfo, "{")), "Mié") > 0 and Weekday(Date) = 4 then goto ContinuaBusqueda2
					If Instr(Mid(vInfo,Instr(vInfo, "{"),Instr(vInfo, "}")-Instr(vInfo, "{")), "Jue") > 0 and Weekday(Date) = 5 then goto ContinuaBusqueda2
					If Instr(Mid(vInfo,Instr(vInfo, "{"),Instr(vInfo, "}")-Instr(vInfo, "{")), "Vie") > 0 and Weekday(Date) = 6 then goto ContinuaBusqueda2
					If Instr(Mid(vInfo,Instr(vInfo, "{"),Instr(vInfo, "}")-Instr(vInfo, "{")), "Sáb") > 0 and Weekday(Date) = 7 then goto ContinuaBusqueda2
					If Instr(Mid(vInfo,Instr(vInfo, "{"),Instr(vInfo, "}")-Instr(vInfo, "{")), "L a V") > 0 and Weekday(Date) > 1 and Weekday(Date) < 7 then goto ContinuaBusqueda2
					InfoTareaCE
					If Msgbox( InfoMostrar+chr(13)+chr(13)+"EL/LOS DIAS ASIGNADOS PARA REALIZAR LA TAREA Nº "+vIdTarea+" NO COINCIDE CON EL ACTUAL.   "+chr(13)+chr(13)+"¿DESEA CURSAR LA TAREA?"+chr(13), 4 + 256 + 32, "IMPORTANTE" ) = 6 then
						
					Else
						Goto SiguienteTarea
					End If
				End If
ContinuaBusqueda2:
				Cell = Sheet.getCellByPosition(8, yIDT)
				Cell.String = "EN CURSO"
				Cell = Sheet.getCellByPosition(12, yIDT)
				Cell.String = vFechaApartir

				'Colorea el fondo de las celdas cursadas
				x = 0
				For x = 1 to 14
					Cell = Sheet.getCellByPosition(x, yIDT)
					Cell.CellBackColor = RGB(102,255,102)'EN CURSO
				Next x

				'Actualiza una Tarea en Expedición Vs Cobros
				Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
				nFila = 0
				For nFila = 5 to 255 'Número
					Cell = Sheet.getCellByPosition(0, nFila)
					If Cell.String = vIdTarea then
						Cell.String = vIdTarea
						Cell = Sheet.getCellByPosition(1, nFila)
						Cell.String = vNroCliente
						Cell = Sheet.getCellByPosition(2, nFila)
						Cell.String = vNombre + chr(13) + vDireccion
						Cell = Sheet.getCellByPosition(3, nFila)
						Cell.String = vZona + chr(13) + vTarea
						Cell = Sheet.getCellByPosition(4, nFila)
						Cell.String = vPrioridad
						Cell = Sheet.getCellByPosition(5, nFila)
						Cell.String = vInfo				
						Cell = Sheet.getCellByPosition(6, nFila)
						Cell.String = "EN CURSO" + chr(13) + vAsignado
						'Colorea fondo cuando es EN CURSO
'						z = 0
'						For z = 0 to 6
'							Cell = Sheet.getCellByPosition(z, nFila)
'							Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
'						Next z
						Exit for
					End If			
				Next nFila
				Sheet = Doc.Sheets.getByName("Carga de Tareas") 
			End If
		End If
SiguienteTarea:
	Next yIDT

	Msgbox "Finalizado.",64,"AVISO"
	If Msgbox( "¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información?", 4 + 32, "Guardar" ) = 6 then 
		Doc.Store()
		HoraUltGuardar = Timer
		Procesando = False
		Exit Sub
	End If
	GuardarPorTiempoTranscurrido
	Procesando = False
End Sub

Sub CargarHojaRutaCT
	Dim vAsig as Integer
	Dim vConc as Integer

	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent

	'CORROBORA QUE HAYA INGRESADO LA CONTRASEÑA DE USUARIO	
	If vUsuario = "" then 
		Sheet = Doc.Sheets.getByName("Usuario")
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
	    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
		otxtPW = oFormulario.getByName("txtPW")
		otxtPWVista = Doc.getCurrentController.getControl( otxtPW )
		otxtPWVista.SetFocus()
		Exit Sub
	End If
	
	BuscaActualizacionesGC
	 	
CargaOtro:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
Paso1:
	'Carga ListCBox dialogo asignado de Dialog3 
	DialogLibraries.LoadLibrary("Standard")
	dlgCT3 = createUnoDialog(DialogLibraries.Standard.Dialog3)
	olstDatos = dlgCT3.getControl("lstCBox2")	
	oHojaDatos = ThisComponent.getSheets.getByName("Datos")	
	oRango = oHojaDatos.getCellRangeByName("C3:C10") 'agregado
	data = oRango.getDataArray()'agregado
	co1 = 0
	Redim src(UBound(data))
	For Each d In data
    	src(co1) = d(0)
      	co1 = co1 + 1
	Next
   	olstDatos.addItems(src, 0)
Paso2:
	vAsignadoHR = ""
	vFechaHR = ""
	'Abre Dialog3 Hoja1
	dlgCT3.Model.DateField1.text = date
	dlgCT3.Model.Step = 1	
	Select Case dlgCT3.Execute()
	Case 1
		vFechaHR = dlgCT3.Model.DateField1.text
		vAsignadoHR =  dlgCT3.Model.lstCBox2.Text	
		If vFechaHR = "" then 
			Msgbox "Ingrese la Fecha de realización de la hoja de ruta" 
			goto Paso2
		End If
		If vAsignadoHR = "" then
			Msgbox "Ingrese el Nombre de la persona que realizó la hoja de ruta" 
			goto Paso2
		End If
		goto Paso3
	Case 0
		dlgCT3.Dispose()
		Exit Sub
	End Select
Paso3:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	vIdTarea = ""	
	yIDT = 0	
	'Abre Dialog3 Hoja2
	dlgCT3.Model.TextField1.Text = ""
	dlgCT3.Model.Step = 2
	Select Case dlgCT3.Execute()
	Case 1
		vIdTarea =  dlgCT3.Model.TextField1.Text		
		'Busca y corrobora que exista el Id.Tarea.
		For yIDT = 10 to 10011 'Número máximo de filas a buscar.
			Cell = Sheet.getCellByPosition(0, yIDT)
			If Cell.Value = CInt(vIdTarea) then
				Cell = Sheet.getCellByPosition(1, yIDT)
				if Cell.String = "" then 
					Msgbox "El número de Id. Tarea no se encuentra registrada."
					goto Paso3
				end if 
				Cell = Sheet.getCellByPosition(8, yIDT)
				if Cell.String = "FINALIZADO" then 
					If Msgbox( "La tarea se encuentra FINALIZADA"+chr(13)+"¿Desea continuar con la carga de otra tarea de la Hoja de Ruta?", 4 + 32, "IMPORTANTE" ) = 6 then
						goto Paso3
					End if
					Exit sub
				end if 
				if Cell.String = "PENDIENTE" then 
					If Msgbox( "La tarea se encuentra PENDIENTE"+chr(13)+"No es una tarea EN CURSO"+chr(13)+"¿Desea continuar y FINALIZARLA?", 4 + 32, "IMPORTANTE" ) = 6 then
						goto Paso4
					End if
					goto Paso3
				end if 
				if Cell.String = "" then 
					If Msgbox( "El estado de la tarea es deconocido"+chr(13)+"¿Desea continuar y FINALIZARLA?", 4 + 32, "IMPORTANTE" ) = 6 then
						goto Paso4
					End if
					goto Paso3
				end if 
				if Cell.String = "EN CURSO" then 
					goto Paso4
				end if 
				Msgbox "No hay coincidencia en el estado de la tarea"
				Exit sub
			End If
		Next yIDT
		Msgbox "El número de Tarea no ha sido encontrado"
		goto Paso3
	Case 0
		dlgCT3.Dispose()
		Doc.Store()
		HoraUltGuardar = Timer
		Exit Sub
	End Select
Paso4:
	Posicionar = yIDT
	PosicionadorCelda
	'Carga ListCBox dialogo asignado de Dialog4 
	DialogLibraries.LoadLibrary("Standard")
	dlgCT4 = createUnoDialog(DialogLibraries.Standard.Dialog4)
	olstDatos = dlgCT4.getControl("lstCBox3")	
	oHojaDatos = ThisComponent.getSheets.getByName("Datos")	
	oRango = oHojaDatos.getCellRangeByName("C3:C6") 'agregado
	data = oRango.getDataArray()'agregado
	co1 = 0
	Redim src(UBound(data))
	For Each d In data
    	src(co1) = d(0)
      	co1 = co1 + 1
	Next
   	olstDatos.addItems(src, 0)

	vNroCliente = ""
	vNombre = ""
	vDireccion = ""
	vZona = ""
	vTarea = ""
	vPrioridad = ""
	vInfo = ""
	vEstado = ""
	vAsignado = ""
	vConcurrio = ""
	vObjetivo = ""
	vFechaCarga = ""
	vFechaApartir = ""

	CadBuscar1 = ""
	
	Sheet = Doc.Sheets.getByName("Carga de Tareas")

	Cell = Sheet.getCellByPosition(1, yIDT)
	vNroCliente = Cell.getString
	Cell = Sheet.getCellByPosition(2, yIDT)
	vNombre = Cell.getString
	Cell = Sheet.getCellByPosition(3, yIDT)
	vDireccion = Cell.getString
	Cell = Sheet.getCellByPosition(4, yIDT)
	vZona = Cell.getString
	Cell = Sheet.getCellByPosition(5, yIDT)
	vTarea = Cell.getString
		dlgCT4.Model.CheckBox1.State = 0
		dlgCT4.Model.CheckBox2.State = 0
		dlgCT4.Model.CheckBox3.State = 0
		dlgCT4.Model.CheckBox4.State = 0
		dlgCT4.Model.CheckBox5.State = 0
		Pos1 = Len(vTarea)
		For z = 1 to Pos1 Step 2
			CadBuscar1 = ""
			CadBuscar1 = Mid(vTarea, z, 1)
			IF CadBuscar1 = "E" THEN dlgCT4.Model.CheckBox1.State = 1
			IF CadBuscar1 = "C" THEN dlgCT4.Model.CheckBox2.State = 1
			IF CadBuscar1 = "D" THEN dlgCT4.Model.CheckBox3.State = 1
			IF CadBuscar1 = "O" THEN dlgCT4.Model.CheckBox4.State = 1
			IF CadBuscar1 = "V" THEN dlgCT4.Model.CheckBox5.State = 1
		Next z	
	Cell = Sheet.getCellByPosition(6, yIDT)
	vPrioridad = Cell.getString
	Cell = Sheet.getCellByPosition(7, yIDT)
	vInfo = Cell.getString
	Cell = Sheet.getCellByPosition(8, yIDT)
	vEstado = Cell.getString
	Cell = Sheet.getCellByPosition(9, yIDT)
	vAsignado = Cell.getString
		If vAsignado <> vAsignadoHR then
			If Msgbox( "No coincide el nombre de la persona asignada en la Hoja de Ruta con lo ingresado en la Carga de Tareas"+chr(13)+"¿Desea modificar la información ingresada en Carga de Tareas?", 4 + 32, "IMPORTANTE" ) = 6 then
				vAsignado = vAsignadoHR
				dlgCT4.Model.lstCBox3.ReadOnly = False
				dlgCT4.Model.lstCBox3.Border = 2 
				dlgCT4.Model.lstCBox3.BackgroundColor = RGB (255,255,255) 
			End if
		End if
'	Cell = Sheet.getCellByPosition(10, yIDT)
'	vConcurrio = Cell.getString
'	Cell = Sheet.getCellByPosition(11, yIDT)
'	vObjetivo = Cell.getString
	Cell = Sheet.getCellByPosition(12, yIDT)
	vFechaApartir = Cell.getString
	Cell = Sheet.getCellByPosition(13, yIDT)
	vFechaCarga = Cell.getString

 
	dlgCT4.Model.TextField1.Text = vIdTarea
	dlgCT4.Model.TextField2.Text = vNroCliente 
	dlgCT4.Model.TextField3.Text = vNombre
	dlgCT4.Model.TextField4.Text = vDireccion
	dlgCT4.Model.TextField5.Text = vZona
	dlgCT4.Model.TextField6.Text = vPrioridad
	dlgCT4.Model.TextField7.Text = vEstado
	dlgCT4.Model.TextField8.Text = vFechaCarga
	dlgCT4.Model.TextField9.Text = vFechaApartir
	dlgCT4.Model.TextField11.Text = vFechaHR
	dlgCT4.Model.TextField10.Text = vInfo
	dlgCT4.Model.TextField12.Text = ""
	dlgCT4.Model.lstCBox3.Text = vAsignado
	
	Select Case dlgCT4.Execute()
	Case 1
		vTarea = ""
		if dlgCT4.Model.CheckBox1.State = 1 then vTarea = vTarea + "E"
		if dlgCT4.Model.CheckBox2.State = 1 then vTarea = vTarea + "C"
		if dlgCT4.Model.CheckBox3.State = 1 then vTarea = vTarea + "D"
		if dlgCT4.Model.CheckBox4.State = 1 then vTarea = vTarea + "O"
		if dlgCT4.Model.CheckBox5.State = 1 then vTarea = vTarea + "V"
		CadBuscar1 = vTarea
		if Len(CadBuscar1) = 1 then vTarea = CadBuscar1
		if Len(CadBuscar1) = 2 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1)
		if Len(CadBuscar1) = 3 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1)
		if Len(CadBuscar1) = 4 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1) + "+" + Mid(CadBuscar1, 4, 1)
		if Len(CadBuscar1) = 5 then vTarea = Mid(CadBuscar1, 1, 1) + "+" + Mid(CadBuscar1, 2, 1) + "+" + Mid(CadBuscar1, 3, 1) + "+" + Mid(CadBuscar1, 4, 1) + "+" + Mid(CadBuscar1, 5, 1)
		Cell = Sheet.getCellByPosition(5, yIDT)
		Cell.String = vTarea

		If dlgCT4.Model.TextField12.Text = "" then
		
		Else
			vInfo = vInfo + "/" + dlgCT4.Model.TextField12.Text
		End If
		Cell = Sheet.getCellByPosition(7, yIDT)
		Cell.String = vInfo
		

		vAsig = 0
		vConc = 0
		If dlgCT4.Model.OptionButton1.State = 1 then 
			vConcurrio = "SI"
			Cell = Sheet.getCellByPosition(16, yIDT)
			vAsig = Cell.getString
			vAsig = vAsig + 1
			Cell.String = vAsig
			Cell = Sheet.getCellByPosition(17, yIDT)
			vConc = Cell.getString
			vConc = vConc + 1
			Cell.String = vConc
		End if
		If dlgCT4.Model.OptionButton2.State = 1 then 
			vConcurrio = "NO"
			Cell = Sheet.getCellByPosition(16, yIDT)
			vAsig = Cell.getString
			vAsig = vAsig + 1
			Cell.String = vAsig
			If vPrioridad = "MEDIA" and vAsig >= 3 then vPrioridad = "ALTA"
			If vPrioridad = "BAJA" and vAsig >= 3 then vPrioridad = "MEDIA"
			Cell = Sheet.getCellByPosition(6, yIDT)
			Cell.String = vPrioridad
			vObjetivo = "NO"
			vEstado = "PENDIENTE"
			' Colorea el fondo de las celdas de Carga de Tareas
			x = 0
			For x = 1 to 14  
				Cell = Sheet.getCellByPosition(x, yIDT)
				Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
			Next x
			goto Paso6
		End if
		If dlgCT4.Model.OptionButton3.State = 1 then 
			vObjetivo = "SI"
			vEstado = "FINALIZADO"
			Cell = Sheet.getCellByPosition(11, yIDT)
			Cell.String = vFechaHR
			' Colorea el fondo de las celdas de Carga de Tareas
			x = 0
			For x = 1 to 14
				Cell = Sheet.getCellByPosition(x, yIDT)
				Cell.CellBackColor = RGB(255,102,102) 'FINALIZADO
			Next x
		End If
		If dlgCT4.Model.OptionButton4.State = 1 then 
			vObjetivo = "NO"
			vEstado = "PENDIENTE"
			' Colorea el fondo de las celdas de Carga de Tareas
			x = 0
			For x = 1 to 14
				Cell = Sheet.getCellByPosition(x, yIDT)
				Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
			Next x
		End if
Paso6:
		Cell = Sheet.getCellByPosition(17+vAsig, yIDT)
		Cell.String = vFechaHR + "-" + vConcurrio + "-" + vObjetivo + "-" + vAsignado
		Cell = Sheet.getCellByPosition(8, yIDT)
		Cell.String = vEstado
		Cell = Sheet.getCellByPosition(9, yIDT)
		Cell.String = vAsignado
		Cell = Sheet.getCellByPosition(10, yIDT)
		Cell.String = ""'vConcurrio
		Cell = Sheet.getCellByPosition(15, yIDT)
		If Right(Cell.String,Len(vUsuario)) = vUsuario then
	
		Else
			If Cell.String = "" then
				Cell.String = vUsuario
			Else
				Cell.String = Cell.String + "/" + vUsuario
			End If
		End If
		FilaActual = yIDT
		FilaNoVisible
	
		'Actualiza una Tarea PENDIENTE o EN CURSO en Expedición Vs Cobros y Borra una FINALIZADA
		Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
		If vEstado = "PENDIENTE" then
			nFila = 0
			For nFila = 5 to 255 'Número
				Cell = Sheet.getCellByPosition(0, nFila)
				If Cell.String = vIdTarea then
					Cell.String = vIdTarea
					Cell = Sheet.getCellByPosition(1, nFila)
					Cell.String = vNroCliente
					Cell = Sheet.getCellByPosition(2, nFila)
					Cell.String = vNombre + chr(13) + vDireccion
					Cell = Sheet.getCellByPosition(3, nFila)
					Cell.String = vZona + chr(13) + vTarea
					Cell = Sheet.getCellByPosition(4, nFila)
					Cell.String = vPrioridad
					Cell = Sheet.getCellByPosition(5, nFila)
					Cell.String = vInfo				
					Cell = Sheet.getCellByPosition(6, nFila)
					Cell.String = vEstado + chr(13) + vAsignado
					'Colorea fondo cuando es PENDIENTE
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
					Next z
					goto FinActualizacion
				End If
			Next nFila
		End If
		If vEstado = "FINALIZADO" then
			nFila = 0
			For nFila = 5 to 255 'Número
				Cell = Sheet.getCellByPosition(0, nFila)
				If Cell.String = vIdTarea then
					'Borra y colorea el fondo
					z = 0
					For z = 0 to 6
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.String = ""
						Cell.CellBackColor = RGB(213,231,234) 'CELESTE CLARO 'VACIA
					Next z
					goto FinActualizacion
				End If
			Next nFila
		End If
FinActualizacion:
		goto Paso5
	Case 0
		dlgCT4.Dispose()
		Doc.Store()
		HoraUltGuardar = Timer
		Exit Sub
	End Select

Paso5:
	If Msgbox( "¿La Hoja de Ruta posee otra tarea que desee ingresar?", 4 + 32, "IMPORTANTE" ) = 6 then
		Goto Paso3
	End if

	If Msgbox( "¿Desea ingresar alguna tarea extra realizada por la persona asiganada?", 4 + 32, "IMPORTANTE" ) = 6 then
		bTareaExtraHR = 1	'VARIABLE = TRUE PERMITE EN CARGA DE TAREAS CARGAR UNA TAREA EXTRA REALIZADA EN LA HOJA DE RUTA 
		CargarTareaCT
		bTareaExtraHR = 0
		goto Paso4
	End if
	Doc.Store()
	HoraUltGuardar = Timer

End Sub




Sub BuscaProxTareaDisponibleCT
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	nProxFila = 0
	vProxTarea = 0
	y = 0
	For y = 10 to 8000 'Número máximo de filas a buscar.
		Cell = Sheet.getCellByPosition(1, y)
		If Cell.String = "" then 
			Cell = Sheet.getCellByPosition(0, y)
			If Cell.String = "" then
				Cell = Sheet.getCellByPosition(2, y)
				If Cell.String = "" then				
					nProxFila = y
					Exit For
				Else
					Posicionar = y
					PosicionadorCelda
					Msgbox "Hay una tarea que no dispone Id.Tarea y Nro. de Cliente.",16,"IMPORTANTE"
					Exit For
				End If
			Else
				Cell = Sheet.getCellByPosition(2, y)
				If Cell.String = "" then				
					nProxFila = y
					Cell = Sheet.getCellByPosition(0, y)
					vProxTarea = CInt(Cell.String)
ControlaSiguiente:
					Cell = Sheet.getCellByPosition(0, y + 1)
					If vProxTarea = CInt(Cell.String) then
						Sheet.getRows.removeByIndex( nProxFila + 1 , 1 )
						Goto ControlaSiguiente
					End If
					Exit For
				Else
					Posicionar = y
					PosicionadorCelda
					Msgbox "Hay una tarea que no dispone Nro. de Cliente.",16,"IMPORTANTE"
					Exit For
				End If
			
			End If	
		End If
	Next y
	If nProxFila > 5010 then MsgBox "Se han superado las 5000 tareas."+chr(13)+"Informe al programador para que las archive y así acelerar el tiempo de respuesta del programa.",,"Importante"
	If vProxTarea > 0 and nProxFila > 0 then
		vIdTarea = CStr(vProxTarea)
	End If
	If vProxTarea = 0 and nProxFila > 0 then NuevaFilaCT
End Sub

Sub NuevaFilaCT
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	vIdTarea = ""
	Cell = Sheet.getCellByPosition(0, nProxFila)
	If Cell.String = "" then
		Cell = Sheet.getCellByPosition(0, nProxFila-1)
		If Cell.String <> "" then
			vProxTarea = Cell.Value + 1
			vIdTarea = CStr(vProxTarea) 
		End IF	
	End If
	If vIdTarea <> "" then
		Sheet.getRows.insertByIndex( nProxFila , 1 ) 'INSERTA UNA NUEVA FILA
 	
		Cell = Sheet.getCellByPosition(0, nProxFila)
		Cell.Value = vProxTarea
		For x = 1 to 50
			Cell = Sheet.getCellByPosition(x, nProxFila)
			Cell.String = ""
			If x > 0 and x < 15 then
				Cell.CellBackColor = RGB(255,255,255) 'BLANCO = PENDIENTE
			End If
		Next x
	Else
		BuscaProxTareaDisponibleCT
	End If
	
End Sub

Sub PosicionadorCelda
	Posicionar = Posicionar + 1
	Posicionador = "B" + Posicionar
	PosCel(0).Name = "ToPoint"
	PosCel(0).Value = Posicionador
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, PosCel())
End Sub

Sub GuardarPorTiempoTranscurrido
	'CORROBORA SI HA TRANSCURRIDO EL TIEMPO PREDETERMINADO EN SEGUNDOS ENTRE LA HORA ACTUAL Y LA ULTIMA ACTUALIZACION.
	' 7200 SEGUNDOS = 2 HORAS
	'EN TAL CASO GUARDA EL DOCUMENTO
	Dim HoraActual As Long
	HoraActual = Timer
	If HoraUltGuardar + 7200 < HoraActual then
		Doc.Store()
		HoraUltGuardar = Timer
	End If
End Sub

Sub BuscaActualizacionesGC
	'CORROBORA ACTUALIZACIONES EN GESTION DE COBROS
	If RutaOrigen = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(11, 1)
		RutaOrigen = ConvertToURL( Cell.String )
	End If
	If ModArchivoGC = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(12, 1)
		ModArchivoGC = Cell.String
	End If
	If cDate(Left(ModArchivoGc, 10)) <> cDate(Left(FileDateTime( RutaOrigen ), 10)) then
		If cDate(Left(ModArchivoGc, 10)) > cDate(Left(FileDateTime( RutaOrigen ), 10)) then
			
		Else
			Msgbox "GESTION DE COBROS a actualizado su información."+chr(13)+chr(13)+"Última Modificación G.C.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior G.C.: "+ModArchivoGc+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	else
		If cDate(Right(ModArchivoGc, 8)) < cDate(Right(FileDateTime( RutaOrigen ), 8)) then
			Msgbox "GESTION DE COBROS a actualizado su información."+chr(13)+chr(13)+"Última Modificación G.C.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior G.C.: "+ModArchivoGc+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	End If
End Sub

Sub InfoTareaCE
	InfoMostrar = ""
	InfoMostrar = "Id.Tarea: " + vIdTarea
	InfoMostrar = InfoMostrar + chr(13) + "Nro.Cliente: " + vNroCliente
	InfoMostrar = InfoMostrar + chr(13) + "Cliente/Destinatario: " + vNombre
	InfoMostrar = InfoMostrar + chr(13) + "Dirección: " + vDireccion
	InfoMostrar = InfoMostrar + chr(13) + "Zona: " + vZona
	InfoMostrar = InfoMostrar + chr(13) + "Tarea: " + vTarea
	InfoMostrar = InfoMostrar + chr(13) + "Prioridad: " + vPrioridad
	InfoMostrar = InfoMostrar + chr(13) + "Comentarios:" + chr(13) + vInfo
	InfoMostrar = InfoMostrar + chr(13) + "Estado: "  + vEstado
	InfoMostrar = InfoMostrar + chr(13) + "Asignado: " + vAsignado
	InfoMostrar = InfoMostrar + chr(13) + "A partir de: "  + vFechaApartir
	InfoMostrar = InfoMostrar + chr(13) + "Fecha de Carga: "  + vFechaCarga
	InfoMostrar = InfoMostrar + chr(13) + "Fecha de Finalizado: "  + vFechaFinalizado
End Sub

'Function CargarDialogo( Nombre As String ) As Object
	'Cargamos la librería Standard en memoria
'	DialogLibraries.LoadLibrary( "Standard" )
	'Cargamos el cuadro de diálogo en memoria
'	CargarDialogo = CreateUnoDialog( DialogLibraries.Standard.getByName( Nombre ) )
'End Function


Sub EstadoTareasCT
	Dim cCBANPend as Integer, cCBACPend as Integer, cCBASPend as Integer, cINTPend as Integer, cSDEFPend as Integer, cCBANEC as Integer
	Dim cCBACEC as Integer, cCBASEC as Integer, cINTEC as Integer, cSDEFEC as Integer, cCBANFin as Integer, cCBACFin as Integer
	Dim cCBASFin as Integer, cINTFin as Integer, cSDEFFin as Integer, cCBANSD as Integer, cCBACSD as Integer, cCBASSD as Integer
	Dim cINTSD as Integer, cSDEFSD as Integer

Inicio:	
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	Doc = thiscomponent
	DialogLibraries.LoadLibrary("Standard")

	'CORROBORA QUE HAYA INGRESADO LA CONTRASEÑA DE USUARIO	
	If vUsuario = "" then 
		Sheet = Doc.Sheets.getByName("Usuario")
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
	    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
		otxtPW = oFormulario.getByName("txtPW")
		otxtPWVista = Doc.getCurrentController.getControl( otxtPW )
		otxtPWVista.SetFocus()
		Exit Sub
	End If
	
	BuscaActualizacionesGC

Paso1:
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 	
	dlgCT11 = createUnoDialog(DialogLibraries.Standard.Dialog11)
Paso2:
	'Busca las tareas que se encuentran en Estado PENDIENTE
	cCBANPend = 0
	cCBACPend = 0
	cCBASPend = 0
	cINTPend = 0
	cSDEFPend = 0
	cCBANEC = 0
	cCBACEC = 0
	cCBASEC = 0
	cINTEC = 0
	cSDEFEC = 0
	cCBANFin = 0
	cCBACFin = 0
	cCBASFin = 0
	cINTFin = 0
	cSDEFFin = 0
	cCBANSD = 0
	cCBACSD = 0
	cCBASSD = 0
	cINTSD = 0
	cSDEFSD = 0
	yIDT = 0
	For yIDT = 10 to 8000
		Cell = Sheet.getCellByPosition(0, yIDT)
		If Cell.String = "" then
			Cell = Sheet.getCellByPosition(0, yIDT + 1)
			If Cell.String = "" then goto Paso3
		End If
		Cell = Sheet.getCellByPosition(8, yIDT)
		If Cell.String = "PENDIENTE" then
			Cell = Sheet.getCellByPosition(4, yIDT)
			If Cell.String = "CBAN" then cCBANPend = cCBANPend + 1
			If Cell.String = "CBAC" then cCBACPend = cCBACPend + 1
			If Cell.String = "CBAS" then cCBASPend = cCBASPend + 1
			If Cell.String = "INT" then cINTPend = cINTPend + 1
			If Cell.String = "Sin Definir" then cSDEFPend = cSDEFPend + 1
			If Cell.String = "" then cSDEFPend = cSDEFPend + 1
		End IF
		If Cell.String = "EN CURSO" then
			Cell = Sheet.getCellByPosition(4, yIDT)
			If Cell.String = "CBAN" then cCBANEC = cCBANEC + 1
			If Cell.String = "CBAC" then cCBACEC = cCBACEC + 1
			If Cell.String = "CBAS" then cCBASEC = cCBASEC + 1
			If Cell.String = "INT" then cINTEC = cINTEC + 1
			If Cell.String = "Sin Definir" then cSDEFEC = cSDEFEC + 1
			If Cell.String = "" then cSDEFEC = cSDEFEC + 1
		End IF
		If Cell.String = "FINALIZADO" then
			Cell = Sheet.getCellByPosition(4, yIDT)
			If Cell.String = "CBAN" then cCBANFin = cCBANFin + 1
			If Cell.String = "CBAC" then cCBACFin = cCBACFin + 1
			If Cell.String = "CBAS" then cCBASFin = cCBASFin + 1
			If Cell.String = "INT" then cINTFin = cINTFin + 1
			If Cell.String = "Sin Definir" then cSDEFFin = cSDEFFin + 1
			If Cell.String = "" then cSDEFFin = cSDEFFin + 1
		End IF
		If Cell.String = "" then
			Cell = Sheet.getCellByPosition(4, yIDT)
			If Cell.String = "CBAN" then cCBANSD = cCBANSD + 1
			If Cell.String = "CBAC" then cCBACSD = cCBACSD + 1
			If Cell.String = "CBAS" then cCBASSD = cCBASSD + 1
			If Cell.String = "INT" then cINTSD = cINTSD + 1
			If Cell.String = "Sin Definir" then cSDEFSD = cSDEFSD + 1
			If Cell.String = "" then cSDEFSD = cSDEFSD + 1
		End IF
	Next yIDT
Paso3:	
	'Carga informacion para el cuadro de dialogo
	dlgCT11.Model.TextField1.Text = cCBANPend
	dlgCT11.Model.TextField2.Text = cCBACPend
	dlgCT11.Model.TextField3.Text = cCBASPend
	dlgCT11.Model.TextField4.Text = cINTPend
	dlgCT11.Model.TextField5.Text = cSDEFPend
	dlgCT11.Model.TextField6.Text = cCBANEC
	dlgCT11.Model.TextField7.Text = cCBACEC	
	dlgCT11.Model.TextField8.Text = cCBASEC	
	dlgCT11.Model.TextField9.Text = cINTEC	
	dlgCT11.Model.TextField10.Text = cSDEFEC
	dlgCT11.Model.TextField11.Text = cCBANFin	
	dlgCT11.Model.TextField12.Text = cCBACFin	
	dlgCT11.Model.TextField13.Text = cCBASFin	
	dlgCT11.Model.TextField14.Text = cINTFin	
	dlgCT11.Model.TextField15.Text = cSDEFFin
	dlgCT11.Model.TextField16.Text = cCBANSD
	dlgCT11.Model.TextField17.Text = cCBACSD	
	dlgCT11.Model.TextField18.Text = cCBASSD	
	dlgCT11.Model.TextField19.Text = cINTSD	
	dlgCT11.Model.TextField20.Text = cSDEFSD
	dlgCT11.Model.TextField21.Text = cCBANPend+cCBACPend+cCBASPend+cINTPend+cSDEFPend
	dlgCT11.Model.TextField22.Text = cCBANEC+cCBACEC+cCBASEC+cINTEC+cSDEFEC
	dlgCT11.Model.TextField23.Text = cCBANFin+cCBACFin+cCBASFin+cINTFin+cSDEFFin
	dlgCT11.Model.TextField24.Text = cCBANSD+cCBACSD+cCBASSD+cINTSD+cSDEFSD
	dlgCT11.Model.TextField25.Text = cCBANPend+cCBANEC+cCBANFin+cCBANSD
	dlgCT11.Model.TextField26.Text = cCBACPend+cCBACEC+cCBACFin+cCBACSD
	dlgCT11.Model.TextField27.Text = cCBASPend+cCBASEC+cCBASFin+cCBASSD
	dlgCT11.Model.TextField28.Text = cINTPend+cINTEC+cINTFin+cINTSD
	dlgCT11.Model.TextField29.Text = cSDEFPend+cSDEFEC+cSDEFFin+cSDEFSD
	dlgCT11.Model.TextField30.Text = cCBANPend+cCBANEC+cCBANFin+cCBANSD+cCBACPend+cCBACEC+cCBACFin+cCBACSD+cCBASPend+cCBASEC+cCBASFin+cCBASSD+cINTPend+cINTEC+cINTFin+cINTSD+cSDEFPend+cSDEFEC+cSDEFFin+cSDEFSD

	Posicionar = yIDT
	PosicionadorCelda
	
	Select Case dlgCT11.Execute()
	Case 1
		Exit sub
	Case 0
		Exit Sub
	End Select
End Sub

Function BotonSiguiente
	cmdBoton = "Siguiente"
	dlgCT10.EndExecute()
	
End Function
