REM  *****  BASIC  *****

Option Explicit

'Contadores y Buscadores
Dim x
Dim xIDT
Dim mIDT
Dim y
Dim z
Dim Pos1
Dim Pos2
Dim nFila
Dim CadBuscar As String
Dim CadBuscar1 As String
Dim CadBuscar2 As String
Dim CadResultado As String

'Hoja de Calculo
Dim Doc As Object
Dim Sheet As Object
Dim Cell As Object
Dim CellRange As Object
Dim Flags As Long

'USUARIO
Dim otxtPWVista As Object

'POSICIONADOR DE CELDA
Dim document   as object
Dim dispatcher as object	
Dim PosCel(0) as new com.sun.star.beans.PropertyValue
Dim Posicionador As String
Dim Posicionar

'Variables ListBox de los dialogos
Dim oHojaDatos As Object
Dim co1 As Long
Dim oRango As Object
Dim data
Dim src
Dim d
Dim olstDatos As Object

Dim chkFEFinalizado As Object
Dim chkFEPendiente As Object
Dim chkFEEnCurso As Object

Global FilaActual
Global vProxTarea As Integer, nProxFila As Integer
Dim vIdTarea as String
Dim oBarraEstado As Object

Private dlgCT9 as Object

Dim vIdTarea as String, vNroCliente as String, vNombre as String, vDireccion as string
Dim vZona as String, vTarea as String, vPrioridad as String, vInfo as String 
Dim vEstado As String, vAsignado As String, vConcurrio As String, vObjetivo As String
Dim vFechaApartir As String, vFechaCarga As String, vFechaFinalizado as String
Dim vUltMod As String

'Boton siguiente
'Dim cmdBotonSig as object

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
'    chkFEEnCurso = oFormulario.getByName( "CVerEstado3" ) 	
 	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	Cell = Sheet.getCellByPosition(9, FilaActual)
	If chkFEFinalizado.State = 0 and Cell.String = "FINALIZADO" then	
		Sheet.getRows.getByindex(FilaActual).IsVisible = False
	End If
	If chkFEPendiente.State = 0 and Cell.String = "PENDIENTE" then	
		Sheet.getRows.getByindex(FilaActual).IsVisible = False
	End If
'	If chkFEEnCurso.State = 0 and Cell.String = "EN CURSO" then	
'		Sheet.getRows.getByindex(FilaActual).IsVisible = False
'	End If
End Sub

Sub FiltroEstado
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	Doc = thiscomponent
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
    chkFEFinalizado = oFormulario.getByName( "CVerEstado1" )
    chkFEPendiente = oFormulario.getByName( "CVerEstado2" )
'    chkFEEnCurso = oFormulario.getByName( "CVerEstado3" ) 
 	
 	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator
 	
 	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	oBarraEstado.start( "Filtrando Estado ", 2000 )
	For y = 11 to 5000
		If y = 100 or y = 300 or y = 600 Then oBarraEstado.setValue( y )
		If y = 800 or y = 1000 or y = 1500 Then oBarraEstado.setValue( y )
		If y = 2000 or y = 3000 or y = 4000 Then oBarraEstado.setValue( y )

		Cell = Sheet.getCellByPosition(9, y)
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
'			If Cell.String = "EN CURSO" then
'				If chkFEEnCurso.State = 1 then
'					Sheet.getRows.getByindex(y).IsVisible = True
'				Else
'					Sheet.getRows.getByindex(y).IsVisible = False
'				End If
'			End If
		Else
			Posicionar = y
			PosicionadorCelda		
			Exit For
		End If
	Next Y
	oBarraEstado.end()
	Procesando = False
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

Sub BuscaProxTareaDisponibleCT
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	nProxFila = 0
	vProxTarea = 0
	y = 0
	For y = 11 to 8000 'Número máximo de filas a buscar.
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
			If x = 10 or x = 11 then
				Cell.CellBackColor = RGB(207,231,245) 'CELESTE CLARO 
			End If
		Next x
	Else
		BuscaProxTareaDisponibleCT
	End If
	
End Sub


Sub CargarTareaGestionCobro
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	Dim vIdTarea as String, vNroCliente as String, vNombre as String, vDireccion as string, vZona as String, vTarea as String, vCadena as String
	Dim vPrioridad as String, vInfo as String, vOtros as String, vEstado As String, vFechaApartir As String
'	Dim vAsignado As String
	'OTROS
	Dim InfoMostrar As String
	'DIALOGOS
	Dim dlgCT2 as Object, dlgCT1 as Object
	
	Dim ctFila

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
	
	'CORROBORA ACTUALIZACIONES EN CONTROL DE EGRESO
	If RutaOrigen = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(11, 2)
		RutaOrigen = ConvertToURL( Cell.String )
	End If
	If ModArchivoCE = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(12, 2)
		ModArchivoCE = Cell.String
	End If
	If cDate(Left(ModArchivoCE, 10)) <> cDate(Left(FileDateTime( RutaOrigen ), 10)) then
		If cDate(Left(ModArchivoCE, 10)) > cDate(Left(FileDateTime( RutaOrigen ), 10)) then
			
		Else
			Msgbox "EXPEDICIÓN a actualizado su información."+chr(13)+chr(13)+"Última Modificación C.E.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior C.E.: "+ModArchivoCE+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	else
		If cDate(Right(ModArchivoCE, 8)) < cDate(Right(FileDateTime( RutaOrigen ), 8)) then
			Msgbox "EXPEDICIÓN a actualizado su información."+chr(13)+chr(13)+"Última Modificación C.E.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior C.E.: "+ModArchivoCE+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	End If
Inicio:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	DialogLibraries.LoadLibrary("Standard")
	dlgCT1 = createUnoDialog(DialogLibraries.Standard.Dialog1)
	dlgCT2 = createUnoDialog(DialogLibraries.Standard.Dialog2)
Paso1:
 
Paso2:
	'Busca la primera celda vacia de Nro. de Cliente en Columna B.
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	ctFila = 0
	If vProxTarea = 0 or nProxFila = 0 then
		BuscaProxTareaDisponibleCT
	Else
		If nProxFila > 0 then NuevaFilaCT
	End If
	Cell = Sheet.getCellByPosition(0, nProxFila)
	vIdTarea = Cell.String
	ctFila = nProxFila
'	yIDT = nProxFila	'eliminar o no
	If nProxFila = 0 or vIdTarea = "" then 
		Procesando = False
		Exit Sub
	End If

Paso3:
	Posicionar = nProxFila 
	PosicionadorCelda

	'carga el listado de clientes en el listbox de dialog2
	dlgCT2 = CargarDialogo("Dialog2")'agregado
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
	'abre dialog2 para ingresar el cliente o destinatario
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
		Procesando = False
		Exit Sub
	End Select

Paso5:
	'BUSCA TAREAS PENDIENTES O EN CURSO EN CONTROL DE EGRESO EN LA HOJA EXPEDICION-COBROS
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	y = 0
	For y = 5 to 255 'Número máximo de filas a buscar en Expedicion-Cobros.
		x = 0 'columna inicial de CONTROL DE EGRESO en EXPEDICION-COBROS
		Cell = Sheet.getCellByPosition(1, y)
		If vNroCliente = "0" then
			If Cell.String = "0" then
			Cell = Sheet.getCellByPosition(2, y)
				If Left(vNombre, 5) = Left(Cell.String, 5) then
					Cell = Sheet.getCellByPosition(6, y) 
					If Left(Cell.String, 8) = "EN CURSO" then
						Gosub InformacionVentanas
						Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
						Msgbox	"EL CLIENTE O DESTINATARIO PODRIA TENER UNA TAREA ¨CURSO¨."+chr(13)+"CORROBORAR CON EL SECTOR DE EXPEDICIÓN:" +chr(13)+chr(13)+chr(13)+ InfoMostrar, 48,"Importante"
					End If
					Cell = Sheet.getCellByPosition(6, y)
					If Left(Cell.String, 9) = "PENDIENTE" then
						Gosub InformacionVentanas
						Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
						Msgbox	"EL CLIENTE O DESTINATARIO PODRIA TENER UNA TAREA ¨PENDIENTE¨."+chr(13)+"CORROBORAR CON EL SECTOR DE EXPEDICIÓN:" +chr(13)+chr(13)+chr(13)+ InfoMostrar, 48,"Importante"
					End If
				End IF
			End If		
		Else
			If vNroCliente = Cell.String then 
				Cell = Sheet.getCellByPosition(6, y) 
				If Left(Cell.String, 8) = "EN CURSO" then
					Gosub InformacionVentanas
					Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
					Msgbox	"EL CLIENTE POSEE UNA TAREA ¨EN CURSO¨."+chr(13)+"CORROBORAR CON EL SECTOR DE EXPEDICIÓN:" +chr(13)+chr(13)+chr(13)+ InfoMostrar, 48,"Importante"
				End if
				Cell = Sheet.getCellByPosition(6, y)
				If Left(Cell.String, 9) = "PENDIENTE" then
					Gosub InformacionVentanas
					Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
					Msgbox	"EL CLIENTE POSEE UNA TAREA ¨PENDIENTE¨."+chr(13)+"CORROBORAR CON EL SECTOR DE EXPEDICIÓN:" +chr(13)+chr(13)+chr(13)+ InfoMostrar, 48,"Importante"
				End if
			End If
		End if
	Next y
	
	'BUSCA TAREAS PENDIENTES DEL CLIENTE INGRESADO EN CURSO EN GESTION DE COBROS EN LA HOJA EXPEDICION-COBROS
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	y = 0
	For y = 5 to 255 'Número máximo de filas a buscar en Expedicion-Cobros.
		x = 8 'columna inicial de GESTION DE COBROS en EXPEDICION-COBROS
		Cell = Sheet.getCellByPosition(9, y)
		If vNroCliente = "0" then
			If Cell.String = "0" then
			Cell = Sheet.getCellByPosition(10, y)
				If Left(vNombre, 5) = Left(Cell.String, 5) then
'					Cell = Sheet.getCellByPosition(14, y) 
'					If Left(Cell.String, 9) = "PENDIENTE" then
						Gosub InformacionVentanas
						Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
						If Msgbox("GESTION DE COBROS PODRIA TENER UN CLIENTE O DESTINATARIO CON UNA TAREA ¨PENDIENTE¨."+chr(13)+chr(13)+chr(13)+ InfoMostrar +chr(13)+chr(13)+"¿DESEA MODIFICAR ESTA TAREA?", 4 + 32, "Importante" ) = 6 then
							Cell = Sheet.getCellByPosition(8, y)
							vIdTarea = Cell.String
							goto Paso9
						End if
'					End If
				End IF
			End If		
		Else
			If vNroCliente = Cell.String then 
'				Cell = Sheet.getCellByPosition(14, y) 
'				If Left(Cell.String, 9) = "PENDIENTE" then
					Gosub InformacionVentanas
					Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
					If Msgbox("GESTION DE COBROS POSEE UN CLIENTE O DESTINATARIO CON UNA TAREA ¨PENDIENTE¨."+chr(13)+chr(13)+chr(13)+ InfoMostrar +chr(13)+chr(13)+"¿DESEA MODIFICAR ESTA TAREA?", 4 + 32, "Importante" ) = 6 then
						Cell = Sheet.getCellByPosition(8, y)
						vIdTarea = Cell.String
						goto Paso9
					End if
'				End if
			End If
		End if
	Next y

Paso6:
	'BUSCA POR NRO. DE CLIENTE EN LA BASE DE DATOS Y CARGA LA INFORMACION DEL MISMO EN LAS VARIABLES.
	vDireccion = ""
	vZona = ""
	vTarea = ""
	vCadena = ""
	vPrioridad = ""
	vInfo = ""
	vOtros = ""
	vEstado = ""
	vFechaApartir = date
	y = 0
	if vNroCliente = "0" then 
		dlgCT1.Model.Step = 2
		goto Paso7 
	End if
	Sheet = Doc.Sheets.getByName("BDClientes")
	For y = 2 to 12002 'Número máximo de filas a buscar en BDClientes.
		Cell = Sheet.getCellByPosition(0, y)
		if Cell.String = vNroCliente then 
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
			goto Paso7 
		End if
	Next y
	MsgBox "Nro. de Cliente no encontrado en la base de datos"
	Goto Paso4 


Paso7:
	' Carga las Variables en el dialogo
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	dlgCT1.Model.TextField2.Text = vIdTarea
	dlgCT1.Model.TextField3.Text = vNroCliente 
	dlgCT1.Model.TextField4.Text = vNombre
	dlgCT1.Model.TextField6.Text = vNombre
	dlgCT1.Model.TextField5.Text = vDireccion
	dlgCT1.Model.ComboBox1.text = vZona
	dlgCT1.Model.CheckBox1.State = 0
	dlgCT1.Model.CheckBox2.State = 1
	dlgCT1.Model.CheckBox3.State = 0
	dlgCT1.Model.CheckBox4.State = 0
	dlgCT1.Model.CheckBox5.State = 0
	dlgCT1.Model.OptionButton2.State = 1
	dlgCT1.Model.OptionButton4.State = 1
	dlgCT1.Model.CheckBox6.State = 1
	dlgCT1.Model.CheckBox7.State = 1
	dlgCT1.Model.CheckBox8.State = 1
	dlgCT1.Model.CheckBox9.State = 1
	dlgCT1.Model.CheckBox10.State = 1
	dlgCT1.Model.CheckBox11.State = 0
	dlgCT1.Model.TimeField1.text = ""
	dlgCT1.Model.TimeField2.text = ""
	dlgCT1.Model.TimeField3.text = ""
	dlgCT1.Model.TimeField4.text = ""
	dlgCT1.Model.TextField1.Text = vInfo
	dlgCT1.Model.DateField1.text = vFechaApartir

'	dlgCT1.Model.ComboBox3.text = vAsignado

Paso8:
	' Abre dialog1 para que se ingrese toda la información restante
	dlgCT1.Model.Step = 1
	If vNroCliente = "0" then 
		dlgCT1.Model.Step = 2
	End If
	Select Case dlgCT1.Execute()
	Case 1
		If vNroCliente = "0" then
			vNombre = dlgCT1.Model.TextField6.Text
			vDireccion = dlgCT1.Model.TextField7.Text
			vZona = dlgCT1.Model.ComboBox2.text
		End if
		if vZona = "" then
			Msgbox "No ha especificado la zona"
			goto Paso8
		End if
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

		vInfo = ""
		If dlgCT1.Model.CheckBox6.State = 1 then vInfo = "Lun"
		If dlgCT1.Model.CheckBox7.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Mar" 
		If dlgCT1.Model.CheckBox7.State = 1 and vInfo = "" then vInfo = "Mar" 
		If dlgCT1.Model.CheckBox8.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Mié" 
		If dlgCT1.Model.CheckBox8.State = 1 and vInfo = "" then vInfo = "Mié" 
		If dlgCT1.Model.CheckBox9.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Jue" 
		If dlgCT1.Model.CheckBox9.State = 1 and vInfo = "" then vInfo = "Jue" 
		If dlgCT1.Model.CheckBox10.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Vie" 
		If dlgCT1.Model.CheckBox10.State = 1 and vInfo = "" then vInfo = "Vie" 
		If dlgCT1.Model.CheckBox11.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Sáb" 
		If dlgCT1.Model.CheckBox11.State = 1 and vInfo = "" then vInfo = "Sáb" 
		If vInfo <> "" then vInfo = vInfo + "."
		If Instr(vInfo, "Lun") > 0 and Instr(vInfo, "Mar") > 0 and Instr(vInfo, "Mié") > 0 and Instr(vInfo, "Jue") > 0 and Instr(vInfo, "Vie") > 0 and Instr(vInfo, "Sáb") = 0 then vInfo = "L a V"
		If dlgCT1.Model.TimeField1.text <> "" and dlgCT1.Model.TimeField2.text <> "" then vInfo = vInfo + " de " + dlgCT1.Model.TimeField1.text + " a " + dlgCT1.Model.TimeField2.text
		If dlgCT1.Model.TimeField3.text <> "" and dlgCT1.Model.TimeField4.text <> "" then 
			If dlgCT1.Model.TimeField1.text <> "" and dlgCT1.Model.TimeField2.text <> "" then
				vInfo = vInfo + " y de " + dlgCT1.Model.TimeField3.text + " a " + dlgCT1.Model.TimeField4.text + "hs"
			Else
				vInfo = vInfo + " de " + dlgCT1.Model.TimeField3.text + " a " + dlgCT1.Model.TimeField4.text + "hs"
			End If
		Else
			If dlgCT1.Model.TimeField1.text <> "" and dlgCT1.Model.TimeField2.text <> "" then vInfo = vInfo + "hs"
		End If
		If vInfo <> "" then vInfo = "{" + vInfo + "}"
		If dlgCT1.Model.TextField1.Text <> "" and vInfo <> "" then vInfo = vInfo + " " + dlgCT1.Model.TextField1.Text
		If dlgCT1.Model.TextField1.Text <> "" and vInfo = "" then vInfo = dlgCT1.Model.TextField1.Text

		vFechaApartir = dlgCT1.Model.DateField1.text
'		vAsignado = dlgCT1.Model.ComboBox3.text
	Case 0
		dlgCT1.Dispose()
		Procesando = False
		Exit Sub
	End Select
	Cell = Sheet.getCellByPosition(1, ctFila)
	Cell.String = vNroCliente
	Cell = Sheet.getCellByPosition(2, ctFila)
	Cell.String = vNombre
	Cell = Sheet.getCellByPosition(3, ctFila)
	Cell.String = vDireccion
	Cell = Sheet.getCellByPosition(4, ctFila)
	Cell.String = vZona
	Cell = Sheet.getCellByPosition(5, ctFila)
	Cell.String = vTarea
	Cell = Sheet.getCellByPosition(6, ctFila)
	Cell.String = vPrioridad
	Cell = Sheet.getCellByPosition(7, ctFila)
	Cell.String = vInfo
	Cell = Sheet.getCellByPosition(8, ctFila)
	Cell.String = vOtros
	Cell = Sheet.getCellByPosition(9, ctFila)
	Cell.String = vEstado

	'COLUMNA CONTROL DE EGRESO
	Cell = Sheet.getCellByPosition(10, ctFila)
	Cell.String = ""
	Cell = Sheet.getCellByPosition(11, ctFila)
	Cell.String = "" 

	Cell = Sheet.getCellByPosition(12, ctFila)
	Cell.String = vFechaApartir
	Cell = Sheet.getCellByPosition(13, ctFila)
	Cell.String = DATE
	Cell = Sheet.getCellByPosition(14, ctFila)
	Cell.String = ""
'	If vEstado = "FINALIZADO" then Cell.String = DATE
	Cell = Sheet.getCellByPosition(15, ctFila)
	If Right(Cell.String,Len(vUsuario)) = vUsuario then
	
	Else
		If Cell.String = "" then
			Cell.String = vUsuario
		Else
			Cell.String = Cell.String + "/" + vUsuario
		End If
	End If
	
	'COLOREA EL FONDO DE LAS CELDAS
	If vEstado = "PENDIENTE" Then					
		z = 0
		For z = 1 to 15
			Cell = Sheet.getCellByPosition(z, ctFila)
			Cell.CellBackColor = RGB(238,238,238) 'PENDIENTE
		Next z
	End If

'	If vEstado = "EN CURSO" Then					
'		z = 0
'		For z = 1 to 14
'			Cell = Sheet.getCellByPosition(z, ctFila)
'			Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
'		Next z
'	End If

'	If vEstado = "FINALIZADO" Then					
'		z = 0
'		For z = 1 to 15
'			Cell = Sheet.getCellByPosition(z, ctFila)
'			Cell.CellBackColor = RGB(255,102,102) 'FINALIZADO
'		Next z
'	End If

	Cell = Sheet.getCellByPosition(8, ctFila)
	Cell.CellBackColor = RGB(255,255,255) 'COLOR COLUMNA OTROS
	Cell = Sheet.getCellByPosition(10, ctFila)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO
	Cell = Sheet.getCellByPosition(11, ctFila)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO
	
	'ACTUALIZA ESTA TAREA EN EXPEDICION-COBROS
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	If vEstado = "PENDIENTE" then
		nFila = 0
		For nFila = 5 to 200 'Número
			Cell = Sheet.getCellByPosition(8, nFila)
			If Cell.String = vIdTarea then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(9, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(10, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(11, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(12, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(13, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(14, nFila)
				Cell.String = vFechaApartir
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 8 to 14
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
'				If vEstado = "EN CURSO" then
'					z = 0
'					For z = 8 to 14
'						Cell = Sheet.getCellByPosition(z, nFila)
'						Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
'					Next z
'				End If			
				Goto FinActualizacion
			End If
		Next nFila
		nFila = 0
		For nFila = 5 to 200 'Número
			Cell = Sheet.getCellByPosition(8, nFila)
			If Cell.String = "" then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(9, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(10, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(11, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(12, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(13, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(14, nFila)
				Cell.String = vFechaApartir
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 8 to 14
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
'				If vEstado = "EN CURSO" then
'					z = 0
'					For z = 8 to 14
'						Cell = Sheet.getCellByPosition(z, nFila)
'						Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
'					Next z
'				End If			
				Goto FinActualizacion
			End If
		Next nFila		
	End If
'	If vEstado = "FINALIZADO" then
'		nFila = 0
'		For nFila = 5 to 200 'Número
'			Cell = Sheet.getCellByPosition(8, nFila)
'			If Cell.String = vIdTarea then
				'Borra y colorea el fondo
'				z = 0
'				For z = 8 to 14
'					Cell = Sheet.getCellByPosition(z, nFila)
'					Cell.String = ""
'					Cell.CellBackColor = RGB(255,255,255) 'VACIA
'				Next z
'				Goto FinActualizacion
'			End If
'		Next nFila
'	End If
FinActualizacion:

	dlgCT1.Dispose()
	dlgCT2.Dispose()
	If Msgbox( "¿Desea cargar otra tarea?", 4 + 32, "" ) = 6 then goto Inicio
	Doc.Store()
	Procesando = False
	Exit Sub
	
Paso9:
	'Modifica una tarea PENDIENTE o EN CURSO
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	ctFila = CInt(vIdTarea) + 10
'	vIdTarea = ""
	vDireccion = ""
	vZona = ""
	vTarea = ""
	vCadena = ""
	vPrioridad = ""
	vInfo = ""
	vEstado = ""
	vFechaApartir = ""
'	vAsignado = ""
'	Cell = Sheet.getCellByPosition(0, ctFila)
'	vIdTarea = Cell.getString
	Cell = Sheet.getCellByPosition(2, ctFila)
	vNombre = Cell.getString
	Cell = Sheet.getCellByPosition(3, ctFila)
	vDireccion = Cell.getString
	Cell = Sheet.getCellByPosition(4, ctFila)
	vZona = Cell.String
	Cell = Sheet.getCellByPosition(5, ctFila)
	vTarea = Cell.getString
	Cell = Sheet.getCellByPosition(6, ctFila)
	vPrioridad = Cell.getString
	Cell = Sheet.getCellByPosition(7, ctFila)
	vInfo = Cell.getString
	Cell = Sheet.getCellByPosition(8, ctFila)
	vOtros = Cell.getString
	Cell = Sheet.getCellByPosition(9, ctFila)
	vEstado = Cell.getString
	Cell = Sheet.getCellByPosition(12, ctFila)
	vFechaApartir = Cell.getString
'	Cell = Sheet.getCellByPosition(15, ctFila)
'	If Cell.getString = vUsuario then vUsuario = Cell.getString
'	If Cell.getString <> vUsuario then vUsuario = Cell.getString + "/" + vUsuario
'	
'
	' Carga el listbox de Asignado
'	dlgCT1 = CargarDialogo("Dialog1")'agregado
'	olstDatos = dlgCT1.getControl("ComboBox3")	
'	oHojaDatos = ThisComponent.getSheets.getByName("Datos")	
'	oRango = oHojaDatos.getCellRangeByName("C3:C7") 'ASIGNADO
'	data = oRango.getDataArray()'agregado
'	co1 = 0
'	Redim src(UBound(data))
'	For Each d In data
 '   	src(co1) = d(0)
  '    	co1 = co1 + 1
'	Next
'   	olstDatos.addItems(src, 0)

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

	dlgCT1.Model.CheckBox6.State = 0
	dlgCT1.Model.CheckBox7.State = 0
	dlgCT1.Model.CheckBox8.State = 0
	dlgCT1.Model.CheckBox9.State = 0
	dlgCT1.Model.CheckBox10.State = 0
	dlgCT1.Model.CheckBox11.State = 0
	If Instr(vInfo, "Lun") > 0 then	dlgCT1.Model.CheckBox6.State = 1
	If Instr(vInfo, "Mar") > 0 then	dlgCT1.Model.CheckBox7.State = 1
	If Instr(vInfo, "Mié") > 0 then	dlgCT1.Model.CheckBox8.State = 1
	If Instr(vInfo, "Jue") > 0 then	dlgCT1.Model.CheckBox9.State = 1
	If Instr(vInfo, "Vie") > 0 then	dlgCT1.Model.CheckBox10.State = 1
	If Instr(vInfo, "Sáb") > 0 then	dlgCT1.Model.CheckBox11.State = 1
	If Instr(vInfo, "L a V") > 0 then
		dlgCT1.Model.CheckBox6.State = 1
		dlgCT1.Model.CheckBox7.State = 1
		dlgCT1.Model.CheckBox8.State = 1
		dlgCT1.Model.CheckBox9.State = 1
		dlgCT1.Model.CheckBox10.State = 1
		dlgCT1.Model.CheckBox11.State = 0
	End If
	dlgCT1.Model.TimeField1.text = ""
	dlgCT1.Model.TimeField2.text = ""
	dlgCT1.Model.TimeField3.text = ""
	dlgCT1.Model.TimeField4.text = ""
	dlgCT1.Model.TextField1.Text = ""
	If Instr(vInfo, "{") > 0 and Instr(vInfo, "}") > 0 then
		If Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), "de ") > 0 then
			dlgCT1.Model.TimeField1.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), "de ")+3, 5)
			dlgCT1.Model.TimeField2.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), "de ")+3+5+3, 5)
		End If
		If Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), " y de ") > 0 then
			dlgCT1.Model.TimeField3.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), " y de ")+6, 5)
			dlgCT1.Model.TimeField4.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), " y de ")+6+5+3, 5)
		End If
		dlgCT1.Model.TextField1.Text = Mid(vInfo, Instr(vInfo, "}")+2, Len(vInfo))
	Else
		dlgCT1.Model.TextField1.Text = vInfo
	End If
'	dlgCT1.Model.ComboBox3.text = vAsignado
	dlgCT1.Model.DateField1.text = vFechaApartir
	Goto Paso8
Procesando = False
Exit sub
InformacionVentanas:
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	InfoMostrar = ""
	Cell = Sheet.getCellByPosition(x, y)
	InfoMostrar = "Id.Tarea: " + Cell.String
	x = x + 1
	Cell = Sheet.getCellByPosition(x, y)
	InfoMostrar = InfoMostrar + chr(13) + "Nro.Cliente: " + Cell.String
	x = x + 1
	Cell = Sheet.getCellByPosition(x, y)
	InfoMostrar = InfoMostrar + chr(13) + "Cliente/Destinatario: " + Cell.String
	x = x + 1
	Cell = Sheet.getCellByPosition(x, y)
	If Instr(Cell.String, "+") = 0 then
		InfoMostrar = InfoMostrar + chr(13) + "Zona: " + Mid (Cell.String, 1, Len(Cell.String)-2)
		InfoMostrar = InfoMostrar + chr(13) + "Tarea: " + Mid (Cell.String, Len(Cell.String), 1)
	Else 
		InfoMostrar = InfoMostrar + chr(13) + "Zona: " + Mid (Cell.String, 1, InStr (Cell.String, "+")-3)
		InfoMostrar = InfoMostrar + chr(13) + "Tarea: " + Mid (Cell.String, InStr (Cell.String, "+") - 1, Len(Cell.String) - InStr (Cell.String, "+")+2)
	End If
	x = x + 1
	Cell = Sheet.getCellByPosition(x, y)
	InfoMostrar = InfoMostrar + chr(13) + "Prioridad: " + Cell.String
	x = x + 1
	Cell = Sheet.getCellByPosition(x, y)
	InfoMostrar = InfoMostrar + chr(13) + "Comentarios:" + chr(13) + Cell.String
	x = x + 1
	Cell = Sheet.getCellByPosition(x, y)
	If x < 7 then
		If Left(Cell.String, 8) = "EN CURSO" then
			InfoMostrar = InfoMostrar + chr(13) + "Estado: EN CURSO"  +chr(13)+ "Asignado: " + Mid (Cell.String, 10, Len(Cell.String))
		Else
			InfoMostrar = InfoMostrar + chr(13) + "Estado: PENDIENTE"  +chr(13)+ "Asignado: " + Mid (Cell.String, 11, Len(Cell.String))
		End If
	End If
	If x > 8 then
		InfoMostrar = InfoMostrar + chr(13) + "A partir de: "  + Cell.String
	End If
Return
End Sub

Sub ModificarTarea
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	Dim vIdTarea as String
	Dim vNroCliente as String
	Dim vNombre as String
	Dim vDireccion as string
	Dim vZona as String
	Dim vTarea as String
	Dim vPrioridad as String
	Dim vInfo as String 
	Dim vEstado As String
'	Dim vAsignado As String
'	Dim vConcurrio As String
'	Dim vObjetivo As String
	Dim vFechaApartir As String
	Dim vFechaCarga As String
	Dim vFechaFinalizado as String
	
	Dim dlgCT7 as Object, dlgCT8 as Object

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
	
	'CORROBORA ACTUALIZACIONES EN CONTROL DE EGRESO
	If RutaOrigen = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(11, 2)
		RutaOrigen = ConvertToURL( Cell.String )
	End If
	If ModArchivoCE = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(12, 2)
		ModArchivoCE = Cell.String
	End If
	If cDate(Left(ModArchivoCE, 10)) <> cDate(Left(FileDateTime( RutaOrigen ), 10)) then
		If cDate(Left(ModArchivoCE, 10)) > cDate(Left(FileDateTime( RutaOrigen ), 10)) then
			
		Else
			Msgbox "EXPEDICIÓN a actualizado su información."+chr(13)+chr(13)+"Última Modificación C.E.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior C.E.: "+ModArchivoCE+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	else
		If cDate(Right(ModArchivoCE, 8)) < cDate(Right(FileDateTime( RutaOrigen ), 8)) then
			Msgbox "EXPEDICIÓN a actualizado su información."+chr(13)+chr(13)+"Última Modificación C.E.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior C.E.: "+ModArchivoCE+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	End If
Inicio:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	DialogLibraries.LoadLibrary("Standard")
	dlgCT7 = createUnoDialog(DialogLibraries.Standard.Dialog7)
	dlgCT8 = createUnoDialog(DialogLibraries.Standard.Dialog8)
Paso1:

Paso2:
	'Abre el Dialogo para que ingrese el Nro. de Id.Tarea
	dlgCT7.Model.TextField1.Text = ""	
	Select Case dlgCT7.Execute()
	Case 1
		vIdTarea = dlgCT7.Model.TextField1.Text
		goto Paso3
	Case 0
		dlgCT7.Dispose()
		Doc.Store()
		Procesando = False
		Exit Sub		
	End Select
Paso3:
	'Busca vIdTarea en Carga de Tareas
	xIDT = 0
	For xIDT = 11 to 10011 'Número máximo de Filas en Carga de Tareas
		Cell = Sheet.getCellByPosition(0, xIDT)
		If Cell.String = vIdTarea then goto Paso4
		If Cell.String = "" then
				Exit For
		End if 
	Next xIDT
	Msgbox "Id. Tarea no encontrado."
	goto Paso2
	Procesando = False	
	Exit Sub
Paso4:
	Posicionar = xIDT
	PosicionadorCelda
	FilaActual = xIDT
	FilaVisible

	' Verifica el Estado de la Tarea
	Cell = Sheet.getCellByPosition(9, xIDT)
	If Cell.String = "PENDIENTE" then
		Cell = Sheet.getCellByPosition(11, xIDT)
		If Left(Cell.String, 8) = "EN CURSO" then
			Msgbox "ESTA TAREA SE ENCUENTRA 'EN CURSO'."+chr(13)+"INFORMAR CUALQUIER CAMBIO QUE REALICE AL SECTOR DE EXPEDICION O A LA PERSONA QUE TIENE ASIGNADA LA TAREA.", 48,"Importante"
			goto Paso5
		End If
		If Left(Cell.String, 9) = "PENDIENTE" then
			Msgbox "ESTA TAREA SE ENCUENTRA 'PENDIENTE' EN EL SECTOR DE EXPEDICION."+chr(13)+"INFORMAR CUALQUIER CAMBIO QUE REALICE.", 48,"Importante"
			goto Paso5
		End If
		goto Paso5
	End if
	If Cell.String = "FINALIZADO" then
		If Msgbox( "ESTA TAREA SE ENCUENTRA 'FINALIZADA'."+chr(13)+"¿DESEA MODIFICARLA?"+chr(13)+"Id.Tarea: "+vIdTarea, 4 + 32, "Importante" ) = 6 then
			goto Paso5	
		End if
		FilaNoVisible
		goto Paso2
	End if
	Msgbox "ERROR:"+chr(13)+"El estado de la Tarea es desconocido."+chr(13)+"Favor de verificar."
	goto Paso2
Paso5:
	' Carga los valores de la Fila xIDT a las variables
	Cell = Sheet.getCellByPosition(1, xIDT)
	vNroCliente = Cell.getString
	Cell = Sheet.getCellByPosition(2, xIDT)
	vNombre = Cell.getString
	Cell = Sheet.getCellByPosition(3, xIDT)
	vDireccion = Cell.getString
	Cell = Sheet.getCellByPosition(4, xIDT)
	vZona = Cell.getString
	Cell = Sheet.getCellByPosition(5, xIDT)
	vTarea = Cell.getString
	Cell = Sheet.getCellByPosition(6, xIDT)
	vPrioridad = Cell.getString
	Cell = Sheet.getCellByPosition(7, xIDT)
	vInfo = Cell.getString
	Cell = Sheet.getCellByPosition(9, xIDT)
	vEstado = Cell.getString
'	Cell = Sheet.getCellByPosition(9, xIDT)
'	vAsignado = Cell.getString
'	Cell = Sheet.getCellByPosition(10, xIDT)
'	vConcurrio = Cell.getString
'	Cell = Sheet.getCellByPosition(11, xIDT)
'	vObjetivo = Cell.getString
	Cell = Sheet.getCellByPosition(12, xIDT)
	vFechaApartir = Cell.getString
	Cell = Sheet.getCellByPosition(13, xIDT)
	vFechaCarga = Cell.getString
	Cell = Sheet.getCellByPosition(14, xIDT)
	vFechaFinalizado = Cell.getString
	
	' Carga el listbox de Asignado a Dialog8
'	dlgCT8 = CargarDialogo("Dialog8")
'	olstDatos = dlgCT8.getControl("ComboBox2")	
'	oHojaDatos = ThisComponent.getSheets.getByName("Datos")	
'	oRango = oHojaDatos.getCellRangeByName("C3:C7") 'ASIGNADO
'	data = oRango.getDataArray()
'	co1 = 0
'	Redim src(UBound(data))
'	For Each d In data
'    	src(co1) = d(0)
'      	co1 = co1 + 1
'	Next
'  	olstDatos.addItems(src, 0)

	' Carga las Variables en Dialog8
	dlgCT8.Model.TextField3.Text = vIdTarea
	dlgCT8.Model.TextField5.Text = vNroCliente 
	dlgCT8.Model.TextField1.Text = vNombre
	dlgCT8.Model.TextField7.Text = vDireccion
	dlgCT8.Model.ComboBox1.text = vZona
	y = 0
'	dlgCT8.Model.CheckBox1.State = 0
	dlgCT8.Model.CheckBox2.State = 0
	dlgCT8.Model.CheckBox3.State = 0
	dlgCT8.Model.CheckBox4.State = 0
'	dlgCT8.Model.CheckBox5.State = 0
	Pos1 = Len(vTarea)
	For y = 1 to Pos1 Step 2
		CadBuscar1 = ""
		CadBuscar1 = Mid(vTarea, y, 1)
'		IF CadBuscar1 = "E" THEN dlgCT8.Model.CheckBox1.State = 1
		IF CadBuscar1 = "C" THEN dlgCT8.Model.CheckBox2.State = 1
		IF CadBuscar1 = "D" THEN dlgCT8.Model.CheckBox3.State = 1
		IF CadBuscar1 = "O" THEN dlgCT8.Model.CheckBox4.State = 1
'		IF CadBuscar1 = "V" THEN dlgCT8.Model.CheckBox5.State = 1
	Next	
	If vPrioridad = "ALTA" then dlgCT8.Model.OptionButton1.State = 1
	If vPrioridad = "MEDIA" then dlgCT8.Model.OptionButton2.State = 1
	If vPrioridad = "BAJA" then dlgCT8.Model.OptionButton3.State = 1	
	If vEstado = "PENDIENTE" then dlgCT8.Model.OptionButton4.State = 1
'	If vEstado = "EN CURSO" then dlgCT8.Model.OptionButton5.State = 1
	If vEstado = "FINALIZADO" then dlgCT8.Model.OptionButton6.State = 1
	
	dlgCT8.Model.CheckBox6.State = 0
	dlgCT8.Model.CheckBox7.State = 0
	dlgCT8.Model.CheckBox8.State = 0
	dlgCT8.Model.CheckBox9.State = 0
	dlgCT8.Model.CheckBox10.State = 0
	dlgCT8.Model.CheckBox11.State = 0
	If Instr(vInfo, "Lun") > 0 then	dlgCT8.Model.CheckBox6.State = 1
	If Instr(vInfo, "Mar") > 0 then	dlgCT8.Model.CheckBox7.State = 1
	If Instr(vInfo, "Mié") > 0 then	dlgCT8.Model.CheckBox8.State = 1
	If Instr(vInfo, "Jue") > 0 then	dlgCT8.Model.CheckBox9.State = 1
	If Instr(vInfo, "Vie") > 0 then	dlgCT8.Model.CheckBox10.State = 1
	If Instr(vInfo, "Sáb") > 0 then	dlgCT8.Model.CheckBox11.State = 1
	If Instr(vInfo, "L a V") > 0 then
		dlgCT8.Model.CheckBox6.State = 1
		dlgCT8.Model.CheckBox7.State = 1
		dlgCT8.Model.CheckBox8.State = 1
		dlgCT8.Model.CheckBox9.State = 1
		dlgCT8.Model.CheckBox10.State = 1
		dlgCT8.Model.CheckBox11.State = 0
	End If
	dlgCT8.Model.TimeField1.text = ""
	dlgCT8.Model.TimeField2.text = ""
	dlgCT8.Model.TimeField3.text = ""
	dlgCT8.Model.TimeField4.text = ""
	dlgCT8.Model.TextField4.Text = ""
	If Instr(vInfo, "{") > 0 and Instr(vInfo, "}") > 0 then
		If Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), "de ") > 0 then
			dlgCT8.Model.TimeField1.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), "de ")+3, 5)
			dlgCT8.Model.TimeField2.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), "de ")+3+5+3, 5)
		End If
		If Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), " y de ") > 0 then
			dlgCT8.Model.TimeField3.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), " y de ")+6, 5)
			dlgCT8.Model.TimeField4.text = Mid(vInfo, Instr(Mid(vInfo, Instr(vInfo, "{"), Instr(vInfo, "}")), " y de ")+6+5+3, 5)
		End If
		dlgCT8.Model.TextField4.Text = Mid(vInfo, Instr(vInfo, "}")+2, Len(vInfo))
	Else
		dlgCT8.Model.TextField4.Text = vInfo
	End If

'	dlgCT8.Model.ComboBox2.text = vAsignado
	dlgCT8.Model.DateField1.text = vFechaApartir
Paso6:
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
			Msgbox "No ha especificado la zona"
			goto Paso5
		End if
		vTarea = ""
'		if dlgCT8.Model.CheckBox1.State = 1 then vTarea = vTarea + "E"
		if dlgCT8.Model.CheckBox2.State = 1 then vTarea = vTarea + "C"
		if dlgCT8.Model.CheckBox3.State = 1 then vTarea = vTarea + "D"
		if dlgCT8.Model.CheckBox4.State = 1 then vTarea = vTarea + "O"
'		if dlgCT8.Model.CheckBox5.State = 1 then vTarea = vTarea + "V"
		CadBuscar1 = vTarea
		if vTarea = "" then
			Msgbox "No ha especificado cual es la tarea a realizar."
			goto Paso5
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
'		if dlgCT8.Model.OptionButton5.State = 1 then vEstado = "EN CURSO"
		if dlgCT8.Model.OptionButton6.State = 1 then vEstado = "FINALIZADO"
		
		vInfo = ""
		If dlgCT8.Model.CheckBox6.State = 1 then vInfo = "Lun"
		If dlgCT8.Model.CheckBox7.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Mar" 
		If dlgCT8.Model.CheckBox7.State = 1 and vInfo = "" then vInfo = "Mar" 
		If dlgCT8.Model.CheckBox8.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Mié" 
		If dlgCT8.Model.CheckBox8.State = 1 and vInfo = "" then vInfo = "Mié" 
		If dlgCT8.Model.CheckBox9.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Jue" 
		If dlgCT8.Model.CheckBox9.State = 1 and vInfo = "" then vInfo = "Jue" 
		If dlgCT8.Model.CheckBox10.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Vie" 
		If dlgCT8.Model.CheckBox10.State = 1 and vInfo = "" then vInfo = "Vie" 
		If dlgCT8.Model.CheckBox11.State = 1 and vInfo <> "" then vInfo = vInfo + "/" + "Sáb" 
		If dlgCT8.Model.CheckBox11.State = 1 and vInfo = "" then vInfo = "Sáb" 
		If vInfo <> "" then vInfo = vInfo + "."
		If Instr(vInfo, "Lun") > 0 and Instr(vInfo, "Mar") > 0 and Instr(vInfo, "Mié") > 0 and Instr(vInfo, "Jue") > 0 and Instr(vInfo, "Vie") > 0 and Instr(vInfo, "Sáb") = 0 then vInfo = "L a V"
		If dlgCT8.Model.TimeField1.text <> "" and dlgCT8.Model.TimeField2.text <> "" then vInfo = vInfo + " de " + dlgCT8.Model.TimeField1.text + " a " + dlgCT8.Model.TimeField2.text
		If dlgCT8.Model.TimeField3.text <> "" and dlgCT8.Model.TimeField4.text <> "" then 
			If dlgCT8.Model.TimeField1.text <> "" and dlgCT8.Model.TimeField2.text <> "" then
				vInfo = vInfo + " y de " + dlgCT8.Model.TimeField3.text + " a " + dlgCT8.Model.TimeField4.text + "hs"
			Else
				vInfo = vInfo + " de " + dlgCT8.Model.TimeField3.text + " a " + dlgCT8.Model.TimeField4.text + "hs"
			End If
		Else
			If dlgCT8.Model.TimeField1.text <> "" and dlgCT8.Model.TimeField2.text <> "" then vInfo = vInfo + "hs"
		End If
		If vInfo <> "" then vInfo = "{" + vInfo + "}"
		If dlgCT8.Model.TextField4.Text <> "" and vInfo <> "" then vInfo = vInfo + " " + dlgCT8.Model.TextField4.Text
		If dlgCT8.Model.TextField4.Text <> "" and vInfo = "" then vInfo = dlgCT8.Model.TextField4.Text

'		vAsignado = dlgCT8.Model.ComboBox2.text
'		if vAsignado = "" then
'			Msgbox "La Tarea no ha sido ASIGNADA a una persona."
'			goto Paso5
'		End if
		If dlgCT8.Model.DateField1.text <> "" then
			vFechaApartir = dlgCT8.Model.DateField1.text
		End If
		If dlgCT8.Model.DateField1.text = "" then
			vFechaApartir = Date
		End If
	Case 0
		dlgCT8.Dispose()
		Procesando = False
		Exit Sub
	End Select
Paso7:
	'Ingresa los datos en la planilla de calculo
'	Cell = Sheet.getCellByPosition(1, xIDT)
'	Cell.String = vNroCliente
'	Cell = Sheet.getCellByPosition(2, xIDT)
'	Cell.String = vNombre
	Cell = Sheet.getCellByPosition(3, xIDT)
	Cell.String = vDireccion
	Cell = Sheet.getCellByPosition(4, xIDT)
	Cell.String = vZona
	Cell = Sheet.getCellByPosition(5, xIDT)
	Cell.String = vTarea
	Cell = Sheet.getCellByPosition(6, xIDT)
	Cell.String = vPrioridad
	Cell = Sheet.getCellByPosition(7, xIDT)
	Cell.String = vInfo
	Cell = Sheet.getCellByPosition(9, xIDT)
	Cell.String = vEstado
'	Cell = Sheet.getCellByPosition(9, xIDT)
'	Cell.String = vAsignado
	Cell = Sheet.getCellByPosition(10, xIDT)
	Cell.String = "" 
	Cell = Sheet.getCellByPosition(11, xIDT)
	Cell.String = "" 
	Cell = Sheet.getCellByPosition(12, xIDT)
	Cell.String = vFechaApartir
	Cell = Sheet.getCellByPosition(13, xIDT)
	Cell.String = vFechaCarga
	Cell = Sheet.getCellByPosition(14, xIDT)
	Cell.String = ""
	If vEstado = "FINALIZADO" then
		Cell.String = Date
	End If
	Cell = Sheet.getCellByPosition(15, xIDT)
	If Right(Cell.String,Len(vUsuario)) = vUsuario then
	
	Else
		If Cell.String = "" then
			Cell.String = vUsuario
		Else
			Cell.String = Cell.String + "/" + vUsuario
		End If
	End If
'	Cell = Sheet.getCellByPosition(16, xIDT)
'	If Cell.String = "" then  Cell.Value = 0
'	Cell = Sheet.getCellByPosition(17, xIDT)
'	If Cell.String = "" then  Cell.Value = 0

	'COLOREA EL FONDO DE LAS CELDAS
	z = 0
	If vEstado = "PENDIENTE" Then					
		z = 0
		For z = 1 to 15
			Cell = Sheet.getCellByPosition(z, xIDT)
			Cell.CellBackColor = RGB(238,238,238) 'PENDIENTE
		Next z
	End If

	If vEstado = "FINALIZADO" Then					
		z = 0
		For z = 1 to 15
			Cell = Sheet.getCellByPosition(z, xIDT)
			Cell.CellBackColor = RGB(255,102,102) 'FINALIZADO
		Next z
		FilaNoVisible
	End If
	If vEstado <> "FINALIZADO" then
		Cell = Sheet.getCellByPosition(8, xIDT)
		Cell.CellBackColor = RGB(255,255,255) 'COLOR COLUMNA OTROS
	End If
	Cell = Sheet.getCellByPosition(10, xIDT)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO
	Cell = Sheet.getCellByPosition(11, xIDT)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO

	'ACTUALIZA ESTA TAREA EN EXPEDICION-COBROS
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	If vEstado = "PENDIENTE" then
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(8, nFila)
			If Cell.String = vIdTarea then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(9, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(10, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(11, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(12, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(13, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(14, nFila)
				Cell.String = vFechaApartir
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 8 to 14
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
'				If vEstado = "EN CURSO" then
'					z = 0
'					For z = 8 to 14
'						Cell = Sheet.getCellByPosition(z, nFila)
'						Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
'					Next z
'				End If			
				Goto FinActualizacion
			End If
		Next nFila
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(8, nFila)
			If Cell.String = "" then
				Cell.String = vIdTarea
				Cell = Sheet.getCellByPosition(9, nFila)
				Cell.String = vNroCliente
				Cell = Sheet.getCellByPosition(10, nFila)
				Cell.String = vNombre + chr(13) + vDireccion
				Cell = Sheet.getCellByPosition(11, nFila)
				Cell.String = vZona + chr(13) + vTarea
				Cell = Sheet.getCellByPosition(12, nFila)
				Cell.String = vPrioridad
				Cell = Sheet.getCellByPosition(13, nFila)
				Cell.String = vInfo				
				Cell = Sheet.getCellByPosition(14, nFila)
				Cell.String = vFechaApartir
				'Colorea fondo cuando es PENDIENTE
				If vEstado = "PENDIENTE" then
					z = 0
					For z = 8 to 14
						Cell = Sheet.getCellByPosition(z, nFila)
						Cell.CellBackColor = RGB(255,255,255) 'PENDIENTE
					Next z
				End If
				'Colorea fondo cuando es EN CURSO
'				If vEstado = "EN CURSO" then
'					z = 0
'					For z = 8 to 14
'						Cell = Sheet.getCellByPosition(z, nFila)
'						Cell.CellBackColor = RGB(102,255,102) 'EN CURSO
'					Next z
'				End If			
				Goto FinActualizacion
			End If
		Next nFila		
	End If
	If vEstado = "FINALIZADO" then
		nFila = 0
		For nFila = 5 to 255 'Número
			Cell = Sheet.getCellByPosition(8, nFila)
			If Cell.String = vIdTarea then
				'Borra y colorea el fondo
				z = 0
				For z = 8 to 14
					Cell = Sheet.getCellByPosition(z, nFila)
					Cell.String = ""
					Cell.CellBackColor = RGB(255,255,255) 'VACIA
				Next z
				Exit For
			End If
		Next nFila
	End If
FinActualizacion:

	dlgCT7.Dispose()
	dlgCT8.Dispose()
	If Msgbox( "¿Desea modificar otra tarea?", 4 + 32, "" ) = 6 then goto Inicio
	Doc.Store()
	Procesando = False
End Sub

Sub EliminarTareas
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	Dim eIdTarea as String
	Dim mIdTarea as String, mNroCliente as String, mNombre as String, mDireccion as string, mZona as String, mTarea as String, mPrioridad as String, mInfo as String, mOtros as String, mEstado As String, mFechaApartir As String, mFechaCarga As String, mFechaFinalizado as String, mControlo as String
	Dim mCEIdTarea as String, mCEEstadoAsignado
	
	Dim gcFila

	Dim dlgCT7 as Object
	
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
	
	'CORROBORA ACTUALIZACIONES EN CONTROL DE EGRESO
	If RutaOrigen = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(11, 2)
		RutaOrigen = ConvertToURL( Cell.String )
	End If
	If ModArchivoCE = "" then
		Sheet = Doc.Sheets.getByName("Datos")
		Cell = Sheet.getCellByPosition(12, 2)
		ModArchivoCE = Cell.String
	End If
	If cDate(Left(ModArchivoCE, 10)) <> cDate(Left(FileDateTime( RutaOrigen ), 10)) then
		If cDate(Left(ModArchivoCE, 10)) > cDate(Left(FileDateTime( RutaOrigen ), 10)) then
			
		Else
			Msgbox "EXPEDICIÓN a actualizado su información."+chr(13)+chr(13)+"Última Modificación C.E.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior C.E.: "+ModArchivoCE+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	else
		If cDate(Right(ModArchivoCE, 8)) < cDate(Right(FileDateTime( RutaOrigen ), 8)) then
			Msgbox "EXPEDICIÓN a actualizado su información."+chr(13)+chr(13)+"Última Modificación C.E.: "+FileDateTime( RutaOrigen )+"hs."+chr(13)+"Actualización Anterior C.E.: "+ModArchivoCE+"hs.", 48,"Importante - Nueva Actualización Disponible"
		End If
	End If
	
Inicio:
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	DialogLibraries.LoadLibrary("Standard")
	dlgCT7 = createUnoDialog(DialogLibraries.Standard.Dialog7)
Paso1:
 
Paso2:
	'Abre el Dialogo para que ingrese el Nro. de Id.Tarea
	dlgCT7.Model.TextField1.Text = ""	
	Select Case dlgCT7.Execute()
	Case 1
		eIdTarea = dlgCT7.Model.TextField1.Text
'		dlgCT7.Dispose()
		goto Paso3
	Case 0
		dlgCT7.Dispose()
		Procesando = False
		Exit Sub		
	End Select
Paso3:
	'Busca vIdTarea en Carga de Tareas
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	xIDT = 0
	For xIDT = 11 to 10011 'Número máximo de Filas en Carga de Tareas
		Cell = Sheet.getCellByPosition(0, xIDT)
		If Cell.String = eIdTarea then goto Paso4
		If Cell.String = "" then
			Msgbox "Id. Tarea no encontrado."
			Exit For
		End if 
	Next xIDT
	Msgbox "Id. Tarea no encontrado."
	goto Paso2
	Procesando = False	
	Exit Sub
Paso4:
	Posicionar = xIDT
	PosicionadorCelda
	FilaActual = xIDT
	FilaVisible

	' Verifica el Estado de la Tarea
	If Msgbox( "Se dispone a ELIMINAR una Tarea."+chr(13)+"¿Esta seguro que desea eliminar la Tarea Nro. "+eIdTarea+" ?", 4 + 32, "IMPORTANTE" ) = 6 then
		Cell = Sheet.getCellByPosition(11, xIDT)
		If Left(Cell.String, 8) = "EN CURSO" then
			Msgbox	"EL CLIENTE POSEE UNA TAREA ¨EN CURSO¨ EN EXPEDICIÓN."+chr(13)+"AL FINALIZAR INFORMAR AL SECTOR DE EXPEDICIÓN O A LA PERSONA ASIGNADA LOS CAMBIOS REALIZADOS.", 48,"Importante"
		End If
		If Left(Cell.String, 9) = "PENDIENTE" then
			Msgbox	"EL CLIENTE POSEE UNA TAREA ¨PENDIENTE¨ EN EXPEDICIÓN."+chr(13)+"AL FINALIZAR FAVOR DE ACTUALIZAR O GUARDAR LOS CAMBIOS.", 48,"Importante"
		End If
		goto Paso5	
	End if
	Procesando = False
	Exit Sub
Paso5:
	' Borra y colorea las celdas de la Fila xIDT
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	z = 0
	For z = 1 to 15
		Cell = Sheet.getCellByPosition(z, xIDT)
		Cell.String = ""
		Cell.CellBackColor = RGB(238,238,238) 'PENDIENTE
	Next z
	Cell = Sheet.getCellByPosition(8, xIDT)
	Cell.CellBackColor = RGB(255,255,255) 'COLOR COLUMNA OTROS
	Cell = Sheet.getCellByPosition(10, xIDT)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO
	Cell = Sheet.getCellByPosition(11, xIDT)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO

	'Elimina la Tarea GC en Expedición Vs Cobros
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	gcFila = 0
	For gcFila = 5 to 255 'Número
		Cell = Sheet.getCellByPosition(8, gcFila)
		If Cell.String = eIdTarea then
			'Borra y colorea el fondo
			z = 0
			For z = 8 to 14
				Cell = Sheet.getCellByPosition(z, gcFila)
				Cell.String = ""
				Cell.CellBackColor = RGB(255,255,255) 'VACIA
			Next z
			Exit For
		End If
	Next gcFila
		
Paso6:
	' Busca la Fila en estado PENDIENTE para mover a la fila eliminada.
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	mIDT = 0
	For mIDT = 10011 to 11 Step -1
		If xIDT = mIDT Then exit sub
		Cell = Sheet.getCellByPosition(1, mIDT)
		If Cell.String <> "" then
			Cell = Sheet.getCellByPosition(9, mIDT)
			If Cell.String = "PENDIENTE" then
				goto Paso7
			End IF
			dlgCT7.Dispose()
			If Msgbox( "¿Desea eliminar otra tarea?", 4 + 32, "Eliminar" ) = 6 then goto Inicio
			If Msgbox( "¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información?", 4 + 32, "Guardar" ) = 6 then 
				Doc.Store()
			End If
			Procesando = False
			Exit Sub
		End If
	Next mIDT
'	Doc.Store()
	Procesando = False
	Exit Sub
Paso7:	
	' Carga en las variables los valores de la fila a mover.
	Cell = Sheet.getCellByPosition(0, mIDT)
	mIdTarea = Cell.getString
	Cell = Sheet.getCellByPosition(1, mIDT)
	mNroCliente = Cell.getString
	Cell = Sheet.getCellByPosition(2, mIDT)
	mNombre = Cell.getString
	Cell = Sheet.getCellByPosition(3, mIDT)
	mDireccion = Cell.getString
	Cell = Sheet.getCellByPosition(4, mIDT)
	mZona = Cell.getString
	Cell = Sheet.getCellByPosition(5, mIDT)
	mTarea = Cell.getString
	Cell = Sheet.getCellByPosition(6, mIDT)
	mPrioridad = Cell.getString
	Cell = Sheet.getCellByPosition(7, mIDT)
	mInfo = Cell.getString
	Cell = Sheet.getCellByPosition(8, mIDT)
	mOtros = Cell.getString
	Cell = Sheet.getCellByPosition(9, mIDT)
	mEstado = Cell.getString
	Cell = Sheet.getCellByPosition(10, mIDT)
	mCEIdTarea = Cell.getString
	Cell = Sheet.getCellByPosition(11, mIDT)
	mCEEstadoAsignado = Cell.getString
	Cell = Sheet.getCellByPosition(12, mIDT)
	mFechaApartir = Cell.getString
	Cell = Sheet.getCellByPosition(13, mIDT)
	mFechaCarga = Cell.getString
	Cell = Sheet.getCellByPosition(14, mIDT)
	mFechaFinalizado = Cell.getString
	Cell = Sheet.getCellByPosition(15, mIDT)
	mControlo = Cell.getString
		
	' Carga la información en la fila eliminada
	Cell = Sheet.getCellByPosition(1, xIDT)
	Cell.String = mNroCliente
	Cell = Sheet.getCellByPosition(2, xIDT)
	Cell.String = mNombre
	Cell = Sheet.getCellByPosition(3, xIDT)
	Cell.String = mDireccion
	Cell = Sheet.getCellByPosition(4, xIDT)
	Cell.String = mZona
	Cell = Sheet.getCellByPosition(5, xIDT)
	Cell.String = mTarea
	Cell = Sheet.getCellByPosition(6, xIDT)
	Cell.String = mPrioridad
	Cell = Sheet.getCellByPosition(7, xIDT)
	Cell.String = mInfo
	Cell = Sheet.getCellByPosition(8, xIDT)
	Cell.String = mOtros
	Cell = Sheet.getCellByPosition(9, xIDT)
	Cell.String = mEstado
	Cell = Sheet.getCellByPosition(10, xIDT)
	Cell.String = mCEIdTarea
	Cell = Sheet.getCellByPosition(11, xIDT)
	Cell.String = mCEEstadoAsignado
	Cell = Sheet.getCellByPosition(12, xIDT)
	Cell.String = mFechaApartir
	Cell = Sheet.getCellByPosition(13, xIDT)
	Cell.String = mFechaCarga
	Cell = Sheet.getCellByPosition(14, xIDT)
	Cell.String = mFechaFinalizado
	Cell = Sheet.getCellByPosition(15, xIDT)
	Cell.String = mControlo

	'Borra y colorea el fondo de las celdas de la Fila mIDT
	z = 0
	For z = 1 to 15
		Cell = Sheet.getCellByPosition(z, mIDT)
		Cell.String = ""
		Cell.CellBackColor = RGB(238,238,238) 'PENDIENTE
	Next z
	Cell = Sheet.getCellByPosition(8, mIDT)
	Cell.CellBackColor = RGB(255,255,255) 'COLOR COLUMNA OTROS
	Cell = Sheet.getCellByPosition(10, mIDT)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO
	Cell = Sheet.getCellByPosition(11, mIDT)
	Cell.CellBackColor = RGB(207,231,245) 'COLOR COLUMNA CONTROL DE EGRESO
	
	'Elimina una Tarea en Expedición Vs Cobros
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	gcFila = 0
	For gcFila = 5 to 255 'Número
		Cell = Sheet.getCellByPosition(8, gcFila)
		If Cell.String = mIdTarea then
			Cell.String = eIdTarea
			Exit For
		End If
	Next gcFila
	 
	dlgCT7.Dispose()
	If Msgbox( "¿Desea eliminar otra tarea?", 4 + 32, "" ) = 6 then goto Inicio
	If Msgbox( "¿Desea guardar los últimos cambios realizados para que otros usuarios puedan ver esta información?", 4 + 32, "Guardar" ) = 6 then 
		Doc.Store()
	End If
	Procesando = False
End Sub

Function CargarDialogo( Nombre As String ) As Object
	'Cargamos la librería Standard en memoria
	DialogLibraries.LoadLibrary( "Standard" )
	'Cargamos el cuadro de diálogo en memoria
	CargarDialogo = CreateUnoDialog( DialogLibraries.Standard.getByName( Nombre ) )
End Function

Sub BuscarTareasCT
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR ESTA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
Dim vObservacion as String
Dim vOtros as String
Dim Encontrado
Dim Encontrado2
Dim yIDT

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
'	Sheet = Doc.Sheets.getByName("Datos")
'	olstDatos = dlgCT9.getControl("lstCBox2")
'  	For d = 1 to 10	
'	 	Cell = Sheet.getCellByPosition(2, d) 	
'	  	vAsignado = Cell.String
'	  	If vAsignado <> "" then olstDatos.addItem( vAsignado, -1 )
'	Next d
Paso2:
	'Abre el Dialogo para que ingrese la información a buscar
	Sheet = Doc.Sheets.getByName("Carga de Tareas") 
	dlgCT9.Model.TextField1.Text = ""
	dlgCT9.Model.TextField2.Text = ""
	dlgCT9.Model.lstCBox1.Text = "" 
	dlgCT9.Model.lstCBox2.Text = "TODOS" 
	vAsignado = "TODOS"
	Select Case dlgCT9.Execute()
	Case 1
		vNroCliente = dlgCT9.Model.TextField1.Text
		vObservacion = dlgCT9.Model.TextField2.Text		
		vOtros = dlgCT9.Model.TextField3.Text
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
		vAsignado = "TODOS"
		vZona = ""
		if dlgCT9.Model.CheckBox1.State = 1 then vZona = vZona + "N"
		if dlgCT9.Model.CheckBox2.State = 1 then vZona = vZona + "C"
		if dlgCT9.Model.CheckBox3.State = 1 then vZona = vZona + "S"
		if dlgCT9.Model.CheckBox4.State = 1 then vZona = vZona + "I"
		vTarea = ""
		'if dlgCT9.Model.CheckBox5.State = 1 then vTarea = vTarea + "E"
		if dlgCT9.Model.CheckBox6.State = 1 then vTarea = vTarea + "C"
		if dlgCT9.Model.CheckBox7.State = 1 then vTarea = vTarea + "D"
		if dlgCT9.Model.CheckBox8.State = 1 then vTarea = vTarea + "O"
		'if dlgCT9.Model.CheckBox9.State = 1 then vTarea = vTarea + "V"
		vEstado = ""
		if dlgCT9.Model.CheckBox10.State = 1 then vEstado = vEstado + "P"
		'if dlgCT9.Model.CheckBox11.State = 1 then vEstado = vEstado + "E"
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
	For yIDT = 11 to 10000 'Número máximo de Filas en Carga de Tareas
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

		If vOtros <> "" then
			Cell = Sheet.getCellByPosition(8, yIDT)
			Encontrado2 = 0
			Encontrado2 = Instr(Cell.String,vOtros)
			If Encontrado2 = 0 then goto Siguiente
		End If
		
		If vNroCliente <> "" then
			Cell = Sheet.getCellByPosition(1, yIDT)
			If vNroCliente <> Cell.String then goto Siguiente
		End If
		If vEstado <> "" then
			Cell = Sheet.getCellByPosition(9, yIDT)
			If Cell.String = "PENDIENTE" and dlgCT9.Model.CheckBox10.State <> 1 then goto Siguiente 
			'If Cell.String = "EN CURSO" and dlgCT9.Model.CheckBox11.State <> 1 then goto Siguiente 
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
'			Pos1 = 0
'			CadBuscar1 = "E"
'			CadBuscar2 = Cell.getString
'			Pos1 = InStr (CadBuscar2, CadBuscar1)
'			If Pos1 > 0 and dlgCT9.Model.CheckBox5.State <> 1 then goto Siguiente 
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
'			Pos1 = 0
'			CadBuscar1 = "V"
'			Pos1 = InStr (CadBuscar2, CadBuscar1)
'			If Pos1 > 0 and dlgCT9.Model.CheckBox9.State <> 1 then goto Siguiente 
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
