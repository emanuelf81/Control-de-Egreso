REM  *****  BASIC  *****

Option Explicit
	
	Dim j
	Dim itTarea as Integer
	Dim olstDatos As Object
	Private dlgCT5 as Object	
	
Sub BuscarTareaCargaControlEgreso
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR UNA NUEVA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
   	Dim vAsignado as String, vVacio as String, vZona as String, vEstado as String, vPrioridad as String, vApartir as Date
   	Dim NroCliente as String, Nombre as String, Direccion as String, Zona as String, Tarea as String, Prioridad as String
   	Dim Observacion as String, Estado as String, Asignado as String, Apartir as String, Bultos as String

	Dim f
	Dim b
	
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "ToPoint"
	args3(0).Value = "$B$2"
	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())

	DialogLibraries.LoadLibrary("Standard")
	
	Doc = thiscomponent
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator
	vAsignado = ""
	vVacio = ""
	vZona = ""
	vEstado = ""
	vPrioridad = ""
	vApartir = ""
	itTarea = 1
   	NroCliente = ""
   	Nombre = ""
   	Direccion = ""
   	Zona = ""
   	Tarea = ""
   	Prioridad = ""
   	Observacion = ""
   	Estado = ""
   	Asignado = ""
   	Apartir = ""
   	Bultos = ""

	'Carga ComboBox2 = Asignado con los nombres
	dlgCT5 = createUnoDialog(DialogLibraries.Standard.Dialog5)
	olstDatos = dlgCT5.getControl("ComboBox2")	
	oHojaDatos = ThisComponent.getSheets.getByName("Datos")	
	oRango = oHojaDatos.getCellRangeByName("C2:C10") 'ASIGNADO
	data = oRango.getDataArray()'agregado
	co1 = 0
	Redim src(UBound(data))
	For Each d In data
    	src(co1) = d(0)
      	co1 = co1 + 1
	Next
   	olstDatos.addItems(src, 0)

	dlgCT5.Model.DateField1.text = date
	dlgCT5.Model.OptionButton1.State = 1
	dlgCT5.Model.OptionButton7.State = 1
	dlgCT5.Model.ComboBox1.text = "TODOS"
	dlgCT5.Model.ComboBox2.text = "TODOS"

	Select Case dlgCT5.Execute()
	Case 1
		'Asigna a las variables los valores a buscar
		vZona = dlgCT5.Model.ComboBox1.text
		If vZona = "" then vZona = "TODOS"
		vAsignado = dlgCT5.Model.ComboBox2.text
		If vAsignado = "" then vAsignado = "TODOS"
		vApartir = dlgCT5.Model.DateField1.text
		if dlgCT5.Model.OptionButton1.State = 1 then vPrioridad = "TODOS"	
		if dlgCT5.Model.OptionButton2.State = 1 then vPrioridad = "ALTA"	
		if dlgCT5.Model.OptionButton3.State = 1 then vPrioridad = "MEDIA"
		if dlgCT5.Model.OptionButton4.State = 1 then vPrioridad = "BAJA"	
		if dlgCT5.Model.OptionButton5.State = 1 then vEstado = "TODOS"	
		if dlgCT5.Model.OptionButton6.State = 1 then vEstado = "PENDIENTE"
		if dlgCT5.Model.OptionButton7.State = 1 then vEstado = "EN CURSO"
		if dlgCT5.Model.OptionButton8.State = 1 then vEstado = "FINALIZADO"
		Goto Buscar

	Case 0
		dlgCT5.Dispose()
		Procesando = False
		Exit Sub
	End Select		


Buscar:
	oBarraEstado.start( "Buscando... ", 120 )
	oBarraEstado.setValue( 5 )
	Procesando = False
	LimpiarHojadeRuta
	Procesando = True
	
	oBarraEstado.setValue( 10 )
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	Cell = Sheet.getCellByPosition(9, 9)
	Cell.String = vAsignado
	Cell = Sheet.getCellByPosition(9, 10)
	Cell.String = vZona
	Cell = Sheet.getCellByPosition(9, 11)
	Cell.String = vEstado
	Cell = Sheet.getCellByPosition(9, 53)
	Cell.String = vAsignado
	Cell = Sheet.getCellByPosition(9, 54)
	Cell.String = vZona
	Cell = Sheet.getCellByPosition(9, 55)
	Cell.String = vEstado
	


	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	j = 14
	For f = 10 to 10000 'f es el número de tareas a buscar en carga de tareas
		
		if vAsignado = "TODOS" then goto Zona 
		Cell = Sheet.getCellByPosition(9, f)
		if vAsignado = Cell.getString then goto Estado
		goto Asignado

Zona:
		if vZona = "TODOS" then goto Estado 
		Cell = Sheet.getCellByPosition(4, f)
		if vZona = Cell.getString then goto Estado
		goto Asignado

Estado:
		if vEstado = "TODOS" then goto Prioridad 
		Cell = Sheet.getCellByPosition(8, f)
		if vEstado = Cell.getString then goto Prioridad
		goto Asignado

Prioridad:
		if vPrioridad = "TODOS" then goto Apartir 
		Cell = Sheet.getCellByPosition(6, f)
		if vPrioridad = Cell.getString then goto Apartir
		goto Asignado

Apartir:
		if vApartir = "TODOS" then goto Nrotarea 
		Cell = Sheet.getCellByPosition(12, f)
		if vApartir >= CDate(Cell.getString) then goto Nrotarea
		goto Asignado
	
Nrotarea:
		Cell = Sheet.getCellByPosition(1, f)
		vVacio = Cell.getString
		if vVacio = "" then
			Cell = Sheet.getCellByPosition(0, f)
			if Cell.String = "" then
				Sheet = Doc.Sheets.getByName("Hoja de Ruta")
				exit for
			else
				Cell = Sheet.getCellByPosition(1, f + 1)
				if Cell.String = "" then
					Sheet = Doc.Sheets.getByName("Hoja de Ruta")
					exit for
				Else
					Cell = Sheet.getCellByPosition(0, f)
					Msgbox "Nro.Cliente o Destinatario desconocido."+chr(13)+chr(13)+"Corroborar la Id.Tarea Nº "+Cell.String,16,"IMPORTANTE"
					Sheet = Doc.Sheets.getByName("Hoja de Ruta")
					Procesando = False
					oBarraEstado.end()
					exit sub
				End If	
			End If
		end if
		Cell = Sheet.getCellByPosition(0, f)
		itTarea = Cell.getValue
		Cell = Sheet.getCellByPosition(1, f)
		NroCliente = Cell.getString
		Cell = Sheet.getCellByPosition(2, f)
		Nombre = Cell.getString
		Cell = Sheet.getCellByPosition(3, f)
		Direccion = Cell.getString
		Cell = Sheet.getCellByPosition(4, f)
		Zona = Cell.getString
		Cell = Sheet.getCellByPosition(5, f)
		Tarea = Cell.getString
		Cell = Sheet.getCellByPosition(6, f)
		Prioridad = Cell.getString
		Cell = Sheet.getCellByPosition(7, f)
		Observacion = Cell.getString
		Cell = Sheet.getCellByPosition(8, f)
		Estado = Cell.getString
		Cell = Sheet.getCellByPosition(9, f)
		Asignado = Cell.getString
		Cell = Sheet.getCellByPosition(12, f)
		Apartir = Cell.getString
		Cell = Sheet.getCellByPosition(10, f)
		Bultos = Cell.getString
		
		
		Sheet = Doc.Sheets.getByName("Hoja de Ruta")
		j = j + 2
		
		oBarraEstado.setValue( j )
		
		if j = 52 then j = j + 8
		if j = 96 then 
			MsgBox "Demasiadas Tareas"
			exit for
		end if
		Sheet = Doc.Sheets.getByName("Hoja de Ruta")
		Cell = Sheet.getCellByPosition(2, j)
		Cell.Value = itTarea
		j = j - 1
		Cell = Sheet.getCellByPosition(2, j)
		Cell.String = NroCliente
		Cell = Sheet.getCellByPosition(3, j)
		Cell.String = Nombre
		Cell = Sheet.getCellByPosition(4, j)
		Cell.String = Tarea
		Cell = Sheet.getCellByPosition(5, j)
		Cell.String = Prioridad
		Cell = Sheet.getCellByPosition(6, j)
		Cell.String = Observacion
		Cell = Sheet.getCellByPosition(7, j)
		Cell.String = "P"
		Cell = Sheet.getCellByPosition(8, j)
		Cell.String = "Si / No"
		Cell = Sheet.getCellByPosition(9, j)
		Cell.String = ""
		Cell = Sheet.getCellByPosition(10, j)
		Cell.String = Bultos
		
		j = j + 1
		Cell = Sheet.getCellByPosition(3, j)
		Cell.String = Direccion
		Cell = Sheet.getCellByPosition(4, j)
		Cell.String = Zona
		Cell = Sheet.getCellByPosition(5, j)
		Cell.String = Apartir
		Cell = Sheet.getCellByPosition(8, j)
		Cell.String = "Si / No"
		Cell = Sheet.getCellByPosition(9, j)
		Cell.String = ""
		
		Sheet = Doc.Sheets.getByName("Carga de Tareas")
Asignado:	
	Next f
	oBarraEstado.setValue( 100 )
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	Cell = Sheet.getCellByPosition(1, 9)
	Cell.STRING = DATE
	Cell = Sheet.getCellByPosition(1, 10)
	Cell.STRING = TIME
	
	If j<=50 then 
		Cell = Sheet.getCellByPosition(1, 11)
		Cell.String = "Pág.1/1"
		Cell = Sheet.getCellByPosition(1, 55)
		Cell.String = "Pág."
	end if
	If j>50 then 
		Cell = Sheet.getCellByPosition(1, 11)
		Cell.String = "Pág.1/2"
		Cell = Sheet.getCellByPosition(1, 55)
		Cell.String = "Pág.2/2"
	end if
	Procesando = False
	oBarraEstado.setValue( 120 )
	oBarraEstado.end()
End Sub

Sub LimpiarHojadeRuta
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR UNA NUEVA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	
	Dim Flags As Long
	
	Doc = thiscomponent

	'BORRA EL RANGO DE CELDAS
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	CellRange = Sheet.getCellRangeByName("C16:G51")
	Flags = com.sun.star.sheet.CellFlags.STRING
	Flags = Flags + com.sun.star.sheet.CellFlags.VALUE
	CellRange.clearContents(Flags)
	CellRange.CellBackColor = RGB(255,255,255) 'BLANCO
	Wait 500
	CellRange = Sheet.getCellRangeByName("K16:K51")
	Flags = com.sun.star.sheet.CellFlags.STRING
	Flags = Flags + com.sun.star.sheet.CellFlags.VALUE
	CellRange.clearContents(Flags)
	CellRange.CellBackColor = RGB(255,255,255) 'BLANCO
	Wait 500
	CellRange = Sheet.getCellRangeByName("C60:G95")
	Flags = com.sun.star.sheet.CellFlags.STRING
	Flags = Flags + com.sun.star.sheet.CellFlags.VALUE
	CellRange.clearContents(Flags)
	CellRange.CellBackColor = RGB(255,255,255) 'BLANCO
	Wait 500	
	CellRange = Sheet.getCellRangeByName("K60:K95")
	Flags = com.sun.star.sheet.CellFlags.STRING
	Flags = Flags + com.sun.star.sheet.CellFlags.VALUE
	CellRange.clearContents(Flags)
	CellRange.CellBackColor = RGB(255,255,255) 'BLANCO
	Wait 500
	
	Cell = Sheet.getCellByPosition(1, 11)
	Cell.String = "Pág."
	Cell = Sheet.getCellByPosition(1, 55)
	Cell.String = "Pág."

	Cell = Sheet.getCellByPosition(9, 9)
	Cell.String = ""
	Cell = Sheet.getCellByPosition(9, 10)
	Cell.String = ""
	Cell = Sheet.getCellByPosition(9, 11)
	Cell.String = ""
	Cell = Sheet.getCellByPosition(9, 53)
	Cell.String = ""
	Cell = Sheet.getCellByPosition(9, 54)
	Cell.String = ""
	Cell = Sheet.getCellByPosition(9, 55)
	Cell.String = ""
	Procesando = False
End Sub

Sub Imprimir
	'VERIFICA SI ESTA CORRIENDO OTRA MACRO ANTES DE EJECUTAR UNA NUEVA
	IF Procesando = True then 
		Msgbox "Cuidado."+chr(13)+"Se están realizando otras tareas."+chr(13)+"Espere a que el sistema le informe que ha finalizado.",16,"IMPORTANTE"
		Exit sub
	Else 
		Procesando = True
	End if
	
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	Doc = thiscomponent
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	
	Cell = Sheet.getCellByPosition(1, 10)
	Cell.STRING = TIME


	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "ToPoint"
	args1(0).Value = "$B$10:$K$96"

	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())

	dispatcher.executeDispatch(document, ".uno:Print", "", 0, Array())

	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "ToPoint"
	args3(0).Value = "$B$2"

	dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args3())
	Doc.Store()
	HoraUltGuardar = Timer

	Procesando = False
End Sub
