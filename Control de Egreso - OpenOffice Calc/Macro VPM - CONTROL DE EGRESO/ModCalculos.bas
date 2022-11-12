REM  *****  BASIC  *****


Option Explicit

Sub ActualizarCalculos

	Dim vMesBuscar As String
	Dim vAnoBuscar As String
	Dim vRealizo As String
	Dim vDia As String
	Dim vMes As String
	Dim vConcurrio As String
	Dim vObjetivo As String
	
	
	Dim cCol, ctFila, ctCol
	Dim tTarea, tDia, tMes
	

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

Inicio:
Paso1:
	Sheet = Doc.Sheets.getByName("Calculos")
	vMesBuscar = ""
	Cell = Sheet.getCellByPosition(1, 7)
	If Cell.String = "ENERO" then vMesBuscar = "01"
	If Cell.String = "FEBRERO" then vMesBuscar = "02"
	If Cell.String = "MARZO" then vMesBuscar = "03"
	If Cell.String = "ABRIL" then vMesBuscar = "04"
	If Cell.String = "MAYO" then vMesBuscar = "05"
	If Cell.String = "JUNIO" then vMesBuscar = "06"
	If Cell.String = "JULIO" then vMesBuscar = "07"
	If Cell.String = "AGOSTO" then vMesBuscar = "08"
	If Cell.String = "SETIEMBRE" then vMesBuscar = "09"
	If Cell.String = "OCTUBRE" then vMesBuscar = "10"
	If Cell.String = "NOVIEMBRE" then vMesBuscar = "11"
	If Cell.String = "DICIEMBRE" then vMesBuscar = "12"
	If Cell.String = "TODOS" then vMesBuscar = "00"
	If Cell.String = "" then Exit sub

	vAnoBuscar = ""
	Cell = Sheet.getCellByPosition(8, 7)
	vAnoBuscar = Cell.String
	If Cell.String = "TODOS" then vAnoBuscar = "0000"
	If Cell.String = "" then Exit Sub

	vRealizo = ""
	Cell = Sheet.getCellByPosition(15, 7)
	vRealizo = Cell.String
	If Cell.String = "TODOS" then vRealizo = "0"
	If Cell.String = "" then Exit Sub
	
	Cell = Sheet.getCellByPosition(0, 8)
	If vMesBuscar = "00" then
		Cell.String = "Mes: "
	Else
		Cell.String = "Día: "
	End If
	For cCol = 1 to 33
		Cell = Sheet.getCellByPosition(cCol, 8)
		Cell.String = ""
		If vMesBuscar = "00" then
			If cCol <= 12 then
				Cell.String = cCol
			End If
			If cCol = 32 then Cell.String = "Total"
			If cCol = 33 then Cell.String = "Promedio Mensual"
		Else
			If cCol <= 31 then
				Cell.String = cCol
			End If
			If cCol = 32 then Cell.String = "Total"
			If cCol = 33 then Cell.String = "Promedio Diario"
		End If
		Cell = Sheet.getCellByPosition(cCol, 9)
		Cell.Value = 0
		Cell = Sheet.getCellByPosition(cCol, 10)
		Cell.Value = 0
		Cell = Sheet.getCellByPosition(cCol, 11)
		Cell.Value = 0
	Next cCol

	x = 0
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	For ctFila = 11 to 10011
		For ctCol = 18 to 21 
			Cell = Sheet.getCellByPosition(ctCol, ctFila)
			If Cell.String = "" then Exit For 
			If Right(Cell.String,Len(vRealizo)) = vRealizo or vRealizo = "0" then
				If Mid(Cell.String,4,2) = vMesBuscar or vMesBuscar = "00" then
					If Mid(Cell.String,7,4) = vAnoBuscar then
						vDia = Mid(Cell.String,1,2)
						vMes = Mid(Cell.String,4,2)
						vConcurrio = Mid(Cell.String,12,2)
						vObjetivo = Mid(Cell.String,15,2)
						
'						If Msgbox( vRealizo+chr(13)+vDia+chr(13)+vMesBuscar+chr(13)+vAnoBuscar+chr(13)+vConcurrio+vObjetivo, 4 + 32, "" ) = 6 then Exit Sub
						Gosub Paso2
					End If		
				End If		
			End If			
		Next ctCol
	Next ctFila
	'CALCULA TOTALES
	Sheet = Doc.Sheets.getByName("Calculos")
	For x = 1 to 3
		tTarea = 0
		tDia = 0
		For cCol = 1 to 31
			Cell = Sheet.getCellByPosition(cCol, 8+x)
			If Cell.Value > 0 then
				tTarea = tTarea + Cell.Value
				tDia = tDia + 1 
			End If
		Next cCol
		Cell = Sheet.getCellByPosition(32, 8+x)
		Cell.Value = tTarea
		Cell = Sheet.getCellByPosition(33, 8+x)
		If tDia <> 0 then
			Cell.Value = tTarea / tDia
		End If
	Next x
	
	Msgbox "FINALIZADO"

'	If Msgbox( Cell.String, 4 + 32, "" ) = 6 then Exit Sub

	Exit Sub

Paso2:
	Sheet = Doc.Sheets.getByName("Calculos")
	If vMesBuscar = "00" then
		Cell = Sheet.getCellByPosition(CInt(vMes), 9)
		Cell.Value = Cell.Value + 1
		If vConcurrio = "SI" then
			Cell = Sheet.getCellByPosition(CInt(vMes), 10)
			Cell.Value = Cell.Value + 1
		End If
		If vObjetivo = "SI" then
			Cell = Sheet.getCellByPosition(CInt(vMes), 11)
			Cell.Value = Cell.Value + 1
		End If
	Else
		Cell = Sheet.getCellByPosition(CInt(vDia), 9)
		Cell.Value = Cell.Value + 1
		If vConcurrio = "SI" then
			Cell = Sheet.getCellByPosition(CInt(vDia), 10)
			Cell.Value = Cell.Value + 1
		End If
		If vObjetivo = "SI" then
			Cell = Sheet.getCellByPosition(CInt(vDia), 11)
			Cell.Value = Cell.Value + 1
		End If
	End If
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
Return

End Sub

Sub GraficarCalculos
' Dim oHojaActiva As Object
' Dim oGraficos As Object
' Dim mRangos(0)
' Dim sNombre As String
' Dim oRec As New com.sun.star.awt.Rectangle
' Dim oDir As New com.sun.star.table.CellRangeAddress
' 
'     'Acceso a la hoja activa
'     oHojaActiva = ThisComponent.getCurrentController.getActiveSheet() 
'     'El nombre de nuestro gráfico
'     sNombre = "Grafico01"
'     'El tamaño y la posición del nuevo gráfico, todas las medidas
'     'en centésimas de milímetro
'     With oRec
'         .X = 200            'Distancia desde la izquierda de la hoja
'         .Y = 10000            'Distancia desde la parte superior
'         .Width = 30000        'El ancho del gráfico
'         .Height = 13000        'El alto del gráfico
'     End With
'     'La dirección del rango de datos para el gráfico
'     With oDir
'         .Sheet = oHojaActiva.getRangeAddress.Sheet
'         .StartColumn = 0
'         .EndColumn = 31
'         .StartRow = 8
'         .EndRow = 11
'     End With
'     'Es una matriz de rangos, pues se pueden establecer más de uno
'     mRangos(0) = oDir
'     'Accedemos al conjunto de todos los gráficos de la hoja
'     oGraficos = oHojaActiva.getCharts()
'     'Verificamos que no exista el nombre
'     If oGraficos.hasByName( sNombre ) Then
'         MsgBox "Ya existe este nombre de gráfico, escoge otro"
'     Else
'         'Si no existe lo agregamos
'         oGraficos.addNewByName(sNombre, oRec, mRangos, True, True)
'     End If
'
     
End Sub

