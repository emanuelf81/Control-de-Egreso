REM  *****  BASIC  *****


Option Explicit
'HOJA DE CALCULO
Dim Doc As Object
Dim Hoja As Object
Dim Celda As Object

Sub CorrecionesDDExcelReporteMGR2
	Dim y, x, y2
	
	Doc = thiscomponent
	If vUsuario <> "EMANUEL" then Exit sub
	MSGBOX "ESPERE A QUE EL SISTEMA ANUNCIE QUE HA FINALIZADO"
	
	'Corrige la Columna Código, quita el ´
	Hoja = Doc.Sheets.getByName("ReporteMGR2")
  	y = 0
  	y2 = 0
  	For y = 0 to 100	
	 	Celda = Hoja.getCellByPosition(0, y) 	
	  	If Celda.String = "Código" then
	  		For y2 = y+1 to 20000
	  			Celda = Hoja.getCellByPosition(0, y2)
	  			If Celda.String <> "" then
		  			Celda.Value = Celda.String
				Else 			
	  				Exit For 
	  			End If
	  		Next y2
	  		Exit For
		End IF
	Next Y
	'Corrige la Columna Zona, quita primer espacio en zona
	Hoja = Doc.Sheets.getByName("ReporteMGR2")
  	y = 0
  	y2 = 0
  	For y = 0 to 100	
	 	Celda = Hoja.getCellByPosition(3, y) 	
	  	If Celda.String = "Zona" then
	  		For y2 = y+1 to 20000
	  			Celda = Hoja.getCellByPosition(3, y2)
	  			If Celda.String <> "" then
	  				If Left(Celda.String,1) = " " then
		  				Celda.String = Right(Celda.String,Len(Celda.String)-1)
		  			End If
	  				If Celda.String = "Interior" then
		  				Celda.String = "INT"
		  			End If
				Else 			
	  				Exit For 
	  			End If
	  		Next y2
	  		Exit For
		End IF
	Next Y

	Doc.store
	MSGBOX "FINALIZADO"
End Sub

Sub ActBDClientesDDReporteMGR2
	Dim y, x, y2
	Dim bCodigo
	Dim SigFila
	Dim oRango As Object
	Dim dReporteMGR2
	Dim dBDClientes
	Dim oBarraEstado As Object
	Dim vProgBar As Integer		
	
	Doc = thiscomponent
	If vUsuario <> "EMANUEL" then Exit sub
	MSGBOX "ESPERE A QUE EL SISTEMA ANUNCIE QUE HA FINALIZADO"
	oBarraEstado = ThisComponent.getCurrentController.StatusIndicator


	'CARGA MATRIZ REPORTEMGR2
    Hoja = Doc.Sheets.getByName("ReporteMGR2")
    oRango = Hoja.getCellRangeByName("A7:N20007")
    dReporteMGR2 = oRango.getDataArray()

	'TRANSFIERE REPORTEMGR2 A TEMPORAL
	Hoja = Doc.Sheets.getByName("Temporal")
    oRango = Hoja.getCellRangeByName("A1:N20001")
    oRango.setDataArray( dReporteMGR2 )	
	
	'CARGA NUEVAMENTE LA MATRIZ
	Hoja = Doc.Sheets.getByName("Temporal")
    oRango = Hoja.getCellRangeByName("A1:O20001")
    dReporteMGR2 = oRango.getDataArray()

	'BORRA TODAS LAS CELDAS VACIAS DE LA MATRIZ TEMPORAL DE REPORTEMGR2
	For y = 0 to 19999
		If dReporteMGR2 (y) (0) = "" then
			If dReporteMGR2 (y+1) (0) = "" then
				oRango.getRows.removeByIndex(y,20000-y)
				dReporteMGR2 = oRango.getDataArray()
				Exit For
			End If
		End if
	Next y

	'ACTUALIZA/TRANSFIERE MGR2 A BDCLIENTES
	oBarraEstado.start( "Actualizando BDClientes ", uBound(dReporteMGR2) )
	Hoja = Doc.Sheets.getByName("BDClientes")
  	For y = 10 to 20009
  		oBarraEstado.setValue( y )
  		bCodigo = ""
	 	Celda = Hoja.getCellByPosition(0, y)
	 	If Celda.String = "" then exit for 	
	  	bCodigo = Celda.Value
	  	For y2 = 0 to uBound(dReporteMGR2)	  	
			If dReporteMGR2 (y2) (0) = bCodigo then
				Celda = Hoja.getCellByPosition(1, y)
				Celda.String = dReporteMGR2 (y2) (1)	'nombre
				Celda = Hoja.getCellByPosition(2, y)
				If Celda.String = "" then
					Celda = Hoja.getCellByPosition(2, y)
					Celda.String = dReporteMGR2 (y2) (5)	'domicilio
					Celda = Hoja.getCellByPosition(3, y)
					Celda.String = dReporteMGR2 (y2) (6)	'barrio
					Celda = Hoja.getCellByPosition(4, y)
					Celda.String = dReporteMGR2 (y2) (7)	'localidad
					Celda = Hoja.getCellByPosition(6, y)
					Celda.String = dReporteMGR2 (y2) (8)	'provincia
				End If	
				Celda = Hoja.getCellByPosition(5, y)
				If Celda.String = "" or Celda.String = "Sin Definir" then
					Celda.String = dReporteMGR2 (y2) (3)	'zona
				End If
								
				Celda = Hoja.getCellByPosition(7, y)
				Celda.String = dReporteMGR2 (y2) (10)	'pago
				Celda = Hoja.getCellByPosition(8, y)
				Celda.String = dReporteMGR2 (y2) (2)	'cuit
				Celda = Hoja.getCellByPosition(9, y)
				Celda.String = dReporteMGR2 (y2) (3)	'zona
				Celda = Hoja.getCellByPosition(10, y)
				Celda.String = dReporteMGR2 (y2) (4)	'estado
				Celda = Hoja.getCellByPosition(11, y)
				Celda.String = dReporteMGR2 (y2) (5)	'domicilio
				Celda = Hoja.getCellByPosition(12, y)
				Celda.String = dReporteMGR2 (y2) (6)	'barrio
				Celda = Hoja.getCellByPosition(13, y)
				Celda.String = dReporteMGR2 (y2) (7)	'localidad				
				Celda = Hoja.getCellByPosition(14, y)
				Celda.String = dReporteMGR2 (y2) (8)	'provincia	
				Celda = Hoja.getCellByPosition(15, y)
				Celda.String = dReporteMGR2 (y2) (13)	'categoria
				dReporteMGR2 (y2) (14) = "X"
				Exit For
			End If
  		Next y2
	Next y
	'BUSCA SIGUIENTE FILA VACIA
'	Hoja = Doc.Sheets.getByName("BDClientes")
'  	SigFila = 0
'  	For y = 10 to 20009
'		Celda = Hoja.getCellByPosition(0, y)
'	  	If Celda.String = "" then
'	  		SigFila = y
'	  		Exit For
'	  	End If
 ' 	Next y
  	'AGREGA LOS NUEVOS CLIENTES A BDCLIENTES
'  	y = SigFila
 ' 	For y2 = 0 to uBound(dReporteMGR2)	  	
'		If dReporteMGR2 (y2) (14) <> "X" then
'			bCodigo = dReporteMGR2 (y2) (0)
'			
'			'Celda = Hoja.getCellByPosition(0, bCodigo+9)
'			'Celda.Value = dReporteMGR2 (y2) (0)	'Código
'			
'			Celda = Hoja.getCellByPosition(1, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (1)	'Nombre
'			Celda = Hoja.getCellByPosition(2, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (5)	'domicilio
'			Celda = Hoja.getCellByPosition(3, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (6)	'barrio
'			Celda = Hoja.getCellByPosition(4, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (7)	'localidad				
'			Celda = Hoja.getCellByPosition(5, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (3)	'zona
'			Celda = Hoja.getCellByPosition(6, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (8)	'provincia	
'			Celda = Hoja.getCellByPosition(7, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (10)	'pago
'			Celda = Hoja.getCellByPosition(8, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (2)	'cuit
'			Celda = Hoja.getCellByPosition(9, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (3)	'zona
'			Celda = Hoja.getCellByPosition(10, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (4)	'estado
'			Celda = Hoja.getCellByPosition(11, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (5)	'domicilio
'			Celda = Hoja.getCellByPosition(12, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (6)	'barrio
'			Celda = Hoja.getCellByPosition(13, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (7)	'localidad				
'			Celda = Hoja.getCellByPosition(14, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (8)	'provincia	
'			Celda = Hoja.getCellByPosition(15, bCodigo+9)
'			Celda.String = dReporteMGR2 (y2) (13)	'categoria
'		End If
'	Next y2
	
	'BORRA TEMPORAL
	Hoja = Doc.Sheets.getByName("Temporal")
    oRango = Hoja.getCellRangeByName("A1:N20001")
	oRango.clearContents( 1023 )

	'CORRIGE NOMBRE DEL VENDEDOR EN ESTADO
	Hoja = Doc.Sheets.getByName("BDClientes")
	For y = 10 to 20009
		Celda = Hoja.getCellByPosition(10, y)
		If Celda.String = "Ventas NORTE" then Celda.String = "MATIAS"
		If Celda.String = "Ventas SUR" then Celda.String = "ESTEBAN"
		If Celda.String = "Ventas CENTRO" then Celda.String = "FEDERICO F."
	Next y

'	'ORDENA POR NOMBRE ASCENDENTE
'	Dim mCamposOrden(0) As New com.sun.star.table.TableSortField
'	Dim mDescriptorOrden()	
'	
'	Hoja = Doc.Sheets.getByName("BDClientes") 'La hoja donde esta el rango a ordenar
'	oRango = Hoja.getCellRangeByName("A10:P20010") 'El rango a ordenar
'	mDescriptorOrden = oRango.createSortDescriptor() 'Descriptor de ordenamiento, o sea, el "como"
'	mCamposOrden(0).Field = 1 'Los campos empiezan en 0
'	mCamposOrden(0).IsAscending = True
'	mCamposOrden(0).IsCaseSensitive = False 'Sensible a MAYUSCULAS/minusculas
'	mCamposOrden(0).FieldType = com.sun.star.table.TableSortFieldType.AUTOMATIC 'Tipo de campo AUTOMATICO
'	mDescriptorOrden(1).Name = "ContainsHeader" 'Indicamos si el rango contiene títulos de campos
'	mDescriptorOrden(1).Value = True
'	mDescriptorOrden(3).Name = "SortFields" 'La matriz de campos a ordenar
'	mDescriptorOrden(3).Value = mCamposOrden
' 	oRango.sort( mDescriptorOrden )	'Ordenamos con los parámetros establecidos
		
	oBarraEstado.end()
	Doc.store
	MSGBOX "FINALIZADO"
	Exit Sub
	
End Sub
