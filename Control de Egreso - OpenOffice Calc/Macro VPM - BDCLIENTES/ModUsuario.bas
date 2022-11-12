REM  *****  BASIC  *****
Option Explicit

'HOJA DE CALCULO
Dim Doc As Object
Dim Sheet As Object
Dim Cell As Object

'dim document   as object
'dim dispatcher as object


'FORMULARIO Y USUARIO
Global vUsuario as String
Dim tUsuario as String
Dim pwUsuario as String
Dim oFormulario As Object
Dim olstUsuario As Object
Dim olstUsuarioVista As Object
Dim otxtPW As Object



Sub ListadoUsuario
	Dim y
	
	Doc = thiscomponent
	Sheet = Doc.Sheets.getByName("Usuario")
    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
	olstUsuario = oFormulario.getByName("lstUsuario")
	olstUsuarioVista = ThisComponent.getCurrentController.getControl( olstUsuario )
	olstUsuarioVista.removeItems( 0, olstUsuarioVista.getItemCount )

	Sheet = Doc.Sheets.getByName("BDClientes")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("BDOtrosDestinatarios")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Datos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Temporal")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("ReporteMGR2")
	Sheet.IsVisible = False


 	'Carga ListBox lstUsuario con los Nombres que hay en la hoja Datos
	Sheet = Doc.Sheets.getByName("Datos")
  	Y = 0
  	For Y = 1 to 10	
	 	Cell = Sheet.getCellByPosition(7, Y) 	
	  	vUsuario = Cell.String
	  	If vUsuario <> "" then olstUsuarioVista.addItem( vUsuario, -1 )
	Next Y
	Sheet = Doc.Sheets.getByName("Usuario")
End Sub

Sub InhabilitarHojas
	Doc = thiscomponent
	Sheet = Doc.Sheets.getByName("Usuario")
    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
	otxtPW = oFormulario.getByName("txtPW")
	otxtPW.Text = ""
	Sheet = Doc.Sheets.getByName("BDClientes")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("BDOtrosDestinatarios")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Datos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Temporal")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("ReporteMGR2")
	Sheet.IsVisible = False
	
End Sub

Sub RecibirFocoTXT
	Doc = thiscomponent
	Sheet = Doc.Sheets.getByName("Usuario")
    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
	otxtPW = oFormulario.getByName("txtPW")
	otxtPW.Text = ""
	Sheet = Doc.Sheets.getByName("BDClientes")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("BDOtrosDestinatarios")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Datos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Temporal")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("ReporteMGR2")
	Sheet.IsVisible = False

End Sub

Sub Ingresar
	Dim Y

	Doc = thiscomponent
	Sheet = Doc.Sheets.getByName("Usuario")
    oFormulario = Doc.getCurrentController.getActiveSheet.getDrawPage.getForms.getByName( "Formulario" ) 
	olstUsuario = oFormulario.getByName("lstUsuario")
	olstUsuarioVista = ThisComponent.getCurrentController.getControl( olstUsuario )
	otxtPW = oFormulario.getByName("txtPW")
	vUsuario = ""
	tUsuario = ""
	pwUsuario = ""
	pwUsuario = otxtPW.Text
	vUsuario =  olstUsuarioVista.getSelectedItem()
	Cell = Sheet.getCellByPosition(25, 0)
	Cell.String = vUsuario
	If vUsuario = "" then Exit sub
	Sheet = Doc.Sheets.getByName("Datos")
  	Y = 0
  	For Y = 1 to 10	
	 	Cell = Sheet.getCellByPosition(7, Y) 	
	  	If vUsuario = Cell.String Then
		 	Cell = Sheet.getCellByPosition(8, Y) 	
		  	If Cell.String = "ADM" Then
		  		tUsuario = "ADMINISTRADOR"
		  		Exit for
		  	Else
		  		tUsuario = "COLABORADOR"
		  		Exit for
		  	End If
	  	End If
	Next Y		

	If tUsuario = "ADMINISTRADOR" and pwUsuario = "VPM" then
		Sheet = Doc.Sheets.getByName("ReporteMGR2")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Datos")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Temporal")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("BDOtrosDestinatarios")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("BDClientes")
		Sheet.IsVisible = True
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
		Exit Sub
	End If
	
	If tUsuario = "COLABORADOR" then
		If vUsuario = "MARTIN" and pwUsuario = "MG" then goto Colaborador
		If vUsuario = "GUILLERMO" and pwUsuario = "A" then goto Colaborador
		If vUsuario = "LORENA" and pwUsuario = "ADM" then goto Colaborador
		If vUsuario = "PAOLA" and pwUsuario = "ADM" then goto Colaborador
		If vUsuario = "ANDREA" and pwUsuario = "ADM" then goto Colaborador
	End If
	Exit Sub
Colaborador:
	
End Sub


