REM  *****  BASIC  *****

Option Explicit

'HOJA DE CALCULO
Dim Doc As Object
Dim Sheet As Object
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

'Dim Cell As Object

Sub CorroboraUsuario
	Dim Doc As Object
	Dim Sheet As Object	
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
End Sub

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
	Sheet = Doc.Sheets.getByName("Datos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Calculos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
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
	Sheet = Doc.Sheets.getByName("Datos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Calculos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
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
	Sheet = Doc.Sheets.getByName("Datos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Calculos")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	Sheet.IsVisible = False
End Sub

Sub Ingresar
	Dim Y

	Doc = thiscomponent
	Sheet = Doc.Sheets.getByName("Usuario")
'	Cell = Sheet.getCellByPosition(0, 20)
'	Cell.String = vUsuario 
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
		Sheet = Doc.Sheets.getByName("BDClientes")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Datos")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Calculos")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Carga de Tareas")
		Sheet.IsVisible = True
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
		Goto Backup
		Exit Sub
	End If
	
	If tUsuario = "COLABORADOR" then
		If vUsuario = "MARTIN" and pwUsuario = "A" then goto Colaborador
		If vUsuario = "GUILLERMO" and pwUsuario = "A" then goto Colaborador
		If vUsuario = "ROXANA" and pwUsuario = "ADM" then goto Colaborador
		If vUsuario = "PAOLA" and pwUsuario = "ADM" then goto Colaborador
		If vUsuario = "ANDREA" and pwUsuario = "ADM" then goto Colaborador
	End If
	Procesando = False
	Exit Sub
Colaborador:
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	Sheet.IsVisible = True
	ThisComponent.getCurrentController.setActiveSheet(Sheet)
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	Sheet.IsVisible = True
	Goto Backup
Exit Sub

Backup:
	Dim Fuente as String
	Dim Destino as String
	Dim Ruta

	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(11, 1)
	Fuente = ConvertToURL( Cell.String )
	If FileExists(Fuente) then
		Cell = Sheet.getCellByPosition(11, 5)
		Ruta = Cell.String+"/("+mid(date,1,2)+"-"+mid(date,4,2)+"-"+mid(date,7,4)+" "+mid(time,1,2)+","+mid(time,4,2)+"hs) VPM - GESTION DE COBROS.ods"
		Destino = ConvertToURL( Ruta )
		FileCopy (Fuente, Destino)
	End If
	Procesando = False
End Sub


