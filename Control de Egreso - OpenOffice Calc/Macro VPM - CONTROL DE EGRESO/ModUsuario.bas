REM  *****  BASIC  *****
Option Explicit

'HOJA DE CALCULO
'Dim Doc As Object
'Sheet As Object
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
Dim otxtPWVista As Object

Sub GuardarDocumento
	Doc = thiscomponent
	Doc.Store
End Sub

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

	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("BDClientes")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Datos")
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
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("BDClientes")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Datos")
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
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("BDClientes")
	Sheet.IsVisible = False
	Sheet = Doc.Sheets.getByName("Datos")
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
		Sheet = Doc.Sheets.getByName("Hoja de Ruta")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("BDClientes")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Datos")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
		Sheet.IsVisible = True
		Sheet = Doc.Sheets.getByName("Carga de Tareas")
		Sheet.IsVisible = True
		ThisComponent.getCurrentController.setActiveSheet(Sheet)
		Procesando = False
		goto Backup
		Exit Sub
	End If
	
	If tUsuario = "COLABORADOR" then
		If vUsuario = "MARTIN" and pwUsuario = "MG" then goto Colaborador
		If vUsuario = "LUIS" and pwUsuario = "LR" then goto Colaborador
		If vUsuario = "NICOLAS" and pwUsuario = "NG" then goto Colaborador
		If vUsuario = "GUILLERMO" and pwUsuario = "GM" then goto Colaborador
		If vUsuario = "MAXIMILIANO" and pwUsuario = "MG" then goto Colaborador
		If vUsuario = "ISMAEL" and pwUsuario = "IL" then goto Colaborador
		If vUsuario = "PAOLA" and pwUsuario = "ADM" then goto Colaborador
		If vUsuario = "ANDREA" and pwUsuario = "ADM" then goto Colaborador
	End If
	Exit Sub
Colaborador:
	Sheet = Doc.Sheets.getByName("Hoja de Ruta")
	Sheet.IsVisible = True
	Sheet = Doc.Sheets.getByName("Carga de Tareas")
	Sheet.IsVisible = True
	ThisComponent.getCurrentController.setActiveSheet(Sheet)
	Sheet = Doc.Sheets.getByName("Expedicion-Cobros")
	Sheet.IsVisible = True
	Procesando = False
	goto Backup
Exit Sub

Backup:
	Dim Fuente as String
	Dim Destino as String
	Dim Ruta

	Sheet = Doc.Sheets.getByName("Datos")
	Cell = Sheet.getCellByPosition(11, 2)
	Fuente = ConvertToURL( Cell.String )
	If FileExists(Fuente) then
		Cell = Sheet.getCellByPosition(11, 5)
		Ruta = Cell.String+"/("+mid(date,1,2)+"-"+mid(date,4,2)+"-"+mid(date,7,4)+" "+mid(time,1,2)+","+mid(time,4,2)+"hs) VPM - CONTROL DE EGRESO.ods"
		Destino = ConvertToURL( Ruta )
		FileCopy (Fuente, Destino)
	End If
	HoraUltGuardar = Timer
	'saco filtrado inicial porque se pone lento en la maquina de expedicion cuando inicia.
	'FiltroEstado
End Sub


