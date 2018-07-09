<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosLog.asp"-->
<%
	Const COLFECHA	= 0
	Const COLUSER	= 1
	Const COLTIPO	= 2
	Const COLMSG	= 3
	
	Const BUSCAR_FECHA	 = 1
	Const BUSCAR_USUARIO = 2
	Const BUSCAR_TIPO	 = 4
	Const BUSCAR_MSG	 = 8
	
	Dim PATHLOGS
	
	PATHLOGS = server.MapPath("logs")
	
	Dim handlers,mensajes
	Dim dia,mes,anio,myfileName,myfile
	Dim myArchivo()
	Dim mostrarFileLog,mostrarDBLog
	Dim busquedaArchivoFecha,busquedaArchivoUsuario,busquedaArchivoTipo,busquedaArchivoMsg,camposABuscar
	Dim auxNombreArchivo,hay
	Dim aux

	mostrarDBLog = false
	
	redim myArchivo(0)
	
	handlers = 0
	mensajes = 0
	
	accion   = GF_PARAMETROS7("accion"  ,"",6)

	archivo  = GF_PARAMETROS7("archivo"	,"",6)
	pantalla = GF_PARAMETROS7("pantalla","",6)
	dataBase = GF_PARAMETROS7("db"		,"",6)	
	info 	 = GF_PARAMETROS7("info"	,"",6)
	errores  = GF_PARAMETROS7("errores"	,"",6)
	warning  = GF_PARAMETROS7("warning"	,"",6)
	debug1 	 = GF_PARAMETROS7("debug"	,"",6)
	
	mostrarFileLog	= GF_PARAMETROS7("filelog"	,"",6)
	fileName 		= GF_PARAMETROS7("fileName"	,"",6)
	
	if (mostrarFileLog = "" ) then 
		mostrarFileLog = False
	else
		mostrarFileLog = True
	end if
	
	busquedaArchivoFecha	= GF_PARAMETROS7("fecha"	,"",6)
	busquedaArchivoUsuario  = GF_PARAMETROS7("usuario"	,"",6)
	busquedaArchivoTipo	    = GF_PARAMETROS7("tipo"		,"",6)
	busquedaArchivoMsg		= GF_PARAMETROS7("msg"		,"",6)
	
	camposABuscar = 0
	if (busquedaArchivoFecha	<> "" ) then camposABuscar = camposABuscar + BUSCAR_FECHA
	if (busquedaArchivoUsuario	<> "" ) then camposABuscar = camposABuscar + BUSCAR_USUARIO
	if (busquedaArchivoTipo		<> "" ) then camposABuscar = camposABuscar + BUSCAR_TIPO
	if (busquedaArchivoMsg		<> "" ) then camposABuscar = camposABuscar + BUSCAR_MSG
	
	if (accion = "GUARDAR") then
		if (archivo  <> "") then handlers = handlers + HND_FILE
		if (pantalla <> "") then handlers = handlers + HND_VIEW
		if (dataBase <> "") then handlers = handlers + HND_DB
		
		if (info     <> "") then mensajes = mensajes + MSG_INF_LOG
		if (errores  <> "") then mensajes = mensajes + MSG_ERR_LOG
		if (warning  <> "") then mensajes = mensajes + MSG_WRN_LOG
		if (debug1   <> "") then mensajes = mensajes + MSG_DBG_LOG
		Call startLog(handlers, mensajes)
	end if
	
	if (accion = "BORRARTODOS") then
		Call startLog(0, 0)
	end if
	if (accion = "ACTIVARTODOS") then
		Call startLog(HND_FILE+HND_VIEW+HND_DB, MSG_INF_LOG+MSG_ERR_LOG+MSG_WRN_LOG+MSG_DBG_LOG)
	end if
	
	if (accion = "delete") then
		deleteFile  = GF_PARAMETROS7("deleteFile","",6)
		if (deleteFile <> "" ) then
			Set myFs1 = Server.CreateObject("Scripting.FileSystemObject")
			myFs1.DeleteFile(PATHLOGS & "\" & deleteFile)
		end if
	end if
	
	archivo  = Session("LOG_HDN_FILE")
	pantalla = Session("LOG_HDN_VIEW")
	dataBase = Session("LOG_HDN_DB"  )
	info 	 = Session("LOG_INF_ENABLED")
	errores  = Session("LOG_ERR_ENABLED")
	warning  = Session("LOG_WRN_ENABLED")
	debug1 	 = Session("LOG_DBG_ENABLED")
	
	'obtengo el archivo del disco para luego mostrarlo en pantalla
	if (mostrarFileLog) then
		if (fileName = "") then
			dia  = day(now())
			mes  = month(now())
			anio = year(now())
			Call GF_STANDARIZAR_FECHA(dia,mes,anio)
			myfileName = anio & mes & dia & ".txt"
			fileName = myfileName
		else
			myfileName = fileName
		end if
		
		Set fso = CreateObject("scripting.filesystemobject")
		if (fso.FileExists(PATHLOGS & "\" & myfileName )) then
				Set myfile = fso.OpenTextFile(PATHLOGS & "\" & myfileName,1,false)
				while not myfile.AtEndOfStream
					redim preserve myArchivo(ubound(myArchivo)+1)
					myArchivo(ubound(myArchivo)) = myfile.ReadLine
				wend
		end if
	end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Administracion de Logs</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript">
	function borrar(){
		document.getElementById('deleteFile').value = document.getElementById('fileName').value
		document.getElementById('accion').value = "delete"
		document.getElementById('form1').submit();
	}
	function submitir(accion){
		document.getElementById('accion').value = accion;
		document.getElementById('form1').submit();
	}
	function submitirBusqueda(e){
		tecla = (document.all) ? e.keyCode : e.which;
		if (tecla==13){
			document.getElementById('accion').value = '';
			document.getElementById('form1').submit();
		}
	}
</script>
</head>
<body>
<form name='form1' id='form1' action="administrarLogs.asp" method="post">
	<input type='hidden' name='accion' id='accion' value='<%=accion%>'>
	<table width="100%" border="0" align="center" class="reg_header" >
	    <tr >
	      <th colspan="4" align="center">  
	        <div align="left" class="reg_header_nav">Handler </div></th>
	    </tr>
	    <tr class="reg_header_navdos">
			<td width="193"><input type="checkbox" name="archivo"  id="archivo"  <%if (archivo <> "False")  then %> checked="checked" <%end if%> />
				Archivo </td>
			<td width="191"><input type="checkbox" name="pantalla" id="pantalla" <%if (pantalla <> "False") then %> checked="checked" <%end if%> />
				Pantalla </td>
			<td width="191"><input type="checkbox" name="db"       id="db"       <%if (dataBase <> "False") then %> checked="checked" <%end if%>/>
				Base de Datos</td>
			<td width="191">&nbsp;</td>
		</tr>
	    <tr>
			<th colspan="4"><div align="left" class="reg_header_nav">Mensajes</div></th>
	    </tr>
	    <tr class="reg_header_navdos">
			<td><input type="checkbox" name="info"    id="info"    <%if (info <> "False") then    %> checked="checked" <%end if%>/>
				Informacion</td>
			<td><input type="checkbox" name="errores" id="errores" <%if (errores <> "False") then %> checked="checked" <%end if%>/>
				Errores</td>
			<td><input type="checkbox" name="warning" id="warning" <%if (warning <> "False") then %> checked="checked" <%end if%>/>
				Warnings</td>
			<td><input type="checkbox" name="debug"   id="debug"   <%if (debug1 <> "False") then  %> checked="checked" <%end if%>/>
				Debug</td>
		</tr>
	</table>
	<p align="left">
	    <input type="button" name="accion" id="accion" value="Guardar"		 onclick="submitir('GUARDAR')"     />
	    <input type="button" name="accion" id="accion" value="Borrar Todos"	 onclick="submitir('BORRARTODOS')" />
	    <input type="button" name="accion" id="accion" value="Activar Todos" onclick="submitir('ACTIVARTODOS')"/>
	</p>
	<td><input type="checkbox" name="filelog" id="filelog" <%if (mostrarFileLog <> "False") then %> checked="checked" <%end if%> onClick="submitir('')"/> Mostrar Log de Archivos
	<% if (mostrarFileLog) then 

		nombre_carpeta = PATHLOGS
		set FSO2     = server.createObject("Scripting.FileSystemObject")
		Set carpeta  = FSO2.GetFolder(nombre_carpeta)
		Set archivos = carpeta.Files

		%>
		<br><br>
		<%
	end if %>
	<table width="100%" border="0" align="center" class="reg_header">
	<%if (mostrarFileLog) then%>
		<tr class="reg_header_nav">
			<th colspan='4'>Log del archivo <select name='fileName' id='fileName' onchange='submitir()'>
											<%
											for each nombre_archivo in archivos
												auxNombreArchivo = split(nombre_archivo,"\")
												if (auxNombreArchivo(ubound(auxNombreArchivo)) = fileName)then 
													auxSeleccion = "selected='selected'"
												else
													auxSeleccion = ""
												end if
												Response.Write "<option " & auxSeleccion & " value='" & auxNombreArchivo(ubound(auxNombreArchivo)) & "'>" & auxNombreArchivo(ubound(auxNombreArchivo)) & "</option>"
											next
											%>
											</select>
				&nbsp;<input type="button" id="delete" name="delete" value="Borrar" onclick='borrar()' >
				<input type="hidden" id="deleteFile" name="deleteFile" value="" >
			</th>

		</tr>
		<tr align='center' class="reg_header_nav">
			<th width='10'><input type='text' name='fecha'		id='fecha'	 value='<%=busquedaArchivoFecha%>'   size='14' onKeyUp="submitirBusqueda(event)"/></th>
			<th width='10'><input type='text' name='usuario'	id='usuario' value='<%=busquedaArchivoUsuario%>' size='4'  onkeyup="submitirBusqueda(event)"/></th>
			<th width='10'><input type='text' name='tipo'		id='tipo'	 value='<%=busquedaArchivoTipo%>'    size='5'  onkeyup="submitirBusqueda(event)"/></th>
			<th><input type='text' name='msg' id='msg' value='<%=busquedaArchivoMsg%>' size='50' onKeyUp="submitirBusqueda(event)"/></th>
		 </tr>
		<tr class="reg_header_nav">
			<th>Fecha  </th>
			<th>Usuario</th>
			<th>Tipo   </th>
			<th>Mensaje</th>
		</tr>
			<% for i = ubound(myArchivo) to 1 step -1 
				aux = split(myArchivo(i),"|")
			%>
			
				<%
				busqueda = 0
				if (camposABuscar > 0) then
					if (busquedaArchivoFecha	<> "")	then 
						if (instr(1,ucase(aux(COLFECHA)),ucase(busquedaArchivoFecha))>0) then
							busqueda = busqueda + BUSCAR_FECHA
						end if
					end if
					if (busquedaArchivoUsuario	<> "")	then 
						if (instr(1,ucase(aux(COLUSER)),ucase(busquedaArchivoUsuario))>0) then
							busqueda = busqueda + BUSCAR_USUARIO
						end if
					end if
					if (busquedaArchivoTipo	<> "")	then 
						if (instr(1,ucase(aux(COLTIPO)),ucase(busquedaArchivoTipo))>0) then
							busqueda = busqueda + BUSCAR_TIPO
						end if
					end if
					if (busquedaArchivoMsg	<> "")	then 
						if (instr(1,ucase(aux(COLMSG)),ucase(busquedaArchivoMsg))>0) then
							busqueda = busqueda + BUSCAR_MSG
						end if
					end if
					
					if (camposABuscar = busqueda ) then
					%>
					<tr class="reg_header_navdos">
						<td><%=aux(COLFECHA)%></td>
						<td><%=aux(COLUSER) %></td>
						<td><%=aux(COLTIPO) %></td>
						<td><%=aux(COLMSG  )%></td>
					</tr>
					<%
					end if
				else%>
				<tr class="reg_header_navdos">
					<td><%=aux(COLFECHA)%></td>
					<td><%=aux(COLUSER) %></td>
					<td><%=aux(COLTIPO) %></td>
					<td><%=aux(COLMSG  )%></td>
				</tr>
				<%
				end if%>
			
			<%next%>
	<%end if%>
	<%if (mostrarDBLog) then%>
		<!--LOG DE BASE DE DATOS-->
	<%end if%>
	</table>
		
</form>
</body>
</html>
