<!--#include file="../Includes/procedimientosParametros.asp"-->
<!--#include file="../Includes/procedimientosFechas.asp"-->
<!--#include file="../Includes/procedimientosLog.asp"-->
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
	
	redim myArchivo(0)
	
	fileName 		= GF_PARAMETROS7("fileName"	,"",6)
	root			= GF_PARAMETROS7("root"	,"",6)
	
	busquedaArchivoFecha	= GF_PARAMETROS7("fecha"	,"",6)
	busquedaArchivoUsuario  = GF_PARAMETROS7("usuario"	,"",6)
	busquedaArchivoTipo	    = GF_PARAMETROS7("tipo"		,"",6)
	busquedaArchivoMsg		= GF_PARAMETROS7("msg"		,"",6)
	
	camposABuscar = 0
	if (busquedaArchivoFecha	<> "" ) then camposABuscar = camposABuscar + BUSCAR_FECHA
	if (busquedaArchivoUsuario	<> "" ) then camposABuscar = camposABuscar + BUSCAR_USUARIO
	if (busquedaArchivoTipo		<> "" ) then camposABuscar = camposABuscar + BUSCAR_TIPO
	if (busquedaArchivoMsg		<> "" ) then camposABuscar = camposABuscar + BUSCAR_MSG
	
	'obtengo el archivo del disco para luego mostrarlo en pantalla	
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
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Consulta de Logs</title>
<link rel="stylesheet" href="../css/ActiSAIntra-1.css" type="text/css">
<script type="text/javascript">
	function submitir() {
		document.getElementById("form1").submit();
	}
</script>
</head>
<body>
<form name='form1' id='form1' action="verLogs.asp" method="post">	
	<% 
		nombre_carpeta = PATHLOGS
		set FSO2     = server.createObject("Scripting.FileSystemObject")
		Set carpeta  = FSO2.GetFolder(nombre_carpeta)
		Set archivos = carpeta.Files

	%>
	<br><br>
	<table width="100%" border="0" align="center" class="reg_header">
		<tr class="reg_header_nav">
			<th colspan='4'>Log del archivo 
					<select name='fileName' id='fileName' onchange='submitir()'>
					<%
					defaultSelected = true
					For each nombre_archivo in archivos
						auxNombreArchivo = split(nombre_archivo,"\")
						'Verifico so el archivo de los encontrado cumple con la raiz comun de archivos a mostrar.
						'Si no indico raiz alguna, se muestran todos los archvos.
						mostrarArchivo = true
						if (root <> "") then
							if (InStr(1, nombre_archivo, root) <= 0) then mostrarArchivo=false
						end if
												
						if (mostrarArchivo) then
							if (auxNombreArchivo(ubound(auxNombreArchivo)) = fileName)then 
								auxSeleccion = "selected='selected'"
								defaultSelected = false
							else
								auxSeleccion = ""								
							end if
							Response.Write "<option " & auxSeleccion & " value='" & auxNombreArchivo(ubound(auxNombreArchivo)) & "'>" & auxNombreArchivo(ubound(auxNombreArchivo)) & "</option>"
						end if
					next
					'No se encontro archivo de Log. Seleccionar el default (el del d�a solicitado).
					if (defaultSelected) then
						Response.Write "<option selected value='" & fileName & "'>" & fileName & "</option>"
					end if
					%>
					</select>
				&nbsp;<input type="submit" id="buscar" name="buscar" value="Ver">				
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
			<%  hayDatos = false
				for i = 1 to ubound(myArchivo)
				hayDatos = true
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
			
			<%next
			if (not hayDatos) then	%>			
				<tr>
					<td class="TDSUCCESS" colspan="4">No hay datos para mostrar</td>
				</tr>
			<%	
			end if			
			%>
	</table>
<input type="hidden" name="root" value="<% =root %>">
</form>
</body>
</html>
