<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/md5.asp"-->

<%
Call initAccessInfo(RES_INV_SM)
Dim idCotizacion, showUplaoder,files,origen,uploadFilesName, accion


'--------------------------------------------------------------------------------------------------
dim myType
idEquipo = GF_PARAMETROS7("idEquipo",0,6)
idEquipoActivado = GF_PARAMETROS7("idEquipoActivado", 0, 6)
myType = GF_PARAMETROS7("pTipo", 0, 6)
idComponent = GF_PARAMETROS7("idComponent",0,6)
idSubComponent = GF_PARAMETROS7("idSubComponent",0,6)
if myType = 0 then myType = 1
if isAdminInAny then showUploader = true
'showUploader = GF_PARAMETROS7("showUploader","",6)

accion 		 = GF_PARAMETROS7("accion","",6)
filePath 	 = GF_PARAMETROS7("filePath","",6)
fileNo 		 = GF_PARAMETROS7("fileNo",0,6)
origen 	 	 = GF_PARAMETROS7("origen","",6)
uploadFilesName = GF_PARAMETROS7("uploadFilesName","",6)
if idEquipoActivado <> 0 then
	myType = 2
	idEquipoDefault = idEquipoActivado
	call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMACTIVEEQUIPMENT_GET_FULL_BY_ID", idEquipoActivado)
	if not rs.eof then
		idEquipo = trim(rs("IDEQUIPMENT"))
		cdEquipo = rs("CDEQUIPMENT")
		dsEquipo = rs("DSEQUIPMENT") 
		idDivision = rs("IDDIVISION") 
		dsDivision = rs("DSDIVISION") 
		idSector = rs("IDSECTOR") 
		dsSector = rs("DSSECTOR") 
		cdActivacion = trim(rs("CDACTIVATION"))	   
		dsActivacion = trim(rs("DSACTIVATION"))	   
		cdActivoFijo = trim(rs("CDACTIVECODE"))	   
	end if			 	
else
	if idEquipo <> 0 then
		idEquipoDefault = idEquipo
		call executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLSMEQUIPMENT_GET_BY_PARAMETERS", idEquipo & "|| || ||0||")
		if not rs.eof then
		   cdEquipo = rs("CDEQUIPMENT")
		   dsEquipo = rs("DSEQUIPMENT") 
		end if			 	
	end if	
end if		
if idComponent <> 0 then 
	idEquipoDefault = idComponent
	myType = 3
end if	
if idSubComponent <> 0 then 
	idEquipoDefault = idSubComponent
	myType = 4
end if	
if (accion = ACCION_GRABAR) then
	Call smFile2Binary(idEquipoDefault, myType , filePath)
elseif (accion = ACCION_BORRAR) then
	call deleteFile(idEquipoDefault, myType, fileNo)
end if
'---------------------------------------------------------------------------------------------------------------


%>

<html>
	<head>
		<link rel="stylesheet" href="css/main.css"	 type="text/css">
		<link rel="stylesheet" href="css/JQueryUpload2.css"	 type="text/css">
        <link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css"	 type="text/css">
		
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

        <script type="text/javascript" src="scripts/channel.js"></script>
        <script type="text/javascript" src="scripts/JQueryUpload2.js"></script>
		
		 <script>
			var ch = new channel();
			var uploader;
			
			function createUploader(){            
				uploader = new qq.FileUploader({
					element: document.getElementById('myFile'),
					action: 'uploadSubmitFile2.asp?accion=upload&folder=<%=PATH_COMPRAS_TEMP %>',
					multiple:false,
					onComplete: function(id, fileName, responseJSON){
									<%if (origen="nuevo") then%>
										parent.uploadFilesName += fileName+","
									<%else%>
										$("#upload").hide();
										$("#loading").show();
										//var auxFileName = fileName.substring(0,uploader.fileName.lastIndexOf("."))
										ch.bind("mantenimientoEquipoFiles.asp?idEquipo=<%=idEquipo%>&isActive=<%=isActive%>&idEquipoActivado=<%=idEquipoActivado%>&idComponent=<%=idComponent%>&idSubComponent=<%=idSubComponent%>&accion=<%=ACCION_GRABAR%>&filePath=<%=PATH_COMPRAS_TEMP%>/"+fileName,"callbackUpload()");
										ch.send();
									<%end if%>
								},
					allowedExtensions:["jpg","gif","png","bmp","doc","xls","pdf","tif","rar","zip","txt", "docx", "xlsx"]
				}); 
			}
	
			function callbackUpload(){
				$("#myform").submit();
			}
			
			function bodyOnLoad(){
				$("#loading").hide();
				<% if (showUploader) then%>
					createUploader();
				<% end if%>
			}
			function deleteFileJS(pIdEquipo, pTipo, pFileNo){
				ch.bind("mantenimientoEquipoFiles.asp?idEquipo=" + pIdEquipo + "&pTipo=" + pTipo + "&idEquipoActivado=<%=idEquipoActivado%>&fileNo=" + pFileNo + "&accion=<%=ACCION_BORRAR%>","callbackUpload()");
				ch.send();
			}
		</script>
</head>
<body onLoad="bodyOnLoad()">
<div class="tableaside size100"> 
	<h3><%=GF_Traducir("Datos del Master")%></h3>
  
    <div class="tableasidecontent">
        <div class="col26 reg_header_navdos"><% =GF_TRADUCIR("ID") %></div>
        <div class="col26"> <% =idEquipo%> </div>
       
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("C�digo") %> </div>
        <div class="col26"> <% =cdEquipo %> </div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Descripci�n") %> </div>
        <div class="col36"> <% =dsEquipo%> </div>
	</div>
</div>

<div class="col66"></div>	 

<%if idEquipoActivado <> 0 then%>
<div class="tableaside size100"> 
	<h3><%=GF_Traducir("Datos de Activaci�n")%></h3>
  
    <div class="tableasidecontent">
        <div class="col26 reg_header_navdos"><% =GF_TRADUCIR("ID") %></div>
        <div class="col26"> <% =idEquipoActivado%> </div>
       
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("C�digo") %> </div>
        <div class="col26"> <% =cdActivacion %> </div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Descripci�n") %> </div>
        <div class="col26"> <% =dsActivacion%> </div>

        <div class="col26 reg_header_navdos"><% =GF_TRADUCIR("Divisi�n") %></div>
        <div class="col26"> <% =dsDivision%> </div>
       
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Sector") %> </div>
        <div class="col26"> <% =dsSector %> </div>
        
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Activo Fijo") %> </div>
        <div class="col26"> <% =cdActivoFijo%> </div>
	</div>	
</div>
<div class="col66"></div>	
<%end if
	if (origen<>"nuevo") then%>
	<table class="datagrid" width="90%" align="center">
		<thead>
		    <tr>
		        <th class="thicon"><%=GF_Traducir("Tipo")%> </th>
		        <th><%=GF_Traducir("Nombre del archivo")%></th>
		        <th class="thiconac" width="80"> - </th>
		    </tr>
		</thead>
		<%
		set files = getFiles(idEquipoDefault, myType)
		if (files.EoF) then%>
			<tbody>	
				<tr>
					<td colspan="3" align="center"><%=GF_Traducir("No se encontraron archivos.")%></td>
				</tr>
			</tbody>	
		<%else%>
			<tbody>	
			<%
					while not files.EoF 
						if files("TYPE") = 1 then
							if flag1 = 0 then
								flag1 = 1 %>
								<tr><td colspan="3"><%=GF_Traducir("Adjuntos del Master")%></td></tr>
								<%
							end if
						elseif files("TYPE") = 2 then	
							if flag2 = 0 then
								flag2 = 1 %>
								<tr><td colspan="3"><%=GF_Traducir("Adjuntos de la Activaci�n")%></td></tr>
								<%
							end if
						elseif files("TYPE") = 3 then	
							if flag3 = 0 then
								flag3 = 1 %>
								<tr><td colspan="3"><%=GF_Traducir("Adjuntos de Componentes")%></td></tr>
								<%
							end if
							if myType = 1 or myType = 2 then
								if cint(files("IDCOMPONENT")) <> cint(myComponentAnt) then
									myComponentAnt = files("IDCOMPONENT") 
									%>
									<tr><td class="" colspan="3"><I><%=files("IDCOMPONENT") & "-" & files("DSCOMPONENT")%><I></td></tr>
									<%
								end if					
							end if
						elseif files("TYPE") = 4 then	
							if flag4 = 0 then
								flag4 = 1 %>
								<tr><td colspan="3"><%=GF_Traducir("Adjuntos de Sub-Componentes")%></td></tr>
								<%
							end if
							if myType = 1 or myType = 2 then
								if cint(files("IDCOMPONENT")) <> cint(mySubComponentAnt) then
									mySubComponentAnt = files("IDCOMPONENT") 
									%>
									<tr><td class="" colspan="3"><i><%=files("IDCOMPONENT") & "-" & files("DSCOMPONENT")%><i></td></tr>
									<%
								end if
							end if
						end if%>	
						<tr>
							<td class="thicon">
								<%=getImageByExt(files("EXT"))%>
							</td>	
							<td>
								<%=files("NAME") & "." & files("EXT")%>
							</td>
							<td class="thiconac">
								<a target='_blank' href='comprasOpenArchivo.asp?id=<%=files("ID")%>&typeO=<%=files("TYPE")%>&secuencia=<%=files("FILENO")%>&type=SM-OPEN'>
									<img width="16" height="16" src="images/download-16.png" title="Descargar Archivo">
								</a>
								<%if (idEquipoActivado = 0 or files("TYPE") = 2) and isAdminInAny  then%>
									<img src="images/cross-16.png" style="cursor:pointer;" onclick="deleteFileJS('<%=files("ID")%>','<%=files("TYPE")%>','<%=files("FILENO")%>')" title="Eliminar Archivo">
								<%else
									Response.Write "."
								end if%>
							</td>
						</tr>
						<%
						files.MoveNext
					wend
				end if
				%>
			</table>
			</tbody>	
		<%end if%>
		<br>
		<div id="upload">
		<% if (showUploader) then%>
			<div id="myFile" style="margin:auto;width:150px"></div>
		<% end if%>
		</div>
		<div style="margin:auto;width:130px" id="loading">
			<img src="images/compras/loading_bar_green.gif" >
		</div>
		<%
		' Ya que el upload no tiene una forma de cargarle los archivos ya subidos los muestro en una lista imitando su funcionamiento
		if ( uploadFilesName <> "") then 
			listFileName = split(uploadFilesName,",")
			response.write "<div style='margin:auto;width:150px'><ul class='qq-upload-list'>"
			for y = 0 to ubound(listFileName)-1 ' el -1 es porque siempre viene con una coma al final
				response.write "<li class=' qq-upload-success'>"&listFileName(y)&"</li>"
			next
			response.write "</ul></div>"
		end if
		%>
		
		<form id="myform">
			<input type="HIDDEN" id="idEquipo" name="idEquipo" value="<%=idEquipo%>">
			<input type="HIDDEN" id="idEquipoActivado" name="idEquipoActivado" value="<%=idEquipoActivado%>">
			<input type="HIDDEN" id="idComponent" name="idComponent" value="<%=idComponent%>">
			<input type="HIDDEN" id="idSubComponent" name="idSubComponent" value="<%=idSubComponent%>">
			<input type="hidden" id="showUploader" name="showUploader" value="<%=showUploader%>">
		</form>
	</body>
</html>