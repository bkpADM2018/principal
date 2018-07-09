<!--#include file="Includes/procedimientosParametros.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosMantenimiento.asp"-->
<!--#include file="Includes/md5.asp"-->

<%
Call initAccessInfo(RES_OT_SM)
Dim idCotizacion, showUplaoder,files,origen,uploadFilesName


'--------------------------------------------------------------------------------------------------
idOT = GF_PARAMETROS7("idOT",0,6)
nroOT = GF_PARAMETROS7("nroOT","",6)
if isAdminInAny then showUploader = true
'showUploader = GF_PARAMETROS7("showUploader","",6)

accion 			= GF_PARAMETROS7("accion","",6)
filePath 		= GF_PARAMETROS7("filePath","",6)
fileNo 			= GF_PARAMETROS7("fileNo",0,6)
origen 	 		= GF_PARAMETROS7("origen","",6)
uploadFilesName = GF_PARAMETROS7("uploadFilesName","",6)

if (accion = ACCION_GRABAR) then
	Call OTFile2Binary(idOT, filePath)
elseif (accion = ACCION_BORRAR) then
	call deleteOTFile(idOT, fileNo)
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
										ch.bind("mantenimientoOTFilesPopUp.asp?idOT=<%=idOT%>&accion=<%=ACCION_GRABAR%>&filePath=<%=PATH_COMPRAS_TEMP%>/"+fileName,"callbackUpload()");
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
			function deleteFile(pIdOT, pFileNo){
				ch.bind("mantenimientoOTFilesPopUp.asp?idOT=" + pIdOT + "&fileNo=" + pFileNo + "&accion=<%=ACCION_BORRAR%>","callbackUpload()");
				ch.send();
			}
		</script>
</head>
<body onLoad="bodyOnLoad()">
	<h3><%=GF_Traducir("Datos de la OT")%></h3>
  
    <div class="tableasidecontent">
        <div class="col26 reg_header_navdos"><% =GF_TRADUCIR("ID") %></div>
        <div class="col46"> <% =idOT%> </div>
       
        <div class="col26 reg_header_navdos"> <% =GF_TRADUCIR("Número") %> </div>
        <div class="col26"> <% =nroOT %> </div>
        
	</div>

<div class="col66"></div>	 

<%
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
		set files = getOTFiles(idOT)
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
					%>
						<tr>
							<td class="thicon">
								<%=getImageByExt(files("EXT"))%>
							</td>	
							<td>
								<%=files("NAME") & "." & files("EXT")%>
							</td>
							<td class="thiconac">
								<a target='_blank' href='comprasOpenArchivo.asp?id=<%=files("ID")%>&secuencia=<%=files("FILENO")%>&type=SM-OT-OPEN'>
									<img width="16" height="16" src="images/download-16.png" title="Descargar Archivo">
								</a>
								<%if isAdminInAny  then%>
									<img src="images/cross-16.png" style="cursor:pointer;" onclick="deleteFile('<%=files("ID")%>','<%=files("FILENO")%>')" title="Eliminar Archivo">
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
			<input type="HIDDEN" id="idOT" name="idOT" value="<%=idOT%>">
			<input type="HIDDEN" id="nroOT" name="nroOT" value="<%=nroOT%>">
			<input type="hidden" id="showUploader" name="showUploader" value="<%=showUploader%>">
		</form>
	</body>
</html>