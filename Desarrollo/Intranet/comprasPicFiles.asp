<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosSeguridad.asp"-->
<!--#include file="Includes/procedimientosfechas.asp"-->
<!--#include file="Includes/procedimientosCTZ.asp"-->
<!--#include file="Includes/md5.asp"-->

<%
Dim idCotizacion, showUplaoder,files,origen,uploadFilesName

'--------------------------------------------------------------------------------------------------
Function getFiles(pIdCotizacion)
	Dim strSQL
	
	strSQL = "select * from TBLCTZBINARYFILES where idcotizacion = " & pidCotizacion & " order by fileno"
	call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	Set getFiles = rs
End Function
'--------------------------------------------------------------------------------------------------

idcotizacion = GF_PARAMETROS7("idcotizacion",0,6)
showUploader = GF_PARAMETROS7("showUploader","",6)
accion 		 = GF_PARAMETROS7("accion","",6)
filePath 	 = GF_PARAMETROS7("filePath","",6)
origen 	 	 = GF_PARAMETROS7("origen","",6)
uploadFilesName = GF_PARAMETROS7("uploadFilesName","",6)


if (accion = ACCION_GRABAR) then
	Call picFile2Binary(idcotizacion,filePath)
end if



%>

<html>
	<head>
		<link rel="stylesheet" href="css/Actisaintra-1.css"	 type="text/css">
		<link rel="stylesheet" href="css/JQueryUpload2.css"	 type="text/css">
        <link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css"	 type="text/css">
        
        <script type="text/javascript" src="scripts/channel.js"></script>
        <script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
        <script type="text/javascript" src="scripts/JQueryUpload2.js"></script>
        <script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>
		
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
										ch.bind("comprasPicFiles.asp?idcotizacion=<%=idcotizacion%>&accion=<%=ACCION_GRABAR%>&filePath=<%=PATH_COMPRAS_TEMP%>/"+fileName,"callbackUpload()");
										ch.send();
									<%end if%>
								},
					allowedExtensions:["jpg","gif","png","bmp","doc","xls","pdf","tif","rar","zip","txt", "docx", "xlsx"]
				}); 
			}
	
			function callbackUpload(){
				//window.parent.guardar();
				$("#myform").submit();
			}
			
			function bodyOnLoad(){
				$("#loading").hide();
				<% if (showUploader) then%>
					createUploader();
				<% end if%>
			}
			
		</script>
	</head>
	 <body onLoad="bodyOnLoad()">
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
		<% if (origen<>"nuevo") then%>
			<table width="320px" align="center" class="reg_header">
				<tr>
					<th class="reg_header_nav ui-corner-tl">Nombre Archivo</th>
					<th class="reg_header_nav ui-corner-tr">.</th>
				</tr>
				<%
				set files = getFiles(idCotizacion)
				
				if (files.EoF) then%>
					<tr>
						<td colspan="2" class="reg_header_navdos" align="center"> No hay Archivos.</td>
					</tr>
				<%else
					while not files.EoF %>
						<tr>
							<td><%=files("name")&"."&files("ext")%></td>
							<td align="center" width="20px">
							<a target='_blank' href='comprasOpenArchivo.asp?idcotizacion=<%=files("idcotizacion")%>&secuencia=<%=files("fileno")%>&type=PIC-OPEN'>
								<img src="images/compras/download.png" title="Descargar">
							</a>
							</td>
						</tr>
						<%
						files.MoveNext
					wend
				end if
				%>
			</table>
		<%end if%>
		<form id="myform">
			<input type="hidden" id="idcotizacion" name="idcotizacion" value="<%=idcotizacion%>">
			<input type="hidden" id="showUploader" name="showUploader" value="<%=showUploader%>">
		</form>
	</body>
</html>