<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosAFE.asp"-->
<%
	Dim idAfe,accion,hayArchivo
	
'--------------------------------------------------------------------------------------------------
Function afeFile2Binary(pIdAfe,pAfeFilePath)
	Dim strSQL,rs,conn
	
	strSQL = "Select * from TBLDATOSAFE where IDAFE = " & pIdAfe
	Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	extension = fso.GetExtensionName(pAfeFilePath)
  
	rs("FILESCAN") = readBinaryFile(pAfeFilePath)
	rs("FILEEXT") = extension
	
	rs.Update
	
End Function
'***********************************************************************
'************		INICIO DE PAGINA 		******************
'***********************************************************************	
	idAfe		= GF_Parametros7("idafe"	  ,0 ,6)	
	accion 		= GF_Parametros7("accion"	  ,"" ,6)	
	afeFilePath = GF_Parametros7("afeFilePath","",6)	
	
	if (idAfe = 0) then
		Response.Redirect "comprasAccesoDenegado.asp"
	else
		strSQL = "select * from tbldatosafe where idafe = " & idafe
        Call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)
		if (not rs.EoF) then
			hayArchivo = false
			if (not isnull(rs("filescan"))) then hayArchivo = true
		end if
		
		if (accion = ACCION_GRABAR) then
			if (afeFilePath <> "") then
				Call afeFile2Binary(idAfe, server.MapPath(".") & "\" & afeFilePath)
				'Si se adjunta una imagen del AFE firmado, se verifica que el monto sea superior a usd 50.000 para aprobarlo(autorizado por Hamburgo)				
				if (necesitaAFEaprobacionHamburgo(idAfe)) then
					strSQL = "UPDATE tbldatosafe SET confirmado = '" & AFE_APROBADO & "' WHERE IDAFE = " & idAfe
					Call executeQueryDb(DBSITE_SQL_INTRA, rsAFE, "UPDATE", strSQL)
				end if
				hayArchivo = true
			end if	
		end if
	end if
%>
<html>
<head>
	<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
	<link rel="stylesheet" href="css/uploadmanager.css"	 type="text/css">
	<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css"	 type="text/css">
	<style>
		.oculto
		{
			visibility:hidden;
		}
		
		.visible
		{
			visibility:visible;
		}
	</style>
    <script type="text/javascript" src="scripts/channel.js"></script>
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
    <script type="text/javascript" src="scripts/JQueryUpload.js"></script>
    <script type="text/javascript" src="scripts/uploadManager.js"></script>
    <script type="text/javascript" src="scripts/botoneraPopUp.js"></script>
    <script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>

	<script>
	
			var ch = new channel();
			var botones = new botonera("botones");
			
			var up1 = new UploadHandler("AfeFile","<% =PATH_COMPRAS_TEMP %>");
			
			function canSubmit() {
				if (up1.isReady) {
					if (!up1.isError) {
						submitForm();
					}
				}
				else {
					setTimeout('canSubmit()', 1000);
				}
			}
	
			function bodyOnLoad()
			{
				<%if (not hayArchivo) then%>
					up1.draw();
					botones.addbutton('Subir','');	
					botones.addbutton('Guardar','canSubmit()');
					botones.show();
				<%end if%>
				
		
			}
			
			function borrarAFE()
			{
				ch.bind("comprasOpenArchivo.asp?id=<%=idAfe%>&type=AFE-DELETE","borrarAfeCallback()");
				ch.send();
			}
			function borrarAfeCallback()
			{
				location.reload();
			}
			
			function submitForm() {			
				if (up1.getFileName() == "")
				{
					alert("Debe subir almenos un archivo.");
				}
				else
				{
					document.getElementById("afeFilePath").value = "<%=PATH_COMPRAS_TEMP%>" +"/"+up1.getFileName();
					document.getElementById("frmSel").submit();
				}
			}
	</script>
    <style>
		
	</style>
	
</head>

<body onLoad="bodyOnLoad()">

<form name="frmSel" id="frmSel" method="post" action="comprasAFEUpload.asp">
	<table width="100%" height="70%" >
		<tr>
			<td align="center" valign="middle">
				<%if (not hayArchivo) then%>
	                <div id="AfeFile"></div>
                <%else
			    	response.write "<img align='absMiddle' src='images/doc.gif'>&nbsp;<a target='_blank' href='comprasOpenArchivo.asp?id="&idafe&"&type=AFE-OPEN'>AFE-"&idAfe&"</a>"
					response.write "&nbsp;<img onclick='borrarAFE()' src='images/delete-16x16.png' style='cursor:pointer' title='Borrar Archivo'>"
			    end if%>
			</td>
		</tr>
	</table>
	
	<div id="botones"></div>
    
    


	<input type="hidden" id="accion" name="accion" value="<%=ACCION_GRABAR%>">
	<input type="hidden" id="idAfe" name="idAfe" value="<%=idAfe%>">
	<input type="hidden" id="afeFilePath" name="afeFilePath" value="">
</form>

</body>

</html>