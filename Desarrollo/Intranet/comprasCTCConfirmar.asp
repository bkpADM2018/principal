<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosCTC.asp"-->
<!--#include file="Includes/procedimientosPCT.asp"-->
<%
'-------------------------------------------------------------------------
'confirmar el contrato, solo se piden dos campos cdcontrato y archivo ----
'adjunto, este ultimo es opcional. Para poder confirmar el usuario debe --
'estar autorizado en el maestro de responsables --------------------------
'-------------------------------------------------------------------------
'***********************************************
'*************  COMIENZO DE PAGINA  ************
'***********************************************
Dim idContrato, accion, cdContrato, file, grabo

idContrato = GF_PARAMETROS7("idContrato",0,6)
accion =  GF_PARAMETROS7("accion","",6)
cdContrato = GF_PARAMETROS7("cdContrato","",6)
file = GF_PARAMETROS7("CTCFile","",6)

if (not canConfirmCTC(session("Usuario"), idContrato)) then
	Response.Redirect "comprasAccesoDenegado.asp"
else
	if (accion <> ACCION_GRABAR) then
		'Para diferenciar si ya confirmó el contrato o si subió un archivo, verifico el estado. Esto es necesario
		'para saber si tengo que agregar los datos adicionales o agregar el archivo
		Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLOBRACONTRATOS_GET_BY_IDCONTRATO", idContrato)
		if (not rs.EoF) then
			if (rs("ESTADO") > ESTADO_CTC_PENDIENTE) Then cdContrato = rs("CDCONTRATO")
		end if
	else
		grabo = confirmarContrato(idContrato, cdContrato, file)
	end if
end if

%>
<html>
	<head>
		<link rel="stylesheet" href="css/ActiSAIntra-1.css"	 type="text/css">
		<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
		<link rel="stylesheet" href="css/uploadManager.css" type="text/css">
		<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
				
		<script type="text/javascript" src="scripts/formato.js"></script>
		<script type="text/javascript" src="scripts/Toolbar.js"></script>
		<script type="text/javascript" src="scripts/controles.js"></script>				
		<script type="text/javascript" src="scripts/channel.js"></script>		
		<script type="text/javascript" src="scripts/uploadManager.js"></script>		
		<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
		<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
		<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
			
		<script defer type="text/javascript" src="scripts/pngfix.js"></script>
		
		<script type="text/javascript">
			
			var up1 = new UploadHandler("fileCTC","<% =PATH_COMPRAS_TEMP %>");
						
			var puw;
			function bodyOnLoad(){
				puw = getObjPopUp('popUpCTCConfirm');
				<% if (grabo) then %>
					puw.hide();
				<% end if %>
				<% if (hayError()) then %>
					document.getElementById("cdContrato").style.backgroundColor = '#FFAA99';
					document.getElementById("cdContrato").style.border = '1px red solid';
				<% end if %>
				up1.draw();
				document.getElementById("cdContrato").focus();
			}
			
			function grabar(){
				document.getElementById("tbLoading").style.visibility = 'visible';
				document.getElementById("btnLoading").style.visibility = 'hidden';
				if (up1.isReady) {
					document.getElementById("CTCFile").value = up1.getFileName();
					document.getElementById("accion").value = '<% =ACCION_GRABAR %>';
					document.getElementById("frmSel").submit();
				} else {					
					setTimeout("grabar()", 1000);	
				}				
			}
		</script>
	</head>
	<body OnLoad="bodyOnLoad()">
		<table><tr><td>
			<img align="absMiddle" src="images/compras/CTC_Confirm-32x32.png">
			<span style="font-size:12px;font-weight:bold;"><% =GF_TRADUCIR("Confirmar Contrato") %></span>
		</td></tr></table>
		<form id=frmSel name=frmSel>
			<table width="90%" align="center">
				<tr>
					<td colspan="2"><% Call showErrors() %></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Id: ") %></td>
					<td><b><% =idContrato %></b></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Contrato: ") %></td>
					<td><input type="text" id="cdContrato" name="cdContrato" value="<% =cdContrato %>"></td>
				</tr>
				<tr>
					<td class="reg_header"><% =GF_TRADUCIR("Adjunto: ") %></td>
					<td><div id="fileCTC"></div></td>
				</tr>
				<tr>
					<td colspan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colspan="2" align=center>
					<%	if (idContrato > 0) then %>
						<input type="button" id="btnLoading" onClick="grabar();" value=<% =GF_TRADUCIR("Aceptar") %>>						
						<table id="tbLoading" border="0" class="reg_header round_border_all" style="visibility:hidden">
							<tr>
								<td align="center"><strong><div><% =GF_TRADUCIR("Confirmando") %>...</div></strong></td>
							</tr>
						</table>																
					<%	end if %>
					</td>
				</tr>
				<input type="hidden" id="idContrato" name="idContrato" value="<% =idContrato %>">
				<input type="hidden" id="CTCFile" name="CTCFile" value="<% =file %>">
				<input type="hidden" id="accion" name="accion" value="">
			</table>
		</form>
	</body>
</html>