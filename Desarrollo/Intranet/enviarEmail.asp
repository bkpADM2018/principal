<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<!-- #include file="Includes/procedimientosCompras.asp"-->
<!-- #include file="Includes/procedimientosMail.asp"-->

<%

Dim email,accion,origen,asunto,mensaje,enviado

email = GF_PARAMETROS7("email","",6)
origen = GF_PARAMETROS7("origen","",6)
asunto = GF_PARAMETROS7("asunto","",6)
mensaje = GF_PARAMETROS7("mensaje","",6)
accion = GF_PARAMETROS7("accion","",6)

mensaje = replace(mensaje,"<br>",chr(10)&chr(13))

enviado = false
if (accion = ACCION_EMAIL) then
	call GP_ENVIAR_MAIL(asunto,mensaje,origen,email)
	enviado = true
end if


%>
<html>
<head>
<title>Familiar Proveedor</title>
	
	<link href="css/jqueryUI/custom-theme/jquery-ui-1.8.15.custom.css" rel="stylesheet" type="text/css">
	<link href="css/ActisaIntra-1.css" rel="stylesheet" type="text/css">
	
	<script type="text/javascript" src="scripts/jquery/jquery-1.5.1.min.js"></script>
    <script type="text/javascript" src="scripts/botoneraPopUp.js"></script>
    <script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.15.custom.min.js"></script>

	<script>
    	var botones = new botonera("botones");
		
		function bodyOnLoad()
		{
			<% if (not enviado) then %>
				botones.addbutton('Enviar','enviar()');
				botones.show();
			<% end if%>
		}
		
		function enviar()
		{
			$("#myForm").submit();
		}
		
	</script>
	
</head>
<body onLoad="bodyOnLoad()">
		<% if (enviado) then %>
			<div class="ui-state-highlight ui-corner-all" style="padding:5px">
					<span class="ui-icon ui-icon-info" style="float: left; margin-right: .3em;"></span>
					El email fue enviado correctamente
				</div><br />
		<% else %>
		<form id="myForm" method="POST" action="enviarEmail.asp">
		<table align="center" class="reg_header" width="350px">
			<tr>
				<td class="reg_header_navdos">
					De:
				</td>
				<td>
					<%=getUserMail(session("Usuario"))%>
					<input type="hidden" id="origen" name="origen" value="<%=getUserMail(session("Usuario"))%>">
				</td>
			</tr>
			<tr>
				<td class="reg_header_navdos">
					Para:
				</td>
				<td>
					<%=email%>
					<input type="hidden" id="email" name="email" value="<%=email%>">
				</td>
			</tr>
			<tr>
				<td class="reg_header_navdos">
					Asunto:
				</td>
				<td>
					<input type="text" id="asunto" name="asunto" value="<%=asunto%>">
				</td>
			</tr>
			<tr>
				<td class="reg_header_navdos">
					mensaje:
				</td>
				<td>
					<textarea name="mensaje" cols="40" rows="5" id="mensaje"><%=mensaje%></textarea>
				</td>
			</tr>
		</table>
		<input type="hidden" id="accion" name="accion" value="<%=ACCION_EMAIL%>">
		<div id="botones"></div>
		</form>
	<% end if %>
</body>
</html>