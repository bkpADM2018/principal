<!-- #include file="Includes/procedimientosMG.asp"-->
<!-- #include file="Includes/procedimientosUser.asp"-->
<!-- #include file="Includes/procedimientosAlmacenes.asp"-->
<!-- #include file="Includes/procedimientosMail.asp"-->

<%

Dim email,accion,origen,asunto,mensaje,enviado,oDiccCantidadesPedidas
Set oDiccCantidadesPedidas  = createObject("Scripting.Dictionary")
'-------------------------------------------------------------------------------------------------'
Function getCantidadPedida(pIdArticulo)
	Dim rtrn

	rtrn = 0
	if (oDiccCantidadesPedidas.Exists(cdbl(pIdArticulo))) then
		rtrn = oDiccCantidadesPedidas.Item(cdbl(pIdArticulo))
	end if

	getCantidadPedida = rtrn
End Function

articulos = GF_PARAMETROS7("articulos","",6)
idAlmacen = GF_PARAMETROS7("idAlmacen",0,6)

Set oDiccCantidadesPedidas = cargarCantidadesPedidas(idAlmacen,0)

strSQL = "select * from TBLMAILSALERTASALMACENES where idalmacen = " &idAlmacen
call executeQueryDb(DBSITE_SQL_INTRA, rs4, "OPEN", strSQL)
email = ""
while not rs4.EoF
	if (trim(cstr(rs4("email"))) <> "") then email = email & rs4("email") & ";"
	rs4.MoveNext
wend
'le quito la ultima coma'

articulos = left(articulos,len(articulos)-1)
	strSQL = 		  "SELECT a.*,( existencia + sobrante ) stock,art.dsarticulo "
	strSQL = strSQL & "FROM   (select * from tblarticulosdatos where idarticulo in ("&articulos&") )a"
	strSQL = strSQL & " INNER JOIN tblarticulos art on a.idarticulo = art.idarticulo"
	strSQL = strSQL & " WHERE  idalmacen = " & idAlmacen
	strSQL = strSQL & " AND ( existencia + sobrante ) <= stockminimo "
	strSQL = strSQL & "AND    stockminimo <> 0 order by art.idarticulo"
	call executeQueryDb(DBSITE_SQL_INTRA, rs1, "OPEN", strSQL)

while not rs1.EoF
	mensaje = mensaje & rs1("idarticulo") & " - " & rs1("dsarticulo") & " | stock Actual: " & rs1("stock") & " - Pedido: "&getCantidadPedida(rs1("idarticulo"))&" - stock Minimo: " & rs1("stockminimo") & chr(10)&chr(13)
	rs1.MoveNext
wend

asunto = "Faltante de Stock"
mensaje = "Los siguientes Articulos poseen faltante de stock: "& chr(10)&chr(13) & mensaje

accion = GF_PARAMETROS7("accion","",6)

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