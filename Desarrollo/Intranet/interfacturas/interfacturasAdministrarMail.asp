<!--#include file="../Includes/procedimientosunificador.asp"-->
<!--#include file="../Includes/procedimientostraducir.asp"-->
<!--#include file="../Includes/procedimientosFormato.asp"-->
<!--#include file="../Includes/procedimientosSeguridad.asp"-->
<!--#include file="../Includes/procedimientosmail.asp"-->
<!--#include file="../Includes/procedimientosparametros.asp"-->
<!--#include file="interfacturas.asp"-->
<% 

Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
end Function
'******************************************************************************************************************
'********************************************	COMIENZO DE LA PAGINA   *******************************************
'******************************************************************************************************************
Dim idProveedor, dsProveedor, accion, pagina,lpp,tipo,factura,flagHayListas, flagToepfer

flagToepfer = false
if (CDbl(session("CuitOrganizacion")) = CDbl(CUIT_TOEPFER)) then  flagToepfer = true

if (not flagToepfer) then
	idProveedor = session("CuitOrganizacion")
else
	idProveedor = GF_PARAMETROS7("idProveedor", "", 6)
end if	
if (idProveedor <> "") then dsProveedor = getDescripcionProveedorCUIT(idProveedor)

call addParam("idProveedor", idProveedor, params)

accion = GF_PARAMETROS7("accion", "", 6)
call addParam("accion", accion, params)
factura = GF_PARAMETROS7("factura", "", 6)
call addParam("factura", factura, params)
flagHayListas = false
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (mostrar = 0) then mostrar = 10
if (paginaActual = 0) then paginaActual = 1
Set sp_ret = executeProcedureDb(DBSITE_SQL_INTRA, rs, "TBLTAREAMAILS_GET_BY_PARAMETERS", TASK_FAC_MAIL_ALERT & "||" & idProveedor )
%>
<html>
<head>
<title>Administrar Mails Facturacion</title>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<meta http-equiv="x-ua-compatible" content="IE=11">
<link rel="stylesheet" href="../css/main.css" type="text/css">
<link rel="stylesheet" href="../css/paginar.css" type="text/css">
<link rel="stylesheet" href="../css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<script type="text/javascript" src="../scripts/controles.js"></script>
<script type="text/javascript" src="../scripts/Toolbar.js"></script>
<script type="text/javascript" src="../scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="../scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="../scripts/channel.js"></script>
<script type="text/javascript" src="../Scripts/jQueryPopUp.js"></script>
<script type="text/javascript">	
	var ch = new channel();
	function bodyOnload(){
		<% if (accion <> ACCION_EMAIL) then %>
			autoCompleteEmpresaResponsable();
		<% end if%>			
	}
	
	function addMail(pIdProveedor,pidx) {
		var newMail = document.getElementById("inputMail_" + pIdProveedor +"_"+ pidx).value;
		serverSave('<%=ACCION_GRABAR%>', pIdProveedor, pidx, newMail, '');		
	}
	
	function saveMail(pIdProveedor,pidx) {		
		var actualMail = document.getElementById("actualMail_" + pIdProveedor +"_" + pidx).value;
		var newMail = document.getElementById("inputMail_" + pIdProveedor +"_"+ pidx).value;
		serverSave('<%=ACCION_PROCESAR%>', pIdProveedor, pidx, newMail, actualMail);
	}
	
	function serverSave(action, pIdProveedor, pidx, newMail, actualMail)	{
		if ((actualMail != newMail) && (newMail != '')) {
			document.getElementById("imgMail_" + pIdProveedor +"_"+ pidx).innerHTML = '<img src="../images/loading_small_green.gif">';
			ch.bind("interfacturasAdministrarMailAjax.asp?idProveedor=" + pIdProveedor + "&orden="+pidx+"&accion=" + action + "&mail="+newMail, "saveMailCallBack("+pIdProveedor+","+ pidx +")");
			ch.send();
		} else {
			alert('Falta ingresar una dirección de mail o la dirección ingresada es igual a la anterior.');
			updateChange(pIdProveedor,pidx, actualMail);
		}
	}
	function saveMailCallBack(pIdProveedor, pidx) {		
		var resp = ch.response();
		if (resp == '<%=RESPUESTA_OK%>') {
			submitInfo();
		} else {
			alert(resp);			
			if (pidx > 0) {
				var actualMail = document.getElementById("actualMail_" + pIdProveedor + "_" + pidx).value;
				updateChange(pIdProveedor,pidx, actualMail);
			} else {
				document.getElementById("imgMail_" + pIdProveedor + "_" + pidx).innerHTML = '<img src="../images/save-16.png" onClick="addMail('+pIdProveedor+','+pidx+')">';
			}
		}
	}
	function updateChange(pIdProveedor,pidx, mail) {
		document.getElementById("inputMail_" + pIdProveedor + "_" + pidx).value = mail;
		document.getElementById("inputMail_" + pIdProveedor + "_" + pidx).style.display = 'none';
		document.getElementById("txtmail_" + pIdProveedor + "_" + pidx).innerHTML = mail;
		document.getElementById("txtmail_" + pIdProveedor + "_" + pidx).style.display = 'block';
		document.getElementById("actualMail_" + pIdProveedor + "_" + pidx).value = mail;		
		document.getElementById("imgMail_" + pIdProveedor + "_" + pidx).innerHTML = '<img src="../images/edit-16.png" onClick="editMail('+pIdProveedor+','+pidx+')">';
	}
	function editMail(pIdProveedor, pidx) {		
		document.getElementById("txtmail_" + pIdProveedor +"_"+pidx).style.display = 'none';
		document.getElementById("inputMail_" + pIdProveedor +"_"+pidx).style.display = 'block';		
		document.getElementById("imgMail_" + pIdProveedor +"_"+pidx).innerHTML = '<img src="../images/save-16.png" onClick="saveMail('+pIdProveedor+','+pidx+')">';
    }
    function deleteMail(pIdProveedor, pidx){
		if(confirm("Desea eliminar el mail ?" )){
			ch.bind("interfacturasAdministrarMailAjax.asp?idProveedor=" + pIdProveedor + "&orden=" + pidx + "&accion=<%=ACCION_BORRAR%>", "eliminarMail_CallBack()");
			ch.send();
		}
    }
    function loadPopUpNew_callback(){
		submitInfo();
    } 
    function submitInfo() {        
        document.getElementById("frmSel").submit();
    }
    function eliminarMail_CallBack(){
		submitInfo();
    }    
    function autoCompleteEmpresaResponsable(){
			$( "#dsProveedor" ).autocomplete({
					minLength: 3,
					source: "../comprasStreamElementos.asp?tipo=JQEmpresas",
					focus: function( event, ui ) {
						$( "#dsProveedor").val(ui.item.dsempresa);
						return false;
					},
					select: function( event, ui ) {
						$( "#dsProveedor"    ).val (ui.item.dsempresa);
						$( "#idProveedor"    ).val (ui.item.cuit);
						return false;
					},
					change: function( event, ui ) {
						if (!ui.item) {
							$( "#dsProveedor").val ("");
							$( "#idProveedor").val ("");
						}
					}
				})
				.data( "autocomplete" )._renderItem = function( ul, item ) {
					return $( "<li></li>" )
						.data( "item.autocomplete", item )
						.append( "<a>" + item.idempresa + " - <font style='font-size:10;'>" + item.dsempresa + "</font></a>" )
						.appendTo( ul );
				};
		}	
	function sendFacturaByMail(){
		if(confirm("Desea enviar la factura por mail al proveedor?")){							
			document.getElementById("divSendMail").innerHTML = "Enviando Mail ...";
			document.getElementById("divSendMail").style.display = "block";
			document.getElementById("divSendMail").className = "TDSUCCESS";	
			document.getElementById("ifrmSendMail").src = "interfacturasEnvioMail.asp?registro=<%=factura%>";
		}
	}	
    function finalizoEnvioMail(){
		document.getElementById("divSendMail").innerHTML = "";
		document.getElementById("divSendMail").style.display = "hidden";
		document.getElementById("divSendMail").className = "";	
		document.getElementById("ifrmSendMail").removeAttribute('src');
	}
</script>
<body onload="bodyOnload()">
	<FORM name="post" id="frmSel" name="frmSel" action="interfacturasAdministrarMail.asp">
		<div class="tableaside size100">
			<div id="searchfilter" class="tableasidecontent">				
				<div class="col16 reg_header_navdos"> <%=GF_Traducir("Proveedor:")%> </div>				
				<div class="col36">
					<% 
					if (not flagToepfer) then
						Response.Write left(idProveedor & "-" & dsProveedor,22) 
					else			%>
						<input name="dsProveedor" size="60" type="<%=pTipo%>" id="dsProveedor" value="<%=dsProveedor%>">
				<%	end if %>
					<input type="hidden" name="idProveedor" id="idProveedor" value="<%=idProveedor%>">
				</div>
				<div class="col16">
					<INPUT type="button" value="Buscar" onclick="submitInfo()" id="buscando" name="buscando" >
				</div>
			</div>	
		</div>
		<div class="col66"><div id="divSendMail" style="display:hidden"></div>
		<div class="col66"></div>
		<table class="datagrid" width="80%" align="center">
			<thead>
				<tr>					
					<th class="thiconac" align="center" nowrap>	<% =GF_TRADUCIR("Mail") %></th>
					<th class="thiconac" align="center" width="5%" nowrap>.</th>
					<th class="thiconac" align="center" width="5%" nowrap>.</th>
				</tr>
			</thead>	
			<tbody> 	
			<% if (not rs.eof) then 
				flagHayListas = true				
				while (not rs.eof) 	%>
						<tr>
							<td>							
								<div id="txtmail_<%=idProveedor & "_" & rs("idTareaMail") %>"><% =rs("EMAIL") %></div>							
								<input type="hidden" id="actualMail_<%=idProveedor & "_" & rs("idTareaMail") %>" value="<% =rs("EMAIL") %>">
								<input type="text" id="inputMail_<%=idProveedor & "_" & rs("idTareaMail") %>" style="display:none;" size="40" value="<% =rs("EMAIL") %>">
							</td>
							<td align="center" ><div id="imgMail_<%=idProveedor & "_" & rs("idTareaMail") %>" style="cursor:pointer;"><img src="../images/edit-16.png" onClick="editMail('<%=idProveedor%>',<%=rs("idTareaMail")%>)" title="Editar"></div></td>
							<td align="center" ><img src="../images/cross-16.png" onClick="deleteMail('<%=idProveedor%>',<%=rs("idTareaMail")%>)" style="cursor:pointer;" title="Eliminar"> </td>
						</tr>
				<% 	rs.MoveNext()			
				wend	
				else %>
					<tr>
						<td align="center"><%=GF_TRADUCIR("No se encontraron resultados")%></td>
					</tr>
			<%	end if	%>			
			</tbody>
			<tfoot>
  				<% if ((accion = ACCION_EMAIL) and (flagHayListas)) then %>
  				<tr><td colspan="5" align="center">
					<input type="button" onclick="sendFacturaByMail();" value="Enviar" />
				</tr></td>	
				<% end if %>
  			</tfoot>
		</table>
<%		if (idProveedor <> "") then			%>	
		<table class="datagrid" width="80%" align="center">
			<thead>
				<tr><th colspan="2">Agregar nuevo mail:</th></tr>
			</thead>
			<tbody
				<tr>
					<td>							
						<input type="text" size="60" id="inputMail_<%=idProveedor & "_0" %>" size="40" value="">
					</td>
					<td align="center" width="32px"><div id="imgMail_<%=idProveedor & "_0" %>" style="cursor:pointer;"><img src="../images/save-16.png" onClick="addMail('<%=idProveedor%>',0)" title="Grabar"></div></td>					
				</tr>			
			</tbody>
		</table>
<%		end if %>			
		<input type="hidden" name="registrosPorPagina" id="registrosPorPagina" value=<%=mostrar%>>
		<input type="hidden" name="numeroPagina" id="numeroPagina" value=<%=paginaActual%>>
		<input type="hidden" name="factura" id="factura" value=<%=factura%>>
		<input type="hidden" id="accion" name="accion" value="<%=accion%>">
		<iframe width="1px" height="1px" style="display:hidden;" id="ifrmSendMail" name="ifrmSendMail" onload="finalizoEnvioMail()"></iframe>
	</form>		
</body>
</html>