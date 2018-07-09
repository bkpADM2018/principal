<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<%
'******************************************************************************************
Function addParam(p_strKey,p_strValue,ByRef p_strParam)
       if (not isEmpty(p_strValue)) then
          if (isEmpty(p_strParam)) then
             p_strParam = "?"
          else
             p_strParam = p_strParam & "&"
          end if
          p_strParam = p_strParam & p_strKey & "=" & p_strValue
       end if
End Function
'******************************************************************************************

Call comprasControlAccesoCM(RES_AUD)


Dim cdNorma, dsNorma, mkWhereNormas, divClass, valor

cdNorma = UCase(GF_PARAMETROS7("cdNorma", "" ,6))
call addParam("cdNorma", cdNorma, params)
dsNorma = GF_PARAMETROS7("dsNorma", "" ,6)
call addParam("dsNorma", dsNorma, params)
pagina = GF_PARAMETROS7("numeroPagina",0,6)
if (pagina = 0) then pagina = 1
regXPag = GF_PARAMETROS7("registrosPorPagina",0,6)
if (regXPag = 0) then regXPag = 10

divClass = "divOculto"

if (cdNorma <> "") then 
	Call mkWhere(mkWhereNormas, "cdNorma", cdNorma, "LIKE", 3)
	divClass = ""
end if
if (dsNorma <> "") then 
	Call mkWhere(mkWhereNormas, "dsNorma", dsNorma, "LIKE", 3)
	divClass = ""
end if

hayBusqueda = false
busquedaActiva = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", busquedaActiva, params)
if busquedaActiva = 1 then hayBusqueda = true

strSQL="Select IDNORMA, CDNORMA, DSNORMA, (VALOR*100) VALOR, UNIDAD from TBLNORMASAUDITORIA " & mkWhereNormas
'response.write strSQL & "<BR>"
call executeQueryDb(DBSITE_SQL_INTRA, rs, "OPEN", strSQL)	
Call setupPaginacion(rs, pagina, regXPag)

%>
<html>
<head>
<title>Normas de Auditor&iacutea</title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<style type="text/css">
.divOculto {
	display: none;
}
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}
</style>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/formato.js"></script>
<script type="text/javascript" src="scripts/controles.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript">
	
	var paginacion;
	
	function loadPopUpNormas(id) {		
		<% if (puedeCrear()) then %>		
			var puw = new winPopUp('popupNorma','comprasPropNorma.asp?idNorma=' + id,'450','230','Propiedades Norma', 'reload()');
		<% end if %>
	}

	function reload() {
		document.getElementById("frmSel").submit();
	}
		
	function buscarOn() {
		document.getElementById("busquedaNormas").className = "";	
		document.getElementById("busquedaActiva").value = "1";			
	}
	
	function buscarOff() {
		document.getElementById("busquedaNormas").className = "divOculto";		
		document.getElementById("busquedaActiva").value = "0";		
	}
	
	function doBuscar() {
		paginacion.solveRequest(1, 10);	
	}
	
	function irHome() {
		location.href = "almacenAuditoria.asp";
	}
	
	function bodyOnLoad() {
		var toolBarNormas = new Toolbar("toolBarNormas", 6, "images/compras/");
		toolBarNormas.addButton("Home-16x16.png", "Home", "irHome()");		
		<% if (puedeCrear()) then %>
			toolBarNormas.addButton("Audit_new-16x16.png", "Nueva", "loadPopUpNormas(0)");		
		<% end if %>
		toolBarNormas.addButtonREFRESH("Refrescar", "reload()");	
		var swt = toolBarNormas.addSwitcher("Search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");						
		toolBarNormas.draw();
		<%	if (hayBusqueda) then %>
				toolBarNormas.changeState(swt);		
		<%	End if 	%>
		<%	if (not rs.eof) then %>
			paginacion = new Paginacion("paginacionNormas");
			paginacion.paginar(<% =pagina %>, <% =rs.RecordCount %>, <% =regXPag %>, 50, "comprasNormasAuditoria.asp<% =params %>");
		<%	End if 	%>
		pngfix();
	}
	
</script>
</head>
<body onLoad="bodyOnLoad()">
<form method="post" id="frmSel">
<table width="100%">
	<tr valign="top">
		<td>
			<form name="frmSel" action="comprasNormasAuditoria.asp">
				<div id="toolBarNormas"></div><br>
					<div id="busquedaNormas" class="<% =divClass %>">
						<table width="70%" cellspacing="0" cellpadding="0" align="center" border="0">
						   <input type="hidden" name="accion" id="accion" value="">
						   <tr>
							   <td width="8"><img src="images/marco_r1_c1.gif"></td>
							   <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
							   <td width="8"><img src="images/marco_r1_c3.gif"></td>
							   <td width="73%"><td>
							   <td></td>
						   </tr>
						   <tr>
							   <td width="8"><img src="images/marco_r2_c1.gif"></td>
							   <td align="center" valign="center"><font class="big" color="#517b4a"><% =GF_TRADUCIR("Busqueda") %></font></td>
							   <td width="8"><img src="images/marco_r2_c3.gif"></td>
							   <td></td>
							   <td></td>
						   </tr>
						   <tr>
							   <td><img src="images/marco_r2_c1.gif" height="8"  width="8"></td>
							   <td></td>
							   <td valign="top" align="right"><img src="images/marco_r1_c2.gif" height="8" width="2"></td>
							   <td><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
							   <td width="8"><img src="images/marco_r1_c3.gif"></td>
						   </tr>
						   <tr>
							   <td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
							   <td colspan="3">								
									<table width="100%" align="center" border="0">
										<tr>
											<td><% = GF_TRADUCIR("Codigo") %>:</td>
											<td><input type="text" size="10" maxlength="10" name="cdNorma" value="<% =cdNorma %>" ></td>
											<td><% = GF_TRADUCIR("Descripcion") %>:</td>
											<td><input type="text" size="50" name="dsNorma" value="<% =dsNorma %>"></td>
										</tr>																		
									</table>	  					
								</td>
								   <td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
							   </tr>
							   <tr>
									<td height="100%"><img src="images/marco_r2_c1.gif" height="100%" width="8"></td>
									<td colspan="3" align="center">
										<input type="button" value="<% =GF_TRADUCIR("Buscar") %>" onClick="javascript:doBuscar()">
									</td>
									<td height="100%"><img src="images/marco_r2_c3.gif" width="8" height="100%"></td>
								</tr>
							   <tr>
								   <td width="8"><img src="images/marco_r3_c1.gif"></td>
								   <td width="100%" align=center colspan="3"><img src="images/marco_r3_c2.gif" width="100%" height="8"></td>
								   <td width="8"><img src="images/marco_r3_c3.gif"></td>
							   </tr>
						</table>
					</div>
					<br>
				<table align="center" width="80%" height="100%" class="reg_header" cellspacing="2" cellpadding="1">				
					<tr><td colspan="4"><div id="paginacionNormas"></div></td></tr>					
					<tr class="reg_header_nav">
						<td align="center">.</td>
						<td align="center" width="15%"><% =GF_TRADUCIR("Codigo") %></td>
						<td><% =GF_TRADUCIR("Descripcion") %></td>
						<td align="center" width="15%"><% =GF_TRADUCIR("Valor") %></td>
					</tr>
					<%		i=0
						while ((not rs.eof)	and (i < regXPag))
							i = i+1
					%>
						<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">
							<td align="center" width="24px" onClick="loadPopUpNormas(<% =rs("IDNORMA") %>)"><img src="images/compras/Audit-16x16.png"></td>
							<td align="center" width="10%" onClick="loadPopUpNormas(<% =rs("IDNORMA") %>)"><b><% =rs("CDNORMA") %></b></td>
							<td onClick="loadPopUpNormas(<% =rs("IDNORMA") %>)"><% =rs("DSNORMA") %></td>
							<td onClick="loadPopUpNormas(<% =rs("IDNORMA") %>)" align="right">
								<%	valor = CDbl(rs("VALOR"))/100
									if (rs("UNIDAD") <> "C") then	'Es un importe!
										valor = getSimboloMoneda(rs("UNIDAD")) & " " & GF_EDIT_DECIMALS(CDbl(rs("VALOR")),2)
									end if 
									response.write valor
								%>
							</td>
						</tr>
								
					<%		rs.MoveNext()
						wend
						if (i = 0) then		
					%>			
						<tr>
							<td class="TDNOHAY" colspan="4"><% =GF_TRADUCIR("No existen normas registradas") %></td>
						</tr>
					<%		end if %>
				</table>
				<input type="hidden" name="busquedaActiva"     id="busquedaActiva"     value=<%=hayBusqueda%>   >

			</form>
		</td>
	</tr>
</table>
</form>
</body>
</html>