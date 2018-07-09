<!--#include file="Includes/procedimientosMG.asp"-->
<!--#include file="Includes/procedimientostraducir.asp"-->
<!--#include file="Includes/procedimientosCompras.asp"-->
<!--#include file="Includes/procedimientosFechas.asp"-->
<!--#include file="Includes/procedimientosObras.asp"-->
<!--#include file="Includes/procedimientosSQL.asp"-->
<!--#include file="Includes/procedimientosFormato.asp"-->
<%
Call comprasControlAccesoCM(RES_OBR)

'---------------------------------------------------------------------------------
'	COMIENZO DE LA PAGINA
'---------------------------------------------------------------------------------
Dim strSQL, rs, rsDivision, conn, i, paginaActual, mostrar, controlObra, isAud
Dim idObra, cdObra, dsObra, idDivision, params, rsAFE,cualesMostrar, simboloMoneda,myRol

call GP_ConfigurarMomentos()

hayBusqueda = GF_PARAMETROS7("busquedaActiva",0,6)
call addParam("busquedaActiva", hayBusqueda, params)
idObra = GF_PARAMETROS7("idObra","",6)
call addParam("idObra", idObra, params)
cdObra = UCase(GF_PARAMETROS7("cdObra","",6))
call addParam("cdObra", cdObra, params)
dsObra = GF_PARAMETROS7("dsObra","",6)
call addParam("dsObra", dsObra, params)
idDivision = GF_PARAMETROS7("idDivision",0,6)
call addParam("idDivision", idDivision, params)
paginaActual = GF_PARAMETROS7("numeroPagina",0,6)
if (paginaActual = 0) then paginaActual=1
mostrar = GF_PARAMETROS7("registrosPorPagina",0,6)
if (mostrar = 0) then mostrar = 10
sql_Order     = GF_PARAMETROS7("sqlOrder","",6)
cualesMostrar = GF_PARAMETROS7("cualesMostrar","",6)
if (cualesMostrar = "") then cualesMostrar=OBRA_ACTIVA
call addParam("cualesMostrar", cualesMostrar, params)

simboloMoneda = getSimboloMoneda(MONEDA_DOLAR)
myRol = getRolFirma(session("Usuario"), SEC_SYS_COMPRAS)

Set rs = obtenerListaObrasOrdenado(idObra, cdObra, dsObra, idDivision, cualesMostrar,sql_Order)
Call setupPaginacion(rs, paginaActual, mostrar)		
%>
<html>
<head>
<title><% = GF_TRADUCIR("Administracion Partidas") %></title>
<link rel="stylesheet" href="css/ActiSAIntra-1.css" type="text/css">
<link rel="stylesheet" href="css/Toolbar.css" type="text/css">
<link rel="stylesheet" href="css/jqueryUI/custom-theme/jquery-ui-1.8.2.custom.css" type="text/css" />
<style type="text/css">
.titleStyle {
	font-weight: bold;
	font-size: 20px;
}

.divOculto {
	display: none;
}
.budgetOK {
	background-color: green;
	color: white;	
}
.budgetWarning {
	background-color: yellow;
	color: black;	
}
.budgetDanger {
	background-color: red;
	color: white;	
	font-weight: bold;
}
</style>
<script type="text/javascript" src="scripts/Toolbar.js"></script>
<script type="text/javascript" src="scripts/jQueryPopUp.js"></script>
<script type="text/javascript" src="scripts/jquery/jquery-1.3.2.min.js"></script>
<script type="text/javascript" src="scripts/jqueryUI/jquery-ui-1.8.2.custom.min.js"></script>
<script type="text/javascript" src="scripts/channel.js"></script>
<script type="text/javascript" src="scripts/paginar.js"></script>
<script defer type="text/javascript" src="scripts/pngfix.js"></script>
<script type="text/javascript">
	var isFirefox = !(navigator.appName == "Microsoft Internet Explorer");	
	
	function irHome() {
		location.href = "comprasIndex.asp";
	}
	function buscarOn() {
		document.getElementById("busqueda").className = "";
		document.getElementById("busquedaActiva").value = "1";
	}
	function buscarOff() {
		document.getElementById("busqueda").className = "divOculto";
		document.getElementById("busquedaActiva").value = "0";
	}
	function irAdministracion() {
		location.href = "comprasAdministracion.asp";
	}
	function irObras() {
		location.href = "comprasObras.asp";
	}
	function irPedidos() {
		location.href = "comprasAdministrarPedidos.asp";
	}
	function loadPopUpObras(id) {
		var w = '600';
		var h = '420';		
		var puw = new winPopUp('popupObra','comprasPropObra.asp?idObra=' + id, w, h,'Propiedades Ptda. Presupuestaria', 'irObras()');	
	}	
	function loadTablero(id) {
		location.href ="comprasTableroObra.asp?idObra=" + id;
	}
	function bodyOnLoad() {	
		var tb = new Toolbar('toolbar', 6, "images/compras/");
		tb.addButton("Home-16x16.png", "Home", "irHome()");		
		<% if puedeCrear then %>		
			tb.addButton("OBR_new-16x16.png", "Nueva", "loadPopUpObras(0)");		
		<% end if %>	
		tb.addButtonREFRESH("Recargar", "irObras()");		
		var swt = tb.addSwitcher("search-16x16.png", "Buscar", "buscarOn()", "buscarOff()");		
		tb.draw();	
		<% if (cint(hayBusqueda) = 1) then %>
			tb.changeState(swt);						
		<% end if
			if (not rs.eof) then		%>								
				var pgn = new Paginacion("paginacion");							
				pgn.paginar(<% =paginaActual %>, <% =rs.RecordCount %>, <% =mostrar %>, 50, "comprasObras.asp<% =params %>&sqlOrder="+document.getElementById("sqlOrder").value);
		<%	end if %>							
		pngfix();
	}
		
	function createBudget(id) {
		window.open('comprasBudgetObra.asp?idObra=' + id);		
	}
	function setOrder(p_campo,p_orden){
		document.getElementById("sqlOrder").value = ' ORDER BY '+p_campo+' '+p_orden;
		document.getElementById("frmSel").submit();
	}
</script>
</head>
<body onLoad="bodyOnLoad()">
	<% call GF_TITULO2("kogge64.gif","Administrar Partidas Presupuestarias") %>	
	<div id="toolbar"></div>
	<br>
	<form name="frmSel" id="frmSel" method="GET">
		<input type="hidden" name="sqlOrder" id="sqlOrder" value="<%=sql_Order%>">	
	<div id="busqueda" class="divOculto">
	<table width="80%" cellspacing="0" cellpadding="0" align="center" border="0">
       <input type="hidden" name="accion" id="accion" value="">
       <tr>
           <td width="8"><img src="images/marco_r1_c1.gif"></td>
           <td width="25%"><img src="images/marco_r1_c2.gif" width="100%" height="8"></td>
           <td width="1%"><img src="images/marco_r1_c3.gif"></td>
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
								<td align="right"><% = GF_TRADUCIR("Codigo") %>:</td>
								<td colspan="2"><input type="text" name="cdObra" value="<% =cdObra %>"></td>
							</tr>														
							<tr>
								<td align="right"><% = GF_TRADUCIR("Descripcion") %>:</td>
								<td><input type="text" name="dsObra" value="<% =dsObra %>"></td>
								<td align="right"><% = GF_TRADUCIR("División") %>:</td>
								<td>
								<%	strSQL="Select * from TBLDIVISIONES"
                                    Call executeQueryDb(DBSITE_SQL_INTRA, rsDivision, "OPEN", strSQL)
								%>
									<select id="idDivision" name="idDivision">							
										<option value="<% =SIN_DIVISION %>" selected="true">- <% =GF_TRADUCIR("Seleccione") %> - 
								<%		while (not rsDivision.eof) 	%>										
										<option value="<% =rsDivision("IDDIVISION") %>" <% if (idDivision = rsDivision("IDDIVISION")) then response.write "selected='true'" %>><% =rsDivision("DSDIVISION") %>
								<%			rsDivision.MoveNext()
										wend	%>								
									</select>
								</td>
							</tr>	
							<tr>
								<td align='right'>
									<% =GF_TRADUCIR("Mostrar") %>:
								</td>
								<td>
									<input type="radio" name="cualesMostrar" id ="cualesMostrar" value="<% =OBRA_ACTIVA %>" <%if (cualesMostrar = OBRA_ACTIVA) then%> checked='true' <%end if%>/> En curso
									<input type="radio" name="cualesMostrar" id ="cualesMostrar" value="X" <%if (cualesMostrar <> OBRA_ACTIVA)  then%> checked='true' <%end if%>/> Todas
								</td>
							</tr>
							<tr>
								<td colspan="4" align="center"><input type="submit" value="Buscar..."></td>							
							<tr>
                     </table>
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
	<input type="hidden" name="busquedaActiva" id="busquedaActiva" value="0">	
	
	<br>
	<table class="reg_header" align="center" width="100%" border="0" cellspacing="1" cellpadding="2">
		<tr><td colspan="11"><div id="paginacion"></div></td></tr>	
		<tr class="reg_header_nav" align="center">					
			<td width="7%">  <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("OBRAS.idobra"        ,"asc")' style="cursor:pointer"><% =GF_TRADUCIR("Codigo")      %><img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("OBRAS.idobra"        ,"desc")' style="cursor:pointer"></td>
			<td>	         <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("OBRAS.dsobra"        ,"asc")' style="cursor:pointer"><% =GF_TRADUCIR("Descripcion") %><img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("OBRAS.dsobra"        ,"desc")' style="cursor:pointer"></td>
			<td>             <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("DIV.dsdivision"      ,"asc")' style="cursor:pointer"><% =GF_TRADUCIR("División")    %><img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("DIV.dsdivision"      ,"desc")' style="cursor:pointer"></td>			
			<td colspan="2"> <% =GF_TRADUCIR("Presupuesto") %></td>
			<td>             <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("OBRAS.fechainicio"   ,"asc")' style="cursor:pointer"><% =GF_TRADUCIR("F. Inicio")   %><img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("OBRAS.fechainicio"   ,"desc")' style="cursor:pointer"></td>
			<td>             <img src="images\compras\arrow_up_12x12.gif" onclick='setOrder("OBRAS.fechafin"      ,"asc")' style="cursor:pointer"><% =GF_TRADUCIR("F. Final")    %><img src="images\compras\arrow_down_12x12.gif" onclick='setOrder("OBRAS.fechafin"      ,"desc")' style="cursor:pointer"></td>			
			<td width="32px" align="center"><% =GF_TRADUCIR(".")           %></td>
			<td width="32px" align="center"><% =GF_TRADUCIR("AFE")         %></td>
			<td width="32px" align="center"><% =GF_TRADUCIR("CTC")           %></td>
			<td width="32px" align="center"><% =GF_TRADUCIR("Fotos")         %></td>
			<td width="32px" align="center"><% =GF_TRADUCIR(".")           %></td>
		</tr>		
		<% 	i=0
		while ((not rs.eof) and (i < mostrar))					
				i=i+1
				controlObra = false
                'Solamente el responsable de compras puede modificar una partida
                if (myRol = FIRMA_ROL_GTE_COMPRAS) then controlObra = true
				presupuestoTotal = calcularCostoEstimadoObra(MONEDA_DOLAR, rs("IDOBRA"),0,0)								
		%>
		<tr class="reg_header_navdos" onMouseOver="this.className='reg_header_navdosHL';" onMouseOut="this.className='reg_header_navdos';">					
			<td onClick="loadTablero(<% =rs("IDOBRA") %>)"><% =rs("CDOBRA") %></td>
			<td onClick="loadTablero(<% =rs("IDOBRA") %>)"><% =left(GF_TRADUCIR(UCASE(rs("DSOBRA"))), 30) %>...</td>
			<td onClick="loadTablero(<% =rs("IDOBRA") %>)" align="center"><% =rs("DSDIVISION") %></td>
			<td onClick="loadTablero(<% =rs("IDOBRA") %>)" align="right">
				<% =simboloMoneda %>&nbsp;&nbsp;<% =GF_EDIT_DECIMALS(presupuestoTotal,2) %>			</td>
			<td align="center" width="16px">
				<%  
					'Call logDebug(controlObra)
					if (controlObra) then
                        if (isBudgetProvisorio(rs("FECHABUDGET"))) then %>
							<a onClick="javascript:createBudget(<% =rs("IDOBRA") %>)"><img style="cursor:pointer" id="imgBudget<% =rs("IDOBRA") %>" src="images/compras/edit-16x16.png" title="<% =GF_TRADUCIR("Cargar/Modificar Presupuesto") %>"></a>
					<%	else %>
					        <a onClick="javascript:location.href='comprasBudgetReasignaciones.asp?idObra=<% =rs("IDOBRA") %>' "><img style="cursor:pointer" id="imgBudget<% =rs("IDOBRA") %>" src="images/compras/budget_view-16x16.png" title="<% =GF_TRADUCIR("Reasignar Presupuesto") %>"></a>
					<%  end if %>
				<%	end if	%>			</td>						
			<td onClick="loadTablero(<% =rs("IDOBRA") %>)" align="center"><% =GF_FN2DTE(rs("FECHAINICIO")) %></td>
			<%
				fechaFin = rs("FECHAFIN")
				if (rs("FECHAAJUSTADA") <> "0") then fechaFin = rs("FECHAAJUSTADA")
			%>
			<td onClick="loadTablero(<% =rs("IDOBRA") %>)" align="center"><% =GF_FN2DTE(fechaFin) %></td>						
			<td align="center"><a href="comprasbudgetobrafilter.asp?idobra=<%=rs("IDOBRA") %>&origen=ComprasObras"> <img src="images/compras/printer-16x16.png" border="0"></a></td>
			<td align="center">						
				<a onClick="window.open('comprasPopUpAFE.asp?idObra=<%=rs("IDOBRA") %>', '_blank','location=no,menubar=no,statusbar=no,height=400,width=500',false);"><img src="images/compras/afe-16x16.png" alt="<% =GF_TRADUCIR("Ver y trabajar con los AFE") %>"></a>
			</td>
			<td align="center">						
				<a onClick="window.open('comprasCTCPopUp.asp?idObra=<%=rs("IDOBRA") %>', '_blank','location=no,menubar=no,statusbar=no,height=400,width=500',false);"><img src="images/compras/CTC-16x16.png" title="<% =GF_TRADUCIR("Ver y trabajar con los Contratos") %>"></a>
			</td>
			<td align="center"><a href="comprasObrasFotos.asp?idObra=<%=rs("IDOBRA") %>&origen=ComprasObras"> <img src="images/compras/Picture-icon-16x16.png" border="0"></a></td>
			<td align="center">			
				<%	if (controlObra) then %>					
						<a onClick="loadPopUpObras(<% =rs("IDOBRA") %>)"><img src="images/compras/edit-16x16.png" title="<% =GF_TRADUCIR("Modificar Partida") %>"></a>
				<%	end if	%>			
			</td>
		</tr>
		<%	'end if
			rs.MoveNext()
		wend	
		if (i = 0) then		
		%>			
		<tr>
			<td class="TDNOHAY" colspan="12"><% =GF_TRADUCIR("No existen obras registradas") %></td>
		</tr>
		<%
		end if 
		%>	
	</table>
</form>
</body>
</html>
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
%>